import os
import pandas as pd
import math

# path
EAC_FOLDER = r"C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/DRM Slides/EAC_analysis/EAC"       
RADET_FOLDER = r"C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/DRM Slides/EAC_analysis/RADET"   
OUTPUT_FOLDER = r"C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/DRM Slides/EAC_analysis/output"  

MERGE_KEY = "Patient ID"  

# Ensure output directory exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Get all valid files from the EAC folder
all_eac_files = [os.path.join(EAC_FOLDER, f) for f in os.listdir(EAC_FOLDER) if f.endswith(('.csv', '.xlsx', '.xls')) and not f.startswith('~$')]

# Storage containers for global combination
all_summaries = []
all_linelists_c = []
all_linelists_d = []
all_linelists_e = []

# handling mislabeled xls files
def smart_read_excel(file_path):
    if file_path.endswith('.csv'):
        return pd.read_csv(file_path, encoding='latin1', on_bad_lines='skip')
    
    if file_path.endswith('.xls'):
        try:
            return pd.read_excel(file_path, engine='xlrd')
        except Exception as e:
            if "xlsx file; not supported" in str(e):
                return pd.read_excel(file_path, engine='openpyxl')
            raise e
    else:
        return pd.read_excel(file_path, engine='openpyxl')

# iterate matching files
for eac_file in all_eac_files:
    try:
        # Read the EAC file 
        df_eac = smart_read_excel(eac_file)
            
        filename_base = os.path.basename(eac_file)
        project_name = filename_base.split('_')[0]
        
        # locate the matching RADET file in the other folder
        radet_match = None
        for f in os.listdir(RADET_FOLDER):
            if f.startswith(project_name) and f.endswith(('.csv', '.xlsx', '.xls')) and not f.startswith('~$'):
                radet_match = os.path.join(RADET_FOLDER, f)
                break
                
        if not radet_match:
            print(f"Warning: No matching RADET file found for Project: {project_name}. Skipping...")
            continue
            
        # Read the matching RADET file 
        df_radet = smart_read_excel(radet_match)
            
        # Clean whitespaces from column headers
        df_eac.columns = df_eac.columns.str.strip()
        df_radet.columns = df_radet.columns.str.strip()
        
        # 4. Bring columns from RADET into EAC data
        radet_cols = [MERGE_KEY, 'Current Viral Load (c/ml)', 'Date of Current Viral Load (yyyy-mm-dd)', 'Current ART Status', 'Last Pickup Date (yyyy-mm-dd)', 'Age']
        for col in ['L.G.A', 'Facility Name']:
            if col not in df_eac.columns and col in df_radet.columns:
                radet_cols.append(col)
                
        df_radet_subset = df_radet[radet_cols].drop_duplicates(subset=[MERGE_KEY])
        df_merged = pd.merge(df_eac, df_radet_subset, on=MERGE_KEY, how='left')
        
        # Add tracking metadata columns
        df_merged['Filename'] = filename_base
        df_merged['IP'] = project_name
        
        
        # STRATEGIC DEDUPLICATION FOR LONGITUDINAL EAC RECORDS
        
        vld_col = 'Date of Unsuppressed Viral Load Result'
        pud_col = 'Last Pickup Date (yyyy-mm-dd)'
        
        if vld_col in df_merged.columns:
            temp_vld = pd.to_datetime(df_merged[vld_col], errors='coerce')
            temp_pud = pd.to_datetime(df_merged[pud_col], errors='coerce')
            
            valid_pud_before_vl = df_merged[pud_col].where(temp_pud < temp_vld)
            temp_valid_pud = pd.to_datetime(valid_pud_before_vl, errors='coerce')
            
            df_merged['_vld_sort'] = temp_vld
            df_merged['_pud_sort'] = temp_valid_pud
            df_merged['_vld_isnull'] = temp_vld.isna()
            
            df_merged = df_merged.sort_values(
                by=[MERGE_KEY, '_vld_isnull', '_vld_sort', '_pud_sort'],
                ascending=[True, True, False, False]
            )
            
            df_merged = df_merged.drop_duplicates(subset=[MERGE_KEY], keep='first')
            df_merged = df_merged.drop(columns=['_vld_sort', '_pud_sort', '_vld_isnull'])
        else:
            print(f"⚠️ Warning: Missing '{vld_col}' in {filename_base}. Standard fallback deduplication applied.")
            df_merged = df_merged.drop_duplicates(subset=[MERGE_KEY], keep='first')
            
        
        # METRIC CALCULATIONS & INDICATORS
        
        # Convert Viral Load to numeric
        df_merged['VL_numeric'] = pd.to_numeric(df_merged['Current Viral Load (c/ml)'].astype(str).str.replace(',', '', regex=False), errors='coerce')
        df_merged['Age'] = pd.to_numeric(df_merged['Age'].astype(str).str.replace(',', '', regex=False), errors='coerce')
        
        # Convert dates to datetime objects 
        df_merged['First_EAC_Date_parsed'] = pd.to_datetime(df_merged['Date of commencement of 1st EAC (yyyy-mm-dd)'], errors='coerce')
        df_merged['VL_Date_parsed'] = pd.to_datetime(df_merged['Date of Current Viral Load (yyyy-mm-dd)'], errors='coerce')
        df_merged['Pickup_Date_parsed'] = pd.to_datetime(df_merged['Last Pickup Date (yyyy-mm-dd)'], errors='coerce')
        df_merged['Second_EAC_Date_parsed'] = pd.to_datetime(df_merged['Date of commencement of 2nd  EAC (yyyy-mm-dd)'], errors='coerce')
        df_merged['Third_EAC_Date_parsed'] = pd.to_datetime(df_merged['Date of commencement of 3rd   EAC (yyyy-mm-dd)'], errors='coerce')

        # Calculate day count deltas for ALL 3 EAC variations against the current Viral Load date
        df_merged['EAC1_minus_VL_days'] = (df_merged['First_EAC_Date_parsed'] - df_merged['VL_Date_parsed']).dt.days
        df_merged['EAC2_minus_VL_days'] = (df_merged['Second_EAC_Date_parsed'] - df_merged['VL_Date_parsed']).dt.days
        df_merged['EAC3_minus_VL_days'] = (df_merged['Third_EAC_Date_parsed'] - df_merged['VL_Date_parsed']).dt.days
        
        # Identify if ANY EAC milestone falls within the expanded window range (-90 to 90 days)
        in_grace_window = (
            df_merged['EAC1_minus_VL_days'].between(-90, 90) |
            df_merged['EAC2_minus_VL_days'].between(-90, 90) |
            df_merged['EAC3_minus_VL_days'].between(-90, 90)
        )

        # Check if 2nd or 3rd EAC is after Last Pickup Date OR after Current Viral Load Date
        eac_2_or_3_after_critical_dates = (
            (df_merged['Second_EAC_Date_parsed'] > df_merged['Pickup_Date_parsed']) |
            (df_merged['Second_EAC_Date_parsed'] > df_merged['VL_Date_parsed']) |
            (df_merged['Third_EAC_Date_parsed'] > df_merged['Pickup_Date_parsed']) |
            (df_merged['Third_EAC_Date_parsed'] > df_merged['VL_Date_parsed'])
        )

        # Logical Conditions
        mask_a = df_merged['Current ART Status'].str.strip().isin(['Active', 'Active Restart']) & (df_merged['VL_numeric'] > 1000) & (df_merged['Age'] <15)
        mask_b = mask_a & (df_merged['First_EAC_Date_parsed'] > df_merged['VL_Date_parsed']) 
        
        # Exclude if ANY EAC date is valid within the 90-day grace windows
        mask_c = mask_a & (df_merged['First_EAC_Date_parsed'].isna() | (df_merged['First_EAC_Date_parsed'] < df_merged['VL_Date_parsed'])) & (~in_grace_window) & (~eac_2_or_3_after_critical_dates)
        mask_d = mask_c & (df_merged['Pickup_Date_parsed'] > df_merged['VL_Date_parsed']) 
        
        mask_e = mask_a & mask_b & ((df_merged['Date of commencement of 1st EAC (yyyy-mm-dd)'] == df_merged['Date of commencement of 2nd  EAC (yyyy-mm-dd)']) | (df_merged['Date of commencement of 2nd  EAC (yyyy-mm-dd)'] == df_merged['Date of commencement of 3rd   EAC (yyyy-mm-dd)']))

        # Map back tracking labels
        df_merged['Active_client_with_VL_gt_1000'] = mask_a
        df_merged['EAC_done_after_Unsuppressed VL_result'] = mask_b
        df_merged['EAC_not_done_after_Unsuppressed_VL_result'] = mask_c
        df_merged['Last_Pickup_Date_after_Unsuppressed_VL_result'] = mask_d
        df_merged['Same_date_for_the_three_EAC_done'] = mask_e
        df_merged['Day_count_btwn_LPUD_&_VLRR'] = (df_merged['Pickup_Date_parsed'] - df_merged['VL_Date_parsed']).dt.days
        
        # Aggregation Summary
        summary_grouped = df_merged.groupby(['IP', 'L.G.A', 'Facility Name'], dropna=False).agg(
            Active_client_with_unsuppressed_VL=('Active_client_with_VL_gt_1000', 'sum'),
            EAC_done_after_Unsuppressed_VL_result=('EAC_done_after_Unsuppressed VL_result', 'sum'),
            EAC_not_done_after_Unsuppressed_VL_result=('EAC_not_done_after_Unsuppressed_VL_result', 'sum'),
            Last_Pickup_Date_after_Unsuppressed_VL_result=('Last_Pickup_Date_after_Unsuppressed_VL_result', 'sum'),
            Same_date_for_the_three_EAC_done=('Same_date_for_the_three_EAC_done', 'sum'),
            Day_count_btwn_LPUD_and_VLRR=('Day_count_btwn_LPUD_&_VLRR', 'sum')
        ).reset_index()
        
        all_summaries.append(summary_grouped)
        
        # Extract raw data rows matching condition C and D for the final Linelist
        # Dropping calculated temporary variables to keep output files organized
        cols_to_drop = ['VL_numeric', 'First_EAC_Date_parsed', 'Second_EAC_Date_parsed','Third_EAC_Date_parsed', 'VL_Date_parsed', 'Pickup_Date_parsed', 'Active_client_with_VL_gt_1000', 'EAC_done_after_Unsuppressed VL_result', 'EAC_not_done_after_Unsuppressed_VL_result', 'Last_Pickup_Date_after_Unsuppressed_VL_result', 'Same_date_for_the_three_EAC_done', 'EAC1_minus_VL_days', 'EAC2_minus_VL_days', 'EAC3_minus_VL_days']
        cols_to_dropp = ['VL_numeric', 'VL_Date_parsed', 'Pickup_Date_parsed', 'Active_client_with_VL_gt_1000', 'EAC_done_after_Unsuppressed VL_result', 'EAC_not_done_after_Unsuppressed_VL_result', 'Last_Pickup_Date_after_Unsuppressed_VL_result', 'Same_date_for_the_three_EAC_done', 'EAC1_minus_VL_days', 'EAC2_minus_VL_days', 'EAC3_minus_VL_days']
        
        all_linelists_c.append(df_merged[mask_c].drop(columns=cols_to_drop, errors='ignore'))
        all_linelists_d.append(df_merged[mask_d].drop(columns=cols_to_drop, errors='ignore'))
        all_linelists_e.append(df_merged[mask_e].drop(columns=cols_to_dropp, errors='ignore'))
        
        print(f"Processed: {project_name}")
        
    except Exception as e:
        print(f"Error processing file {eac_file}: {e}")

# generating outputs
if all_summaries:
    combined_summary = pd.concat(all_summaries, ignore_index=True)
    combined_linelist_c = pd.concat(all_linelists_c, ignore_index=True)
    combined_linelist_d = pd.concat(all_linelists_d, ignore_index=True)
    combined_linelist_e = pd.concat(all_linelists_e, ignore_index=True)
    
    # Save Summary Document
    summary_out_path = os.path.join(OUTPUT_FOLDER, "Facility_Summary_Counts_lt_15.xlsx")
    combined_summary.to_excel(summary_out_path, index=False)
    print(f"\nSummary matrix written to: {summary_out_path}")
    
    # Save Multi-sheet linelists Document
    linelist_out_path = os.path.join(OUTPUT_FOLDER, "Flag_linelist_Sheets_lt_15.xlsx")
    with pd.ExcelWriter(linelist_out_path, engine='openpyxl') as writer:
        combined_linelist_c.to_excel(writer, sheet_name='EAC_not_done_after_Unsuppressed_VL_result', index=False)
        combined_linelist_d.to_excel(writer, sheet_name='LastPickupDate_after_Unsuppressed_VL_result', index=False)
        combined_linelist_e.to_excel(writer, sheet_name='Same_date_for_the_three_EAC_done', index=False)
    print(f"Individual Linelists written to: {linelist_out_path}")
    
else:
    print("No operational files processed successfully.") 