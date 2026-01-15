import pandas as pd
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import re

# Path to the directory containing the CSV files
folder_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/FY26Q1_PrEP'

# Define output directory for projects
output_file_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Output_by_facility/PrEP_Aggregate.xlsx'

# Defining Periods
Start_of_quarter = pd.to_datetime('2025-04-01')
End_of_quarter = pd.to_datetime('2025-06-30')

# Combine all CSV files into one DataFrame, specifying 'latin1' encoding
all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(('.csv', '.xlsx', '.xls'))]

# Combine all files into one DataFrame
combined_data = pd.DataFrame()

for file in all_files:
    try:
        if file.endswith('.csv'):
            data = pd.read_csv(file, encoding='latin1', on_bad_lines='skip')
        elif file.endswith('.xlsx'):
            data = pd.read_excel(file, engine='openpyxl')
        elif file.endswith('.xls'):
            data = pd.read_excel(file)
        else:
            continue
        
        # Add Filename column
        data['Filename'] = os.path.basename(file)
        combined_data = pd.concat([combined_data, data], ignore_index=True)
    except Exception as e:
        print(f"Error processing file {file}: {e}")

if combined_data.empty:
    print("No valid files found or data could not be combined.")
    exit()

# Extract project name from the 'Filename' column
# IMPORTANT: Ensure 'Filename' column exists before attempting to split
if 'Filename' not in combined_data.columns:
    print("FATAL ERROR: 'Filename' column is missing from the combined data. Cannot extract ProjectName.")
    exit()

combined_data['ProjectName'] = combined_data['Filename'].str.split('_').str[0]


# Check if 'ProjectName' and 'Facility Name' columns exist and are not entirely empty
required_cols_for_aggregation = ['ProjectName', 'Facility Name', 'Facility Id (Datim)']
for col in required_cols_for_aggregation:
    if col not in combined_data.columns:
        print(f"FATAL ERROR: Required column '{col}' is missing from the combined data.")
        if col == 'ProjectName':
            print("This usually means the 'Filename' column couldn't be processed to extract ProjectName.")
            print(f"First 5 filenames: {combined_data['Filename'].head().tolist()}")
        exit()
    
    # Convert to string and handle potential NaNs before further processing
    # This helps prevent type-related errors in groupby/merge
    combined_data[col] = combined_data[col].astype(str).fillna('UNKNOWN')

    if (combined_data[col] == 'UNKNOWN').all():
        print(f"Warning: Column '{col}' is entirely 'UNKNOWN' (or NaN in original data). This might affect grouping and filtering.")
    elif combined_data[col].nunique() == 1 and combined_data[col].iloc[0] == 'UNKNOWN':
         print(f"Warning: Column '{col}' contains only 'UNKNOWN' values. This might indicate an issue with data extraction or input.")


# Ensure date columns are in datetime format
date_columns = [
    'Date Of Commencement (yyyy-mm-dd)', 
    'Date Of Last Pickup (yyyy-mm-dd)'
]
for col in date_columns:
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

# Check for invalid dates
invalid_dates = combined_data[date_columns].isnull().any(axis=1)
if invalid_dates.any():
    print(f"Warning: Invalid dates found in the following rows:\n{combined_data[invalid_dates]}")


# Filters for PrEP_CT
prep_ct = combined_data[
    (combined_data['Date Of Commencement (yyyy-mm-dd)'] < Start_of_quarter) &
    ((combined_data['Date Of Last Pickup (yyyy-mm-dd)']>= Start_of_quarter) & (combined_data['Date Of Last Pickup (yyyy-mm-dd)'] <= End_of_quarter)) &  
    (
        (combined_data['HIV status at PrEP Initiation'].str.contains('Negative', na=False)) | # na=False ensures NaN values don't match 'Negative'
        (combined_data['HIV status at PrEP Initiation'].isna()) | # Catches np.nan
        (combined_data['HIV status at PrEP Initiation'] == '') | # Catches explicit empty strings
        (combined_data['HIV status at PrEP Initiation'].str.strip() == '') # Catches strings with only whitespace
    ) &
    (combined_data['Age'] >=15) &
    (~combined_data['Sex'].isna())
]

# Filters for PrEP_CT_Type
prep_ct_type = combined_data[
    (combined_data['Date Of Commencement (yyyy-mm-dd)'] < Start_of_quarter) &
    ((combined_data['Date Of Last Pickup (yyyy-mm-dd)']>= Start_of_quarter) & (combined_data['Date Of Last Pickup (yyyy-mm-dd)'] <= End_of_quarter)) &  
    (
        (combined_data['HIV status at PrEP Initiation'].str.contains('Negative', na=False)) | # na=False ensures NaN values don't match 'Negative'
        (combined_data['HIV status at PrEP Initiation'].isna()) | # Catches np.nan
        (combined_data['HIV status at PrEP Initiation'] == '') | # Catches explicit empty strings
        (combined_data['HIV status at PrEP Initiation'].str.strip() == '') # Catches strings with only whitespace
    ) &
    ((combined_data['Current Prep Type'].str.contains('Oral')) | 
     (combined_data['Current Prep Type'].isna()) | # Catches np.nan
        (combined_data['Current Prep Type'] == '') | # Catches explicit empty strings
        (combined_data['Current Prep Type'].str.strip() == '') # Catches strings with only whitespace) &
    ) &
    (combined_data['Age'] >=15) &
    (~combined_data['Sex'].isna())
]

# Filters for PrEP_CT_Distribution
prep_ct_distribution = combined_data[
    (combined_data['Date Of Commencement (yyyy-mm-dd)'] < Start_of_quarter) &
    ((combined_data['Date Of Last Pickup (yyyy-mm-dd)']>= Start_of_quarter) & (combined_data['Date Of Last Pickup (yyyy-mm-dd)'] <= End_of_quarter)) &  
    (
        (combined_data['HIV status at PrEP Initiation'].str.contains('Negative', na=False)) | # na=False ensures NaN values don't match 'Negative'
        (combined_data['HIV status at PrEP Initiation'].isna()) | # Catches np.nan
        (combined_data['HIV status at PrEP Initiation'] == '') | # Catches explicit empty strings
        (combined_data['HIV status at PrEP Initiation'].str.strip() == '') # Catches strings with only whitespace
    ) &
    ((combined_data['Current Prep Distribution Setting'].str.contains('Facility|Community')) & (~combined_data['Current Prep Distribution Setting'].isna())) &
    (combined_data['Age'] >=15) &
    (~combined_data['Sex'].isna())
]

# Filters for PrEP_CT_Pregnant and Breastfeeding
prep_ct_PBF = combined_data[
    (combined_data['Date Of Commencement (yyyy-mm-dd)'] < Start_of_quarter) &
    ((combined_data['Date Of Last Pickup (yyyy-mm-dd)']>= Start_of_quarter) & (combined_data['Date Of Last Pickup (yyyy-mm-dd)'] <= End_of_quarter)) &  
    (
        (combined_data['HIV status at PrEP Initiation'].str.contains('Negative', na=False)) | # na=False ensures NaN values don't match 'Negative'
        (combined_data['HIV status at PrEP Initiation'].isna()) | # Catches np.nan
        (combined_data['HIV status at PrEP Initiation'] == '') | # Catches explicit empty strings
        (combined_data['HIV status at PrEP Initiation'].str.strip() == '') # Catches strings with only whitespace
    ) &
    (combined_data['Age'] >=15) &
    (combined_data['Sex'] == 'Female') &
    ((combined_data['Pregnancy Status'].str.contains('Pregnant|Breastfeeding')) & (~combined_data['Pregnancy Status'].isin(['Not Pregnant'])))
]


# Filters for PrEP_CT_TestResult
common_condition = (
        (combined_data['Date Of Commencement (yyyy-mm-dd)'] < Start_of_quarter) &
        ((combined_data['Date Of Last Pickup (yyyy-mm-dd)']>= Start_of_quarter) & (combined_data['Date Of Last Pickup (yyyy-mm-dd)'] <= End_of_quarter)) &  
        (
        (combined_data['HIV status at PrEP Initiation'].str.contains('Negative', na=False)) | # na=False ensures NaN values don't match 'Negative'
        (combined_data['HIV status at PrEP Initiation'].isna()) | # Catches np.nan
        (combined_data['HIV status at PrEP Initiation'] == '') | # Catches explicit empty strings
        (combined_data['HIV status at PrEP Initiation'].str.strip() == '') # Catches strings with only whitespace
    ) &
        (combined_data['Age'] >=15) &
        (~combined_data['Sex'].isna())
)

negative = (combined_data['Current HIV Status'].str.contains('Negative'))
positive = (combined_data['Current HIV Status'].str.contains('Positive'))
other = ((combined_data['Current HIV Status'].isna()) | # Catches np.nan
        (combined_data['Current HIV Status'] == '') | # Catches explicit empty strings
        (combined_data['Current HIV Status'].str.strip() == ''))

prep_ct_negative =combined_data[common_condition & (negative)]
prep_ct_positive =combined_data[common_condition & (positive)]
prep_ct_other =combined_data[common_condition & (other)]

prep_ct_test_result = combined_data[common_condition & (negative|positive|other)]


# Filters for PrEP_NEW
prep_new = combined_data[
    ((combined_data['Date Of Commencement (yyyy-mm-dd)'] >= Start_of_quarter) & 
     (combined_data['Date Of Commencement (yyyy-mm-dd)'] <= End_of_quarter)) &
    (
        (combined_data['HIV status at PrEP Initiation'].str.contains('Negative', na=False)) | # na=False ensures NaN values don't match 'Negative'
        (combined_data['HIV status at PrEP Initiation'].isna()) | # Catches np.nan
        (combined_data['HIV status at PrEP Initiation'] == '') | # Catches explicit empty strings
        (combined_data['HIV status at PrEP Initiation'].str.strip() == '') # Catches strings with only whitespace
    ) &
    (combined_data['Age'] >= 15) &
    (~combined_data['Sex'].isna())
]



# Filters for PrEP_NEW_Pregnant and Breastfeeding
prep_new_PBF = combined_data[
    ((combined_data['Date Of Commencement (yyyy-mm-dd)']>= Start_of_quarter) & (combined_data['Date Of Commencement (yyyy-mm-dd)'] <= End_of_quarter)) &  
    (
        (combined_data['HIV status at PrEP Initiation'].str.contains('Negative', na=False)) | # na=False ensures NaN values don't match 'Negative'
        (combined_data['HIV status at PrEP Initiation'].isna()) | # Catches np.nan
        (combined_data['HIV status at PrEP Initiation'] == '') | # Catches explicit empty strings
        (combined_data['HIV status at PrEP Initiation'].str.strip() == '') # Catches strings with only whitespace
    ) &
    (combined_data['Age'] >=15) &
    (combined_data['Sex']=='Female') &
    ((combined_data['Pregnancy Status'].str.contains('Pregnant|Breastfeeding')) & (~combined_data['Pregnancy Status'].isin(['Not Pregnant'])))
]

# Filters for PrEP_NEW_Type
prep_new_type = combined_data[
    ((combined_data['Date Of Commencement (yyyy-mm-dd)']>= Start_of_quarter) & (combined_data['Date Of Commencement (yyyy-mm-dd)'] <= End_of_quarter)) &  
    (
        (combined_data['HIV status at PrEP Initiation'].str.contains('Negative', na=False)) | # na=False ensures NaN values don't match 'Negative'
        (combined_data['HIV status at PrEP Initiation'].isna()) | # Catches np.nan
        (combined_data['HIV status at PrEP Initiation'] == '') | # Catches explicit empty strings
        (combined_data['HIV status at PrEP Initiation'].str.strip() == '') # Catches strings with only whitespace
    ) &
    ((combined_data['Prep Type'].str.contains('Oral', na=False)) | 
     (combined_data['Prep Type'].isna()) | # Catches np.nan
        (combined_data['Prep Type'] == '') | # Catches explicit empty strings
        (combined_data['Prep Type'].str.strip() == '') # Catches strings with only whitespace) &
    ) &
    (combined_data['Age'] >=15) &
    (~combined_data['Sex'].isna())
]

# Filters for PrEP_NEW_Distribution
prep_new_distribution = combined_data[
    ((combined_data['Date Of Commencement (yyyy-mm-dd)']>= Start_of_quarter) & (combined_data['Date Of Commencement (yyyy-mm-dd)'] <= End_of_quarter)) &  
    (
        (combined_data['HIV status at PrEP Initiation'].str.contains('Negative', na=False)) | # na=False ensures NaN values don't match 'Negative'
        (combined_data['HIV status at PrEP Initiation'].isna()) | # Catches np.nan
        (combined_data['HIV status at PrEP Initiation'] == '') | # Catches explicit empty strings
        (combined_data['HIV status at PrEP Initiation'].str.strip() == '') # Catches strings with only whitespace
    ) &
    ((combined_data['Prep Distribution Setting'].str.contains('Facility|Community')) & (~combined_data['Prep Distribution Setting'].isna())) &
    (combined_data['Age'] >=15) &
    (~combined_data['Sex'].isna())
]




# Pivot_data function to group by 'ProjectName' AND 'Facility Name'
# This ensures we can later filter by project but still have facility-level aggregates
def pivot_data(data, value_name):
    # Ensure the data passed to pivot_data has the required columns
    for col in ['ProjectName', 'Facility Name','Facility Id (Datim)']:
        if col not in data.columns:
            print(f"Error: '{col}' not found in data passed to pivot_data. Skipping pivot.")
            return pd.DataFrame(columns=['ProjectName', 'Facility Name', 'Facility Id (Datim)', value_name]) # Return empty DataFrame

    return data.groupby(['ProjectName', 'Facility Name', 'Facility Id (Datim)']).size().reset_index(name=value_name)

# Apply the pivot, aggregated by ProjectName and Facility Name
prep_ct_pivot = pivot_data(prep_ct, 'PrEP_CT')
prep_ct_type_pivot = pivot_data(prep_ct_type, 'PrEP_CT_Type')
prep_ct_distribution_pivot = pivot_data(prep_ct_distribution, 'PrEP_CT_Distribution')
prep_ct_test_result_pivot = pivot_data(prep_ct_test_result, 'PrEP_CT_TestResult')
prep_ct_PBF_pivot = pivot_data(prep_ct_PBF, 'PrEP_CT_PregnantandBreastfeeding')
prep_new_pivot = pivot_data(prep_new, 'PrEP_NEW')
prep_new_type_pivot = pivot_data(prep_new_type, 'PrEP_NEW_Type')
prep_new_distribution_pivot = pivot_data(prep_new_distribution, 'PrEP_NEW_Distribution')
prep_new_PBF_pivot = pivot_data(prep_new_PBF, 'PrEP_NEW_PregnantandBreastfeeding')

# List of all pivots to be merged into the master summary DataFrame
all_pivots_for_summary = [
    prep_ct_pivot,
    prep_ct_type_pivot,
    prep_ct_distribution_pivot,
    prep_ct_test_result_pivot,
    prep_ct_PBF_pivot,
    prep_new_pivot,
    prep_new_type_pivot,
    prep_new_distribution_pivot,
    prep_new_PBF_pivot
    ] 

# Get all unique combinations of ProjectName and Facility Name from the original combined data
# This ensures all facilities are represented, even if they have no metrics for a pivot
if not combined_data.empty:
    unique_project_facility_combinations = combined_data[['ProjectName', 'Facility Name','Facility Id (Datim)']].drop_duplicates()
else:
    unique_project_facility_combinations = pd.DataFrame(columns=['ProjectName', 'Facility Name', 'Facility Id (Datim)'])

# Initialize a master DataFrame that will contain all aggregated data (by ProjectName and Facility Name)
master_aggregated_df = unique_project_facility_combinations.copy()

# Merge all calculated pivots onto the master aggregated DataFrame
for pivot_df in all_pivots_for_summary:
    # Ensure pivot_df is not empty and has the keys before merging
    if not pivot_df.empty and 'ProjectName' in pivot_df.columns and 'Facility Name' in pivot_df.columns:
        master_aggregated_df = pd.merge(master_aggregated_df, pivot_df, 
                                        on=['ProjectName', 'Facility Name', 'Facility Id (Datim)'], how='left')
    else:
        print(f"Warning: An aggregated pivot is empty or missing key columns (ProjectName/Facility Name/Facility Id (Datim)) and will not be merged.")
        


# Fill NaN values (for facilities with no data for a particular metric) with 0 for counts
for col in master_aggregated_df.columns:
    if col not in ['ProjectName', 'Facility Name', 'Facility Id (Datim)']: # Only fill for metric columns
        master_aggregated_df[col] = master_aggregated_df[col].fillna(0).astype(int)

# Get unique project names for creating separate sheets
project_names = master_aggregated_df['ProjectName'].unique()

# Create the Excel writer object
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    # Loop through each unique ProjectName and save its aggregated data to a separate sheet
    if master_aggregated_df.empty:
        print("Master aggregated DataFrame is empty. No project sheets will be created.")
    else:
        for project in project_names:
            # Sanitize project name for Excel sheet name (max 31 chars, no invalid chars)
            sheet_name = str(project)[:31]
            sheet_name = re.sub(r'[\\/*?[\]:]', '', sheet_name) # Remove invalid characters

            # Filter the master aggregated data for the current project
            project_aggregated_data = master_aggregated_df[master_aggregated_df['ProjectName'] == project].copy()
            
            # Drop the ProjectName column before saving, as it's redundant on a project-specific sheet
            if 'ProjectName' in project_aggregated_data.columns:
                project_aggregated_data = project_aggregated_data.drop(columns=['ProjectName'])
            
            # Save project-specific aggregated data to its own sheet
            if not project_aggregated_data.empty:
                project_aggregated_data.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Aggregated data for Project '{project}' (by Facility Name and Facility Id (Datim)) saved to sheet '{sheet_name}'.")
            else:
                print(f"No aggregated data found for Project: {project} after filtering, skipping sheet creation.")

print(f"Analysis complete. Results saved to: {output_file_path}")
