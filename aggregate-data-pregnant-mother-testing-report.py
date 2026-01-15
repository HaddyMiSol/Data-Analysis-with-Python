import pandas as pd
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import re

# Path to the directory containing the CSV files
folder_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/HTS_IPP/pmtct_hts'

# Output path for the final Excel file
output_file_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Output_by_facility/PMTCT_HTS_Aggregates by Facility_IP.xlsx'

# Output path for troubleshooting HIV Testing Setting and Modality Output
pmtct_setting_output_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Extracted_HTS_setting.xlsx'

# Defining Periods
Start_of_quarter = pd.to_datetime('2025-07-01')
End_of_quarter = pd.to_datetime('2025-09-30')

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


# Check if 'ProjectName' and 'Facility' columns exist and are not entirely empty
required_cols_for_aggregation = ['ProjectName', 'Facility']
for col in required_cols_for_aggregation:
    if col not in combined_data.columns:
        print(f"FATAL ERROR: Required column '{col}' is missing from the combined data.")
        if col == 'ProjectName':
            print("This usually means the 'Filename' column couldn't be processed to extract ProjectName.")
            print(f"First 5 filenames: {combined_data['Filename'].head().tolist()}")
        exit()
    
    # Convert to string and handle potential NaNs before further processing
    combined_data[col] = combined_data[col].astype(str).fillna('UNKNOWN')

    if (combined_data[col] == 'UNKNOWN').all():
        print(f"Warning: Column '{col}' is entirely 'UNKNOWN' (or NaN in original data). This might affect grouping and filtering.")
    elif combined_data[col].nunique() == 1 and combined_data[col].iloc[0] == 'UNKNOWN':
         print(f"Warning: Column '{col}' contains only 'UNKNOWN' values. This might indicate an issue with data extraction or input.")



# Ensure date columns are in datetime format
date_columns = [
    'Date Tested for HIV'
]

for col in date_columns:
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

# Check for invalid dates
invalid_dates = combined_data[date_columns].isnull().any(axis=1)
if invalid_dates.any():
    print(f"Warning: Invalid dates found in the following rows:\n{combined_data[invalid_dates]}")

# Ensure Numeric field is numeric
combined_data['Age'] = pd.to_numeric(combined_data['Age'], errors='coerce')



PMTCT_ANC_Facility = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Facility'])) &
    (combined_data['ANC Setting'].isin(['ANC', 'Spoke health facility', 'PMTCT (ANC1 Only)'])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)'])) &
    (combined_data['HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]

PMTCT_ANC_Facility_Negative = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Facility'])) &
    (combined_data['ANC Setting'].isin(['ANC', 'Spoke health facility', 'PMTCT (ANC1 Only)'])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)'])) &
    (combined_data['HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]

PMTCT_ANC_Facility_Positive = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Facility'])) &
    (combined_data['ANC Setting'].isin(['ANC', 'Spoke health facility', 'PMTCT (ANC1 Only)'])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)'])) &
    (combined_data['HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]


PMTCT_ANC_Community = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Community'])) &
    (combined_data['ANC Setting'].isin(['Congregational setting', 'Delivery homes','TBA Orthodx', 'TBA Orthodox', 'TBA rt-HCW'])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)'])) &
    (combined_data['HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]

PMTCT_ANC_Community_Negative = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Community'])) &
    (combined_data['ANC Setting'].isin(['Congregational setting', 'Delivery homes','TBA Orthodx', 'TBA Orthodox', 'TBA rt-HCW'])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)'])) &
    (combined_data['HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]

PMTCT_ANC_Community_Positive = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Community'])) &
    (combined_data['ANC Setting'].isin(['Congregational setting', 'Delivery homes','TBA Orthodx', 'TBA Orthodox', 'TBA rt-HCW'])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)'])) &
    (combined_data['HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]


PMTCT_LD = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Facility'])) &
    (combined_data['ANC Setting'].isin(['L&D'])) &
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Pregnancy/L&D)'])) &
    (combined_data['HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]

PMTCT_LD_Negative = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Facility'])) &
    (combined_data['ANC Setting'].isin(['L&D'])) &
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Pregnancy/L&D)'])) &
    (combined_data['HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]

PMTCT_LD_Positive = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Facility'])) &
    (combined_data['ANC Setting'].isin(['L&D'])) &
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Pregnancy/L&D)'])) &
    (combined_data['HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]


PMTCT_Breastfeeding = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Facility'])) &
    (combined_data['ANC Setting'].isin(['Post Natal Ward/Breastfeeding'])) &
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Breastfeeding)'])) &
    (combined_data['HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]

PMTCT_Breastfeeding_Negative = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Facility'])) &
    (combined_data['ANC Setting'].isin(['Post Natal Ward/Breastfeeding'])) &
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Breastfeeding)'])) &
    (combined_data['HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]

PMTCT_Breastfeeding_Positive = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['Point of Entry'].isin(['Facility'])) &
    (combined_data['ANC Setting'].isin(['Post Natal Ward/Breastfeeding'])) &
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Breastfeeding)'])) &
    (combined_data['HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Age'].isna()) &
    (combined_data['Age'] != '')
]



Setting_no_Modality = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (~combined_data['ANC Setting'].isna()) &
    (combined_data['Modality'].isna()) &
    (combined_data['HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

Modality_no_Setting = combined_data[
    (combined_data['Date Tested for HIV'] >= Start_of_quarter) &
    (combined_data['Date Tested for HIV'] <= End_of_quarter) &
    (combined_data['ANC Setting'].isna()) &
    (~combined_data['Modality'].isna()) &
    (combined_data['HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]



# Pivot_data function to group by 'ProjectName' AND 'Facility'
# This ensures we can later filter by project but still have facility-level aggregates
def pivot_data(data, value_name):
    # Ensure the data passed to pivot_data has the required columns
    for col in ['ProjectName', 'Facility']:
        if col not in data.columns:
            print(f"Error: '{col}' not found in data passed to pivot_data. Skipping pivot.")
            return pd.DataFrame(columns=['ProjectName', 'Facility', value_name]) # Return empty DataFrame

    return data.groupby(['ProjectName', 'Facility']).size().reset_index(name=value_name)

# Apply the pivot, aggregated by ProjectName and Facility
PMTCT_ANC_Facility_pivot = pivot_data(PMTCT_ANC_Facility, 'PMTCT_ANC_Facility')
PMTCT_ANC_Facility_Negative_pivot = pivot_data(PMTCT_ANC_Facility_Negative, 'PMTCT_ANC_Facility_Negative')
PMTCT_ANC_Facility_Positive_pivot = pivot_data(PMTCT_ANC_Facility_Positive, 'PMTCT_ANC_Facility_Positive')
PMTCT_ANC_Community_pivot = pivot_data(PMTCT_ANC_Community, 'PMTCT_ANC_Community')
PMTCT_ANC_Community_Negative_pivot = pivot_data(PMTCT_ANC_Community_Negative, 'PMTCT_ANC_Community_Negative')
PMTCT_ANC_Community_Positive_pivot = pivot_data(PMTCT_ANC_Community_Positive, 'PMTCT_ANC_Community_Positive')
PMTCT_Breastfeeding_pivot = pivot_data(PMTCT_Breastfeeding, 'PMTCT_Breastfeeding')
PMTCT_Breastfeeding_Negative_pivot = pivot_data(PMTCT_Breastfeeding_Negative, 'PMTCT_Breastfeeding_Negative')
PMTCT_Breastfeeding_Positive_pivot = pivot_data(PMTCT_Breastfeeding_Positive, 'PMTCT_Breastfeeding_Positive')
PMTCT_LD_pivot = pivot_data(PMTCT_LD, 'PMTCT_LD')
PMTCT_LD_Negative_pivot = pivot_data(PMTCT_LD_Negative, 'PMTCT_LD_Negative')
PMTCT_LD_Positive_pivot = pivot_data(PMTCT_LD_Positive, 'PMTCT_LD_Positive')
Setting_no_Modality_pivot = pivot_data(Setting_no_Modality, 'Setting_no_Modality')
Modality_no_Setting_pivot = pivot_data(Modality_no_Setting, 'Modality_no_Setting')



# List of all pivots to be merged into the master summary DataFrame
all_pivots_for_summary = [
    PMTCT_ANC_Facility_pivot,
    PMTCT_ANC_Facility_Negative_pivot,
    PMTCT_ANC_Facility_Positive_pivot,
    PMTCT_ANC_Community_pivot,
    PMTCT_ANC_Community_Negative_pivot,
    PMTCT_ANC_Community_Positive_pivot,
    PMTCT_Breastfeeding_pivot,
    PMTCT_Breastfeeding_Negative_pivot,
    PMTCT_Breastfeeding_Positive_pivot,
    PMTCT_LD_pivot,
    PMTCT_LD_Negative_pivot,
    PMTCT_LD_Positive_pivot,
    Setting_no_Modality_pivot,
    Modality_no_Setting_pivot
    ] 

# Get all unique combinations of ProjectName and Facility from the original combined data
# This ensures all facilities are represented, even if they have no metrics for a pivot
if not combined_data.empty:
    unique_project_facility_combinations = combined_data[['ProjectName', 'Facility']].drop_duplicates()
else:
    unique_project_facility_combinations = pd.DataFrame(columns=['ProjectName', 'Facility'])

# Initialize a master DataFrame that will contain all aggregated data (by ProjectName and Facility)
master_aggregated_df = unique_project_facility_combinations.copy()

# Merge all calculated pivots onto the master aggregated DataFrame
for pivot_df in all_pivots_for_summary:
    # Ensure pivot_df is not empty and has the keys before merging
    if not pivot_df.empty and 'ProjectName' in pivot_df.columns and 'Facility' in pivot_df.columns:
        master_aggregated_df = pd.merge(master_aggregated_df, pivot_df, 
                                        on=['ProjectName', 'Facility'], how='left')
    else:
        print(f"Warning: An aggregated pivot is empty or missing key columns (ProjectName/Facility) and will not be merged.")
        


# Create the Excel writer object
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    if not master_aggregated_df.empty:
        # Sort the DataFrame for better readability in the final output
        master_aggregated_df.sort_values(by=['ProjectName', 'Facility'], inplace=True)
        
        # Save the entire master aggregated data to a single sheet
        sheet_name = 'Facility_Aggregates'
        master_aggregated_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"All aggregated data saved to a single sheet '{sheet_name}'.")
    else:
        print("Master aggregated DataFrame is empty. No data will be saved.")

print(f"Analysis complete. Results saved to: {output_file_path}")



pmtct_hts_columns = ['ANC Setting', 'Modality', 'Date Tested for HIV', 'Facility']
if all(col in combined_data.columns for col in pmtct_hts_columns):
    cutoff_date = Start_of_quarter  #pd.to_datetime('2025-06-30')
    filtered_data = combined_data[combined_data['Date Tested for HIV'] >= cutoff_date].copy()
    pmtct_hts_setting_df = filtered_data[pmtct_hts_columns].copy()
    distinct_pmtct_hts_setting_df = pmtct_hts_setting_df  #.drop_duplicates()

    try:
        distinct_pmtct_hts_setting_df.to_excel(pmtct_setting_output_path, index=False)
        print(f"\nSuccessfully saved distinct hts setting data to: {pmtct_setting_output_path}")
    except Exception as e:
        print(f"\nError saving hts setting data: {e}")
else:
    print("\nWarning: 'ANC Setting' or 'Modality' column not found. Skipping hts setting data export.")
