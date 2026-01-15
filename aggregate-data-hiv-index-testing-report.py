import pandas as pd
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import re

# Path to the directory containing the CSV files
folder_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/HTS_IPP/hts 3rd Oct/CS_Filtered2'  

# Output path for the final Excel file
output_file_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Output_by_facility/HTS_Aggregates by Facility_CS_3rd Oct_ACE-1b.xlsx'


# Defining Periods
Start_of_quarter = pd.to_datetime('2025-01-01')
End_of_quarter = pd.to_datetime('2025-03-31')


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
    combined_data[col] = combined_data[col].astype(str).fillna('UNKNOWN')

    if (combined_data[col] == 'UNKNOWN').all():
        print(f"Warning: Column '{col}' is entirely 'UNKNOWN' (or NaN in original data). This might affect grouping and filtering.")
    elif combined_data[col].nunique() == 1 and combined_data[col].iloc[0] == 'UNKNOWN':
         print(f"Warning: Column '{col}' contains only 'UNKNOWN' values. This might indicate an issue with data extraction or input.")



# Ensure date columns are in datetime format
date_columns = [
    'Date offered index testing'
]
for col in date_columns:
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

# Check for invalid dates
invalid_dates = combined_data[date_columns].isnull().any(axis=1)
if invalid_dates.any():
    print(f"Warning: Invalid dates found in the following rows:\n{combined_data[invalid_dates]}")

# Offered Index
offered_index_fac = combined_data[
    (combined_data['Index client entry point'].isin(['Facility'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]

# Accepted Index
Accepted_index_fac = combined_data[
    (combined_data['Index client entry point'].isin(['Facility'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]

# Elicited Index
Elicited_index_fac = combined_data[
    (combined_data['Index client entry point'].isin(['Facility'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]

HTS_Index_total_fac = combined_data[
    (combined_data['Index client entry point'].isin(['Facility'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (combined_data['Date of HTS'] >= Start_of_quarter) &
    (combined_data['Date of HTS'] <= End_of_quarter) &
    (combined_data['HIV Test Result'].isin(['Positive', 'Negative'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]

HTS_Index_knownpositive_fac = combined_data[
    (combined_data['Index client entry point'].isin(['Facility'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (combined_data['Date of HTS'] >= Start_of_quarter) &
    (combined_data['Date of HTS'] <= End_of_quarter) &
    (combined_data['elicitedclientknownpositive'].isin(['Yes'])) &
    (combined_data['HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]


HTS_Index_newpositive_fac = combined_data[
    (combined_data['Index client entry point'].isin(['Facility'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (combined_data['Date of HTS'] >= Start_of_quarter) &
    (combined_data['Date of HTS'] <= End_of_quarter) &
    (combined_data['elicitedclientknownpositive'].isin(['No'])) &
    (combined_data['HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]


HTS_Index_newnegative_fac = combined_data[
    (combined_data['Index client entry point'].isin(['Facility'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (combined_data['Date of HTS'] >= Start_of_quarter) &
    (combined_data['Date of HTS'] <= End_of_quarter) &
    (combined_data['Age'] >=2) &
    (combined_data['elicitedclientknownpositive'].isin(['No'])) &
    (combined_data['HIV Test Result'].isin(['Negative'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]


HTS_Index_docnegative_fac = combined_data[
    (combined_data['Index client entry point'].isin(['Facility'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (combined_data['Date of HTS'] >= Start_of_quarter) &
    (combined_data['Date of HTS'] <= End_of_quarter) &
    (combined_data['Age'] <2) &
    (combined_data['elicitedclientknownpositive'].isin(['No'])) &
    (combined_data['HIV Test Result'].isin(['Negative'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]

offered_index_comm = combined_data[
    (combined_data['Index client entry point'].isin(['Community'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]

# Filters for TX_CURR-ARV Dispense
Accepted_index_comm = combined_data[
    (combined_data['Index client entry point'].isin(['Community'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]


Elicited_index_comm = combined_data[
    (combined_data['Index client entry point'].isin(['Community'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]

HTS_Index_total_comm = combined_data[
    (combined_data['Index client entry point'].isin(['Community'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (combined_data['Date of HTS'] >= Start_of_quarter) &
    (combined_data['Date of HTS'] <= End_of_quarter) &
    (combined_data['HIV Test Result'].isin(['Positive', 'Negative'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]

HTS_Index_knownpositive_comm = combined_data[
    (combined_data['Index client entry point'].isin(['Community'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (combined_data['Date of HTS'] >= Start_of_quarter) &
    (combined_data['Date of HTS'] <= End_of_quarter) &
    (combined_data['elicitedclientknownpositive'].isin(['Yes'])) &
    (combined_data['HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]


HTS_Index_newpositive_comm = combined_data[
    (combined_data['Index client entry point'].isin(['Community'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (combined_data['Date of HTS'] >= Start_of_quarter) &
    (combined_data['Date of HTS'] <= End_of_quarter) &
    (combined_data['elicitedclientknownpositive'].isin(['No'])) &
    (combined_data['HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]


HTS_Index_newnegative_comm = combined_data[
    (combined_data['Index client entry point'].isin(['Community'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (combined_data['Date of HTS'] >= Start_of_quarter) &
    (combined_data['Date of HTS'] <= End_of_quarter) &
    (combined_data['Age'] >=2) &
    (combined_data['elicitedclientknownpositive'].isin(['No'])) &
    (combined_data['HIV Test Result'].isin(['Negative'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
]


HTS_Index_docnegative_comm = combined_data[
    (combined_data['Index client entry point'].isin(['Community'])) &
    (combined_data['Date offered index testing'] >= Start_of_quarter) &
    (combined_data['Date offered index testing'] <= End_of_quarter) &
    (combined_data['Accepted Index Testing'] == 'Yes') &
    (combined_data['Date of Elicitation'] >= Start_of_quarter) &
    (combined_data['Date of Elicitation'] <= End_of_quarter) &
    (combined_data['Date of HTS'] >= Start_of_quarter) &
    (combined_data['Date of HTS'] <= End_of_quarter) &
    (combined_data['Age'] <2) &
    (combined_data['elicitedclientknownpositive'].isin(['No'])) &
    (combined_data['HIV Test Result'].isin(['Negative'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna())
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
offered_index_fac_pivot = pivot_data(offered_index_fac, 'Offered_Index_Fac')
Accepted_index_fac_pivot = pivot_data(Accepted_index_fac, 'Accepted_Index_Fac')
Elicited_index_fac_pivot = pivot_data(Elicited_index_fac, 'Elicited_Index_Fac')
HTS_Index_total_fac_pivot = pivot_data(HTS_Index_total_fac, 'HTS_Index_Total_Fac')
HTS_Index_knownpositive_fac_pivot = pivot_data(HTS_Index_knownpositive_fac, 'HTS_Index_KnownPos_Fac')
HTS_Index_newpositive_fac_pivot = pivot_data(HTS_Index_newpositive_fac, 'HTS_Index_NewPos_Fac')
HTS_Index_newnegative_fac_pivot = pivot_data(HTS_Index_newnegative_fac, 'HTS_Index_NewNeg_Fac')
HTS_Index_docnegative_fac_pivot = pivot_data(HTS_Index_docnegative_fac, 'HTS_Index_DocNeg_Fac')
offered_index_comm_pivot = pivot_data(offered_index_comm, 'Offered_Index_Fac')
Accepted_index_comm_pivot = pivot_data(Accepted_index_comm, 'Accepted_Index_Fac')
Elicited_index_comm_pivot = pivot_data(Elicited_index_comm, 'Elicited_Index_Fac')
HTS_Index_total_comm_pivot = pivot_data(HTS_Index_total_comm, 'HTS_Index_Total_Comm')
HTS_Index_knownpositive_comm_pivot = pivot_data(HTS_Index_knownpositive_comm, 'HTS_Index_KnownPos_Comm')
HTS_Index_newpositive_comm_pivot = pivot_data(HTS_Index_newpositive_comm, 'HTS_Index_NewPos_Comm')
HTS_Index_newnegative_comm_pivot = pivot_data(HTS_Index_newnegative_comm, 'HTS_Index_NewNeg_Comm')
HTS_Index_docnegative_comm_pivot = pivot_data(HTS_Index_docnegative_comm, 'HTS_Index_DocNeg_Comm')


# List of all pivots to be merged into the master summary DataFrame
all_pivots_for_summary = [
    offered_index_fac_pivot,
    Accepted_index_fac_pivot,
    Elicited_index_fac_pivot,
    HTS_Index_total_fac_pivot,
    HTS_Index_knownpositive_fac_pivot,
    HTS_Index_newpositive_fac_pivot,
    HTS_Index_docnegative_fac_pivot,
    HTS_Index_newnegative_fac_pivot,
    offered_index_comm_pivot,
    Accepted_index_comm_pivot,
    Elicited_index_comm_pivot,
    HTS_Index_total_comm_pivot,
    HTS_Index_knownpositive_comm_pivot,
    HTS_Index_newpositive_comm_pivot,
    HTS_Index_docnegative_comm_pivot,
    HTS_Index_newnegative_comm_pivot
    
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
        print(f"Warning: An aggregated pivot is empty or missing key columns (ProjectName/Facility Name/Datim Id) and will not be merged.")
        


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
                print(f"Aggregated data for Project '{project}' (by Facility Name and Datim Id) saved to sheet '{sheet_name}'.")
            else:
                print(f"No aggregated data found for Project: {project} after filtering, skipping sheet creation.")

print(f"Analysis complete. Results saved to: {output_file_path}")
