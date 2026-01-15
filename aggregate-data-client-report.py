import pandas as pd
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import re

# Path to the directory containing the CSV files
folder_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/FY25Q4_RADET/FY25Q4 Reporting/IP_RADET/new'#/Updated'

# Output path for the final Excel file
output_file_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Output_by_facility/IP_CARE2_Treatment_data_IP_FY25Q4_updateddd.xlsx'

# Path for the separate viral load output file
viral_load_output_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Cleaned_Viral_Load_Values.xlsx'

# Path for the unique CD4 values output file
unique_cd4_output_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Cleaned_CD4_Values.xlsx'


# Defining Periods
Start_of_quarter = pd.to_datetime('2025-07-01')
End_of_quarter = pd.to_datetime('2025-09-30')
vl_start = End_of_quarter - pd.DateOffset(months=12, days=-1) #12 months from the end of the quarter
six_months_ago = pd.to_datetime('2025-04-01')
End_of_vl_month = Start_of_quarter + pd.DateOffset(months=4, days=-1) #taking into consideration the one month of result received for vl samples collected within the 12 months period

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

# --- NEW: Robust column validation ---
# Check if 'ProjectName' and 'Facility Name' columns exist and are not entirely empty
required_cols_for_aggregation = ['ProjectName', 'Facility Name', 'DatimId']
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
    'ART Start Date (yyyy-mm-dd)', 
    'Date of Current ViralLoad Result Sample (yyyy-mm-dd)', 
    'Date of Current Viral Load (yyyy-mm-dd)',
    'Confirmed Date of Previous ART Status',
    'Date of Current ART Status',
    'Date of Precancerous Lesions Treatment (yyyy-mm-dd)',
    'Date of Cervical Cancer Screening (yyyy-mm-dd)',
    'Date of TB Screening (yyyy-mm-dd)',
    'Date of TB Sample Collection (yyyy-mm-dd)',
    'Date of TB Diagnostic Result Received (yyyy-mm-dd)',
    'Date of Start of TB Treatment (yyyy-mm-dd)',
    'Date of TPT Start (yyyy-mm-dd)',
    'TPT Completion date (yyyy-mm-dd)',
    'Date of Last CD4 Count'
]
for col in date_columns:
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

# Checks for invalid dates
invalid_dates = combined_data[date_columns].isnull().any(axis=1)
if invalid_dates.any():
    print(f"Warning: Invalid dates found in the following rows:\n{combined_data[invalid_dates]}")


# --- Updated Viral Load Cleaning Function ---
def clean_viral_load(value):
    """
    Cleans and maps viral load values based on specific rules,
    returning a string 'NULL' for missing or non-mappable values.
    """
    if pd.isna(value) or str(value).strip() == '':
        return 'NULL'  # Return 'NULL' for NaN/blank values
    
    val_str = str(value).strip().lower().replace(' ', '') # Normalize by removing ALL spaces

    # Rule: Undetected/Not Detected/TNF/TND variants
    # The regex patterns are now simplified to match the normalized string
    undetected_patterns = [
        r'undetected', r'notdetected', r'tnf', r'tnd', r'nd', r't\.n\.d', r'notdet', r'not/d'
    ]
    if any(re.search(pattern, val_str) for pattern in undetected_patterns):
        return 0
    
    # Rule: Handle integer/float patterns, including those with special characters
    match = re.search(r'^[<>=]?(-?[0-9,]+(?:(?:\.|,)[0-9]+)?)', val_str)
    
    if match:
        try:
            number_str = match.group(1).replace(',', '').replace('..', '.')
            return float(number_str)
        except (ValueError, TypeError):
            pass

    # All other values not caught by the above rules will return 'NULL'
    return 'NULL'

# Apply the cleaning function to create the new column
combined_data['Cleaned Current Viral Load (c/ml)'] = combined_data['Current Viral Load (c/ml)'].apply(clean_viral_load)

# --- END OF SECTION ---



# Convert columns to integers where possible
def convert_to_integer_columns(df, columns):
    for column in columns:
        df[column] = pd.to_numeric(df[column], errors='coerce')  # Convert to numeric, invalid entries become NaN
        df[column] = df[column].fillna(0).astype(int)  # Replace NaN with 0 and convert to integer
        
    return df

# Example usage
columns_to_clean = ['Last CD4 Count',  'Months of ARV Refill', 'Age'] #'Current Viral Load (c/ml)',
combined_data = convert_to_integer_columns(combined_data, columns_to_clean)


def standardize_art_status(df):
    df.loc[df['Current ART Status'].str.contains('STOPPED TREATMENT', case=False, na=False), 'Current ART Status'] = 'Stopped Treatment'
    df.loc[df['Previous ART Status'].str.contains('STOPPED TREATMENT', case=False, na=False), 'Previous ART Status'] = 'Stopped Treatment'
    
    return df

combined_data = standardize_art_status(combined_data)

def screentype(df):
    df.loc[df['TB Screening Type'].str.contains('None', case=False, na=False), 'TB Screening Type'] = 'Others'
    
    return df

combined_data = screentype(combined_data)

# Filters for TX_CURR
tx_curr = combined_data[
    (combined_data['Current ART Status'].str.contains('Active')) &
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) | #str.strip() == 'valid' ) |#
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &  
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]



# Filters for TX_CURR-ARV Dispense
tx_curr_ARV_Disp = combined_data[
    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    (~(combined_data['Months of ARV Refill'].isna()) | ~(combined_data['Months of ARV Refill'] == 0.0)) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


vl_eligibility = (combined_data[
      (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
     ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    ((combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 180)&
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
    ])

# Filters for TX_PVLS_D
tx_pvls_d = (combined_data[
    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | 
     (combined_data['Client Verification Outcome'] == '') | 
     (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] >= vl_start) &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] >= vl_start) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] <= End_of_vl_month) &
    ((combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 180) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
])


tx_pvls_d_pbf = (combined_data[
      (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
     ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
     (combined_data['Pregnancy Status'].isin(['Pregnant', 'Breastfeeding'])) &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] >= vl_start) &       
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] >=vl_start) & 
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] <= End_of_vl_month) &
    ((combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 180) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')

    ])


tx_pvls_d_pregnant = (combined_data[
      (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
     ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
     (combined_data['Pregnancy Status'] =='Pregnant') &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] >= vl_start) &        
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] >= vl_start) &         
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] <= End_of_vl_month) &
    ((combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 180)&
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')

    ])

tx_pvls_d_breastfeeding = (combined_data[
      (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
     ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
     (combined_data['Pregnancy Status'] =='Breastfeeding') &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] >= vl_start) &         
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] >= vl_start) &              
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] <= End_of_vl_month) &
    ((combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 180) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')

    ])


# Filters for TX_PVLS_N
tx_pvls_n = (combined_data[
    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) |
     (combined_data['Client Verification Outcome'] == '') |
     (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] >= vl_start) &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] >= vl_start) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] <= End_of_vl_month) &
    ((combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 180) &
    (combined_data['Cleaned Current Viral Load (c/ml)'] != 'NULL') &
    (pd.to_numeric(combined_data['Cleaned Current Viral Load (c/ml)'], errors='coerce').between(0, 999, inclusive='both').fillna(False)) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
])

tx_pvls_n_pbf = (combined_data[
      (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
     ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
     (combined_data['Pregnancy Status'].isin(['Pregnant', 'Breastfeeding'])) &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] >= vl_start)&  
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] >= vl_start)&  
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] <= End_of_vl_month) &
    ((combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 180) &
    (combined_data['Cleaned Current Viral Load (c/ml)'] != 'NULL') &
    (pd.to_numeric(combined_data['Cleaned Current Viral Load (c/ml)'], errors='coerce').between(0, 999, inclusive='both').fillna(False)) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
    ])

tx_pvls_n_pregnant = (combined_data[
      (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
     ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
     (combined_data['Pregnancy Status'] =='Pregnant') &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] >= vl_start)&             
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] >= vl_start)&                        
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] <= End_of_vl_month) &
    ((combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 180) &
    (combined_data['Cleaned Current Viral Load (c/ml)'] != 'NULL') &
    (pd.to_numeric(combined_data['Cleaned Current Viral Load (c/ml)'], errors='coerce').between(0, 999, inclusive='both').fillna(False)) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
    ])

tx_pvls_n_breastfeeding = (combined_data[
      (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
     ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
     (combined_data['Pregnancy Status'] =='Breastfeeding') &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] >= vl_start)&             
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] >= vl_start)&  
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] <= End_of_vl_month) &
    ((combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 180) &
    (combined_data['Cleaned Current Viral Load (c/ml)'] != 'NULL') &
    (pd.to_numeric(combined_data['Cleaned Current Viral Load (c/ml)'], errors='coerce').between(0, 999, inclusive='both').fillna(False)) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
    ])



# --- CORRECTED CD4 COUNT MAPPING FUNCTION (MODIFIED FOR TEXT VALUES) ---
def map_cd4_count(value):
    s = str(value).strip()
    
    # Handle blanks first, returning pd.NA to keep them blank in output
    if pd.isna(value) or s == "":
        return pd.NA
    
    # Handle specific text values, returning the string 'NULL'
    if re.search(r'copies/ml|cells|positive|P0sitive|suggestive|failed|N/A', s, re.IGNORECASE):
        return 'NULL'
    
    # Replace common non-numeric characters that look like numbers
    s = s.replace('O', '0').replace('o', '0')
    
    # Regex to find the first integer or float
    match = re.search(r'[\d]+(?:\.[\d]+)?', s)
    
    if match:
        try:
            # Extract the matched string and convert to float
            cleaned_value = float(match.group(0))
            return cleaned_value
        except (ValueError, IndexError):
            return 'NULL'
    else:
        return 'NULL'

# Apply the new mapping to create the 'Cleaned Last CD4 Count' column
if 'Last CD4 Count' in combined_data.columns:
    combined_data['Cleaned Last CD4 Count'] = combined_data['Last CD4 Count'].apply(map_cd4_count)
else:
    print("Warning: 'Last CD4 Count' column not found. Skipping CD4 mapping.")
    combined_data['Cleaned Last CD4 Count'] = 'NULL'

# ----------------- MODIFIED SECTION -----------------
# Create a DataFrame of unique values for original and cleaned CD4 counts
if 'Last CD4 Count' in combined_data.columns:
    unique_cd4_values_df = combined_data[['Last CD4 Count', 'Cleaned Last CD4 Count']].drop_duplicates().sort_values(by='Last CD4 Count').reset_index(drop=True)
    unique_cd4_values_df.to_excel(unique_cd4_output_path, index=False)
    print(f"Unique CD4 Count values saved to: {unique_cd4_output_path}")

# ----------------- END OF MODIFIED SECTION -----------------


def is_valid_cd4(value):
    # Handle blank (NaN) or empty string.
    if pd.isna(value) or str(value).strip() == "":
        return False
    
    try:
        numeric_value = float(value)
        return numeric_value <= 1600
    except (ValueError, TypeError):
        return False

def grt_cd4(value):
    if pd.isna(value) or str(value).strip() == "":
        return False
    
    try:
        numeric_value = float(value)
        return numeric_value > 1600
    except (ValueError, TypeError):
        return False

    


# Filters for TX_NEW
# add all cd4 >1600 to unknown as well as blank
common_condition_tx_new = (
    ((combined_data['ART Start Date (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['ART Start Date (yyyy-mm-dd)'] <= End_of_quarter)) &
    (combined_data['Care Entry Point'] != 'Transfer-in') &
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
)

# Use the 'Cleaned Last CD4 Count' for analysis
first_condition = ((combined_data['Age'] >4) & (~combined_data['Cleaned Last CD4 Count'].isna()) & (combined_data['Cleaned Last CD4 Count'].apply(is_valid_cd4)))
second_condition = ((combined_data['Age'] <5) | ((combined_data['Age'] >4) & (combined_data['Cleaned Last CD4 Count'].apply(grt_cd4))) | ((combined_data['Age'] >4) & (combined_data['Cleaned Last CD4 Count'].isna())))


tx_new_first_condition = combined_data[common_condition_tx_new & first_condition]
tx_new_second_condition = combined_data[common_condition_tx_new & second_condition]
tx_new = combined_data[common_condition_tx_new & (first_condition | second_condition)]

# Filters for TX_NEW(P/BF)
Breastfeeding = (
    (combined_data['Pregnancy Status'].isin(['Breastfeeding'])) &
    (combined_data['Sex'] == 'Female')) 

Pregnant = (
    (combined_data['Pregnancy Status'].isin(['Pregnant'])) &
    (combined_data['Sex'] == 'Female')) 

Pregnant_and_BF = (
    (combined_data['Pregnancy Status'].isin(['Pregnant', 'Breastfeeding'])) &
    (combined_data['Sex'] == 'Female')) 

tx_new_BF = combined_data[common_condition_tx_new & ((first_condition | second_condition) & Breastfeeding)]
tx_new_Pregnant = combined_data[common_condition_tx_new & ((first_condition | second_condition) & Pregnant)]
tx_new_Pregnant_and_BF = combined_data[common_condition_tx_new & ((first_condition | second_condition) & Pregnant_and_BF)]

tx_rtt = combined_data[(
    (combined_data['Current ART Status'].isin(['Active Restart'])) &
    ((combined_data['Date of Current ART Status'] >= Start_of_quarter) & (combined_data['Date of Current ART Status'] <= End_of_quarter)) &
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    (combined_data['Previous ART Status'].isin(['IIT', 'Stopped Treatment', 'Stopped treatment', 'Invalid - Long-term IIT', 'STOPPED TREATMENT', 'Invalid – Long-term IIT'])) &#.str.contains('IIT', 'Stopped Treatment', 'STOPPED TREATMENT',)) & #, 'Invalid - Long-term IIT', 'Invalid - Long-term IIT','STOPPED TREATMENT' 'Died' 'Stopped Treatment'
    (combined_data['Confirmed Date of Previous ART Status'] < Start_of_quarter) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
    )]


# TX_ML
# Filters for TX_ML (Transferred out)
# Filters for TX_ML (Died)
# Filters for TX_ML (Stopped Treatment)
common_condition_tx_ml = (
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    ((combined_data['Date of Current ART Status'] >= Start_of_quarter) & (combined_data['Date of Current ART Status'] <= End_of_quarter)) &
    (((combined_data['Previous ART Status'].str.contains('Active')) & (combined_data['Confirmed Date of Previous ART Status'] < Start_of_quarter)) |
    ((combined_data['Previous ART Status'].isna()) & ((combined_data['ART Start Date (yyyy-mm-dd)'] >= Start_of_quarter) & (combined_data['ART Start Date (yyyy-mm-dd)'] <= End_of_quarter)))))



Transfer_out = (combined_data['Current ART Status'].str.contains('Transfer')) 
Died = (combined_data['Current ART Status'].str.contains('Died')) 
Stopped = (combined_data['Current ART Status'].str.contains('Stop')) 

tx_ml_Transfer_out = combined_data[common_condition_tx_ml & (Transfer_out) & (~combined_data['Sex'].isna()) & ((~combined_data['Age'].isna()) |(~combined_data['Age'] == ''))]
tx_ml_Died = combined_data[common_condition_tx_ml & (Died) & (~combined_data['Sex'].isna()) & ((~combined_data['Age'].isna()) |(~combined_data['Age'] == ''))]
tx_ml_Stopped_TX = combined_data[common_condition_tx_ml & (Stopped) & (~combined_data['Sex'].isna()) & ((~combined_data['Age'].isna()) |(~combined_data['Age'] == ''))]


# TX_ML(tx_lt_three)
# TX_ML(tx_btwn_three_to_five)
# TX_ML(tx_gt_six)
tx_lt_three = (
    (((combined_data['Date of Current ART Status'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 0) &
    ((combined_data['Date of Current ART Status'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days < 90)) &
    (combined_data['Current ART Status'].isin(['IIT'])))


tx_btwn_three_to_five = (
    (((combined_data['Date of Current ART Status'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 90) &
    ((combined_data['Date of Current ART Status'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days <= 179)) &
    (combined_data['Current ART Status'].isin(['IIT'])))


tx_gt_six = (
    ((combined_data['Date of Current ART Status'] - combined_data['ART Start Date (yyyy-mm-dd)']).dt.days >= 180) &
(combined_data['Current ART Status'].isin(['IIT'])))


tx_ml_IIT = (combined_data[common_condition_tx_ml & (combined_data['Current ART Status'].isin(['IIT'])) & (~combined_data['Sex'].isna())])
tx_ml_IIT_lt_three = combined_data[common_condition_tx_ml & (tx_lt_three)] 
tx_ml_IIT_btwn_three_to_five = combined_data[common_condition_tx_ml & (tx_btwn_three_to_five)] 
tx_ml_IIT_gt_six = combined_data[common_condition_tx_ml & (tx_gt_six)] 


# Filter TX_ML_Died Cause of death
combined_data['Cause of Death'] = combined_data['Cause of Death'].str.lower() #.fillna()
tx_ml_Died_Unknown = (combined_data[(
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    ((combined_data['Date of Current ART Status'] >= Start_of_quarter) & (combined_data['Date of Current ART Status'] <= End_of_quarter)) &
    (((combined_data['Previous ART Status'].str.contains('Active')) & (combined_data['Confirmed Date of Previous ART Status'] < Start_of_quarter)) |
    ((combined_data['Previous ART Status'].isna()) & ((combined_data['ART Start Date (yyyy-mm-dd)'] >= Start_of_quarter) & (combined_data['ART Start Date (yyyy-mm-dd)'] <= End_of_quarter)))) &
    (combined_data['Current ART Status'].str.contains('Died'))  &
    (combined_data['Cause of Death'].isin(['Unknown', 'unknown', 'uknown', 'unknown cause'])))
    ])

tx_ml_Died_Non_natural = (combined_data[(
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    ((combined_data['Date of Current ART Status'] >= Start_of_quarter) & (combined_data['Date of Current ART Status'] <= End_of_quarter)) &
    (((combined_data['Previous ART Status'].str.contains('Active')) & (combined_data['Confirmed Date of Previous ART Status'] < Start_of_quarter)) |
    ((combined_data['Previous ART Status'].isna()) & ((combined_data['ART Start Date (yyyy-mm-dd)'] >= Start_of_quarter) & (combined_data['ART Start Date (yyyy-mm-dd)'] <= End_of_quarter)))) &
    (combined_data['Current ART Status'].str.contains('Died')) &
    (combined_data['Cause of Death'].isin(['Non-natural causes', 'non-natural causes']))) # 'Suspected ARV Side effect (Specify)',
    ])

tx_ml_Died_Other_HIV_Disease = (combined_data[(
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    ((combined_data['Date of Current ART Status'] >= Start_of_quarter) & (combined_data['Date of Current ART Status'] <= End_of_quarter)) &
    (((combined_data['Previous ART Status'].str.contains('Active')) & (combined_data['Confirmed Date of Previous ART Status'] < Start_of_quarter)) |
    ((combined_data['Previous ART Status'].isna()) & ((combined_data['ART Start Date (yyyy-mm-dd)'] >= Start_of_quarter) & (combined_data['ART Start Date (yyyy-mm-dd)'] <= End_of_quarter)))) &
    (combined_data['Current ART Status'].str.contains('Died'))  &
    (combined_data['Cause of Death'].isin(['other hiv disease resulting in other disease or conditions leading to death', 'suspected arv side effect (speciify)'])))
    ])

tx_ml_Died_other_infectious = (combined_data[(
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    ((combined_data['Date of Current ART Status'] >= Start_of_quarter) & (combined_data['Date of Current ART Status'] <= End_of_quarter)) &
    (((combined_data['Previous ART Status'].str.contains('Active')) & (combined_data['Confirmed Date of Previous ART Status'] < Start_of_quarter)) |
    ((combined_data['Previous ART Status'].isna()) & ((combined_data['ART Start Date (yyyy-mm-dd)'] >= Start_of_quarter) & (combined_data['ART Start Date (yyyy-mm-dd)'] <= End_of_quarter)))) &
    (combined_data['Current ART Status'].str.contains('Died'))  &
    (combined_data['Cause of Death'].isin(['suspected opportunistic infection (specify)', 'hiv disease resulting in other infectious and parasitic disease'])))
    ])

tx_ml_Died_TB = (combined_data[(
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    ((combined_data['Date of Current ART Status'] >= Start_of_quarter) & (combined_data['Date of Current ART Status'] <= End_of_quarter)) &
    (((combined_data['Previous ART Status'].str.contains('Active')) & (combined_data['Confirmed Date of Previous ART Status'] < Start_of_quarter)) |
    ((combined_data['Previous ART Status'].isna()) & ((combined_data['ART Start Date (yyyy-mm-dd)'] >= Start_of_quarter) & (combined_data['ART Start Date (yyyy-mm-dd)'] <= End_of_quarter)))) &
    (combined_data['Current ART Status'].str.contains('Died'))  &
    (combined_data['Cause of Death'].isin(['hiv disease resulting in tb', 'tuberculosis'])))
    ])

tx_ml_Died_cancer = (combined_data[(
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    ((combined_data['Date of Current ART Status'] >= Start_of_quarter) & (combined_data['Date of Current ART Status'] <= End_of_quarter)) &
    (((combined_data['Previous ART Status'].str.contains('Active')) & (combined_data['Confirmed Date of Previous ART Status'] < Start_of_quarter)) |
    ((combined_data['Previous ART Status'].isna()) & ((combined_data['ART Start Date (yyyy-mm-dd)'] >= Start_of_quarter) & (combined_data['ART Start Date (yyyy-mm-dd)'] <= End_of_quarter)))) &
    (combined_data['Current ART Status'].str.contains('Died')) &
    (combined_data['Cause of Death'].isin(['hiv-related (cancer,parasitic disease)'])))
    ])

tx_ml_Died_Other_natural = (combined_data[(
    ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    ((combined_data['Date of Current ART Status'] >= Start_of_quarter) & (combined_data['Date of Current ART Status'] <= End_of_quarter)) &
    (((combined_data['Previous ART Status'].str.contains('Active')) & (combined_data['Confirmed Date of Previous ART Status'] < Start_of_quarter)) |
    ((combined_data['Previous ART Status'].isna()) & ((combined_data['ART Start Date (yyyy-mm-dd)'] >= Start_of_quarter) & (combined_data['ART Start Date (yyyy-mm-dd)'] <= End_of_quarter)))) &
    (combined_data['Current ART Status'].str.contains('Died')) &
    (combined_data['Cause of Death'].isin(['other cause of death', 'natural cause', 'other natural causes'])))
    ])


# Filters for CXCA_SCRN
cxca_scrn = combined_data[
   (combined_data['Sex'] == 'Female') &
   (combined_data['Age'] >= 15) &
   ((combined_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'] <= End_of_quarter)) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Cervical Cancer Screening Type'].isin(['First Time Screening', 'Follow-up after previous negative result or suspected cancer', 'Post-treatment Follow-up'])) &
   (~combined_data['Cervical Cancer Screening Method'].isin([''])) &
   (combined_data['Result of Cervical Cancer Screening'].isin(['Negative', 'Positive', 'Suspicious for cancer']))]


# Filters for CXCA_TX
cxca_tx = combined_data[
   (combined_data['Sex'] == 'Female') &
   (combined_data['Age'] >= 15) &
   ((combined_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'] <= End_of_quarter)) &
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    (combined_data['Current ART Status'].str.contains('Active')) &
   (combined_data['Cervical Cancer Screening Type'].isin(['First Time Screening', 'Follow-up after previous negative result or suspected cancer', 'Post-treatment Follow-up'])) &
   (~combined_data['Cervical Cancer Screening Method'].isin([''])) &
   (combined_data['Result of Cervical Cancer Screening'].str.contains('Positive')) &
   ((combined_data['Date of Precancerous Lesions Treatment (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of Precancerous Lesions Treatment (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['Cervical Cancer Screening Method'].isna())
#    (combined_data['Precancerous Lesions Treatment Methods'].str.contains('cryotherapy', 'LEEP', 'Thermal'))
]


#Filters for TX_TB_D (include this if needed:Disaggregated by already/new on ART and TB Status)
tx_tb_d = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['TB status'].str.contains('Presumptive TB|No signs or symptoms of TB|No sign or symptoms of TB|Presumptive TB and referred for evaluation|TB Suspected and referred for evaluation|Confirmed TB|Currently on TPT')) |
   ((combined_data['TB status'].isin(['Currently on TB treatment'])) & ((combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] <=End_of_quarter)))) &
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray with CAD','Chest X-ray', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
   (~combined_data['TB status'].isna())
   ]



tx_tb_d_old_scrnpos = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['TB status'].str.contains('Presumptive TB|Presumptive TB and referred for evaluation|TB Suspected and referred for evaluation|Confirmed TB')) |
   ((combined_data['TB status'].isin(['Currently on TB treatment'])) & ((combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] <=End_of_quarter)))) &
   (combined_data['ART Start Date (yyyy-mm-dd)'] < six_months_ago) &
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray with CAD', 'Chest X-ray','Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
   (~combined_data['TB status'].isna())
   ]

tx_tb_d_new_scrnpos = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['TB status'].str.contains('Presumptive TB|Presumptive TB and referred for evaluation|TB Suspected and referred for evaluation|Confirmed TB')) |
   ((combined_data['TB status'].isin(['Currently on TB treatment'])) & ((combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] <=End_of_quarter)))) &
   ((combined_data['ART Start Date (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['ART Start Date (yyyy-mm-dd)'] <=End_of_quarter)) &
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray with CAD', 'Chest X-ray','Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
   (~combined_data['TB status'].isna())
   ]

tx_tb_d_old_scrnneg = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   (combined_data['TB status'].str.contains('No signs or symptoms of TB|No sign or symptoms of TB|Currently on TPT')) &
   (combined_data['ART Start Date (yyyy-mm-dd)'] < six_months_ago) &
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD','Chest X-ray', 'Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
   (~combined_data['TB status'].isna())
   ]


tx_tb_d_new_scrnneg = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   (combined_data['TB status'].str.contains('No signs or symptoms of TB|No sign or symptoms of TB|Currently on TPT')) &
   ((combined_data['ART Start Date (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['ART Start Date (yyyy-mm-dd)'] <=End_of_quarter)) &
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray with CAD','Chest X-ray', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
   (~combined_data['TB status'].isna())
   ]

 # Filters for TX_TB_D (Screening type)
tx_tb_d_Screening_type = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['TB status'].str.contains('Presumptive TB|No signs or symptoms of TB|No sign or symptoms of TB|TB Suspected and referred for evaluation|Presumptive TB and referred for evaluation|Confirmed TB|Currently on TPT')) |
   ((combined_data['TB status'].isin(['Currently on TB treatment'])) & ((combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] <=End_of_quarter)))) &
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray with CAD','Chest X-ray', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
   (~combined_data['TB status'].isna())
#    ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
   ]

# Filters for TX_TB_D (Specimen sent)
tx_tb_d_Specimen_sent = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) &
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray with CAD','Chest X-ray', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
   (combined_data['TB status'].str.contains('Presumptive TB|Presumptive TB and referred for evaluation|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter))
   ]


# 'TB Diagnostic Result' to lowercase for case-insensitive matching
combined_data['TB Diagnostic Result'] = combined_data['TB Diagnostic Result'].str.lower().fillna('')


# Define the pattern for positive and negative results
pos_neg_pattern = r'pos|neg|\+|\-|\_|\+ve|\-ve|nag|p0s|nrg|pso|ng|ned'
chest_x_ray_pos_neg_pattern = r'not sugestive|suggestive|mbt detectected|mt detected|mt not detected|mtb n0t detected|dectected|mtb not detectd|not detected|detected|dtected|detectted|detectd|dedected|detect|mtbd|deteted|dectected|\-mtb'
mtb_pos_neg_pattern = r'mbt detectected|mt detected|mt not detected|mtb n0t detected|dectected|mtb not detectd|not detected|detected|dtected|detectted|detectd|dedected|detect|mtbd|deteted|mtb trace|dectected|\-mtb|error|incomplete|invalid'

tb_exclusion_pattern = r'(?:tb\s+(?:positive|negative|pos|neg|ned))'
afb_exclusion_pattern = r'(?:afb\s+(?:positive|negative|pos|neg|ned))'

pos_pattern = r'pos|\+|\+ve|p0s|pso'
chest_x_ray_pos_pattern = r'^(suggestive|x-ray suggestive|mtb detected|mbt detectected|mt detected|detected|mtbdetect|mtb detectected|mtb detectted|detectted|mtb detectd|detectd|mtb dectected|dectected|mtb dedected|dedected|mtb dtected|dtected|mtd detected|mtb detectted|detectted|mtbd|mtb trace|detect|ptb detect|dedected|ptb suspect|deteted|\+mtb)'
mtb_pos_pattern = r'^(mtb detected|mbt detectected|mt detected|detected|mtbdetect|mtb detectected|mtb detectted|detectted|mtb detectd|detectd|mtb dectected|dectected|mtb dedected|dedected|mtb trace|mtb dtected|dtected|mtd detected|mtb detectted|detectted|mtbd|mtb trace|detect|ptb detect|dedected|ptb suspect|deteted|\+mtb)'

tb_pos_exclusion_pattern = r'(?:TB\s+(?:positive|pos|p0s))'
afb_pos_exclusion_pattern = r'(?:AFB\s+(?:positive|pos|p0s))'


tx_tb_d_TB_Test_Type = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Presumptive TB and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
    ((combined_data['TB Diagnostic Test Type'].str.contains('LF-LAM|TB LAMP|Clinical evaluation only', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_exclusion_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_exclusion_pattern, na=False)))  |

    ((combined_data['TB Diagnostic Test Type'].str.contains('AFB Smear Microscopy', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_exclusion_pattern, na=False)))    |
    
    ((combined_data['TB Diagnostic Test Type'].str.contains('TB-LAM', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_exclusion_pattern, na=False)))   |
     
    ((combined_data['TB Diagnostic Test Type'].str.contains('Chest X-ray', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(chest_x_ray_pos_neg_pattern, na=False)))   |
     
    ((combined_data['TB Diagnostic Test Type'].str.contains('Gene Xpert|TrueNAT', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(mtb_pos_neg_pattern, na=False)))
    )
    ]

tx_tb_d_TB_Test_Type_Xpert = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
         
    ((combined_data['TB Diagnostic Test Type'].str.contains('Gene Xpert', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(mtb_pos_neg_pattern, na=False)))
    )
    ]

tx_tb_d_TB_Test_Type_TrueNAT = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
         
    ((combined_data['TB Diagnostic Test Type'].str.contains('TrueNAT', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(mtb_pos_neg_pattern, na=False)))
    )
    ]

tx_tb_d_TB_Test_Type_Xray = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
    
    ((combined_data['TB Diagnostic Test Type'].str.contains('Chest X-ray', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(chest_x_ray_pos_neg_pattern, na=False)))  
    )
    ]

tx_tb_d_TB_Test_Type_LAM = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
        
    ((combined_data['TB Diagnostic Test Type'].str.contains('TB-LAM', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_exclusion_pattern, na=False))) |

    ((combined_data['TB Diagnostic Test Type'].str.contains('LF-LAM|TB LAMP', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_exclusion_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_exclusion_pattern, na=False))) 
     
    )
    ]

tx_tb_d_TB_Test_Type_AFB = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
    
    ((combined_data['TB Diagnostic Test Type'].str.contains('AFB Smear Microscopy', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_exclusion_pattern, na=False))) 
    )
    ]

tx_tb_d_TB_Test_Type_Clinical = combined_data[
   ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
    ((combined_data['TB Diagnostic Test Type'].str.contains('Clinical evaluation only', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_exclusion_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_exclusion_pattern, na=False)))  
    )
    ]


# Filters for TX_TB_D (Result Returned)
tx_tb_d_Result_returned = combined_data[
    ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
    ((combined_data['TB Diagnostic Test Type'].str.contains('LF-LAM|TB LAMP|Clinical evaluation only', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_pos_exclusion_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_pos_exclusion_pattern, na=False)))  |

    ((combined_data['TB Diagnostic Test Type'].str.contains('AFB Smear Microscopy', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_pos_exclusion_pattern, na=False)))    |
    
    ((combined_data['TB Diagnostic Test Type'].str.contains('TB-LAM', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_pos_exclusion_pattern, na=False)))   |
     
    ((combined_data['TB Diagnostic Test Type'].str.contains('Chest X-ray', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(chest_x_ray_pos_pattern, na=False)))   |
     
    ((combined_data['TB Diagnostic Test Type'].str.contains('Gene Xpert|TrueNAT', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(mtb_pos_pattern, na=False)))
    )
    ]


tx_tb_d_Result_returned_Xray = combined_data[
    ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
    ((combined_data['TB Diagnostic Test Type'].str.contains('Chest X-ray', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(chest_x_ray_pos_pattern, na=False)))  
    )
    ]


tx_tb_d_Result_returned_Xpert = combined_data[
    ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
         
    ((combined_data['TB Diagnostic Test Type'].str.contains('Gene Xpert', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(mtb_pos_pattern, na=False)))
    )
    ]

tx_tb_d_Result_returned_TrueNAT = combined_data[
    ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
         
    ((combined_data['TB Diagnostic Test Type'].str.contains('TrueNAT', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(mtb_pos_pattern, na=False)))
    )
    ]


tx_tb_d_Result_returned_LAM = combined_data[
    ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
    ((combined_data['TB Diagnostic Test Type'].str.contains('LF-LAM|TB LAMP', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_pos_exclusion_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_pos_exclusion_pattern, na=False)))  |
    
    ((combined_data['TB Diagnostic Test Type'].str.contains('TB-LAM', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_pos_exclusion_pattern, na=False)))   
   )
    ]


tx_tb_d_Result_returned_Clinical = combined_data[
    ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
    ((combined_data['TB Diagnostic Test Type'].str.contains('Clinical evaluation only', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_pos_exclusion_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_pos_exclusion_pattern, na=False)))   
   )
    ]


tx_tb_d_Result_returned_AFB = combined_data[
    ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
    ((combined_data['TB Diagnostic Test Type'].str.contains('AFB Smear Microscopy', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_pos_exclusion_pattern, na=False)))
    )
    ]


# Filters for TX_TB_N (Started on TB Treatment)
tx_tb_n = combined_data[
    ((combined_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((combined_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
   (combined_data['Current ART Status'].str.contains('Active')) &
   ((combined_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (combined_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~combined_data['TB Diagnostic Test Type'].isna()) &
   ((combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   ((combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
    ((combined_data['TB Diagnostic Test Type'].str.contains('LF-LAM|TB LAMP|Clinical evaluation only', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_pos_exclusion_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_pos_exclusion_pattern, na=False)))  |

    ((combined_data['TB Diagnostic Test Type'].str.contains('AFB Smear Microscopy', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(tb_pos_exclusion_pattern, na=False)))    |
    
    ((combined_data['TB Diagnostic Test Type'].str.contains('TB-LAM', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(pos_pattern, na=False)) &
    (~combined_data['TB Diagnostic Result'].str.contains(afb_pos_exclusion_pattern, na=False)))   |
     
    ((combined_data['TB Diagnostic Test Type'].str.contains('Chest X-ray', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(chest_x_ray_pos_pattern, na=False)))   |
     
    ((combined_data['TB Diagnostic Test Type'].str.contains('Gene Xpert|TrueNAT', na=False)) &
    (combined_data['TB Diagnostic Result'].str.contains(mtb_pos_pattern, na=False)))
    )]
    
    
 
# Define the target start and completion date ranges
#Print the date out just to confirm again
threehp_period1_start = Start_of_quarter - pd.DateOffset(months=9) #six months from the start of the new semi-annual period
threehp_period1_end = Start_of_quarter - pd.DateOffset(months=6, days=1) #end of last six months from the start of the new semi-annual period

threehp_period2_start = Start_of_quarter - pd.DateOffset(months=6)  #six months from the start of the new semi-annual period
threehp_period2_end = Start_of_quarter - pd.DateOffset(months=3, days=1) #end of last six months from the start of a new semi-annual period


inh_period1_start = Start_of_quarter - pd.DateOffset(months=9) #six months from the start of the new semi-annual period
inh_period1_end = Start_of_quarter - pd.DateOffset(months=3, days=1) #end of last three months from the start of a new semi-annual period

completion_period1_start = Start_of_quarter - pd.DateOffset(months=6)  #six months from the start of the new semi-annual period  
completion_period2_start = Start_of_quarter - pd.DateOffset(months=3)  #three months from the start of the new semi-annual period  
completion_period1_end = End_of_quarter
 



# Filters for TB_PREV_N
# Filter the DataFrame based on the conditions
# Condition for 3HP 
tb_prev_n_3HP_6mths =(        
        (combined_data['Date of TPT Start (yyyy-mm-dd)'] >= threehp_period1_start) & #Oct 2024
        (combined_data['Date of TPT Start (yyyy-mm-dd)'] <= threehp_period1_end) & #Dec 2024
        (combined_data['TPT Completion date (yyyy-mm-dd)'] >= completion_period1_start) & #Jan 2025
        (combined_data['TPT Completion date (yyyy-mm-dd)'] <= completion_period1_end) & #Sept 2025
        (combined_data['TPT Type'].str.contains('(3HP)|(3HR)')) &  #'Isoniazid and Rifapentine-(3HP)', 'Isoniazid and Rifampicin-(3HR)'
        (combined_data['TPT Completion status'].str.contains('Treatment Completed|Treatment success|Completed|completed')))
        

# Condition for 3HP      
tb_prev_n_3HP_3mths = (        
        ((combined_data['Date of TPT Start (yyyy-mm-dd)'] >= threehp_period2_start) & #Jan 2025
        (combined_data['Date of TPT Start (yyyy-mm-dd)'] <= threehp_period2_end)) & #Mar 2025
        ((combined_data['TPT Completion date (yyyy-mm-dd)'] >= completion_period2_start) & #April 2025
        (combined_data['TPT Completion date (yyyy-mm-dd)'] <= completion_period1_end)) & #Sept 2025
        (combined_data['TPT Type'].str.contains('(3HP)|(3HR)')) &
        (combined_data['TPT Completion status'].str.contains('Treatment Completed|Treatment success|Completed|completed')))


# Condition for INH       
tb_prev_n_INH = (        
        ((combined_data['Date of TPT Start (yyyy-mm-dd)'] >= inh_period1_start) & #Oct 2024
        (combined_data['Date of TPT Start (yyyy-mm-dd)'] <= inh_period1_end)) & #Mar 2025
        ((combined_data['TPT Completion date (yyyy-mm-dd)'] >= six_months_ago) & #April 2025
        (combined_data['TPT Completion date (yyyy-mm-dd)'] <= End_of_quarter)) & #September 2025
        (~((combined_data['TPT Type'].str.contains('(3HP)|(3HR)')) & (combined_data['TPT Type'].isna()))) &
        (combined_data['TPT Completion status'].str.contains('Treatment Completed|Treatment success|Completed|completed')))


common_conditions_tb = (
    ((combined_data['Client Verification Outcome'].fillna('').isin(['valid'])) |
     (combined_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (combined_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (combined_data['Client Verification Outcome'].str.strip() == '') ) &
    (combined_data['Current ART Status'].str.contains('Active|IIT|Stopped Treatment|Transferred Out|Died')) 
    )

Tranferin_condition = ((combined_data['Care Entry Point'].str.contains('Transfer-in')) & (combined_data['Date of Registration'] < six_months_ago))


# Filters for TB_PREV_D
# Filter the DataFrame based on the conditions
tb_prev_d_3HP_6mths =(        
        (combined_data['Date of TPT Start (yyyy-mm-dd)'] >= threehp_period1_start) &  #Oct 2024
        (combined_data['Date of TPT Start (yyyy-mm-dd)'] <= threehp_period1_end) & #Dec 2024
        (combined_data['TPT Type'].str.contains('(3HP)|(3HR)')))

# Condition for 'Date of TPT Start' 
tb_prev_d_3HP_3mths = (        
        ((combined_data['Date of TPT Start (yyyy-mm-dd)'] >= threehp_period2_start) & #Jan 2025
        (combined_data['Date of TPT Start (yyyy-mm-dd)'] <= threehp_period2_end)) & #Mar 2025
        (combined_data['TPT Type'].str.contains('(3HP)|(3HR)')))


tb_prev_d_INH = (        
        ((combined_data['Date of TPT Start (yyyy-mm-dd)'] >= inh_period1_start) & #Oct 2024
        (combined_data['Date of TPT Start (yyyy-mm-dd)'] <= inh_period1_end)) & #Mar 2025
        (~((combined_data['TPT Type'].str.contains('(3HP)|(3HR)')) & (combined_data['TPT Type'].isna())))) 

tb_prev_n = combined_data[common_conditions_tb & (tb_prev_n_3HP_6mths | tb_prev_n_3HP_3mths | tb_prev_n_INH)]
tb_prev_d = combined_data[common_conditions_tb & (tb_prev_d_3HP_6mths | tb_prev_d_3HP_3mths | tb_prev_d_INH)] #Tranferin_condition |


# Pivot_data function to group by 'ProjectName' AND 'Facility Name'
def pivot_data(data, value_name):
    # Ensure the data passed to pivot_data has the required columns
    for col in ['ProjectName', 'Facility Name','DatimId']:
        if col not in data.columns:
            print(f"Error: '{col}' not found in data passed to pivot_data. Skipping pivot.")
            return pd.DataFrame(columns=['ProjectName', 'Facility Name', 'DatimId', value_name]) # Return empty DataFrame

    return data.groupby(['ProjectName', 'Facility Name', 'DatimId']).size().reset_index(name=value_name)

# Apply the pivot, aggregated by ProjectName and Facility Name
tx_curr_pivot = pivot_data(tx_curr, 'TX_CURR')
tx_curr_ARV_Disp_pivot = pivot_data(tx_curr_ARV_Disp, 'TX_CURR_ARV_DISP')
vl_eligibility_pivot = pivot_data(vl_eligibility, 'VL_ELIGIBILITY')
tx_pvls_d_pivot = pivot_data(tx_pvls_d, 'TX_PVLS_D')
tx_pvls_d_pbf_pivot = pivot_data(tx_pvls_d_pbf, 'TX_PVLS_D_PBF')
tx_pvls_d_pregnant_pivot = pivot_data(tx_pvls_d_pregnant, 'TX_PVLS_D_Pregnant')
tx_pvls_d_breastfeeding_pivot = pivot_data(tx_pvls_d_breastfeeding, 'TX_PVLS_D_Breastfeeding')
tx_pvls_n_pivot = pivot_data(tx_pvls_n, 'TX_PVLS_N')
tx_pvls_n_pbf_pivot = pivot_data(tx_pvls_n_pbf, 'TX_PVLS_N_PBF')
tx_pvls_n_pregnant_pivot = pivot_data(tx_pvls_n_pregnant, 'TX_PVLS_N_Pregnant')
tx_pvls_n_breastfeeding_pivot = pivot_data(tx_pvls_n_breastfeeding, 'TX_PVLS_N_Breastfeeding')
tx_new_pivot = pivot_data(tx_new, 'TX_NEW')
tx_new_BF_pivot = pivot_data(tx_new_BF, 'TX_NEW_BF')
tx_rtt_pivot = pivot_data(tx_rtt, 'TX_RTT')
tx_ml_Stopped_pivot = pivot_data(tx_ml_Stopped_TX, 'TX_ML_Stopped_TX')
tx_ml_Died_pivot = pivot_data(tx_ml_Died, 'TX_ML_Died')
tx_ml_Transfer_out_pivot = pivot_data(tx_ml_Transfer_out, 'TX_ML_Transfer_out')
tx_ml_IIT_pivot = pivot_data(tx_ml_IIT, 'TX_ML_IIT') 
tx_ml_IIT_lt_three_pivot = pivot_data(tx_ml_IIT_lt_three, 'IIT<3')
tx_ml_IIT_btwn_three_to_five_pivot = pivot_data(tx_ml_IIT_btwn_three_to_five, 'IIT3-5')
tx_ml_IIT_gt_six_pivot = pivot_data(tx_ml_IIT_gt_six, 'IIT>=6')
tx_ml_Died_Unknown_pivot = pivot_data(tx_ml_Died_Unknown, 'TX_ML_Died_Unknown')
tx_ml_Died_cancer_pivot = pivot_data(tx_ml_Died_cancer, 'TX_ML_Died_cancer')
tx_ml_Died_Non_natural_pivot = pivot_data(tx_ml_Died_Non_natural, 'TX_ML_Died_Non_natural')
tx_ml_Died_Other_HIV_Disease_pivot = pivot_data(tx_ml_Died_Other_HIV_Disease, 'TX_ML_Died_Other_HIV_Disease')
tx_ml_Died_other_infectious_pivot = pivot_data(tx_ml_Died_other_infectious, 'TX_ML_Died_Other_infectious')
tx_ml_Died_Other_natural_pivot = pivot_data(tx_ml_Died_Other_natural, 'TX_ML_Died_Other_natural')
tx_ml_Died_TB_pivot = pivot_data(tx_ml_Died_TB, 'TX_ML_Died_TB')
cxca_scrn_pivot = pivot_data(cxca_scrn, 'CXCA_SCRN')
cxca_tx_pivot = pivot_data(cxca_tx, 'CXCA_TX')
tx_tb_d_pivot = pivot_data(tx_tb_d, 'TX_TB_D')
tx_tb_d_old_scrnpos_pivot = pivot_data(tx_tb_d_old_scrnpos, 'TX_TB_D_AlreadyonART_ScreenedPositive')
tx_tb_d_new_scrnpos_pivot = pivot_data(tx_tb_d_new_scrnpos, 'TX_TB_D_NewonART_ScreenedPositive')
tx_tb_d_old_scrnneg_pivot = pivot_data(tx_tb_d_old_scrnneg, 'TX_TB_D_AlreadyonART_ScreenedNegative')
tx_tb_d_new_scrnneg_pivot = pivot_data(tx_tb_d_new_scrnneg, 'TX_TB_D_NewonART_ScreenedNegative')
tx_tb_d_Screening_type_pivot = pivot_data(tx_tb_d_Screening_type, 'TX_TB_D(Screening type)')
tx_tb_d_Specimen_sent_pivot = pivot_data(tx_tb_d_Specimen_sent, 'TX_TB_D(Specimen sent)')
tx_tb_d_TB_Test_Type_pivot = pivot_data(tx_tb_d_TB_Test_Type, 'TX_TB_D(TB Test Type)')
tx_tb_d_TB_Test_Type_Xpert_pivot =pivot_data(tx_tb_d_TB_Test_Type_Xpert, 'TX_TB_D(TB Test Type)_Xpert')
tx_tb_d_TB_Test_Type_TrueNAT_pivot = pivot_data(tx_tb_d_TB_Test_Type_TrueNAT, 'TX_TB_D(TB Test Type)_TrueNAT')
tx_tb_d_TB_Test_Type_LAM_pivot = pivot_data(tx_tb_d_TB_Test_Type_LAM, 'TX_TB_D(TB Test Type)_LAM')
tx_tb_d_TB_Test_Type_Xray_pivot =pivot_data(tx_tb_d_TB_Test_Type_Xray, 'TX_TB_D(TB Test Type)_Xray')
tx_tb_d_TB_Test_Type_AFB_pivot =pivot_data(tx_tb_d_TB_Test_Type_AFB, 'TX_TB_D(TB Test Type)_AFB')
tx_tb_d_TB_Test_Type_Clinical_pivot =pivot_data(tx_tb_d_TB_Test_Type_Clinical, 'TX_TB_D(TB Test Type)_Clinical')
tx_tb_d_Result_returned_pivot = pivot_data(tx_tb_d_Result_returned, 'TX_TB_D(Result Returned)')
tx_tb_d_Result_returned_Xpert_pivot = pivot_data(tx_tb_d_Result_returned_Xpert, 'TX_TB_D(Result Returned)_Xpert')
tx_tb_d_Result_returned_TrueNAT_pivot= pivot_data(tx_tb_d_Result_returned_TrueNAT, 'TX_TB_D(Result Returned)_TrueNAT')
tx_tb_d_Result_returned_LAM_pivot= pivot_data(tx_tb_d_Result_returned_LAM, 'TX_TB_D(Result Returned)_LAM')
tx_tb_d_Result_returned_Xray_pivot= pivot_data(tx_tb_d_Result_returned_Xray, 'TX_TB_D(Result Returned)_Xray')
tx_tb_d_Result_returned_AFB_pivot= pivot_data(tx_tb_d_Result_returned_AFB, 'TX_TB_D(Result Returned)_AFB')
tx_tb_d_Result_returned_Clinical_pivot = pivot_data(tx_tb_d_Result_returned_Clinical, 'TX_TB_D(Result Returned)_Clinical')
tx_tb_n_pivot = pivot_data(tx_tb_n, 'TX_TB_N')
tb_prev_n_pivot =pivot_data(tb_prev_n, 'TB_PREV_N')
tb_prev_d_pivot =pivot_data(tb_prev_d, 'TB_PREV_D')

# List of all pivots to be merged into the master summary DataFrame
all_pivots_for_summary = [
    tx_curr_pivot,
    tx_curr_ARV_Disp_pivot,
    vl_eligibility_pivot,
    tx_pvls_d_pivot,
    tx_pvls_d_pbf_pivot,
    tx_pvls_d_pregnant_pivot,
    tx_pvls_d_breastfeeding_pivot,
    tx_pvls_n_pivot,
    tx_pvls_n_pbf_pivot,
    tx_pvls_n_pregnant_pivot,
    tx_pvls_n_breastfeeding_pivot,
    tx_new_pivot,
    tx_new_BF_pivot,
    tx_rtt_pivot,
    tx_ml_Stopped_pivot,
    tx_ml_Died_pivot,
    tx_ml_Transfer_out_pivot,
    tx_ml_IIT_pivot,
    tx_ml_IIT_lt_three_pivot,
    tx_ml_IIT_btwn_three_to_five_pivot,
    tx_ml_IIT_gt_six_pivot,
    tx_ml_Died_Unknown_pivot,
    tx_ml_Died_cancer_pivot,
    tx_ml_Died_Non_natural_pivot,
    tx_ml_Died_Other_HIV_Disease_pivot,
    tx_ml_Died_other_infectious_pivot,
    tx_ml_Died_Other_natural_pivot,
    tx_ml_Died_TB_pivot,
    cxca_scrn_pivot,
    cxca_tx_pivot,
    tx_tb_d_pivot,
    tx_tb_d_old_scrnpos_pivot,
    tx_tb_d_new_scrnpos_pivot,
    tx_tb_d_old_scrnneg_pivot,
    tx_tb_d_new_scrnneg_pivot,
    tx_tb_d_Screening_type_pivot,
    tx_tb_d_Specimen_sent_pivot,
    tx_tb_d_TB_Test_Type_pivot,
    tx_tb_d_TB_Test_Type_Xpert_pivot,
    tx_tb_d_TB_Test_Type_TrueNAT_pivot,
    tx_tb_d_TB_Test_Type_LAM_pivot,
    tx_tb_d_TB_Test_Type_Xray_pivot,
    tx_tb_d_TB_Test_Type_AFB_pivot,
    tx_tb_d_TB_Test_Type_Clinical_pivot,
    tx_tb_d_Result_returned_pivot,
    tx_tb_d_Result_returned_Xpert_pivot,
    tx_tb_d_Result_returned_TrueNAT_pivot,
    tx_tb_d_Result_returned_LAM_pivot,
    tx_tb_d_Result_returned_Xray_pivot,
    tx_tb_d_Result_returned_AFB_pivot,
    tx_tb_d_Result_returned_Clinical_pivot,
    tx_tb_n_pivot,
    tb_prev_n_pivot,
    tb_prev_d_pivot
    ] 

# Get all unique combinations of ProjectName and Facility Name from the original combined data
# This ensures all facilities are represented, even if they have no metrics for a pivot
if not combined_data.empty:
    unique_project_facility_combinations = combined_data[['ProjectName', 'Facility Name','DatimId']].drop_duplicates()
else:
    unique_project_facility_combinations = pd.DataFrame(columns=['ProjectName', 'Facility Name', 'DatimId'])

# Initialize a master DataFrame that will contain all aggregated data (by ProjectName and Facility Name)
master_aggregated_df = unique_project_facility_combinations.copy()

# Merge all calculated pivots onto the master aggregated DataFrame
for pivot_df in all_pivots_for_summary:
    # Ensure pivot_df is not empty and has the keys before merging
    if not pivot_df.empty and 'ProjectName' in pivot_df.columns and 'Facility Name' in pivot_df.columns:
        master_aggregated_df = pd.merge(master_aggregated_df, pivot_df, 
                                        on=['ProjectName', 'Facility Name', 'DatimId'], how='left')
    else:
        print(f"Warning: An aggregated pivot is empty or missing key columns (ProjectName/Facility Name/Datim Id) and will not be merged.")
        


# Fill NaN values (for facilities with no data for a particular metric) with 0 for counts
for col in master_aggregated_df.columns:
    if col not in ['ProjectName', 'Facility Name', 'DatimId']: # Only fill for metric columns
        master_aggregated_df[col] = master_aggregated_df[col].fillna(0).astype(int)

# Get unique project names for creating separate sheets
project_names = master_aggregated_df['ProjectName'].unique()

# # Create the Excel writer object
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



# --- Save original and cleaned viral load data to a new file ---
viral_load_columns = ['Current Viral Load (c/ml)', 'Cleaned Current Viral Load (c/ml)']
if all(col in combined_data.columns for col in viral_load_columns):
    viral_load_df = combined_data[viral_load_columns].copy()
    distinct_viral_load_df = viral_load_df.drop_duplicates()

    try:
        distinct_viral_load_df.to_excel(viral_load_output_path, index=False)
        print(f"\nSuccessfully saved distinct viral load data to: {viral_load_output_path}")
    except Exception as e:
        print(f"\nError saving viral load data: {e}")
else:
    print("\nWarning: 'Current Viral Load (c/ml)' or 'Cleaned Current Viral Load (c/ml)' column not found. Skipping viral load data export.")

