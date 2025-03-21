import pandas as pd
import os
from datetime import datetime, timedelta
import re

# Path to the directory containing the CSV files
folder_path = 'C:/Users/DELL/Documents/DataFi/Data Review Meeting/CS/CS_RADET_Files'

# Output path for the final Excel file
output_file_path = 'C:/Users/DELL/Documents/DataFi/Data Review Meeting/CS/Analysis_Output(CS).xlsx'

# Combine all CSV files into one DataFrame, specifying 'latin1' encoding
all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(('.csv', '.xlsx'))]

# Combine all files into one DataFrame
combined_data = pd.DataFrame()

for file in all_files:
    try:
        if file.endswith('.csv'):
            data = pd.read_csv(file, encoding='latin1', on_bad_lines='skip')
        elif file.endswith('.xlsx'):
            data = pd.read_excel(file, engine='openpyxl')
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
combined_data['ProjectName'] = combined_data['Filename'].str.split('_').str[0]
#combined_data['ProjectName'] = combined_data['Facility Name']

# Get the date for "last week ending"
last_week_ending = datetime.now() - timedelta(days=7)
End_of_quarter = pd.to_datetime('2024-12-31')
End_date = '2024-12-31'

def month_difference(start_date, end_date):
    """Calculates the difference in full months between two dates.

    Args:
        start_date: Pandas Timestamp or datetime object.
        end_date: Pandas Timestamp or datetime object.

    Returns:
        The difference in full months (integer), or 0 if either date is NaN.
    """
    if pd.isna(start_date) or pd.isna(end_date):
        return 0

    diff_years = end_date.year - start_date.year
    diff_months = end_date.month - start_date.month

    # Check if the day of the end date is less than the day of the start date
    if end_date.day < start_date.day:
        diff_months -= 1  # Subtract one month if not a full month has passed

    return diff_years * 12 + diff_months

# Ensure date columns are in datetime format
date_columns = [
    'ART Start Date (yyyy-mm-dd)', 
    'Date of Current ViralLoad Result Sample (yyyy-mm-dd)', 
    'Date of Current Viral Load (yyyy-mm-dd)',
    'Confirmed Date of Previous ART Status',
    'Date of Current ART Status',
    'Date of Precancerous Lesions Treatment (yyyy-mm-dd)',
    'Date of Cervical Cancer Screening (yyyy-mm-dd)'
]
for col in date_columns:
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

# Check for invalid dates
invalid_dates = combined_data[date_columns].isnull().any(axis=1)
if invalid_dates.any():
    print(f"Warning: Invalid dates found in the following rows:\n{combined_data[invalid_dates]}")

# Ensure 'Current Viral Load (c/ml)' is numeric
combined_data['Current Viral Load (c/ml)'] = pd.to_numeric(
    combined_data['Current Viral Load (c/ml)'], errors='coerce'
)

# Filters for TX_CURR
tx_curr = combined_data[
    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) #'Client Verification Outcome'
]


# Filters for TX_PVLS_D
condition_d_a = (
    (combined_data['Current ART Regimen'].str.contains('DTG', na=False) | combined_data['Current ART Regimen'].str.contains('Dolutegravir', na=False)) &
    (combined_data.apply(lambda row: month_difference(row['ART Start Date (yyyy-mm-dd)'], row['Date of Current ViralLoad Result Sample (yyyy-mm-dd)']), axis=1) >= 3)
    )

condition_d_b = (
    ~combined_data['Current ART Regimen'].str.contains('DTG', na=False) &
    ~combined_data['Current ART Regimen'].str.contains('Dolutegravir', na=False) &
    (combined_data.apply(lambda row: month_difference(row['ART Start Date (yyyy-mm-dd)'], row['Date of Current ViralLoad Result Sample (yyyy-mm-dd)']), axis=1) >= 6)
)

common_conditions = (
    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] >= (End_of_quarter - timedelta(days=365))) & 
    (combined_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] >= (End_of_quarter - timedelta(days=365))) &
    (combined_data['Date of Current Viral Load (yyyy-mm-dd)'] <= End_of_quarter)
)

tx_pvls_d = combined_data[common_conditions & (condition_d_a | condition_d_b)]

# Filters for TX_PVLS_N
condition_n_a = (
    (combined_data['Current ART Regimen'].str.contains('DTG', na=False) | combined_data['Current ART Regimen'].str.contains('Dolutegravir', na=False)) &
    (combined_data.apply(lambda row: month_difference(row['ART Start Date (yyyy-mm-dd)'], row['Date of Current ViralLoad Result Sample (yyyy-mm-dd)']), axis=1) >= 3) &
    (combined_data['Current Viral Load (c/ml)'].fillna(float('inf')) < 1000)
)

condition_n_b = (
    ~combined_data['Current ART Regimen'].str.contains('DTG', na=False) &
    ~combined_data['Current ART Regimen'].str.contains('Dolutegravir', na=False) &
    (combined_data.apply(lambda row: month_difference(row['ART Start Date (yyyy-mm-dd)'], row['Date of Current ViralLoad Result Sample (yyyy-mm-dd)']), axis=1) >= 6) &
    (combined_data['Current Viral Load (c/ml)'].fillna(float('inf')) < 1000)
)

tx_pvls_n = combined_data[common_conditions & (condition_n_a | condition_n_b)]

#Function for TX_NEW & TX_RTT for CD4 Count
def is_valid_cd4(value):
    # Check for blank (NaN)
    if pd.isna(value):
        return False  # Blank is valid
    
    # Convert the value to a string for pattern matching
    value_str = str(value).strip()
    
    # Define valid patterns (expandable for additional valid values)
    valid_patterns = [
        r"^\d+$",  # Integer (e.g., 200)
        r"^\d+\.\d+$",  # Float (e.g., 200.5)
        #r"^[>]=?\d+$",  # Comparisons with numbers (e.g., >=200)
        #r"^[<]?\d+$",  # Comparisons with numbers (e.g., <200)
        r"^(>=200|<200)$",  # Correct regex for >=200 OR <200
        r"^\s*$" #blank
    ]
    
    # Check if the value matches any valid pattern
    for pattern in valid_patterns:
        if re.fullmatch(pattern, value_str):
            return True  # Value is valid
        
    return True



# Filters for TX_NEW
filter_date = '2024-10-01'
End_date = '2024-12-31'
tx_new = combined_data[
    ((combined_data['ART Start Date (yyyy-mm-dd)'] >= filter_date) &
    (combined_data['ART Start Date (yyyy-mm-dd)'] <= End_date)) &
    (combined_data['Care Entry Point'] != 'Transfer-in') &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Last CD4 Count'].apply(is_valid_cd4) | combined_data['Last CD4 Count'].isna())
]

# Filters for TX_RTT
tx_rtt = combined_data[
    (combined_data['Current ART Status'].isin(['Active Restart'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['IIT', 'Stopped Treatment'])) | (~combined_data['Previous ART Status'].isna())) &
    (combined_data['Confirmed Date of Previous ART Status'] < filter_date) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    #(combined_data['Previous ART Status'].isin(['IIT'])) &
    (combined_data['Last CD4 Count'].apply(is_valid_cd4) | combined_data['Last CD4 Count'].isna())
]

# Filters for TX_RTT(iit<3)
caution_date = '2024-09-30'
tx_rtt_iit_lt_three = combined_data[
    (combined_data['Current ART Status'].isin(['Active Restart'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    ((combined_data['Previous ART Status'].isin(['IIT', 'Stopped Treatment'])) | (~combined_data['Previous ART Status'].isna())) &
    #(combined_data['Previous ART Status'].isin(['IIT'])) &
    (combined_data['Confirmed Date of Previous ART Status'] < filter_date) &
    (combined_data.apply(lambda row: month_difference(row['Confirmed Date of Previous ART Status'], row['Date of Current ART Status']), axis=1) < 3) &
    (combined_data['Last CD4 Count'].apply(is_valid_cd4) | combined_data['Last CD4 Count'].isna())
]

# Filters for TX_RTT(iit3-5)
tx_rtt_btwn_three_to_five = combined_data[
    (combined_data['Current ART Status'].isin(['Active Restart'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    ((combined_data['Previous ART Status'].isin(['IIT', 'Stopped Treatment'])) | (~combined_data['Previous ART Status'].isna())) &
    #(combined_data['Previous ART Status'].isin(['IIT'])) &
    (combined_data['Confirmed Date of Previous ART Status'] < filter_date) &
    ((combined_data.apply(lambda row: month_difference(row['Confirmed Date of Previous ART Status'], row['Date of Current ART Status']), axis=1) >= 3) & 
     (combined_data.apply(lambda row: month_difference(row['Confirmed Date of Previous ART Status'], row['Date of Current ART Status']), axis=1) <= 5)) &
    (combined_data['Last CD4 Count'].apply(is_valid_cd4) | combined_data['Last CD4 Count'].isna())
]

# Filters for TX_RTT(iit>=6)
tx_rtt_iit_gt_six = combined_data[
    (combined_data['Current ART Status'].isin(['Active Restart'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    ((combined_data['Previous ART Status'].isin(['IIT', 'Stopped Treatment'])) | (~combined_data['Previous ART Status'].isna())) &
    #(combined_data['Previous ART Status'].isin(['IIT'])) &
    (combined_data['Confirmed Date of Previous ART Status'] < filter_date) &
    (combined_data.apply(lambda row: month_difference(row['Confirmed Date of Previous ART Status'], row['Date of Current ART Status']), axis=1) >= 6) &
    (combined_data['Last CD4 Count'].apply(is_valid_cd4) | combined_data['Last CD4 Count'].isna())
]

# Filters for TX_ML (Transferred out)
tx_ml_Transfer_out = combined_data[
    (combined_data['Current ART Status'].isin(['Transferred Out'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', '']))
]

# Filters for TX_ML (Died)
tx_ml_Died = combined_data[
    (combined_data['Current ART Status'].isin(['Died'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', '']))
]

# Filters for TX_ML (IIT)
tx_ml_IIT = combined_data[
    (combined_data['Current ART Status'].isin(['IIT'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', '']))
]

# Filters for TX_ML (Invalid)
Invalid = combined_data[
    #(combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['invalid'])) &
     ((combined_data['Date of Current ART Status'] >= filter_date) &
    (combined_data['Date of Current ART Status'] <= End_date)) #'Client Verification Outcome'
]

# Filters for TX_ML (Verification ongoing)
Verification_ongoing = combined_data[
    #(combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['Verification Ongoing'])) &
     ((combined_data['Date of Current ART Status'] >= filter_date) &
    (combined_data['Date of Current ART Status'] <= End_date)) #'Client Verification Outcome'
]

iit_lt_three = combined_data[
    (combined_data['Current ART Status'].isin(['IIT'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart'])) | (combined_data['Previous ART Status'].isna())) &
    (combined_data['Confirmed Date of Previous ART Status'] < filter_date) &
    (combined_data.apply(lambda row: month_difference(row['Confirmed Date of Previous ART Status'], row['Date of Current ART Status']), axis=1) < 3) 
    
]

# Filters for TX_ML(iit3-5)
iit_btwn_three_to_five = combined_data[
    (combined_data['Current ART Status'].isin(['IIT'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart'])) | (combined_data['Previous ART Status'].isna())) &
    (combined_data['Confirmed Date of Previous ART Status'] < filter_date) &
    ((combined_data.apply(lambda row: month_difference(row['Confirmed Date of Previous ART Status'], row['Date of Current ART Status']), axis=1) >= 3) & 
     (combined_data.apply(lambda row: month_difference(row['Confirmed Date of Previous ART Status'], row['Date of Current ART Status']), axis=1) <= 5))
    
]

# Filters for TXML(iit>=6)
iit_gt_six = combined_data[
    (combined_data['Current ART Status'].isin(['IIT'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart'])) | (combined_data['Previous ART Status'].isna())) &
    (combined_data['Confirmed Date of Previous ART Status'] < filter_date) &
    (combined_data.apply(lambda row: month_difference(row['Confirmed Date of Previous ART Status'], row['Date of Current ART Status']), axis=1) >= 6) 
    
]


# Filters for TX_ML (Stopped Treatment)
tx_ml_Stopped_TX = combined_data[
    (combined_data['Current ART Status'].isin(['Stopped Treatment'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', '']))
]


# Filters for TX_ML (Died) - COD (Unknown)
tx_ml_Died_Unknown = combined_data[
    (combined_data['Current ART Status'].isin(['Died'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Cause of Death'].isin(['Unknown', 'Unknown cause', 'Unknown Cause']))
]


# Filters for TX_ML (Died) - COD (Other natural)
tx_ml_Died_Other_natural = combined_data[
    (combined_data['Current ART Status'].isin(['Died'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Cause of Death'].isin(['Other natural causes']))
]

# Filters for TX_ML (Died) - COD (Other cause of death)
tx_ml_Died_Other_cause = combined_data[
    (combined_data['Current ART Status'].isin(['Died'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Cause of Death'].isin(['Other cause of death']))
]

# Filters for TX_ML (Died) - COD (Other HIV Disease)
tx_ml_Died_Other_HIV_Disease = combined_data[
    (combined_data['Current ART Status'].isin(['Died'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Cause of Death'].isin(['Other HIV disease resulting in other disease or conditions leading to death']))
]

# Filters for TX_ML (Died) - COD (Non natural)
tx_ml_Died_Non_natural = combined_data[
    (combined_data['Current ART Status'].isin(['Died'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Cause of Death'].isin(['Non-natural causes']))
]


# Filters for TX_ML (Died) - COD (Natural)
tx_ml_Died_Natural = combined_data[
    (combined_data['Current ART Status'].isin(['Died'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Cause of Death'].isin(['Natural Cause', 'Natural cause']))
]

# Filters for TX_ML (Died) - COD (Resulting in other infectious)
tx_ml_Died_other_infectious = combined_data[
    (combined_data['Current ART Status'].isin(['Died'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Cause of Death'].isin(['HIV disease resulting in other infectious and parasitic disease']))
]

# Filters for TX_ML (Died) - COD (cancer)
tx_ml_Died_cancer = combined_data[
    (combined_data['Current ART Status'].isin(['Died'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Cause of Death'].isin(['HIV disease resulting in cancer', 'HIV related (Cancer', 'HIV-related (Cancer']))
]

# Filters for TX_ML (Died) - COD (TB)
tx_ml_Died_TB = combined_data[
    (combined_data['Current ART Status'].isin(['Died'])) &
    ((combined_data['Date of Current ART Status'] >= filter_date) & (combined_data['Date of Current ART Status'] <= End_date)) &
    ((combined_data['Previous ART Status'].isin(['Active', 'Active Restart']) | (combined_data['Previous ART Status'].isna()))) &
    ((combined_data['Confirmed Date of Previous ART Status'] < filter_date) | combined_data['Confirmed Date of Previous ART Status'].isna()) &
    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
    (combined_data['Cause of Death'].isin(['HIV disease resulting in TB', 'Tuberculosis']))
]

# Filters for CXCA_SCRN
#CXCA_SCRN = combined_data[
 #   (combined_data['Sex'] == 'Female') &
 #   (combined_data['Age'] >= 15) &
   # (combined_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'] >= filter_date) &
   # (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
  #  (combined_data['Cervical Cancer Screening Type'].isin(['First Time Screening', 'Follow-up after previous negative result or suspected cancer', 'Post-treatment Follow-up'])) &
  #  (combined_data['Result of Cervical Cancer Screening'].isin(['Negative', 'Positive', 'Suspicious for cancer']))
#]

# Filters for CXCA_TX
#CXCA_TX = combined_data[
#    (combined_data['Sex'] == 'Female') &
 #   (combined_data['Age'] >= 15) &
 #   (combined_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'] >= filter_date) &
#    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
 #   (combined_data['Cervical Cancer Screening Type'].isin(['First Time Screening', 'Follow-up after previous negative result or suspected cancer', 'Post-treatment Follow-up'])) &
 #   (combined_data['Result of Cervical Cancer Screening'].isin(['Negative', 'Positive', 'Suspicious for cancer'])) &
 #   (combined_data['Date of Precancerous Lesions Treatment (yyyy-mm-dd)'] >= filter_date) &
 #   (combined_data['Precancerous Lesions Treatment Methods'].isin(['cryotherapy', 'LEEP', 'Thermal']))
#]

# Filters for TX_TB_D (include this if needed:Disaggregated by already/new on ART and TB Status)
# TB_filter_date = '2024-10-01'
#TX_TB_D = combined_data[
 #   (combined_data['Date of TB Screening (yyyy-mm-dd)'] >= TB_filter_date) & 
#    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
#    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
 #   (~combined_data['TB Screening Type'].isna()) & 
 #   (~combined_data['TB status'].isna())
 #   ]

 # Filters for TX_TB_D (Screening type)
# TB_filter_date = '2024-10-01'
#TX_TB_D(Screening type) = combined_data[
 #   (combined_data['Date of TB Screening (yyyy-mm-dd)'] >= TB_filter_date) & 
#    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
#    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
 #   (~combined_data['TB Screening Type'].isna())
 #   ]

# Filters for TX_TB_D (Specimen sent)
# TB_filter_date = '2024-10-01'
#TX_TB_D(Specimen sent) = combined_data[
 #   (combined_data['Date of TB Screening (yyyy-mm-dd)'] >= TB_filter_date) & 
#    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
#    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
 #   (~combined_data['TB Screening Type'].isna()) &
 #   (combined_data['TB status'].isin('Presumptive TB', 'Confirmed TB', 'Currently on TB treatment')) &
 #   (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= TB_filter_date)
 #   ]

 # Filters for TX_TB_D (TB Test Type)
#TX_TB_D(TB Test Type) = combined_data[
 #   (combined_data['Date of TB Screening (yyyy-mm-dd)'] >= TB_filter_date) & 
#    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
#    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
 #   (~combined_data['TB Screening Type'].isna()) &
 #    (combined_data['TB status'].isin('Presumptive TB', 'Confirmed TB', 'Currently on TB treatment')) &
 #   (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= TB_filter_date) &
 #   (~combined_data['TB Diagnostic Test Type'].isna())
 #   ]


# Filters for TX_TB_D (Diagnosed with TB)
#TX_TB_D(Diagnosed with TB) = combined_data[
 #   (combined_data['Date of TB Screening (yyyy-mm-dd)'] >= TB_filter_date) & 
#    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
#    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
 #   (~combined_data['TB Screening Type'].isna()) &
 #    (combined_data['TB status'].isin('Presumptive TB', 'Confirmed TB', 'Currently on TB treatment')) &
 #   (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= TB_filter_date) &
 #   (~combined_data['TB Diagnostic Test Type'].isna()) &
 #    (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= TB_filter_date) & 
 #    (combined_data['TB Diagnostic Result'].str.contains(['Positive', 'MTB Detected'])) 
 #   ]


 # Filters for TX_TB_N (Started on TB Treatment)
#TX_TB_N = combined_data[
 #   (combined_data['Date of TB Screening (yyyy-mm-dd)'] >= TB_filter_date) & 
#    (combined_data['Client Verification Outcome'].fillna('').isin(['valid', ''])) &
#    (combined_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
 #   (~combined_data['TB Screening Type'].isna()) &
 #    (combined_data['TB status'].isin('Presumptive TB', 'Confirmed TB', 'Currently on TB treatment')) &
 #   (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= TB_filter_date) &
 #   (~combined_data['TB Diagnostic Test Type'].isna()) &
 #    (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= TB_filter_date) & 
 #    (combined_data['TB Diagnostic Result'].str.contains(['Positive', 'MTB Detected'])) &
 #     (combined_data['Date of Start of TB Treatment (yyyy-mm-dd)'] >= TB_filter_date)
 #   ]


# Pivoting the results
def pivot_data(data, value_name):
    return data.groupby('ProjectName').size().reset_index(name=value_name)#Datim Id

tx_curr_pivot = pivot_data(tx_curr, 'TX_CURR')
tx_pvls_d_pivot = pivot_data(tx_pvls_d, 'TX_PVLS_D')
tx_pvls_n_pivot = pivot_data(tx_pvls_n, 'TX_PVLS_N')
tx_new_pivot = pivot_data(tx_new, 'TX_NEW')
tx_rtt_pivot = pivot_data(tx_rtt, 'TX_RTT')
#tx_rtt_iit_lt_three_pivot = pivot_data(tx_rtt_iit_lt_three, 'TX_RTT_IIT<3')
#tx_rtt_btwn_three_to_five_pivot = pivot_data(tx_rtt_btwn_three_to_five, 'TX_RTT_IIT3-5')
#tx_rtt_iit_gt_six_pivot = pivot_data(tx_rtt_iit_gt_six, 'TX_RTT_IIT>=6')
tx_ml_Died_pivot = pivot_data(tx_ml_Died, 'TX_ML_Died')
tx_ml_Transfer_out_pivot = pivot_data(tx_ml_Transfer_out, 'TX_ML_Transfer_out')
tx_ml_IIT_pivot = pivot_data(tx_ml_IIT, 'TX_ML_IIT')
Invalid_pivot = pivot_data(Invalid, 'INVALID')
verification_ongoing_pivot = pivot_data(Verification_ongoing, 'VERIFICATION_ONGOING')
iit_lt_three_pivot = pivot_data(iit_lt_three, 'IIT<3')
iit_btwn_three_to_five_pivot = pivot_data(iit_btwn_three_to_five, 'IIT3-5')
iit_gt_six_pivot = pivot_data(iit_gt_six, 'IIT>=6')
tx_ml_Stopped_pivot = pivot_data(tx_ml_Stopped_TX, 'TX_ML_Stopped_TX')
tx_ml_Died_Unknown_pivot = pivot_data(tx_ml_Died_Unknown, 'TX_ML_Died_Unknown')
tx_ml_Died_cancer_pivot = pivot_data(tx_ml_Died_cancer, 'TX_ML_Died_cancer')
tx_ml_Died_Non_natural_pivot = pivot_data(tx_ml_Died_Non_natural, 'TX_ML_Died_Non_natural')
tx_ml_Died_Natural_pivot = pivot_data(tx_ml_Died_Natural, 'TX_ML_Died_Natural')
tx_ml_Died_Other_HIV_Disease_pivot = pivot_data(tx_ml_Died_Other_HIV_Disease, 'TX_ML_Died_Other_HIV_Disease')
tx_ml_Died_other_infectious_pivot = pivot_data(tx_ml_Died_other_infectious, 'TX_ML_Died_Other_infectious')
tx_ml_Died_Other_natural_pivot = pivot_data(tx_ml_Died_Other_natural, 'TX_ML_Died_Other_natural')
tx_ml_Died_TB_pivot = pivot_data(tx_ml_Died_TB, 'TX_ML_Died_TB')
tx_ml_Died_Other_cause_pivot = pivot_data(tx_ml_Died_Other_cause, 'TX_ML_Died_Other_cause')
#cxca_scrn_pivot = pivot_data(cxca_scrn, 'CXCA_SCRN')
#cxca_tx_pivot = pivot_data(cxca_tx, 'CXCA_TX')
#tx_tb_d_pivot = pivot_data(tx_tb_d, 'TX_TB_D')
#tx_tb_d(Screening type)_pivot = pivot_data(tx_tb_d(Screening type), 'TX_TB_D(Screening type)')
#tx_tb_d(Specimen sent)_pivot = pivot_data(tx_tb_d(Specimen sent), 'TX_TB_D(Specimen sent)')
#tx_tb_d(TB Test Type)_pivot = pivot_data(tx_tb_d(Screening type), 'TX_TB_D(TB Test Type)')
#tx_tb_d(Diagnosed with TB)_pivot = pivot_data(tx_tb_d(Diagnosed with TB), 'TX_TB_D(Diagnosed with TB)')
#tx_tb_n_pivot = pivot_data(tx_tb_n, 'TX_TB_N')

# Save to Excel
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    tx_curr_pivot.to_excel(writer, sheet_name='TX_CURR_Pivot', index=False)
    tx_pvls_d_pivot.to_excel(writer, sheet_name='TX_PVLS_D_Pivot', index=False)
    tx_pvls_n_pivot.to_excel(writer, sheet_name='TX_PVLS_N_Pivot', index=False)
    tx_new_pivot.to_excel(writer, sheet_name='TX_NEW_Pivot', index=False)
    tx_rtt_pivot.to_excel(writer, sheet_name='TX_RTT_Pivot', index=False)
    #tx_rtt_iit_lt_three_pivot.to_excel(writer, sheet_name='TX_RTT_IIT<3_Pivot', index=False)
    #tx_rtt_btwn_three_to_five_pivot.to_excel(writer, sheet_name='TX_RTT_IIT3-5_Pivot', index=False)
    #tx_rtt_iit_gt_six_pivot.to_excel(writer, sheet_name='TX_RTT_IIT>=6_Pivot', index=False)
    tx_ml_Died_pivot.to_excel(writer, sheet_name='TX_ML_Died_Pivot', index=False)
    tx_ml_Transfer_out_pivot.to_excel(writer, sheet_name='TX_Transferred_out_Pivot', index=False)
    tx_ml_IIT_pivot.to_excel(writer, sheet_name='TX_ML_IIT_Pivot', index=False)
    Invalid_pivot.to_excel(writer, sheet_name='Invalid_Pivot', index=False)
    verification_ongoing_pivot.to_excel(writer, sheet_name='Verification_ongoing_Pivot', index=False)
    iit_lt_three_pivot.to_excel(writer, sheet_name='IIT<3_Pivot', index=False)
    iit_btwn_three_to_five_pivot.to_excel(writer, sheet_name='IIT3-5_Pivot', index=False)
    iit_gt_six_pivot.to_excel(writer, sheet_name='IIT>=6_Pivot', index=False)
    tx_ml_Stopped_pivot.to_excel(writer, sheet_name='TX_ML_Stopped_Pivot', index=False)
    tx_ml_Died_Unknown_pivot.to_excel(writer, sheet_name='TX_ML_Died_Unknown_Pivot', index=False)
    tx_ml_Died_cancer_pivot.to_excel(writer, sheet_name='TX_ML_Died_cancer_Pivot', index=False)
    tx_ml_Died_Non_natural_pivot.to_excel(writer, sheet_name='TX_ML_Died_Non_natural_Pivot', index=False)
    tx_ml_Died_Natural_pivot.to_excel(writer, sheet_name='TX_ML_Died_Natural_Pivot', index=False)
    tx_ml_Died_Other_HIV_Disease_pivot.to_excel(writer, sheet_name='TX_ML_Died__Other_HIV_Disease_Pivot', index=False)
    tx_ml_Died_other_infectious_pivot.to_excel(writer, sheet_name='TX_ML_Died_Other_infectious_Pivot', index=False)
    tx_ml_Died_Other_natural_pivot.to_excel(writer, sheet_name='TX_ML_Died_Other_natural_Pivot', index=False)
    tx_ml_Died_TB_pivot.to_excel(writer, sheet_name='TX_ML_Died_TB_Pivot', index=False)
    tx_ml_Died_Other_cause_pivot.to_excel(writer, sheet_name='TX_ML_Died_Other_cause_Pivot', index=False)
    #cxca_scrn_pivot .to_excel(writer, sheet_name='CXCA_SCRN_Pivot', index=False)
    #cxca_tx_pivot.to_excel(writer, sheet_name='CXCA_TX_Pivot', index=False)
    #tx_tb_d_pivot.to_excel(writer, sheet_name='TX_TB_D_Pivot', index=False)
    #tx_tb_d(Screening type)_pivot.to_excel(writer, sheet_name='TX_TB_D(Screening type)_Pivot', index=False)
    #tx_tb_d(Specimen sent)_pivot.to_excel(writer, sheet_name='TX_TB_D(Specimen sent)_Pivot', index=False)
    #tx_tb_d(TB Test Type)_pivot.to_excel(writer, sheet_name='TX_TB_D(TB Test Type)_Pivot', index=False)
    #tx_tb_d(Diagnosed with TB)_pivot.to_excel(writer, sheet_name='TX_TB_D(Diagnosed with TB)_Pivot', index=False)
    #tx_tb_n_pivot.to_excel(writer, sheet_name='TX_TB_N_Pivot', index=False)

print(f"Analysis complete. Results saved to: {output_file_path}")
