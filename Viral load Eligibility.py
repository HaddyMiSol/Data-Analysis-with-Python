import pandas as pd
from datetime import datetime, timedelta
import chardet
import logging

# Configure logging
logging.basicConfig(filename='data_processing.log', level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

# Paths (Use raw strings to handle special characters)
input_file_path = r"/Users/DELL/Documents/ACE1/ACE1_RADET FY25Q1.csv"
output_file_path = r"/Users/DELL/Documents/ACE1//eligible_for_vl_sample.csv"

# Step 1: Detect the encoding
try:
    with open(input_file_path, 'rb') as f:
        raw_data = f.read()  # Read the entire file for better detection
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    confidence = result['confidence']
    print(f"Detected encoding: {encoding} with confidence {confidence}")
except Exception as e:
    print(f"Error detecting encoding: {e}")
    encoding = 'utf-8'  # Fallback encoding

# Step 2: Load the dataset with the detected encoding
try:
    df = pd.read_csv(
        input_file_path,
        encoding=encoding,
        sep=',',  # Specify the delimiter; change if different
        engine='python',  # Use the Python engine for better handling of complex files
        on_bad_lines='warn',  # Handle bad lines gracefully
        dtype=str  # Read all columns as strings to prevent dtype issues
    )
    print(f"Successfully read the file with {len(df)} rows.")
except UnicodeDecodeError as e:
    print(f"UnicodeDecodeError encountered: {e}")
    print("Attempting to read with 'cp1252' encoding and skipping bad lines.")
    try:
        df = pd.read_csv(
            input_file_path,
            encoding='cp1252',
            sep=',',
            engine='python',
            on_bad_lines='skip',  # Silently skip bad lines
            dtype=str
        )
        print(f"Successfully read the file with 'cp1252' encoding with {len(df)} rows.")
    except Exception as e2:
        print(f"Failed to read the file with 'cp1252' encoding: {e2}")
        df = pd.DataFrame()  # Create an empty DataFrame as a fallback

# Optional: Verify the DataFrame
print(df.head())
print(df.info())
logging.info("DataFrame head and info displayed.")

# Step 3: Save the DataFrame to a new CSV (if needed)
try:
    df.to_csv(output_file_path, index=False)
    print(f"Data successfully saved to {output_file_path}")
    logging.info(f"Data successfully saved to {output_file_path}")
except Exception as e:
    print(f"Error saving the file: {e}")
    logging.error(f"Error saving the file: {e}")

# Data Cleaning
columns_to_clean = ['Current ART Status']
for col in columns_to_clean:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.title()
        logging.info(f"Cleaned column: {col}")


# Ensure date columns are in datetime format
date_columns = [
    'Date of Current Viral Load (yyyy-mm-dd)',
    'Date of Viral Load Sample Collection (yyyy-mm-dd)',
    'ART Start Date (yyyy-mm-dd)', 
    'Confirmed Date of Previous ART Status',
    'Date of Current ART Status',
    'Date of Current ViralLoad Result Sample (yyyy-mm-dd)'
]

for col in date_columns:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')

# Function to calculate quarter start and end dates based on fiscal quarters
def get_quarter_dates(reference_date):
    year = reference_date.year
    month = reference_date.month

    if 10 <= month <= 12:  # Q1
        return (datetime(year, 10, 1), datetime(year, 12, 31))
    elif 1 <= month <= 3:  # Q2
        return (datetime(year, 1, 1), datetime(year, 3, 31))
    elif 4 <= month <= 6:  # Q3
        return (datetime(year, 4, 1), datetime(year, 6, 30))
    elif 7 <= month <= 9:  # Q4
        return (datetime(year, 7, 1), datetime(year, 9, 30))

# Current date and quarter calculations
current_date = datetime.now()
current_q_start, current_q_end = get_quarter_dates(current_date)

# Calculate dates for previous and two quarters ago
previous_q_start, previous_q_end = get_quarter_dates(current_q_start - timedelta(days=1))
two_q_ago_start, two_q_ago_end = get_quarter_dates(previous_q_start - timedelta(days=1))

# Calculate previous year dates for current quarter range
previous_year_start = current_q_start.replace(year=current_q_start.year - 1)
previous_year_end = current_q_end.replace(year=current_q_end.year - 1)


# Function to get the quarter two quarters ago before the last completed quarter
def get_two_quarters_ago(previous_year, previous_quarter):
    # Calculate the quarter and year two quarters back
    if previous_quarter > 2:
        two_quarters_ago = previous_quarter - 2
        two_quarters_ago_year = previous_year
    else:
        # Adjust year if we need to go back to the previous year
        two_quarters_ago = previous_quarter + 2  # e.g., if Q1 or Q2, we go to Q3 or Q4 of last year
        two_quarters_ago_year = previous_year - 1
    
    return get_quarter_dates(two_quarters_ago_year, two_quarters_ago)



# Count active records after cleaning
total_active = df[df['Current ART Status'] == 'Active'].shape[0]
print(f"Total active records after cleaning: {total_active}")

# ============================
# Define Criteria for Cohorts
# ============================

# Criteria 1: Active patients (>19yrs) with VL sample collected 1 year ago but not within 6 months before the current quarter
df['Age'] = pd.to_numeric(df['Age'], errors='coerce')
df['Current Viral Load (c/ml)'] = pd.to_numeric(df['Current Viral Load (c/ml)'], errors='coerce')
df['Current ART Status'] = df['Current ART Status'].str.strip()#.str.lower()

criteria1 = df[
    (df['Current ART Status'].isin(['Active'])) &
    (df['Current Viral Load (c/ml)'] < 1000) &
    (df['Age'] > 19) &
    (df['Date of Current Viral Load (yyyy-mm-dd)'].between(previous_year_start, previous_year_end)) &
    (
        ((df['Date of Viral Load Sample Collection (yyyy-mm-dd)'] < (current_q_start - timedelta(days=92)))) |
         #(df['Date of Viral Load Sample Collection (yyyy-mm-dd)'] >= previous_year_start)) |
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'].isna())
    )
]
# Add the cohort label
criteria1 = criteria1.copy()
criteria1['Cohort'] = 'One_year_Anniversary'

# Output for verification
print(f"Total records meeting criteria1: {len(criteria1)}")


# Criteria 2: Patients started on ART 6 months before the current quarter, but VL sample not collected 3 months before the current quarter
criteria2 = df[
    (df['Current ART Status'].isin(['Active'])) & 
    (df['Care Entry Point'] != 'Transfer-in') &  
    (df['ART Start Date (yyyy-mm-dd)'] >= two_q_ago_start) &
    (df['ART Start Date (yyyy-mm-dd)'] <= two_q_ago_end) &
    (
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'].isna()) |
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'] < (two_q_ago_end - timedelta(days=91)))
    )
]
# Add cohort label
criteria2 = criteria2.copy()
criteria2['Cohort'] = 'TX_NEW (6months_ago)'

# Output for verification
print(f"Total criteria2 records after filtering: {len(criteria2)}")
#print(criteria2.head())


# Criteria 3: Active patients <=19 years with last VL sample collected at least 6 months ago and not within 3 months before the current quarter
criteria3 = df[
    (df['Current ART Status'].isin(['Active'])) &
    (df['Age'].astype(float) <= 19) &  
    (df['Date of Current Viral Load (yyyy-mm-dd)'] >= two_q_ago_start) &
     (df['Date of Current Viral Load (yyyy-mm-dd)'] <= two_q_ago_end) &
    (
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'] < (two_q_ago_end - timedelta(days=91))) |
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'].isna())
    )
]
# Add cohort label
criteria3 = criteria3.copy()
criteria3['Cohort'] = 'Pediatrics_<15 (6mths_ago)'

# Output total records for this criteria
print(f"Total records meeting criteria3: {len(criteria3)}")


#Criteria 4: All clients who were eligible in the previous quarter but VL was not collected
# Calculate the equivalent dates for the same quarter in the previous year
same_quarter_prev_year_start = previous_q_start.replace(year=previous_q_start.year - 1)
same_quarter_prev_year_end = previous_q_end.replace(year=previous_q_end.year - 1)
# Pending Criteria 1: Active patients (>19yrs) with VL sample collected 1 year ago but not within 6 months before the previous quarter 
Pending_criteria1 = df[
    (df['Current ART Status'].isin(['Active'])) &
    (df['Current Viral Load (c/ml)'] < 1000) &
    (df['Age'] > 19) &
    (df['Date of Current Viral Load (yyyy-mm-dd)'] >= same_quarter_prev_year_start) &
    (df['Date of Current Viral Load (yyyy-mm-dd)'] <= same_quarter_prev_year_end) &
    (
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'] < (previous_q_start - timedelta(days=92))) |
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'].isna())
    )
]
Pending_criteria1 = Pending_criteria1.copy()
Pending_criteria1['Cohort'] = 'Pending_previous_q (A_year_Anniversary)'
print(f"Total Pending_criteria1 records after cleaning: {len(Pending_criteria1)}")


# Pending Criteria 2: Patients started on ART two quarters ago from previous quarter, but VL sample not collected within 3 months from ART Start Date
Pending_criteria2 = df[
    (df['Current ART Status'].isin(['Active'])) &
    (df['Care Entry Point'] != 'Transfer-in') &
    (df['ART Start Date (yyyy-mm-dd)'] >= (two_q_ago_start - timedelta(days=92))) &
    (df['ART Start Date (yyyy-mm-dd)'] <= (two_q_ago_end - timedelta(days=92))) &
    (
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'].isna()) | 
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'] < (two_q_ago_end - timedelta(days=182)))
    )
]

Pending_criteria2 = Pending_criteria2.copy()
Pending_criteria2['Cohort'] = 'Pending_previous_q (TX_NEW-6months_ago)'
print(f"Total Pending_criteria2 records after cleaning: {len(Pending_criteria2)}")


# Pending Criteria 3: Active patients <=19 years with last VL sample collected 6 months ago and not within 3 months before the previous quarter
Pending_criteria3 = df[
    (df['Current ART Status'].isin(['Active'])) &
    (df['Age'] <= 19) &
    (df['Date of Current Viral Load (yyyy-mm-dd)'] >= (two_q_ago_start - timedelta(days=92))) &
    (df['Date of Current Viral Load (yyyy-mm-dd)'] <=  (two_q_ago_end - timedelta(days=92))) &
    (
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'] < (two_q_ago_start - timedelta(days=92))) |
        (df['Date of Viral Load Sample Collection (yyyy-mm-dd)'].isna())
    )
]

Pending_criteria3 = Pending_criteria3.copy()
Pending_criteria3['Cohort'] = 'Pending_previous_q (Pediatrics_<15-6mths_ago)'
print(f"Total Pending_criteria3 records after cleaning: {len(Pending_criteria3)}")


# Last Criteria: Active restart patients with specific conditions (Should not be added to eligibility list but to be shared with Professional case managers to decide if VL should be collected)
Last_Criteria = df[
    (df['Current ART Status'] == 'Active Restart') &
    (df['Previous ART Status'] != 'Active') &
    (df['Date of Current ART Status'] >= previous_q_start) &
    (df['Date of Current ART Status'] <= previous_q_end) 
    
]
Last_Criteria = Last_Criteria.copy()
Last_Criteria['Cohort'] = 'Active_Restart_previous_quarter'
print(f"Total Last_Criteria records after cleaning: {len(Last_Criteria)}")

# ============================
# Combine All Criteria Results
# ============================
eligible_for_vl = pd.concat([
    criteria1,
    criteria2, 
    criteria3,
    Pending_criteria1, 
    Pending_criteria2, 
    Pending_criteria3,
    Last_Criteria
]).drop_duplicates()


# ============================
# Output the Final List
# ============================
eligible_for_vl.to_csv(output_file_path, index=False, encoding='utf-8')

print(f"CSV file saved at: {output_file_path}")
print(f"Total eligible patients: {eligible_for_vl.shape[0]}")

