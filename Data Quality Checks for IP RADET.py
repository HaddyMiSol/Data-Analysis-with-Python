import pandas as pd
import os
from datetime import datetime
import re

# Path to the directory containing the CSV files
folder_path = 'C:/Users/DELL/Documents/DataFi/Data Review Meeting/CS/CS_RADET_Files'

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
#combined_data['ProjectName'] = combined_data['IP']

# Convert date columns to datetime format
for col in ['Date of Birth (yyyy-mm-dd)', 'Last Pickup Date (yyyy-mm-dd)', 'ART Start Date (yyyy-mm-dd)', 'Date of Current ART Status', 'Date of Current Viral Load (yyyy-mm-dd)',
             'Confirmed Date of Previous ART Status', 'Date of Current ViralLoad Result Sample (yyyy-mm-dd)', 'Date of Start of Current ART Regimen',
               'Date of TB Screening (yyyy-mm-dd)', 'Date of TB Sample Collection (yyyy-mm-dd)', 'Enrollment  Date (yyyy-mm-dd)','Date of TPT Start (yyyy-mm-dd)', 
               'Date of Last CD4 Count', 'Date of Registration', 'Enrollment  Date (yyyy-mm-dd)', 'TPT Completion date (yyyy-mm-dd)']:
    
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

# Convert Last CD4 Count and Current Viral Load to integers
#combined_data['Last CD4 Count'] = pd.to_numeric(combined_data['Last CD4 Count'], errors='coerce')
#combined_data['Current Viral Load (c/ml)'] = pd.to_numeric(combined_data['Current Viral Load (c/ml)'], errors='coerce')

# Define filter date
filter_date = datetime(2024, 10, 1)

quality_issue_counts = []
line_lists = {}  

# Define a function to check if a value is invalid
def is_invalid_cd4(value):
    # Check for blank (NaN)
    if pd.isna(value):
        return False  # Blank is valid
    
    # Convert the value to a string for pattern matching
    value_str = str(value).strip()
    
    # Define valid patterns (expandable for additional valid values)
    valid_patterns = [
        r"^\d+$",  # Integer (e.g., 200)
        r"^\d+\.\d+$",  # Float (e.g., 200.5)
        r"^[><]=?\d+$",  # Comparisons with numbers (e.g., >=200, <=200)
    ]
    
    # Check if the value matches any valid pattern
    for pattern in valid_patterns:
        if re.fullmatch(pattern, value_str):
            return False  # Value is valid
    
    # Define invalid patterns
    invalid_patterns = [
        r"[a-zA-Z]",  # Contains letters
        r"[+@`\]\^#%_,/\\|\s]",  # Contains specific invalid characters or spaces
        r"cp/ml|cell/pl|cell/mms|cp / ml",  # Contains invalid units
        r"[><]=?\s*\d",  # Improper comparison with numbers
        r"\d\s+\d",  # Numbers with spaces in between
        r"<|>|=",  # Standalone comparison symbols
        r"^\d{1,4}-\d{1,2}-\d{1,4}$",  # Matches date format (e.g., 2024-12-15)
        r"[.]{2,}",  # Multiple periods in a row
        r"^\d*\s+[a-zA-Z]+$",  # Number followed by a letter or word
        r"\s{2,}",  # Multiple spaces
        r"^\d*[+-/*|]$",  # Number followed by invalid operator
        r"^\d*[><=]+$",  # Number followed by comparison alone
        r"^\d+\.$",  # Number ending with a period
        r"[><]=?\d+\s*\d"  # Comparison sign followed by two numbers (e.g., <20, >30)
    ]
    
    # Check if the value matches any invalid patterns
    for pattern in invalid_patterns:
        if re.search(pattern, value_str):
            return True  # Value is invalid
    
    # Default case: If it doesn't match any valid pattern and doesn't match specific invalid ones
    return True


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
        r"^[><]=?\d+$",  # Comparisons with numbers (e.g., >=200, <=200)
    ]
    
    # Check if the value matches any valid pattern
    for pattern in valid_patterns:
        if re.fullmatch(pattern, value_str):
            return True  # Value is valid
        
    return True


# Define a function to check if a value is invalid
def is_invalid_viralload(value):

    #combined_data['Current Viral Load (c/ml)'] = pd.to_numeric(combined_data['Current Viral Load (c/ml)'], errors='coerce')

    # Check for blank (NaN)
    if pd.isna(value):
        return False  # Blank is valid
    
    # Convert the value to a string for pattern matching
    value_str = str(value).strip()
    
    # Define valid patterns (expandable for additional valid values)
    valid_patterns = [
        r"^\d+$",  # Integer (e.g., 200)
        r"^\d+\.\d+$",  # Float (e.g., 200.5)
        #r"^[><]=?\d+$",  # Comparisons with numbers (e.g., >=200, <=200)
    ]
    
    # Check if the value matches any valid pattern
    for pattern in valid_patterns:
        if re.fullmatch(pattern, value_str):
            return False  # Value is valid
    
    # Define invalid patterns
    invalid_patterns = [
        r"[a-zA-Z]",  # Contains letters
        r"[+@`\]\^#%_,/\\|\s]",  # Contains specific invalid characters or spaces
        r"cp/ml|cell/pl|cell/mms|cp / ml",  # Contains invalid units
        r"[><]=?\s*\d",  # Improper comparison with numbers
        r"\d\s+\d",  # Numbers with spaces in between
        r"<|>|=",  # Standalone comparison symbols
        r"^\d{1,4}-\d{1,2}-\d{1,4}$",  # Matches date format (e.g., 2024-12-15)
        r"[.]{2,}",  # Multiple periods in a row
        r"^\d*\s+[a-zA-Z]+$",  # Number followed by a letter or word
        r"\s{2,}",  # Multiple spaces
        r"^\d*[+-/*|]$",  # Number followed by invalid operator
        r"^\d*[><=]+$",  # Number followed by comparison alone
        r"^\d+\.$",  # Number ending with a period
        r"[><]=?\d+\s*\d",  # Comparison sign followed by two numbers (e.g., <20, >30)
        r"^[><]=?\d+$"  # Comparisons with numbers (e.g., >=200, <=200)
    ]
    
    # Check if the value matches any invalid patterns
    for pattern in invalid_patterns:
        if re.search(pattern, value_str):
            return True  # Value is invalid
    
    # Default case: If it doesn't match any valid pattern and doesn't match specific invalid ones
    return True

# Define output directory for projects
output_base_dir = 'C:/Users/DELL/Documents/DataFi/Data Review Meeting/CS/Project_Export_Quality_Check'
os.makedirs(output_base_dir, exist_ok=True)

# Iterate through projects
for project_name, project_data in combined_data.groupby('ProjectName'):
    # Directory for the current project
    project_dir = os.path.join(output_base_dir, project_name)
    os.makedirs(project_dir, exist_ok=True)

    # Initialize dictionary for project issues
    project_issues = {}

    # Line lists for quality issues
    line_lists = {}

    # Condition 1: Blank Date of Birth
    blank_dob = project_data[project_data['Date of Birth (yyyy-mm-dd)'].isna()]
    project_issues['Blank Date of Birth'] = blank_dob.shape[0]
    line_lists[f"{project_name}_BlankDOB"] = blank_dob

    # Condition 1: Blank Age
    blank_age = project_data[project_data['Age'].isna()]
    project_issues['Blank Age'] = blank_dob.shape[0]
    line_lists[f"{project_name}_BlankAge"] = blank_age

    # Condition 2: Blank Last Pickup Date
    blank_last_pickup = project_data[project_data['Last Pickup Date (yyyy-mm-dd)'].isna()]
    project_issues['Blank Last Pickup Date'] = blank_last_pickup.shape[0]
    line_lists[f"{project_name}_BlankLastPickup"] = blank_last_pickup

    # Condition : LPUD < ART start Date
    LPUD_lt_ARTStart = project_data[
        (project_data['Last Pickup Date (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['LPUD_lt_ARTStart'] = LPUD_lt_ARTStart.shape[0]
    line_lists[f"{project_name}_LPUD_lt_ARTStart"] = LPUD_lt_ARTStart

    # Condition 3: Blank ART Start Date
    blank_ARTStart = project_data[project_data['ART Start Date (yyyy-mm-dd)'].isna()]
    project_issues['Blank ART Start Date'] = blank_ARTStart.shape[0]
    line_lists[f"{project_name}_BlankARTStart"] = blank_ARTStart

    # Condition : Gender =' Male' and Pregnancy status = ' Pregnant' or 'Breastfeeding'
    Male_with_pregnancystatus = project_data[
        (project_data['Sex'].isin(['Male'])) & (project_data['Pregnancy Status'].isin(['Pregnant', 'Breastfeeding']))]
    project_issues['Male_with_pregnancystatus'] = Male_with_pregnancystatus.shape[0]
    line_lists[f"{project_name}_Male_with_pregnancystatus"] = Male_with_pregnancystatus

    # Condition : Date_of start of Current Regimen < ART start Date
    Date_start_regimen_lt_ARTStart = project_data[
        (project_data['Date of Start of Current ART Regimen'] < project_data['ART Start Date (yyyy-mm-dd)']) &
          (project_data['Client Verification Outcome'].isin(['valid', ''])) &
          (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['Date_start_regimen_lt_ARTStart'] = Date_start_regimen_lt_ARTStart.shape[0]
    line_lists[f"{project_name}_Date_start_regimen_lt_ARTStart"] = Date_start_regimen_lt_ARTStart

    # Condition : Date_of Enrollment < Date of registration
    Date_enroll_lt_Date_Registratn = project_data[
        (project_data['Enrollment  Date (yyyy-mm-dd)'] < project_data['Date of Registration']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['Date_enroll_lt_Date_Registratn'] = Date_enroll_lt_Date_Registratn.shape[0]
    line_lists[f"{project_name}_Date_enroll_lt_Date_Registratn"] = Date_enroll_lt_Date_Registratn

    # Condition 4: Invalid CD4 values
    cd4_invalid = project_data[
        (project_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Date of Current ART Status'] >= filter_date) &
        project_data['Last CD4 Count'].apply(is_invalid_cd4)]

    # Count invalid rows
    project_issues['Invalid CD4'] = cd4_invalid.shape[0]

    # Save invalid rows for review
    line_lists[f"{project_name}_InvalidCD4"] = cd4_invalid


    # Condition : CD4 Count for <5yrs
    CD4_Count_for_lt_five = project_data[
        (project_data['Age'] < 5) & (project_data['Last CD4 Count'].apply(is_valid_cd4)) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart'])) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['CD4_Count_for_lt_five'] = CD4_Count_for_lt_five.shape[0]
    line_lists[f"{project_name}_CD4_Count_for_lt_five"] = CD4_Count_for_lt_five


    # Condition 5: Current ART Status "Died" but blank Cause of Death
    died_blank_cause = project_data[
        (project_data['Current ART Status'] == 'Died') & project_data['Cause of Death'].isna() &
        (project_data['Previous ART Status'].isin(['Active', 'Active Restart'])) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Date of Current ART Status'] >=filter_date)]
    project_issues['Died with Blank Cause of Death'] = died_blank_cause.shape[0]
    line_lists[f"{project_name}_DiedBlankCause"] = died_blank_cause

    # Condition 6: Blank Care Entry Point for ART Start Date >= October 2024
    blank_care_entry = project_data[(project_data['ART Start Date (yyyy-mm-dd)'] >= filter_date) & (project_data['Care Entry Point'].isna())]
    project_issues['Blank Care Entry Point'] = blank_care_entry.shape[0]
    line_lists[f"{project_name}_BlankCareEntry"] = blank_care_entry

    # Condition 7: Non-integer Current Viral Load
    viralload_invalid = project_data[
        (project_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
        project_data['Current Viral Load (c/ml)'].apply(is_invalid_viralload)]

    # Count invalid rows
    project_issues['Invalid viralload'] = viralload_invalid.shape[0]

    # Save invalid rows for review
    line_lists[f"{project_name}_InvalidViralLoad"] = viralload_invalid

    # Condition 8: ART Start Date < DOB
    ARTStart_Date_lt_DOB = project_data[
        (project_data['ART Start Date (yyyy-mm-dd)'] < project_data['Date of Birth (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['ARTStart_Date_lt_DOB'] = ARTStart_Date_lt_DOB.shape[0]
    line_lists[f"{project_name}_ARTStart_Date_lt_DOB"] = ARTStart_Date_lt_DOB

    # Condition 9: LPUD < DOB
    LPUD_lt_DOB = project_data[
        (project_data['Last Pickup Date (yyyy-mm-dd)'] < project_data['Date of Birth (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['LPUD_lt_DOB'] = LPUD_lt_DOB.shape[0]
    line_lists[f"{project_name}_LPUD_lt_DOB"] = LPUD_lt_DOB


    # Condition 9: LPUD < ART Start Date
    LPUD_lt_ARTStart = project_data[
        (project_data['Last Pickup Date (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['LPUD_lt_ARTStart'] = LPUD_lt_ARTStart.shape[0]
    line_lists[f"{project_name}_LPUD_lt_ARTStart"] = LPUD_lt_ARTStart

    # Condition 9: Current ART Status Date < ART Start Date
    ART_Status_Date_lt_ARTStart = project_data[
        (project_data['Date of Current ART Status'] < project_data['ART Start Date (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['ART_Status_Date_lt_ARTStart'] = ART_Status_Date_lt_ARTStart.shape[0]
    line_lists[f"{project_name}_ART_Status_Date_lt_ARTStart"] = ART_Status_Date_lt_ARTStart

    # Condition 9: Date of Last CD4 < ART Start Date
    lastCD4CountDate_lt_ARTStart = project_data[
        (project_data['Date of Last CD4 Count'] < project_data['ART Start Date (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['lastCD4CountDate_lt_ARTStart'] = lastCD4CountDate_lt_ARTStart.shape[0]
    line_lists[f"{project_name}_lastCD4CountDate_lt_ARTStart"] = lastCD4CountDate_lt_ARTStart

    # Condition 9: Current ART Status Date < Last Pickup Date
    ART_Status_Date_lt_LPUD = project_data[
        (project_data['Date of Current ART Status'] < project_data['Last Pickup Date (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['ART_Status_Date_lt_LPUD'] = ART_Status_Date_lt_LPUD.shape[0]
    line_lists[f"{project_name}_ART_Status_Date_lt_LPUD"] = ART_Status_Date_lt_LPUD

    # Condition 10: Date of Previous ART_Status > Date of Current ART Status
    PreviousARTStatusDate_gt_CurrentARTStatusDate = project_data[
        (project_data['Confirmed Date of Previous ART Status'] > project_data['Date of Current ART Status']) & (project_data['Date of Current ART Status'] >= filter_date) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['PreviousARTStatusDate_gt_CurrentARTStatusDate'] = PreviousARTStatusDate_gt_CurrentARTStatusDate.shape[0]
    line_lists[f"{project_name}_PreviousARTStatusDate_gt_CurrentARTStatusDate"] = PreviousARTStatusDate_gt_CurrentARTStatusDate


    # Condition 11: Date of Current Viral Load (yyyy-mm-dd) without Current Viral Load (c/ml)
    CurrentVLDate_without_CurrentVL = project_data[
        (project_data['Date of Current Viral Load (yyyy-mm-dd)'] >= filter_date) & (project_data['Current Viral Load (c/ml)'].isna()) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['CurrentVLDate_without_CurrentVL'] = CurrentVLDate_without_CurrentVL.shape[0]
    line_lists[f"{project_name}_CurrentVLDate_without_CurrentVL"] = CurrentVLDate_without_CurrentVL


    # Condition : Date of Viral Load Sample Collection (yyyy-mm-dd) < ART start Date
    VLsamplecollectionDate_lt_ARTStart = project_data[
        (project_data['Date of Viral Load Sample Collection (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['VLsamplecollectionDate_lt_ARTStart'] = VLsamplecollectionDate_lt_ARTStart.shape[0]
    line_lists[f"{project_name}_VLsamplecollectionDate_lt_ARTStart"] = VLsamplecollectionDate_lt_ARTStart

    # Condition : Date of Current ViralLoad Result Sample (yyyy-mm-dd) < ART start Date
    CurrentVLResultSampleDate_lt_ARTStart = project_data[
        (project_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['CurrentVLResultSampleDate_lt_ARTStart'] = CurrentVLResultSampleDate_lt_ARTStart.shape[0]
    line_lists[f"{project_name}_CurrentVLResultSampleDate_lt_ARTStart"] = CurrentVLResultSampleDate_lt_ARTStart


    # Condition : Date of Current Viral Load (yyyy-mm-dd) < ART start Date
    CurrentVLDate_lt_ARTStart = project_data[
        (project_data['Date of Current Viral Load (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['CurrentVLDate_lt_ARTStart'] = CurrentVLDate_lt_ARTStart.shape[0]
    line_lists[f"{project_name}_CurrentVLDate_lt_ARTStart"] = CurrentVLDate_lt_ARTStart


    # Condition : Wrong ART Enrollment setting 
    WrongARTEnrollmentSetting = project_data[
    (~project_data['ART Enrollment Setting'].isin(['Facility', 'Community'])) & (~project_data['ART Enrollment Setting'].isna()) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['WrongARTEnrollmentSetting'] = WrongARTEnrollmentSetting.shape[0]
    line_lists[f"{project_name}_WrongARTEnrollmentSetting"] = WrongARTEnrollmentSetting


    # Condition 12: Active Restart without previous current ART Status
    TX_RTT_without_PreviousARTStatus = project_data[
        (project_data['Current ART Status'].isin(['Active Restart'])) &
        ((~project_data['Previous ART Status'].isin(['Active', 'Active Restart'])) | project_data['Previous ART Status'].isna())  &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Date of Current ART Status'] >= filter_date)&
        (project_data['Confirmed Date of Previous ART Status'] < filter_date)]
    project_issues['TX_RTT_without_PreviousARTStatus'] = TX_RTT_without_PreviousARTStatus.shape[0]
    line_lists[f"{project_name}_TX_RTT_without_PreviousARTStatus"] = TX_RTT_without_PreviousARTStatus

    

    # Append the project issues to the list
    quality_issue_counts.append(project_issues)

# Convert the list of dictionaries to a DataFrame
quality_issues_df = pd.DataFrame(quality_issue_counts)


# Save all line lists to an Excel file with separate worksheets
output_file_path = os.path.join(project_dir, f"{project_name}_Quality_Check.xlsx")
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    
    for issue_name, data in line_lists.items():
        data.to_excel(writer, index=False, sheet_name=issue_name[:100])
        quality_issues_df.to_excel(writer, sheet_name='Quality Issues', index=False) 

