import os
import pandas as pd
import re
from datetime import datetime

# Function to combine all Excel files from a folder
def combine_documents(folder_path):
    all_files = os.listdir(folder_path)
    combined_data = []
    
    for file in all_files:
        file_path = os.path.join(folder_path, file)  # Define file_path for all files
        if file.endswith('.xlsx'):  # Process Excel files
            data = pd.read_excel(file_path)
        else:
            data = pd.read_csv(file_path, encoding='latin1', engine='python', on_bad_lines='skip')
        
        combined_data.append(data)
    
    # Concatenate all dataframes into one
    return pd.concat(combined_data, ignore_index=True)


# Function to clean up facility names (remove initials)
def clean_facility_name(facility_name):
    # Remove initials like 'ad', 'bo', 'yo', etc. from the start of the facility name
    cleaned_name = re.sub(r'^[a-zA-Z]{2}\s?', '', facility_name)  # Assumes initials are 2 letters, followed by space or no space
    return cleaned_name


# Function to clean and replace blank cells in a date column with a default value
# Function to clean all date columns dynamically
def clean_all_date_columns(df, default_date='1900-01-01'):
    # Identify columns with 'date' in their name (case-insensitive)
    date_columns = [col for col in df.columns if 'date' in col.lower()]
    
    # Replace NaN and empty strings in date columns with the default date
    for column in date_columns:
        df[column] = pd.to_datetime(df[column], format='%d/%m/%Y', errors='coerce')  # Convert to datetime, handle errors
        df[column] = df[column].fillna(pd.to_datetime(default_date))  # Fill NaT with default date
    
    return df


def clean_all_text_columns(df, default_text='no_data'):
    # Identify columns with 'text' or 'str' in their name (case-insensitive)
    text_columns = [col for col in df.columns if not 'date' in col.lower()]
    #text_columns = [
    #col for col in df.columns
    #if 'date' not in col.lower() and col not in ['Last CD4 Count', 'Current Viral Load (c/ml)']]

    
    # Replace blanks or NaN values in each identified text column
    for column in text_columns:
        df[column] = df[column].fillna(default_text)  # Replace NaN with default text
        df[column] = df[column].replace(['', ' ', '\xa0'], default_text)  # Replace empty strings or non-breaking spaces
        #df['Last CD4 Count'] = df['Last CD4 Count'].astype(int)
    
    return df

# Convert columns to integers where possible
def convert_to_integer_columns(df, columns):
    for column in columns:
        df[column] = pd.to_numeric(df[column], errors='coerce')  # Convert to numeric, invalid entries become NaN
        df[column] = df[column].fillna(0).astype(int)  # Replace NaN with 0 and convert to integer
    return df

columns_to_clean = ['Last CD4 Count', 'Current Viral Load (c/ml)']



# Function to standardize one or multiple date columns
def standardize_date_columns(df, columns, output_format='%d/%m/%Y'):
    for column in columns:
        df[column] = pd.to_datetime(df[column], errors='coerce').dt.strftime(output_format)
    return df

#def alt_name(df):
 #   alternative_names = ['Patient ID', 'PatientID', 'Patient_Id', 'patientid', 'Patient_ID', 'patient_id', 'uniquepersonuuid']
  #  for name in alternative_names:
   #     if name in df.columns:
    #        df.rename(columns={name: 'Patient ID'}, inplace=True)
        


# Function to compare the columns between the two combined documents
def compare_documents(df1, df2):

    rename_dict = {
        'state': 'State',
        'facilityname': 'Facility Name',
        'uniquepersonuuid': 'Patient ID',
        'hospitalnumber': 'Hospital Number',
        'datim_id': 'DatimId',
        'Datim Id': 'DatimId',
        'uniqueid': 'Unique Id',
        'lga': 'L.G.A',
        'gender': 'Sex',
        'age': 'Age',
        'Date Birth (yyyy-mm-dd)': 'Date of Birth (yyyy-mm-dd)',
        #'Date of Birth (yyyy-mm-dd)': 'Date Of Birth (yyyy-mm-dd)',
        #'dateofbirth': 'Date Of Birth (yyyy-mm-dd)',
        #'dateofbirth': 'date of birth (yyyy-mm-dd)',
        'targetgroup': 'Target group',
        'ovcuniqueid': 'OVC Unique ID',
        'artstartdate': 'ART Start Date (yyyy-mm-dd)',
        'stateofresidence': 'State Of Residence',
        'householduniqueno': 'Household Unique No',
        'regimenatstart': 'Regimen at ART Start',
        'dateofregistration': 'Date of Registration',
        'dateofenrollment': 'Enrollment  Date (yyyy-mm-dd)',
        'enrollmentsetting': 'ART Enrollment Setting',
        'careentry': 'Care Entry Point',
        'regimenlineatstart': 'Regimen Line at ART Start',
        'ndrpatientidentifier': 'NDR Patient Identifier',
        'dateofviralloadsamplecollection': 'Date of Viral Load Sample Collection (yyyy-mm-dd)',
        'dateofcurrentviralloadsample': 'Date of Current ViralLoad Result Sample (yyyy-mm-dd)',
        'viralloadindication': 'Viral Load Indication',
        'currentviralload': 'Current Viral Load (c/ml)',
        'dateofcurrentviralload': 'Date of Current Viral Load (yyyy-mm-dd)',
        'currentdsdmodel': 'Current DSD model',
        'dateofstartofcurrentartregimen': 'Date of Start of Current ART Regimen',
        'lastpickupdate': 'Last Pickup Date (yyyy-mm-dd)',
        'currentartregimen': 'Current ART Regimen',
        'currentregimenline': 'Current Regimen Line',
        'monthsofarvrefill': 'Months of ARV Refill',
        'datebiometricsenrolled': 'Date Biometrics Enrolled (yyyy-mm-dd)',
        'numberoffingerscaptured': 'Number of Fingers Captured',
        'dateofcommencementofeac': 'Date of commencement of EAC (yyyy-mm-dd)',
        'numberofeacsessioncompleted': 'Number of EAC Sessions Completed',
        'dateoflasteacsessioncompleted': 'Date of last EAC Session Completed',
        'dateofextendeaccompletion': 'Date of Extended EAC Completion (yyyy-mm-dd)',
        'dateofrepeatviralloadresult': 'Date of Repeat Viral load result- POST EAC VL',
        'repeatviralloadresult': 'Repeat Viral load result (c/ml)- POST EAC',
        'dateofiptstart': 'Date of TPT Start (yyyy-mm-dd)',
        'iptcompletiondate': 'TPT Completion date (yyyy-mm-dd)',
        'iptcompletionstatus': 'TPT Completion status',
        'ipttype': 'TPT Type',
        'dateofcervicalcancerscreening': 'Date of Cervical Cancer Screening (yyyy-mm-dd)',
        'treatmentmethoddate': 'Date of Precancerous Lesions Treatment (yyyy-mm-dd)',
        'cervicalcancerscreeningtype': 'Cervical Cancer Screening Type',
        'cervicalcancerscreeningmethod': 'Cervical Cancer Screening Method',
        'cervicalcancertreatmentscreened': 'Precancerous Lesions Treatment Methods',
        'resultofcervicalcancerscreening': 'Result of Cervical Cancer Screening',
        'tbtreatmenttype': 'TB Type (new, relapsed etc)',
        'tbtreatmentstartdate': 'Date of Start of TB Treatment (yyyy-mm-dd)',
        'tbtreatmentoutcome': 'TB Treatment Outcome',
        'tbcompletiondate': 'Date of Completion of TB Treatment (yyyy-mm-dd)',
        'dateoftbsamplecollection': 'Date of TB Sample Collection (yyyy-mm-dd)',
        'dateoftbdiagnosticresultreceived': 'Date of TB Diagnostic Result Received (yyyy-mm-dd)',
        'tbdiagnosticresult': 'TB Diagnostic Result',
        'tbdiagnostictesttype': 'TB Diagnostic Test Type',
        'dateoftbscreened': 'Date of TB Screening (yyyy-mm-dd)',
        'tbstatus': 'TB status',
        'causeofdeath': 'Cause of Death',
        'vacauseofdeath': 'VA Cause of Death',
        'previousstatus': 'Previous ART Status',
        'previousstatusdate': 'Confirmed Date of Previous ART Status',
        'currentstatus': 'Current ART Status',
        'currentstatusdate': 'Date of Current ART Status',
        'vleligibilitystatus': 'Viral Load Eligibility Status',
        'dateofvleligibilitystatus': 'Date of Viral Load Eligibility Status',
        'lastcd4count': 'Last CD4 Count',
        'dateoflastcd4count': 'Date of Last CD4 Count',
        'casemanager': 'Case Manager',
        'clientverificationoutcome': 'Client Verification Outcome',
        'currentweight': 'Current Weight (kg)',
        'pregnancystatus': 'Pregnancy Status',
        'modeldevolvedto': 'Model devolved to',
        'dateofdevolvement': 'Date of Devolvement',
        #'Date of devolvement':'Date of Devolvement',
        #'Date of Devolvement':'Date of Devolvement',
        #'TB Type (new, relapsed etc)':'TB Treatment Type (new, relapsed etc)',
        'Model Devolved To' : 'Model devolved to',
        'Current DSD Model':'Current DSD model',
        'dateofcurrentdsd': 'Date of current DSD',
        'datereturntosite': 'Date of Return of DSD Client to Facility (yyyy-mm-dd)',
        'dateofrepeatviralloadeacsamplecollection': 'Date of Repeat Viral Load - Post EAC VL Sample collected (yyyy-mm-dd)',
        'tbscreeningtype': 'TB Screening Type',
        'currentclinicalstage': 'Clinical Staging at Last Visit',
        'datebiometricsrecaptured': 'Date Biometrics Recapture (yyyy-mm-dd)',
        'numberoffingersrecaptured': 'Number of Fingers Recaptured',
        'Date of Repeat Viral Load - Post EAC VL Sample collected (yyyy-' : 'Date of Repeat Viral Load - Post EAC VL Sample collected (yyyy-mm-dd)'
    }

    rename_dict2 = {'TB Type (new, relapsed etc)':'TB Treatment Type (new, relapsed etc)',
                    'Date of devolvement':'Date of Devolvement',
                    'Client Verification Status': 'Client Verification Outcome'}
    df1.rename(columns=rename_dict2, inplace=True)
    df2.rename(columns=rename_dict, inplace=True)

    df1 = convert_to_integer_columns(df1, columns_to_clean)
    df2 = convert_to_integer_columns(df2, columns_to_clean)

        # Check if 'Patient ID' exists in both dataframes
    if 'Patient ID' not in df1.columns:
        raise KeyError("'NDR Patient Identifier' not found in df1")
    if 'Patient ID' not in df2.columns:
        raise KeyError("'NDR Patient Identifier' not found in df2")


    # Duplicate the Patient ID column and rename them before merging
    df1['NDR Patient Identifier_RADET'] = df1['NDR Patient Identifier']
    df2['NDR Patient Identifier_Centralsync'] = df2['NDR Patient Identifier']
    
    # Clean the facility names in the RADET document before merging
    #df1['Facility'] = df1['Facility'].apply(clean_facility_name)

    # Clean all date columns in both dataframes
    df1 = clean_all_date_columns(df1)
    df2 = clean_all_date_columns(df2)

    # Clean all date columns in both dataframes
    df1 = clean_all_text_columns(df1)
    df2 = clean_all_text_columns(df2)

    # Clean patientid in both dataframes
    #df1 = alt_name(df1)
    #df2 = alt_name(df2)


    # Standardize date columns explicitly
    date_columns = [col for col in df1.columns if 'date' in col.lower()]
    date_columns = [col for col in df2.columns if 'date' in col.lower()]
    df1 = standardize_date_columns(df1, date_columns)
    df2 = standardize_date_columns(df2, date_columns)

    df1.columns = df1.columns.str.strip()#.str.lower()
    df2.columns = df2.columns.str.strip()#.str.lower()

    
    # Merge the two dataframes based on 'Patient ID' using an outer join
    merged_df = pd.merge(df1, df2, on='NDR Patient Identifier', how='outer', suffixes=('_RADET', '_Centralsync'))

    merged_df.fillna('N/A', inplace=True)

    merged_df = clean_all_date_columns(merged_df)
    
    # Handle State, Facility Name, and LGA to get the non-null values for each Patient ID
    #merged_df['State'] = merged_df['State_RADET'].combine_first(merged_df['State_Centralsync'])
    #merged_df['Facility'] = merged_df['Facility_RADET'].combine_first(merged_df['Facility_Centralsync'])
    #merged_df['LGA'] = merged_df['L.G.A_RADET'].combine_first(merged_df['L.G.A_Centralsync'])
    
    # Prepare the result DataFrame with Patient ID, State, Facility Name, and LGA
    result_df = merged_df[['NDR Patient Identifier']].drop_duplicates()

    result_df = result_df.copy()


    # Explicitly create the Patient_ID_RADET and Patient_ID_Centralsync columns, return N/A if no match
    result_df['NDR Patient Identifier_RADET'] = merged_df['NDR Patient Identifier_RADET'].fillna('N/A')
    result_df['NDR Patient Identifier_Centralsync'] = merged_df['NDR Patient Identifier_Centralsync'].fillna('N/A')

    result_df['NDR Patient Identifier_RADET'] = merged_df['NDR Patient Identifier_RADET'].replace(['', ' ', '\xa0'], 'N/A')
    result_df['NDR Patient Identifier_Centralsync'] = merged_df['NDR Patient Identifier_Centralsync'].replace(['', ' ', '\xa0'], 'N/A')

    
    # Add comparison results for Patient ID existence in RADET and Centralsync
    result_df['Match'] = result_df.apply(
        lambda row: 'Match' if row['NDR Patient Identifier_RADET'] != 'N/A' and row['NDR Patient Identifier_Centralsync'] != 'N/A' else 'No Match', axis=1
    )
    
    
    # Add comparison results for other columns (Age, Sex, ART Start Date, etc.)
    comparison_columns = ['State',	'L.G.A','Facility Name','DatimId','Hospital Number','Patient ID', 'Date of Birth (yyyy-mm-dd)',	'Unique Id',	'Household Unique No',	'OVC Unique ID','Sex','Target group'	,
                          'Current Weight (kg)',	'Pregnancy Status',	'Age',	'Care Entry Point',	'Date of Registration',	'Enrollment  Date (yyyy-mm-dd)',	'ART Start Date (yyyy-mm-dd)',	'Last Pickup Date (yyyy-mm-dd)',
                          'Months of ARV Refill',	'Regimen Line at ART Start'	,'Regimen at ART Start'	,'Date of Start of Current ART Regimen',	'Current Regimen Line'	,'Current ART Regimen',	'Clinical Staging at Last Visit',	'Date of Last CD4 Count',	
                          'Last CD4 Count'	,'Date of Viral Load Sample Collection (yyyy-mm-dd)'	,'Date of Current ViralLoad Result Sample (yyyy-mm-dd)',	'Current Viral Load (c/ml)',	'Date of Current Viral Load (yyyy-mm-dd)',	'Viral Load Indication',	
                          'Viral Load Eligibility Status','Date of Viral Load Eligibility Status',	'Current ART Status',	'Date of Current ART Status',	'Client Verification Outcome',	'Cause of Death',	'VA Cause of Death',	'Previous ART Status',	'Confirmed Date of Previous ART Status',
                          'ART Enrollment Setting','Date of TB Screening (yyyy-mm-dd)'	,'TB Screening Type'	,'TB status',	'Date of TB Sample Collection (yyyy-mm-dd)',	'TB Diagnostic Test Type',	'Date of TB Diagnostic Result Received (yyyy-mm-dd)',	'TB Diagnostic Result',
                          'Date of Start of TB Treatment (yyyy-mm-dd)',	'TB Treatment Type (new, relapsed etc)',#'TB Type (new, relapsed etc)'	,
                          'Date of Completion of TB Treatment (yyyy-mm-dd)',	'TB Treatment Outcome',	'Date of TPT Start (yyyy-mm-dd)',	'TPT Type',	'TPT Completion date (yyyy-mm-dd)',	'TPT Completion status',	
                          'Date of commencement of EAC (yyyy-mm-dd)',	'Number of EAC Sessions Completed',	'Date of last EAC Session Completed',	'Date of Extended EAC Completion (yyyy-mm-dd)',	'Date of Repeat Viral Load - Post EAC VL Sample collected (yyyy-mm-dd)',	'Repeat Viral load result (c/ml)- POST EAC',	
                          'Date of Repeat Viral load result- POST EAC VL',	'Date of Devolvement',	'Model devolved to',	'Date of current DSD',	'Current DSD model', 'Date of Return of DSD Client to Facility (yyyy-mm-dd)',
                         'Date of Cervical Cancer Screening (yyyy-mm-dd)',	'Cervical Cancer Screening Type', 'Cervical Cancer Screening Method',	'Result of Cervical Cancer Screening',	'Date of Precancerous Lesions Treatment (yyyy-mm-dd)',	'Precancerous Lesions Treatment Methods',	'Date Biometrics Enrolled (yyyy-mm-dd)',
                        'Number of Fingers Captured',	'Date Biometrics Recapture (yyyy-mm-dd)',	'Number of Fingers Recaptured',	'Case Manager'
    ]

    #comparison_columns = {
     #   col.replace('_RADET', '') for col in merged_df.columns if col.endswith('_RADET')
    #}.intersection(
     #   col.replace('_Centralsync', '') for col in merged_df.columns if col.endswith('_Centralsync')
    #)

    for col in comparison_columns:
        # Adding columns for the actual values from RADET and Centralsync
        result_df[f'{col}_RADET'] = merged_df[f'{col}_RADET']
        result_df[f'{col}_Centralsync'] = merged_df[f'{col}_Centralsync']
        
        # Adding match/no match result
        result_df[f'{col}_match'] = merged_df.apply(
    lambda row: 'N/A' if row[f'{col}_RADET'] == 'N/A' or row[f'{col}_Centralsync'] == 'N/A' or 
                          row['NDR Patient Identifier_RADET'] == 'N/A' or row['NDR Patient Identifier_Centralsync'] == 'N/A'
    else ('Match' if row[f'{col}_RADET'] == row[f'{col}_Centralsync'] else 'No Match'), 
    axis=1
)

    return result_df


# Paths to the folders
folder_radet = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis/RADETt'
folder_centralsync = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis/Centralsyncc'

# Combine documents
df_radet = combine_documents(folder_radet)
df_centralsync = combine_documents(folder_centralsync)

# Compare the two combined documents
client_level_analysis = compare_documents(df_radet, df_centralsync)

# Save the result to an Excel file
output_path = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis/client_level_analysis_ACE2_Radet.xlsx'
client_level_analysis.to_excel(output_path, index=False)

print(f'Comparison result saved to {output_path}') 


