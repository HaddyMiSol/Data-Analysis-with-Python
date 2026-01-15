import os
import pandas as pd
import re
from datetime import datetime


# Paths to the folders
folder_client = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis/RADET'
folder_centralsync = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis/Centralsync'

output_path = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis/ACE-2_FY25Q4_RADET_client_level_analysis.xlsx'



# Function to combine all Excel files from a folder
def combine_documents(folder_path):
    all_files = os.listdir(folder_path)
    combined_data = []
    
    for file in all_files:
        file_path = os.path.join(folder_path, file)  # Define file_path for all files
        if file.endswith('.xlsx'):  # Process Excel files
            data = pd.read_excel(file_path, engine= "openpyxl")
        elif file.endswith('xls'):
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


columns_to_clean = ['Last CD4 Count', 'Current Viral Load (c/ml)', 'Repeat Viral load result (c/ml)- POST EAC']



# Function to standardize one or multiple date columns
def standardize_date_columns(df, columns, output_format='%d/%m/%Y'):
    for column in columns:
        df[column] = pd.to_datetime(df[column], errors='coerce').dt.strftime(output_format)
    return df



def standardize_art_status(df):
    df.loc[df['Current ART Status'].str.contains('STOPPED TREATMENT', case=False, na=False), 'Current ART Status'] = 'Stopped Treatment'
    df.loc[df['Previous ART Status'].str.contains('STOPPED TREATMENT', case=False, na=False), 'Previous ART Status'] = 'Stopped Treatment'
    df.loc[df['Model Devolved To'].str.contains('CBM2 - Community ART Refill Group: Healthcare Worker â€“ led', case=False, na=False), 'Model Devolved To'] = 'CBM2 - Community ART Refill Group: Healthcare Worker – led'
    df.loc[df['Model Devolved To'].str.contains('"CBM3 - Community ART Refill Group: PLHIV â€“ led "', case=False, na=False), 'Model Devolved To'] = '"CBM3 - Community ART Refill Group: PLHIV – led "'
    df.loc[df['Current DSD Model'].str.contains('CBM2 - Community ART Refill Group: Healthcare Worker â€“ led', case=False, na=False), 'Current DSD Model'] = 'CBM2 - Community ART Refill Group: Healthcare Worker – led'
    df.loc[df['Current DSD Model'].str.contains('"CBM3 - Community ART Refill Group: PLHIV â€“ led "', case=False, na=False), 'Current DSD Model'] = '"CBM3 - Community ART Refill Group: PLHIV – led "'

    return df

       


# Function to compare the columns between the two combined documents
def compare_documents(df1, df2):

    rename_dict = {
        
        # 'state': 'State',
        'Date of Additional TB Diagnosis Result using XRAY (for client w' : 'Date of Additional TB Diagnosis Result using XRAY (for client with negative lab results with CAD score of 40 & above)',
        'Additional TB Diagnosis Result using XRAY (for client with nega' : 'Additional TB Diagnosis Result using XRAY (for client with negative lab results with CAD score of 40 & above)'
        
    }

    rename_dict2 = {
                    'Date of Birth (yyyy-mm-dd)' : 'Date Of Birth (yyyy-mm-dd)',
                    'Datim Id' : 'Datim Id',
                    'DatimId'  : 'Datim Id',
                    'Date of Repeat Viral Load - Post EAC VL Sample collected (yyyy-mm-dd)' : 'DateOfRepeatViralLoad-Post-EAC-VL-SampleCollected (yyyy-mm-dd)',
                    'Date of devolvement':'Date of Devolvement', 
                    'Model devolved to': 'Model Devolved To',
                    'Current DSD model': 'Current DSD Model',
                    
                    }
    df1.rename(columns=rename_dict2, inplace=True) #client
    df2.rename(columns=rename_dict, inplace=True)  #centralsync

    df1 = convert_to_integer_columns(df1, columns_to_clean)
    df2 = convert_to_integer_columns(df2, columns_to_clean)

    df1 = standardize_art_status(df1)
    df2 = standardize_art_status(df2)

        # Check if 'Patient ID' exists in both dataframes
    if 'Patient ID' not in df1.columns:
        raise KeyError("'NDR Patient Identifier' not found in df1")
    if 'Patient ID' not in df2.columns:
        raise KeyError("'NDR Patient Identifier' not found in df2")


    # Duplicate the Patient ID column and rename them before merging
    df1['NDR Patient Identifier_Client'] = df1['NDR Patient Identifier']
    df2['NDR Patient Identifier_Centralsync'] = df2['NDR Patient Identifier']
    
    

    # Clean all date columns in both dataframes
    df1 = clean_all_date_columns(df1)
    df2 = clean_all_date_columns(df2)

    # Clean all date columns in both dataframes
    df1 = clean_all_text_columns(df1)
    df2 = clean_all_text_columns(df2)

    
    # Standardize date columns 
    date_columns = [col for col in df1.columns if 'date' in col.lower()]
    date_columns = [col for col in df2.columns if 'date' in col.lower()]
    df1 = standardize_date_columns(df1, date_columns)
    df2 = standardize_date_columns(df2, date_columns)

    df1.columns = df1.columns.str.strip()#.str.lower()
    df2.columns = df2.columns.str.strip()#.str.lower()

    
    # Merge the two dataframes based on 'Patient ID' using an outer join
    merged_df = pd.merge(df1, df2, on='NDR Patient Identifier', how='outer', suffixes=('_Client', '_Centralsync'))

    merged_df.fillna('N/A', inplace=True)

    merged_df = clean_all_date_columns(merged_df)
    
        
    # Prepare the result DataFrame with Patient ID, State, Facility Name, and LGA
    result_df = merged_df[['NDR Patient Identifier']].drop_duplicates()

    result_df = result_df.copy()


    # Create the Patient_ID_Client and Patient_ID_Centralsync columns, return N/A if no match
    result_df['NDR Patient Identifier_Client'] = merged_df['NDR Patient Identifier_Client'].fillna('N/A')
    result_df['NDR Patient Identifier_Centralsync'] = merged_df['NDR Patient Identifier_Centralsync'].fillna('N/A')

    result_df['NDR Patient Identifier_Client'] = merged_df['NDR Patient Identifier_Client'].replace(['', ' ', '\xa0'], 'N/A')
    result_df['NDR Patient Identifier_Centralsync'] = merged_df['NDR Patient Identifier_Centralsync'].replace(['', ' ', '\xa0'], 'N/A')

    
    # Add comparison results for Patient ID existence in Client and Centralsync
    result_df['Match'] = result_df.apply(
        lambda row: 'Match' if row['NDR Patient Identifier_Client'] != 'N/A' and row['NDR Patient Identifier_Centralsync'] != 'N/A' else 'No Match', axis=1
    )
    
    
    
    #columns to compare
    comparison_columns = ['State', 'L.G.A', 'LGA Of Residence', 'Facility Name','Datim Id',  
                          'Patient ID', 'NDR Patient Identifier', 'Hospital Number', 'Unique Id',
                            'Household Unique No', 'OVC Unique ID', 'Sex', 'Target group', 
                            'Current Weight (kg)', 'Pregnancy Status', 'Date Of Birth (yyyy-mm-dd)', # 'Date of Birth (yyyy-mm-dd)', #
                            'Age', 'Care Entry Point', 'Date of Registration', 'Enrollment  Date (yyyy-mm-dd)',
                            'ART Start Date (yyyy-mm-dd)', 'Last Pickup Date (yyyy-mm-dd)', 'Months of ARV Refill',
                            'Regimen Line at ART Start', 'Regimen at ART Start', 'Date of Start of Current ART Regimen',
                            'Current Regimen Line', 'Current ART Regimen', 'Clinical Staging at Last Visit',
                            'Date of Last CD4 Count', 'Last CD4 Count', 'Date of Viral Load Sample Collection (yyyy-mm-dd)',
                            'Date of Current ViralLoad Result Sample (yyyy-mm-dd)', 'Current Viral Load (c/ml)',
                            'Date of Current Viral Load (yyyy-mm-dd)', 'Viral Load Indication',
                            'Viral Load Eligibility Status', 'Date of Viral Load Eligibility Status',
                            'Current ART Status', 'Date of Current ART Status', 'Client Verification Outcome',
                            'Cause of Death', 'VA Cause of Death', 'Previous ART Status',
                            'Confirmed Date of Previous ART Status', 'ART Enrollment Setting',
                            'Date of TB Screening (yyyy-mm-dd)', 'TB Screening Type', 'CAD Score',
                            'TB status', 'Date of TB Sample Collection (yyyy-mm-dd)',
                            'TB Diagnostic Test Type', 'Date of TB Diagnostic Result Received (yyyy-mm-dd)',
                            'TB Diagnostic Result',  'Date of Additional TB Diagnosis Result using XRAY (for client with negative lab results with CAD score of 40 & above)',
                            'Additional TB Diagnosis Result using XRAY (for client with negative lab results with CAD score of 40 & above)',
                            'Date of Start of TB Treatment (yyyy-mm-dd)', 'TB Type (new, relapsed etc)', 
                            'Date of Completion of TB Treatment (yyyy-mm-dd)', 'TB Treatment Outcome', 
                            'Date of TPT Start (yyyy-mm-dd)', 'TPT Type', 'TPT Completion date (yyyy-mm-dd)', 
                            'TPT Completion status', 'Date of commencement of EAC (yyyy-mm-dd)', 
                            'Number of EAC Sessions Completed', 'Date of last EAC Session Completed', 
                            'Date of Extended EAC Completion (yyyy-mm-dd)', 
                            'DateOfRepeatViralLoad-Post-EAC-VL-SampleCollected (yyyy-mm-dd)', #'Date of Repeat Viral Load - Post EAC VL Sample collected (yyyy-mm-dd)', # 
                            'Repeat Viral load result (c/ml)- POST EAC', 'Date of Repeat Viral load result- POST EAC VL', 
                            'Date of Devolvement',  #'Date of devolvement', #
                            'Model Devolved To',  #'Model devolved to', #
                            'Date of current DSD', 
                            'Current DSD Model',  #'Current DSD model', # 
                            'Current DSD Outlet', 'Date of Return of DSD Client to Facility (yyyy-mm-dd)', 
                            'Screening for Chronic Conditions', 'Co-morbidities', 'Date of Cervical Cancer Screening (yyyy-mm-dd)', 
                            'Cervical Cancer Screening Type', 'Cervical Cancer Screening Method', 'Result of Cervical Cancer Screening', 
                            'Date of Precancerous Lesions Treatment (yyyy-mm-dd)', 'Precancerous Lesions Treatment Methods', 
                            'Date Biometrics Enrolled (yyyy-mm-dd)', 'Number of Fingers Captured', 
                            'Date Biometrics Recapture (yyyy-mm-dd)', 'Number of Fingers Recaptured', 'Case Manager'
    ]

    

    for col in comparison_columns:
        # Adding columns for the actual values from Client and Centralsync
        result_df[f'{col}_Client'] = merged_df[f'{col}_Client']
        result_df[f'{col}_Centralsync'] = merged_df[f'{col}_Centralsync']
        
        # Adding match/no match result
        result_df[f'{col}_match'] = merged_df.apply(
    lambda row: 'N/A' if row[f'{col}_Client'] == 'N/A' or row[f'{col}_Centralsync'] == 'N/A' or 
                          row['NDR Patient Identifier_Client'] == 'N/A' or row['NDR Patient Identifier_Centralsync'] == 'N/A'
    else ('Match' if row[f'{col}_Client'] == row[f'{col}_Centralsync'] else 'No Match'), 
    axis=1
)

    return result_df


def generate_summary_sheet(df, summary_output_path):

    summary_data = []

    match_cols = [c for c in df.columns if str(c).lower.endswith('_match')]

    for col in match_cols:
        clean_name = str(col).replace('_match', '').replace('_', '').strip()

        counts = df[col].value_counts()

        match_count = counts.get('Match', 0)
        no_match_count = counts.get('No Match', 0)

        summary_data.append({
            'Column Name': clean_name,
            'Count of Match': match_count,
            'Count of No Match': no_match_count
        })

    #creating dataframe and save
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(summary_output_path, index=False)
    print(f"Summary report saved to {summary_output_path}")



# Combine documents
df_Client = combine_documents(folder_Client)
df_centralsync = combine_documents(folder_centralsync)

# Compare the two combined documents
client_level_analysis = compare_documents(df_Client, df_centralsync)


#Generate the summary Dataframe
def get_summary_df(df):

    summary_data = []

    match_cols = [c for c in df.columns if str(c).lower().endswith('_match')]

    for col in match_cols:
        clean_name = str(col).replace('_match', '').replace('_', '').strip()
        counts = df[col].value_counts()

        summary_data.append({
            'Column Name': clean_name,
            'Count of Match': counts.get('Match', 0),
            'Count of No Match': counts.get('No Match', 0),
            'Count of N/A' : counts.get('N/A', 0)
        })
    return pd.DataFrame(summary_data)

summary_df = get_summary_df(client_level_analysis)

#Using Excelwriter to write to different sheets
with pd.ExcelWriter(output_path, engine ='openpyxl') as writer:

    client_level_analysis.to_excel(writer, sheet_name = 'Comparison', index=False)

    summary_df.to_excel(writer, sheet_name = 'Match Summary', index=False)

print(f'Comparison result saved to {output_path}') 



