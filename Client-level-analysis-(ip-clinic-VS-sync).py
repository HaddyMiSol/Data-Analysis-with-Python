import os
import pandas as pd
import re
from datetime import datetime


# Paths to the folders
folder_Client = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis/Clinic Client'
folder_centralsync = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis/Clinic CS'

output_path = 'C:/Users/DELL/Documents/DataFi/Client_level_analysis/Clinic Client level analysis/UMTH_Clinic_client_level_analysis.xlsx'

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

    
    # Replace blanks or NaN values in each identified text column
    for column in text_columns:
        df[column] = df[column].fillna(default_text)  # Replace NaN with default text
        df[column] = df[column].replace(['', ' ', '\xa0'], default_text)  # Replace empty strings or non-breaking spaces
    
    return df

# Convert columns to integers where possible
def convert_to_integer_columns(df, columns):
    for column in columns:
        df[column] = pd.to_numeric(df[column], errors='coerce')  # Convert to numeric, invalid entries become NaN
        df[column] = df[column].fillna(0).astype(int)  # Replace NaN with 0 and convert to integer
    return df

# Function to standardize one or multiple date columns
def standardize_date_columns(df, columns, output_format='%d/%m/%Y'):
    for column in columns:
        df[column] = pd.to_datetime(df[column], errors='coerce').dt.strftime(output_format)
    return df



# Function to compare the columns between the two combined documents
def compare_documents(df1, df2):

    # rename_dict = {'DATIM Id' : 'Datim Id',
    #                'Hospital Num' : 'Hospital Number',
    #                'Regimens(Include supported Drugs)' : 'Regimens'}

    rename_dict2 = {'Facility' : 'Facility Name'
                        }
    df1.rename(columns=rename_dict2, inplace=True) #client
    # df2.rename(columns=rename_dict, inplace=True) #centralsync

    # Check if 'Patient Id' exists in both dataframes
    if 'Patient Id' not in df1.columns:
        raise KeyError("'Patient Identifier' not found in df1")
    if 'Patient Id' not in df2.columns:
        raise KeyError("'Patient Identifier' not found in df2")

    # Duplicate the Patient ID column and rename them before merging
    df1['Patient Id_Client'] = df1['Patient Id']
    df2['Patient Id_Centralsync'] = df2['Patient Id']

    # Clean all date columns in both dataframes
    df1 = clean_all_date_columns(df1)
    df2 = clean_all_date_columns(df2)

    # Clean all text columns in both dataframes
    df1 = clean_all_text_columns(df1)
    df2 = clean_all_text_columns(df2)

    # Standardize date columns explicitly
    date_columns_df1 = [col for col in df1.columns if 'date' in col.lower()]
    date_columns_df2 = [col for col in df2.columns if 'date' in col.lower()]
    df1 = standardize_date_columns(df1, date_columns_df1)
    df2 = standardize_date_columns(df2, date_columns_df2)

    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    # Merge the two dataframes based on 'Patient Id' using an outer join
    merged_df = pd.merge(df1, df2, on='Patient Id', how='outer', suffixes=('_Client', '_Centralsync'))

    merged_df.fillna('N/A', inplace=True)

    merged_df = clean_all_date_columns(merged_df)

    # Prepare the result DataFrame with Patient ID
    result_df = merged_df[['Patient Id']].copy() #.drop_duplicates()

    # Explicitly create the Patient_ID columns
    result_df['Patient Id_Client'] = merged_df['Patient Id_Client'].replace(['', ' ', '\xa0'], 'N/A')
    result_df['Patient Id_Centralsync'] = merged_df['Patient Id_Centralsync'].replace(['', ' ', '\xa0'], 'N/A')

    # Add comparison results for Patient ID existence
    result_df['Match'] = result_df.apply(
        lambda row: 'Match' if row['Patient Id_Client'] != 'N/A' and row['Patient Id_Centralsync'] != 'N/A' else 'No Match', axis=1
    )

    # columns to compare
    comparison_columns = [
        'Facility Name','Datim Id','Patient Id','Hospital Number','Date Visit',
        'Clinic Stage','Function Status','TB Status','Body Weight (kg)','Height (cm)',
        'BP (mmHg)','Systolic','Diastolic','Pregnant Status','Next Appointment'
    ]

    for col in comparison_columns:
        # Adding columns for the actual values from Client and Centralsync
        result_df[f'{col}_Client'] = merged_df[f'{col}_Client']
        result_df[f'{col}_Centralsync'] = merged_df[f'{col}_Centralsync']

        # Adding match/no match result
        result_df[f'{col}_match'] = merged_df.apply(
            lambda row: 'N/A' if row[f'{col}_Client'] == 'N/A' or row[f'{col}_Centralsync'] == 'N/A' or
                                    row['Patient Id_Client'] == 'N/A' or row['Patient Id_Centralsync'] == 'N/A'
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


#Generate the summary Dtaframe
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


# Save the result to an Excel file with multiple sheets
# Determine the maximum number of rows per sheet (Excel limit is around 1,048,576)
rows_per_sheet = 1000000  # Adjust this value as needed

# Calculate the number of sheets required
num_sheets = (len(client_level_analysis) // rows_per_sheet) + (1 if len(client_level_analysis) % rows_per_sheet > 0 else 0)

# Excel writer object
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    for i in range(num_sheets):
        start_row = i * rows_per_sheet
        end_row = min((i + 1) * rows_per_sheet, len(client_level_analysis))
        sheet_name = f'Sheet{i+1}'
        client_level_analysis[start_row:end_row].to_excel(writer, sheet_name='Comparison', index=False)

    summary_df.to_excel(writer, sheet_name = 'Match Summary', index=False)

print(f'Comparison result saved to {output_path}') 
