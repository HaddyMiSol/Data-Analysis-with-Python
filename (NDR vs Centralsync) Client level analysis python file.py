import pandas as pd
import re
from datetime import datetime

# File paths
ndr_file_path = 'C:/Users/DELL/Documents/Ideas for Automation in DataFi/Client_level_analysis/RADET/HANN.xlsx'
centralsync_file_path = 'C:/Users/DELL/Documents/Ideas for Automation in DataFi/Client_level_analysis/Centralsync/HAN.csv'


output_path = 'C:/Users/DELL/Documents/Ideas for Automation in DataFi/Client_level_analysis/client_level_analysis.xlsx'

# Function to clean up facility names (remove initials)
def clean_facility_name(facility_name):
    # Remove initials like 'ad', 'bo', 'yo', etc. from the start of the facility name
    cleaned_name = re.sub(r'^[a-zA-Z]{2}\s?', '', facility_name)  # Assumes initials are 2 letters, followed by space or no space
    return cleaned_name

def clean_blanks(df, columns_to_clean):
    for column in columns_to_clean:
        if column in df.columns:
            # Convert the column to string to handle invisible characters
            df[column] = df[column].astype(str)
            
            # Replace empty strings, non-breaking spaces, and other whitespace characters with NaN
            df[column] = df[column].replace(['', '\xa0', ' ', '\t', '\n', 'NULL'], pd.NA)
            
            # Strip leading/trailing whitespaces
            df[column] = df[column].str.strip()
            
            # Replace all NaN values with consistent pandas NA type
            df[column] = df[column].fillna(pd.NA)
    return df

# Function to clean invisible or non-standard blank cells in date columns
def clean_date_column(df, column_name):
    # Convert date column to string format to clean up any invisible characters
    df[column_name] = df[column_name].astype(str)
    
    # Replace 'NaT' or empty string and non-breaking spaces with NaT
    df[column_name] = df[column_name].replace(['', '\xa0', ' ', '\t', '\n'], pd.NaT)
    
    # Convert back to datetime, handling any other non-date values gracefully
    df[column_name] = pd.to_datetime(df[column_name], format='%Y-%m-%d', errors='coerce')
    
    return df

# Function to format 'current_viral_load_Q4' column
def format_viral_load(viral_load):
    try:
        # Convert to float and then remove .0 if present
        return str(int(float(viral_load))) if float(viral_load).is_integer() else viral_load
    except ValueError:
        # Return the original value if it cannot be converted
        return viral_load



def compare_document(ndr_file_path, centralsync_file_path):
    # Load the NDR dataset
    try:
        ndr_df = pd.read_excel(ndr_file_path)
    except Exception as e:
        print(f"Error loading NDR file: {e}")
        raise

    # Load the Centralsync dataset with robust settings
    try:
        centralsync_df = pd.read_csv(
            centralsync_file_path,
            encoding='latin1',  # Use a fallback encoding
            engine='python',  # Use the Python engine for flexibility
            on_bad_lines='skip'  # Skip problematic lines
        )
    except Exception as e:
        print(f"Error loading Centralsync file: {e}")
        raise

    # Rename column in Centralsync dataset for consistency
    if 'NDR Patient Identifier' in centralsync_df.columns:
        centralsync_df.rename(columns={'NDR Patient Identifier': 'patient_identifier'}, inplace=True)

    # Rename columns in NDR dataset for clarity
    rename_dict = {
        'state_name': 'State',
        'facility_name': 'Facility Name',
        # 'patient_identifier': 'patient_identifier',  
        'art_start_date': 'ART Start Date (yyyy-mm-dd)',
        'last_drug_pickup_date_Q4': 'Last Pickup Date (yyyy-mm-dd)',
        'current_viral_load_Q4': 'Current Viral Load (c/ml)',
        'date_of_current_viral_load_Q4': 'Date of Current Viral Load (yyyy-mm-dd)',
        'date_of_current_sample_collection_Q4': 'Date of Viral Load Sample Collection (yyyy-mm-dd)',
        'final_outcome': 'Client Verification Status'
    }
    ndr_df.rename(columns=rename_dict, inplace=True)

    # Standardize values
    ndr_df['Client Verification Status'] = ndr_df['Client Verification Status'].str.lower().astype(str)
    centralsync_df['Client Verification Status'] = centralsync_df['Client Verification Status'].str.lower().astype(str)

    # Format 'Current Viral Load (c/ml)' column
    ndr_df['Current Viral Load (c/ml)'] = ndr_df['Current Viral Load (c/ml)'].astype(str).apply(format_viral_load)
    centralsync_df['Current Viral Load (c/ml)'] = centralsync_df['Current Viral Load (c/ml)'].astype(str).apply(format_viral_load)

    # Create separate identifiers for clarity during the merge
    ndr_df['patient_identifier_ndr'] = ndr_df['patient_identifier']
    centralsync_df['patient_identifier_centralsync'] = centralsync_df['patient_identifier']

    # Merge the datasets on 'patient_identifier'
    merged_df = pd.merge(ndr_df, centralsync_df, on='patient_identifier', how='outer', suffixes=('_ndr', '_centralsync'))

    # Prepare the result DataFrame
    result_df = merged_df[['patient_identifier']].drop_duplicates()  # , 'State', 'Facility Name'

    # Fill NA values in patient identifiers
    result_df['patient_identifier_ndr'] = merged_df['patient_identifier_ndr'].fillna('N/A')
    result_df['patient_identifier_centralsync'] = merged_df['patient_identifier_centralsync'].fillna('N/A')

        # List of columns to clean
    columns_to_clean = [
        'Last Pickup Date (yyyy-mm-dd)_ndr',
        'Last Pickup Date (yyyy-mm-dd)_centralsync',
        'ART Start Date (yyyy-mm-dd)_ndr',
        'ART Start Date (yyyy-mm-dd)_centralsync',
        'Date of Current Viral Load (yyyy-mm-dd)_ndr',
        'Date of Current Viral Load (yyyy-mm-dd)_centralsync',
        'Date of Viral Load Sample Collection (yyyy-mm-dd)_ndr',
        'Date of Viral Load Sample Collection (yyyy-mm-dd)_centralsync',
        'Current Viral Load (c/ml)',
        'Client Verification Status'
    ]

    # Apply the cleaning function to the DataFrame
    merged_df = clean_blanks(merged_df, columns_to_clean)


    # List of date columns to clean
    date_columns = [
        'Last Pickup Date (yyyy-mm-dd)_ndr',
        'Last Pickup Date (yyyy-mm-dd)_centralsync',
        'ART Start Date (yyyy-mm-dd)_ndr',
        'ART Start Date (yyyy-mm-dd)_centralsync',
        'Date of Current Viral Load (yyyy-mm-dd)_ndr',
        'Date of Current Viral Load (yyyy-mm-dd)_centralsync',
        'Date of Viral Load Sample Collection (yyyy-mm-dd)_ndr',
        'Date of Viral Load Sample Collection (yyyy-mm-dd)_centralsync'
    ]

    # Clean each date column in the list
    for col in date_columns:
        merged_df = clean_date_column(merged_df, col)

    # Columns to compare
    columns_to_compare = [
        'patient_identifier',
        'ART Start Date (yyyy-mm-dd)',
        'Last Pickup Date (yyyy-mm-dd)',
        'Current Viral Load (c/ml)',
        'Date of Current Viral Load (yyyy-mm-dd)',
        'Date of Viral Load Sample Collection (yyyy-mm-dd)',
        'Client Verification Status'
    ]

    # Add comparison columns
    for col in columns_to_compare:
        # Add columns for values from NDR and Centralsync
        result_df[f'{col}_ndr'] = merged_df.get(f'{col}_ndr', None)
        result_df[f'{col}_centralsync'] = merged_df.get(f'{col}_centralsync', None)

        # Add match/no match result
        result_df[f'{col}_match'] = merged_df.apply(
            lambda row: 'Match' if row.get(f'{col}_ndr') == row.get(f'{col}_centralsync') else 'No Match', axis=1
        )

    return result_df

# Compare the two documents
client_level_analysis = compare_document(ndr_file_path, centralsync_file_path)

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

#Using Excelwriter to write to different sheets
with pd.ExcelWriter(output_path, engine ='openpyxl') as writer:

    client_level_analysis.to_excel(writer, sheet_name = 'Comparison', index=False)

    summary_df.to_excel(writer, sheet_name = 'Match Summary', index=False)

print(f'Comparison result saved to {output_path}') 


