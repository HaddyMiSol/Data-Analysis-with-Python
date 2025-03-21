import pandas as pd
import os
from datetime import datetime

# Function to read all CSV files in a folder and return a combined DataFrame
def load_and_combine_csv(folder_path):
    all_files = os.listdir(folder_path)
    csv_files = [f for f in all_files if f.endswith('.csv')]
    df_list = []
    
    for file in csv_files:
        file_path = os.path.join(folder_path, file)
        df = pd.read_csv(file_path)
        # Extract IP Name from file name
        df['IP Name'] = file.split('_')[0]
        df_list.append(df)
        
    combined_df = pd.concat(df_list, ignore_index=True)
    return combined_df

# Function to clean up date columns to date format
def clean_date_columns(df, date_columns):
    for col in date_columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

# Function to calculate the metrics
def calculate_metrics(df, filter_date):
    # Convert filter_date to datetime
    filter_date = pd.to_datetime(filter_date)
    
    # Calculate each metric based on the provided conditions
    df['TX_NEW'] = ((df['ART Start Date (yyyy-mm-dd)'] >= filter_date) &
                    (df['Care Entry Point'] != 'Transfer-in') &
                    ((df['Client Verification Status'].isna()) | (df['Client Verification Status'] == 'valid')))
    
    df['TX_RTT'] = ((df['Current ART Status'] == 'Active Restart') &
                    (df['Date of Current ART Status'] >= filter_date) &
                    ((df['Client Verification Status'].isna()) | (df['Client Verification Status'] == 'valid')))
    
    df['TX_ML-Died'] = ((df['Current ART Status'] == 'Died') &
                        (df['Date of Current ART Status'] >= filter_date) &
                        ((df['Client Verification Status'].isna()) | (df['Client Verification Status'] == 'valid')))
    
    df['TX_ML-IIT<3mth'] = ((df['Current ART Status'] == 'IIT') &
                            (df['Date of Current ART Status'] >= filter_date) &
                            ((df['ART Start Date (yyyy-mm-dd)'] - df['Last Pickup Date (yyyy-mm-dd)']).dt.days < 90) &
                            ((df['Client Verification Status'].isna()) | (df['Client Verification Status'] == 'valid')))
    
    df['TX_ML-IIT3-5mths'] = ((df['Current ART Status'] == 'IIT') &
                               (df['Date of Current ART Status'] >= filter_date) &
                               ((df['ART Start Date (yyyy-mm-dd)'] - df['Last Pickup Date (yyyy-mm-dd)']).dt.days >= 90) &
                               ((df['ART Start Date (yyyy-mm-dd)'] - df['Last Pickup Date (yyyy-mm-dd)']).dt.days < 150) &
                               ((df['Client Verification Status'].isna()) | (df['Client Verification Status'] == 'valid')))
    
    df['TX_ML-IIT>=6mths'] = ((df['Current ART Status'] == 'IIT') &
                              (df['Date of Current ART Status'] >= filter_date) &
                              ((df['ART Start Date (yyyy-mm-dd)'] - df['Last Pickup Date (yyyy-mm-dd)']).dt.days >= 180) &
                              ((df['Client Verification Status'].isna()) | (df['Client Verification Status'] == 'valid')))
    
    df['TX_ML-Transferred_out'] = ((df['Current ART Status'].str.contains('Transfer')) &
                                   (df['Date of Current ART Status'] >= filter_date) &
                                   ((df['Client Verification Status'].isna()) | (df['Client Verification Status'] == 'valid')))
    
    df['TX_ML-Stopped_TX'] = ((df['Current ART Status'].str.contains('Stop')) &
                              (df['Date of Current ART Status'] >= filter_date) &
                              ((df['Client Verification Status'].isna()) | (df['Client Verification Status'] == 'valid')))
    
    df['TX_ML-Invalid'] = ((df['Date of Current ART Status'] >= filter_date) &
                           (df['Client Verification Status'] == 'Invalid'))
    
    df['TX_ML-Verification_ongoing'] = ((df['Date of Current ART Status'] >= filter_date) &
                                        (df['Client Verification Status'].str.contains('Verification')))
    
    df['Current_TX_CURR'] = ((df['Current ART Status'].str.contains('Active')) &
                             ((df['Client Verification Status'].isna()) | (df['Client Verification Status'] == 'valid')))
    
    return df

# Function to generate the pivot table
def generate_pivot_table(df):
    # Group by the required columns and aggregate the counts for each metric
    pivot_df = df.groupby(['IP Name', 'State', 'Facility Name', 'Datim Id']).agg(
        TX_NEW=('TX_NEW', 'sum'),
        TX_RTT=('TX_RTT', 'sum'),
        TX_ML_Died=('TX_ML-Died', 'sum'),
        TX_ML_IIT_3mth=('TX_ML-IIT<3mth', 'sum'),
        TX_ML_IIT_3_5mths=('TX_ML-IIT3-5mths', 'sum'),
        TX_ML_IIT_6mths=('TX_ML-IIT>=6mths', 'sum'),
        TX_ML_Transferred_out=('TX_ML-Transferred_out', 'sum'),
        TX_ML_Stopped_TX=('TX_ML-Stopped_TX', 'sum'),
        TX_ML_Invalid=('TX_ML-Invalid', 'sum'),
        TX_ML_Verification_ongoing=('TX_ML-Verification_ongoing', 'sum'),
        Current_TX_CURR=('Current_TX_CURR', 'sum')
    ).reset_index()
    
    return pivot_df

# Export to Excel
def export_to_excel(pivot_df, output_path):
    pivot_df.to_excel(output_path, index=False)

# Main function to run the process
def process_csv_files(folder_path, filter_date, output_path):
    df = load_and_combine_csv(folder_path)
    
    # Clean date columns
    date_columns = ['ART Start Date (yyyy-mm-dd)', 'Date of Current ART Status', 'Last Pickup Date (yyyy-mm-dd)']
    df = clean_date_columns(df, date_columns)
    
    # Calculate metrics
    df = calculate_metrics(df, filter_date)
    
    # Generate pivot table
    pivot_df = generate_pivot_table(df)
    
    # Export to Excel
    export_to_excel(pivot_df, output_path)
    print(f"Pivot table saved to {output_path}")

# usage:
folder_path = 'C:/Users/DELL/Documents/DataFi/Waterfall Analysis/Current_quarter RADET'
filter_date = '2024-10-01'  # filter date
output_path = 'C:/Users/DELL/Documents/DataFi/Waterfall Analysis/Current_quarter RADET/output_pivot_table.xlsx'

process_csv_files(folder_path, filter_date, output_path)


# For Previous FY TX_CURR, use thesame code above while commenting out every other indicator and using the folder path below
# usage:
folder_path = 'C:/Users/DELL/Documents/DataFi/Waterfall Analysis/Prev_quarter RADET'
filter_date = '2024-10-01'  # filter date
output_path = 'C:/Users/DELL/Documents/DataFi/Waterfall Analysis/Prev_quarter RADET/output_pivot_table2.xlsx'

process_csv_files(folder_path, filter_date, output_path)
