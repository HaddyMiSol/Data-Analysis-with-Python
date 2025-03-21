import os
import pandas as pd

def aggregate_data_from_folder(folder_path, output_path, output_excel_file):
    """
    Aggregates data from CSV files in a folder, grouping by data element, IP, and Period.

    Args:
        folder_path (str): The path to the folder containing the CSV files.
        output_path (str): The path to the output folder.
        output_excel_file (str): The name of the output Excel file.
    """

    all_data = {}  # Dictionary to store dataframes for each period

    for filename in os.listdir(folder_path):
        if filename.endswith((".csv", ".xlsx")):  # Check for both CSV and XLSX
            file_path = os.path.join(folder_path, filename)
            try:
                if filename.endswith(".csv"):
                    df = pd.read_csv(file_path)
                elif filename.endswith(".xlsx"):
                    df = pd.read_excel(file_path) #read excel files

                if 'IP' in df.columns and 'Facility Name' in df.columns and 'Data Element Name' in df.columns and 'Value' in df.columns and 'Period' in df.columns:
                    for ip_value, ip_group in df.groupby('IP'):
                        for period_value, period_group in ip_group.groupby('Period'): # group by period also within each IP group.
                            grouped = period_group.groupby([ 'Facility Name','Data Element Name'])['Value'].sum().reset_index()
                            for index, row in grouped.iterrows():
                                data_row = {
                                    'IP': ip_value,
                                    'Facility Name': row['Facility Name'],
                                    'Data Element Name': row['Data Element Name'],
                                    'Aggregated Data': row['Value'],
                                    'Period': period_value
                                }
                                if period_value not in all_data:
                                    all_data[period_value] = []
                                all_data[period_value].append(data_row)
                else:
                    print(f"Warning: 'IP', 'Facility Name' ,'Data Element Name', 'Value' or 'Period' column not found in {filename}. Skipping.")

            except Exception as e:
                print(f"Error processing {filename}: {e}")

    # Write data to Excel, each period in a separate sheet
    output_full_path = os.path.join(output_path, output_excel_file) 
    with pd.ExcelWriter(output_full_path) as writer:
        for period, data in all_data.items():
            pd.DataFrame(data).to_excel(writer, sheet_name=f"Period {period}", index=False)

    print(f"Data aggregated and written to {output_full_path}")


folder = "C:/Users/DELL/Documents/DataFi/CS Flatfile Aggregation/Flatfile"
output_path = "C:/Users/DELL/Documents/DataFi/CS Flatfile Aggregation" 
output_file = "aggregated_data1.xlsx"
aggregate_data_from_folder(folder, output_path, output_file) 
