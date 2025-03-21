import zipfile
import os
import shutil
import pandas as pd

# Set the path to the directory containing your zip files
zip_dir = 'C:/Users/DELL/Documents/DataFi/NDR Flatfile generation/Zipped_files'
output_dir = 'C:/Users/DELL/Documents/DataFi/NDR Flatfile generation/Output_Directory'

# Initialize empty lists for each group of documents
pvls_den_data = []
pvls_num_data = []
tx_curr_data = []
tx_new_data = []
tx_rtt_data = []
tx_ml_data = []
tx_ml_iit_lt_3_data = []
tx_ml_iit_3_to_5_data = []
tx_ml_iit_6_data = []
tx_ml_died_data = []
tx_ml_transfer_data = []
tx_ml_stopped_data = []

# Loop through the zip files in the directory
for zip_filename in os.listdir(zip_dir):
    if zip_filename.endswith('.zip'):
        zip_filepath = os.path.join(zip_dir, zip_filename)

        # Extract project name from the zip file name
        project_name = zip_filename.split('_')[-3]

        # Unzip the folder to a temporary directory
        with zipfile.ZipFile(zip_filepath, 'r') as zip_ref:
            # Define the temporary directory path
            temp_dir = os.path.join(output_dir, zip_filename.replace('.zip', ''))

            # Remove the existing temporary directory if it exists, to overwrite it
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

            # Create the directory again (now it's clean and ready)
            os.makedirs(temp_dir, exist_ok=True)
            zip_ref.extractall(temp_dir)

            # Loop through the files in the unzipped directory
            for file in os.listdir(temp_dir):
                if file.endswith('.xlsx'):  # Process only Excel files
                    file_path = os.path.join(temp_dir, file)

                    # Read the Excel file
                    df = pd.read_excel(file_path, engine='openpyxl')

                    # Add the project name to the DataFrame
                    df['ProjectName'] = project_name

                    # Categorize files based on their prefix and process accordingly
                    if file.startswith('PVLSDen'):
                        pvls_den_data.append(df)
                    elif file.startswith('PVLSNum'):
                        pvls_num_data.append(df)
                    elif file.startswith('TXNEW'):
                        tx_new_data.append(df)
                    elif file.startswith('TXRTT'):
                        tx_rtt_data.append(df)
                    elif file.startswith('TXML'):
                        if any('Died' in item for item in df['CategoryOptionComboName']):
                            df_filtered = df[df['CategoryOptionComboName'].str.contains('Died', na=False)]
                            tx_ml_died_data.append(df_filtered)
                        if any('Transferred Out' in item for item in df['CategoryOptionComboName']):
                            df_filtered = df[df['CategoryOptionComboName'].str.contains('Transferred Out', na=False)]
                            tx_ml_transfer_data.append(df_filtered)
                        if any('Refused (Stopped) Treatment, Positive' in item for item in df['CategoryOptionComboName']):
                            df_filtered = df[df['CategoryOptionComboName'].str.contains('Refused (Stopped) Treatment, Positive', na=False, regex=False)]  # The fix: regex=False
                            tx_ml_stopped_data.append(df_filtered)
                    elif file.startswith('TXML_IIT'):
                        if any('Interruption in Treatment (<3 Months Treatment), Positive' in item for item in df['CategoryOptionComboName']):
                            df_filtered = df[df['CategoryOptionComboName'].str.contains('Interruption in Treatment (<3 Months Treatment), Positive', na=False, regex=False)]
                            tx_ml_iit_lt_3_data.append(df_filtered)
                        if any('Interruption in Treatment (3-5 Months Treatment), Positive' in item for item in df['CategoryOptionComboName']):
                            df_filtered = df[df['CategoryOptionComboName'].str.contains('Interruption in Treatment (3-5 Months Treatment), Positive', na=False, regex=False)]
                            tx_ml_iit_3_to_5_data.append(df_filtered)
                        if any('Interruption in Treatment (6+ Months Treatment), Positive' in item for item in df['CategoryOptionComboName']):
                            df_filtered = df[df['CategoryOptionComboName'].str.contains('Interruption in Treatment (6+ Months Treatment), Positive', na=False, regex=False)]
                            tx_ml_iit_6_data.append(df_filtered)
                    elif file.startswith('TXCURR'):
                        # Filter out rows in TXCURR where DataElementName contains 'ARVDispense'
                        df = df[~df['DataElementName'].str.contains('ARVDispense', na=False)]
                        tx_curr_data.append(df)

# Combine the data for each group
pvls_den_combined = pd.concat(pvls_den_data, ignore_index=True) if pvls_den_data else pd.DataFrame()
pvls_num_combined = pd.concat(pvls_num_data, ignore_index=True) if pvls_num_data else pd.DataFrame()
tx_curr_combined = pd.concat(tx_curr_data, ignore_index=True) if tx_curr_data else pd.DataFrame()
tx_new_combined = pd.concat(tx_new_data, ignore_index=True) if tx_new_data else pd.DataFrame()
tx_rtt_combined = pd.concat(tx_rtt_data, ignore_index=True) if tx_rtt_data else pd.DataFrame()
tx_ml_combined = pd.concat(tx_ml_data, ignore_index=True) if tx_ml_data else pd.DataFrame()
tx_ml_died_combined = pd.concat(tx_ml_died_data, ignore_index=True) if tx_ml_died_data else pd.DataFrame()
tx_ml_transfer_combined = pd.concat(tx_ml_transfer_data, ignore_index=True) if tx_ml_transfer_data else pd.DataFrame()
tx_ml_stopped_combined = pd.concat(tx_ml_stopped_data, ignore_index=True) if tx_ml_stopped_data else pd.DataFrame()
tx_ml_iit_lt_3_combined = pd.concat(tx_ml_iit_lt_3_data, ignore_index=True) if tx_ml_iit_lt_3_data else pd.DataFrame()
tx_ml_iit_3_to_5_combined = pd.concat(tx_ml_iit_3_to_5_data, ignore_index=True) if tx_ml_iit_3_to_5_data else pd.DataFrame()
tx_ml_iit_6_combined = pd.concat(tx_ml_iit_6_data, ignore_index=True) if tx_ml_iit_6_data else pd.DataFrame()

# Pivot the data for each group based on 'ProjectName' and 'Value'
pvls_den_pivot = pvls_den_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not pvls_den_combined.empty else pd.DataFrame()
pvls_num_pivot = pvls_num_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not pvls_num_combined.empty else pd.DataFrame()
tx_curr_pivot = tx_curr_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not tx_curr_combined.empty else pd.DataFrame()
tx_new_pivot = tx_new_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not tx_new_combined.empty else pd.DataFrame()
tx_rtt_pivot = tx_rtt_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not tx_rtt_combined.empty else pd.DataFrame()
tx_ml_pivot = tx_ml_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not tx_ml_combined.empty else pd.DataFrame()
tx_ml_died_pivot = tx_ml_died_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not tx_ml_died_combined.empty else pd.DataFrame()
tx_ml_transfer_pivot = tx_ml_transfer_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not tx_ml_transfer_combined.empty else pd.DataFrame()
tx_ml_stopped_pivot = tx_ml_stopped_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not tx_ml_stopped_combined.empty else pd.DataFrame()
tx_ml_iit_lt_3_pivot = tx_ml_iit_lt_3_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not tx_ml_iit_lt_3_combined.empty else pd.DataFrame()
tx_ml_iit_3_to_5_pivot = tx_ml_iit_3_to_5_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not tx_ml_iit_3_to_5_combined.empty else pd.DataFrame()
tx_ml_iit_6_pivot = tx_ml_iit_6_combined.pivot_table(index='ProjectName', values='Value', aggfunc='sum') if not tx_ml_iit_6_combined.empty else pd.DataFrame()

# Save the combined data and pivoted data to separate Excel files
combined_filepath = os.path.join(output_dir, 'combined_data.xlsx')
with pd.ExcelWriter(combined_filepath, engine='openpyxl') as writer:
    if not pvls_den_combined.empty:
        pvls_den_combined.to_excel(writer, sheet_name='PVLSDen_Combined', index=False)
    if not pvls_num_combined.empty:
        pvls_num_combined.to_excel(writer, sheet_name='PVLSNum_Combined', index=False)
    if not tx_curr_combined.empty:
        tx_curr_combined.to_excel(writer, sheet_name='TXCURR_Combined', index=False)
    if not tx_new_combined.empty:
        tx_new_combined.to_excel(writer, sheet_name='TXNEW_Combined', index=False)
    if not tx_rtt_combined.empty:
        tx_rtt_combined.to_excel(writer, sheet_name='TXRTT_Combined', index=False)
    if not tx_ml_combined.empty:
        tx_ml_combined.to_excel(writer, sheet_name='TXML_Combined', index=False)
    if not tx_ml_died_combined.empty:
        tx_ml_died_combined.to_excel(writer, sheet_name='TXML_DIED_Combined', index=False)
    if not tx_ml_transfer_combined.empty:
        tx_ml_transfer_combined.to_excel(writer, sheet_name='TXML_TRANSFER_Combined', index=False)
    if not tx_ml_stopped_combined.empty:
        tx_ml_stopped_combined.to_excel(writer, sheet_name='TXML_STOPPED_Combined', index=False)
    if not tx_ml_iit_lt_3_combined.empty:
        tx_ml_iit_lt_3_combined.to_excel(writer, sheet_name='TXML_IIT_LT_3_Combined', index=False)
    if not tx_ml_iit_3_to_5_combined.empty:
        tx_ml_iit_3_to_5_combined.to_excel(writer, sheet_name='TXML_IIT_3_TO_5_Combined', index=False)
    if not tx_ml_iit_6_combined.empty:
        tx_ml_iit_6_combined.to_excel(writer, sheet_name='TXML_IIT_6_Combined', index=False)

pivot_filepath = os.path.join(output_dir, 'pivot_data.xlsx')
with pd.ExcelWriter(pivot_filepath, engine='openpyxl') as writer:
    if not pvls_den_pivot.empty:
        pvls_den_pivot.to_excel(writer, sheet_name='PVLSDen_Pivot', index=True)
    if not pvls_num_pivot.empty:
        pvls_num_pivot.to_excel(writer, sheet_name='PVLSNum_Pivot', index=True)
    if not tx_curr_pivot.empty:
        tx_curr_pivot.to_excel(writer, sheet_name='TXCURR_Pivot', index=True)
    if not tx_new_pivot.empty:
        tx_new_pivot.to_excel(writer, sheet_name='TXNEW_Pivot', index=True)
    if not tx_rtt_pivot.empty:
        tx_rtt_pivot.to_excel(writer, sheet_name='TXRTT_Pivot', index=True)
    if not tx_ml_pivot.empty:
        tx_ml_pivot.to_excel(writer, sheet_name='TXML_Pivot', index=True)
    if not tx_ml_died_pivot.empty:
        tx_ml_died_pivot.to_excel(writer, sheet_name='TXML_DIED_Pivot', index=True)
    if not tx_ml_transfer_pivot.empty:
        tx_ml_transfer_pivot.to_excel(writer, sheet_name='TXML_TRANSFER_Pivot', index=True)
    if not tx_ml_stopped_pivot.empty:
        tx_ml_stopped_pivot.to_excel(writer, sheet_name='TXML_STOPPED_Pivot', index=True)
    if not tx_ml_iit_lt_3_pivot.empty:
        tx_ml_iit_lt_3_pivot.to_excel(writer, sheet_name='TXML_IIT_LT_3_Pivot', index=True)
    if not tx_ml_iit_3_to_5_pivot.empty:
        tx_ml_iit_3_to_5_pivot.to_excel(writer, sheet_name='TXML_IIT_3_TO_5_Pivot', index=True)
    if not tx_ml_iit_6_pivot.empty:
        tx_ml_iit_6_pivot.to_excel(writer, sheet_name='TXML_IIT_6_Pivot', index=True)

print(f'Combined data saved to: {combined_filepath}')
print(f'Pivoted data saved to: {pivot_filepath}')
