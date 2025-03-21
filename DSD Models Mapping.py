import pandas as pd
import numpy as np
import os

# Function to combine all Excel files from a folder
def combine_documents(folder_path):
    # Get all Excel files in the specified folder
    all_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    # Read and combine the files into one DataFrame
    combined_data = []
    for file in all_files:
        file_path = os.path.join(folder_path, file)
        data = pd.read_excel(file_path)
        combined_data.append(data)
    
    # Concatenate all dataframes into one
    return pd.concat(combined_data, ignore_index=True)

# Folder path where your Excel files are stored
folder_path = 'C:/Users/DELL/Documents/DataFi/DSD Model/RADET'

# Combine documents in the folder
df = combine_documents(folder_path)

# Clean up invisible or non-standard blank cells in 'DSD Model'
df['DSD Model'] = df['DSD Model'].str.strip().replace('', np.nan).replace('\xa0', np.nan)#(.replace(['', '\xa0'], np.nan)

# Normalize the 'State' column (strip spaces, convert to title case, and remove hidden characters)
df['State'] = df['State'].str.strip()  # Remove leading and trailing spaces
df['State'] = df['State'].apply(lambda x: x.strip() if isinstance(x, str) else x)  # Remove hidden spaces
df['State'] = df['State'].str.title()  # Convert to title case (e.g., 'sokoto' becomes 'Sokoto')

# Filter the data based on 'Current ART Status' and 'Client Verification Status'
filtered_df = df[(df['Current ART Status'].str.contains('Active', na=False))]#&  # Contains 'Active'
                 #(df['Client Verification Status'].isin(['valid', '']))]  # Valid or blank 

# Count the rows where 'DSD Model' is blank (NaN or empty)
filtered_df['blank_currentdsdmodel_count'] = filtered_df['DSD Model'].isna() | (filtered_df['DSD Model'] == '')

# Group by 'State' and sum the 'blank_currentdsdmodel_count' to get the count of blank 'DSD Model' for each state
blank_count_by_state = filtered_df.groupby('State')['blank_currentdsdmodel_count'].sum().reset_index()

# Create the pivot table based on 'State' and 'DSD Model'
pivot_table = filtered_df.pivot_table(
    index='State',  # State in one column
    columns='DSD Model',  # Different columns for current DSD model
    values='Patient ID',  # Count values of Patient ID
    aggfunc='count',  # Count occurrences
    fill_value=0  # Fill empty cells with 0 (or another value)
)

# Merge the blank count for each state into the pivot table
pivot_table = pivot_table.merge(blank_count_by_state, on='State', how='left')

# Rename the new column for clarity
pivot_table = pivot_table.rename(columns={'blank_currentdsdmodel_count': 'Blank Current DSD Model Count'})

# Optional: Save the pivoted data to a new Excel file
pivot_table.to_excel('pivoted_output.xlsx')

# Display the pivot table
print(pivot_table)
