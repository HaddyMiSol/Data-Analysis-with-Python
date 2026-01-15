import pandas as pd
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import re

# Path to the directory containing the CSV files
folder_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/HTS_IPP/hts 3rd Oct/CS_Filtered2'  

# Output path for the final Excel file
output_file_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Output_by_facility/HTS_Aggregates by Facility_CS_3rd Oct_ACE-1b.xlsx'

# Output path for troubleshooting HIV Testing Setting and Modality Output
hts_setting_output_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Extracted_HTS_setting.xlsx'


# Defining Periods
Start_of_quarter = pd.to_datetime('2025-07-01')
End_of_quarter = pd.to_datetime('2025-09-30')

# Combine all CSV files into one DataFrame, specifying 'latin1' encoding
all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(('.csv', '.xlsx', '.xls'))]


# Combine all files into one DataFrame
combined_data = pd.DataFrame()

for file in all_files:
    try:
        if file.endswith('.csv'):
            data = pd.read_csv(file, encoding='latin1', on_bad_lines='skip')
        elif file.endswith('.xlsx'):
            data = pd.read_excel(file, engine='openpyxl')
        elif file.endswith('.xls'):
            data = pd.read_excel(file)
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
# IMPORTANT: Ensure 'Filename' column exists before attempting to split
if 'Filename' not in combined_data.columns:
    print("FATAL ERROR: 'Filename' column is missing from the combined data. Cannot extract ProjectName.")
    exit()

combined_data['ProjectName'] = combined_data['Filename'].str.split('_').str[0]

# --- Robust column validation ---
# Check if 'ProjectName' and 'Facility' columns exist and are not entirely empty
required_cols_for_aggregation = ['ProjectName', 'Facility', 'Facility Id (Datim)']
for col in required_cols_for_aggregation:
    if col not in combined_data.columns:
        print(f"FATAL ERROR: Required column '{col}' is missing from the combined data.")
        if col == 'ProjectName':
            print("This usually means the 'Filename' column couldn't be processed to extract ProjectName.")
            print(f"First 5 filenames: {combined_data['Filename'].head().tolist()}")
        exit()
    
    # Convert to string and handle potential NaNs before further processing
    combined_data[col] = combined_data[col].astype(str).fillna('UNKNOWN')

    if (combined_data[col] == 'UNKNOWN').all():
        print(f"Warning: Column '{col}' is entirely 'UNKNOWN' (or NaN in original data). This might affect grouping and filtering.")
    elif combined_data[col].nunique() == 1 and combined_data[col].iloc[0] == 'UNKNOWN':
         print(f"Warning: Column '{col}' contains only 'UNKNOWN' values. This might indicate an issue with data extraction or input.")

# --- End of NEW column validation ---


# Ensure date columns are in datetime format
date_columns = [
    'Date Of HIV Testing (yyyy-mm-dd)'
]

for col in date_columns:
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

# Check for invalid dates
invalid_dates = combined_data[date_columns].isnull().any(axis=1)
if invalid_dates.any():
    print(f"Warning: Invalid dates found in the following rows:\n{combined_data[invalid_dates]}")

# Ensure Numeric field is numeric
combined_data['Age'] = pd.to_numeric(combined_data['Age'], errors='coerce')


HTS_TST_Emergency = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Emergency'])) &
    (combined_data['Modality'].isin(['Emergency'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '') 
]


HTS_Emergency_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Emergency'])) &
    (combined_data['Modality'].isin(['Emergency'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_Emergency_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Emergency'])) &
    (combined_data['Modality'].isin(['Emergency'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_Index = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility', 'Community'])) &
    (combined_data['Testing Setting'].isin(['Index'])) &
    (combined_data['Modality'].isin(['Index'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_Index_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility', 'Community'])) &
    (combined_data['Testing Setting'].isin(['Index'])) &
    (combined_data['Modality'].isin(['Index'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_Index_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility', 'Community'])) &
    (combined_data['Testing Setting'].isin(['Index'])) &
    (combined_data['Modality'].isin(['Index'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_Inpatient = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Inpatient', 'Ward/Inpatient', 'Ward'])) &
    (combined_data['Modality'].isin(['Inpatient'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_Inpatient_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Inpatient', 'Ward/Inpatient', 'Ward'])) &
    (combined_data['Modality'].isin(['Inpatient'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_Inpatient_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Inpatient', 'Ward/Inpatient', 'Ward'])) &
    (combined_data['Modality'].isin(['Inpatient'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_Malnutrition = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Modality'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Positive'])) &
    (combined_data['Age'] <5) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_Malnutrition_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Modality'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative'])) &
    (combined_data['Age'] <5) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_Malnutrition_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Modality'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (combined_data['Age'] <5) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_MobileMod = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['Outreach'])) &
    (combined_data['Modality'].isin(['Mobile'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive','Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_MobileMod_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['Outreach'])) &
    (combined_data['Modality'].isin(['Mobile'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_MobileMod_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['Outreach'])) &
    (combined_data['Modality'].isin(['Mobile'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_OtherMod = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['Others', 'Standalone', 'Standalone HTS', 'OVC'])) & #, 'Outreach (Community)', 'Outreach',
    (combined_data['Modality'].isin(['Other Community Platforms'])) & #, 'Other (Community)'
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_OtherMod_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['Others', 'Standalone', 'Standalone HTS', 'OVC'])) & #, 'Outreach (Community)', 'Outreach'
    (combined_data['Modality'].isin(['Other Community Platforms'])) & #, 'Other (Community)'
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_OtherMod_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['Others', 'Standalone', 'Standalone HTS', 'OVC'])) & #, 'Outreach (Community)', 'Outreach' 
    (combined_data['Modality'].isin(['Other Community Platforms'])) & #, 'Other (Community)'
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_OtherPITC = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (
    ((combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Others', 'FP', 'BloodBank', 'Standalone', 'Standalone HTS', 'TB', 'Others (Specify)'])) &
    (combined_data['Modality'].isin(['Other PITC']))) |

    ((combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['TB'])) &
    (combined_data['Modality'].isin(['TB_STAT/OtherPITC']))) |

    ((combined_data['Entry Point'].isin(['Facility'])) &
     (combined_data['Testing Setting'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Modality'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Age'] >=5) ) |

    ((combined_data['Entry Point'].isin(['Facility'])) &
     (combined_data['Testing Setting'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Modality'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Age'] >=5))
    ) &

    (combined_data['Final HIV Test Result'].isin(['Negative', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_OtherPITC_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (
    ((combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Others', 'FP', 'BloodBank', 'Standalone', 'Standalone HTS', 'TB', 'Others (Specify)'])) &
    (combined_data['Modality'].isin(['Other PITC']))) |

    ((combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['TB'])) &
    (combined_data['Modality'].isin(['TB_STAT/OtherPITC']))) |

    ((combined_data['Entry Point'].isin(['Facility'])) &
     (combined_data['Testing Setting'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Modality'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Age'] >=5) ) |

    ((combined_data['Entry Point'].isin(['Facility'])) &
     (combined_data['Testing Setting'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Modality'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Age'] >=5))
    ) &

    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]



HTS_OtherPITC_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (
    ((combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Others', 'FP', 'BloodBank','Blood Bank','Standalone', 'Standalone HTS', 'TB', 'Others (Specify)'])) &
    (combined_data['Modality'].isin(['Other PITC']))) |

    ((combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['TB'])) &
    (combined_data['Modality'].isin(['TB_STAT/OtherPITC']))) |

    ((combined_data['Entry Point'].isin(['Facility'])) &
     (combined_data['Testing Setting'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Modality'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Age'] >=5) ) |

    ((combined_data['Entry Point'].isin(['Facility'])) &
     (combined_data['Testing Setting'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Modality'].isin(['Malnutrition', 'Malnutrition Clinic'])) &
    (combined_data['Age'] >=5))
    ) &

    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_Pediatric = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Modality'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (combined_data['Age'] <5) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_Pediatric_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Modality'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (combined_data['Age'] <5) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_Pediatric_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Modality'].isin(['Pediatrics <5 Clinic', 'Pediatric'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (combined_data['Age'] <5) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


PMTCT_ANC = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (
    ((combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['ANC', 'Spoke health facility'])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)']))) |
     ((combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['Congregational setting', 'Delivery homes','TBA Orthodx', 'TBA Orthodox', 'TBA rt-HCW'])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)']))) 
     )
     &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


PMTCT_ANC_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (
    ((combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['ANC', 'Spoke health facility'])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)']))) |

     ((combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['Congregational setting', 'Delivery homes','TBA Orthodx', 'TBA Orthodox', 'TBA rt-HCW' ])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)']))) 
     )
     &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

PMTCT_ANC_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (
    ((combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['ANC', 'Spoke health facility'])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)']))) |
     ((combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['Congregational setting', 'Delivery homes','TBA Orthodx', 'TBA Orthodox','TBA rt-HCW' ])) &
    (combined_data['Modality'].isin(['PMTCT (ANC1 Only)']))) 
     )
     &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_PMTCT_Post_ANC1_Breastfeeding = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Post Natal Ward/Breastfeeding'])) &
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Breastfeeding)'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative','Negetive', 'Positive'])) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_PMTCT_Post_ANC1_Breastfeeding_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Post Natal Ward/Breastfeeding'])) &
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Breastfeeding)'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative'])) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_PMTCT_Post_ANC1_Breastfeeding_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['Post Natal Ward/Breastfeeding'])) &
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Breastfeeding)'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_TST_PMTCT_PostANC1_Pregnant_Labour_and_Delivery = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['L&D', 'Retesting'])) &  
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Pregnancy/L&D)'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Positive'])) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['L&D', 'Retesting'])) & #, Retesting'
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Pregnancy/L&D)'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['L&D', 'Retesting'])) &  #, Retesting'
    (combined_data['Modality'].isin(['PMTCT (Post ANC1: Pregnancy/L&D)'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (combined_data['Sex'] == 'Female') &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_SNS = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['SNS'])) &
    (combined_data['Modality'].isin(['SNS'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_SNS_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['SNS'])) &
    (combined_data['Modality'].isin(['SNS'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_SNS_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['SNS'])) &
    (combined_data['Modality'].isin(['SNS'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_TST_SNSMod = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['SNS'])) &
    (combined_data['Modality'].isin(['SNS'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive','Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_SNSMod_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['SNS'])) &
    (combined_data['Modality'].isin(['SNS'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_SNSMod_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['SNS'])) &
    (combined_data['Modality'].isin(['SNS'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_STI = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['STI'])) &
    (combined_data['Modality'].isin(['STI'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_STI_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['STI'])) &
    (combined_data['Modality'].isin(['STI'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_STI_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['STI'])) &
    (combined_data['Modality'].isin(['STI'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_VCT = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['CT'])) &
    (combined_data['Modality'].isin(['VCT', 'CT'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_VCT_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['CT'])) &
    (combined_data['Modality'].isin(['VCT', 'CT'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_VCT_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['CT'])) &
    (combined_data['Modality'].isin(['VCT', 'CT'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_PrEP = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['PrEP Testing'])) &
    (combined_data['Modality'].isin(['PrEP_CT HTS'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_PrEP_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['PrEP Testing'])) &
    (combined_data['Modality'].isin(['PrEP_CT HTS'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_PrEP_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['PrEP Testing'])) &
    (combined_data['Modality'].isin(['PrEP_CT HTS'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_TB = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['TB'])) &
    (combined_data['Modality'].isin(['TB'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TB_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['TB'])) &
    (combined_data['Modality'].isin(['TB'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_TB_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Facility'])) &
    (combined_data['Testing Setting'].isin(['TB'])) &
    (combined_data['Modality'].isin(['TB'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_TST_VCTMod = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['CT'])) &
    (combined_data['Modality'].isin(['VCT'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


HTS_VCTMod_Negative = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['CT'])) &
    (combined_data['Modality'].isin(['VCT'])) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]

HTS_VCTMod_Positive = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isin(['Community'])) &
    (combined_data['Testing Setting'].isin(['CT'])) &
    (combined_data['Modality'].isin(['VCT'])) &
    (combined_data['Final HIV Test Result'].isin(['Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


Blank_Entry_Point = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (combined_data['Entry Point'].isna()) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


Setting_no_Modality = combined_data[
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= Start_of_quarter) &
    (combined_data['Date Of HIV Testing (yyyy-mm-dd)'] <= End_of_quarter) &
    (~combined_data['Testing Setting'].isna()) &
    (combined_data['Modality'].isna()) &
    (combined_data['Final HIV Test Result'].isin(['Negative', 'Negetive', 'Positive'])) &
    (~combined_data['Sex'].isna()) &
    (~combined_data['Age'].isna()) &
     (combined_data['Age'] != '')
]


# Pivot_data function to group by 'ProjectName' AND 'Facility'
# This ensures we can later filter by project but still have facility-level aggregates
def pivot_data(data, value_name):
    # Ensure the data passed to pivot_data has the required columns
    for col in ['ProjectName', 'Facility','Facility Id (Datim)']:
        if col not in data.columns:
            print(f"Error: '{col}' not found in data passed to pivot_data. Skipping pivot.")
            return pd.DataFrame(columns=['ProjectName', 'Facility', 'Facility Id (Datim)', value_name]) # Return empty DataFrame

    return data.groupby(['ProjectName', 'Facility', 'Facility Id (Datim)']).size().reset_index(name=value_name)

# Apply the pivot, aggregated by ProjectName and Facility
# HTS_TST_pivot = pivot_data(HTS_TST, 'HTS_TST')
HTS_TST_Emergency_pivot = pivot_data(HTS_TST_Emergency, 'HTS_TST_Emergency')
HTS_Emergency_Negative_pivot = pivot_data(HTS_Emergency_Negative, 'HTS_Emergency_Negative')
HTS_Emergency_Positive_pivot = pivot_data(HTS_Emergency_Positive, 'HTS_Emergency_Positive')
HTS_TST_Index_pivot = pivot_data(HTS_TST_Index, 'HTS_TST_Index')
HTS_Index_Negative_pivot = pivot_data(HTS_TST_Index_Negative, 'HTS_Index_Negative')
HTS_Index_Positive_pivot = pivot_data(HTS_TST_Index_Positive, 'HTS_Index_Positive')
HTS_TST_Inpatient_pivot = pivot_data(HTS_TST_Inpatient, 'HTS_TST_Inpatient')
HTS_Inpatient_Negative_pivot = pivot_data(HTS_Inpatient_Negative, 'HTS_Inpatient_Negative')
HTS_Inpatient_Positive_pivot = pivot_data(HTS_Inpatient_Positive, 'HTS_Inpatient_Positive')
HTS_TST_Malnutrition_pivot = pivot_data(HTS_TST_Malnutrition, 'HTS_TST_Malnutrition')
HTS_Malnutrition_Negative_pivot = pivot_data(HTS_Malnutrition_Negative, 'HTS_Malnutrition_Negative')
HTS_Malnutrition_Positive_pivot = pivot_data(HTS_Malnutrition_Positive, 'HTS_Malnutrition_Positive')
HTS_TST_MobileMod_pivot = pivot_data(HTS_TST_MobileMod, 'HTS_TST_MobileMod')
HTS_MobileMod_Negative_pivot = pivot_data(HTS_MobileMod_Negative, 'HTS_MobileMod_Negative')
HTS_MobileMod_Positive_pivot = pivot_data(HTS_MobileMod_Positive, 'HTS_MobileMod_Positive')
HTS_TST_OtherMod_pivot = pivot_data(HTS_TST_OtherMod, 'HTS_TST_OtherMod')
HTS_OtherMod_Negative_pivot = pivot_data(HTS_OtherMod_Negative, 'HTS_OtherMod_Negative')
HTS_OtherMod_Positive_pivot = pivot_data(HTS_OtherMod_Positive, 'HTS_OtherMod_Positive')
HTS_TST_OtherPITC_pivot = pivot_data(HTS_TST_OtherPITC, 'HTS_TST_OtherPITC')
HTS_OtherPITC_Negative_pivot = pivot_data(HTS_OtherPITC_Negative, 'HTS_OtherPITC_Negative')
HTS_OtherPITC_Positive_pivot = pivot_data(HTS_OtherPITC_Positive, 'HTS_OtherPITC_Positive')
HTS_TST_Pediatric_pivot = pivot_data(HTS_TST_Pediatric, 'HTS_TST_Pediatric')
HTS_Pediatric_Negative_pivot = pivot_data(HTS_Pediatric_Negative, 'HTS_Pediatric_Negative')
HTS_Pediatric_Positive_pivot = pivot_data(HTS_Pediatric_Positive, 'HTS_Pediatric_Positive')
PMTCT_ANC_pivot = pivot_data(PMTCT_ANC, 'PMTCT_ANC')
PMTCT_ANC_Negative_pivot = pivot_data(PMTCT_ANC_Negative, 'PMTCT_ANC_Negative')
PMTCT_ANC_Positive_pivot = pivot_data(PMTCT_ANC_Positive, 'PMTCT_ANC_Positive')
HTS_TST_PMTCT_Post_ANC1_Breastfeeding_pivot = pivot_data(HTS_TST_PMTCT_Post_ANC1_Breastfeeding, 'HTS_TST_PMTCT_Post_ANC1_Breastfeeding')
HTS_PMTCT_Post_ANC1_Breastfeeding_Negative_pivot = pivot_data(HTS_PMTCT_Post_ANC1_Breastfeeding_Negative, 'HTS_PMTCT_Post_ANC1_Breastfeeding_Negative')
HTS_PMTCT_Post_ANC1_Breastfeeding_Positive_pivot = pivot_data(HTS_PMTCT_Post_ANC1_Breastfeeding_Positive, 'HTS_PMTCT_Post_ANC1_Breastfeeding_Positive')
HTS_TST_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_pivot = pivot_data(HTS_TST_PMTCT_PostANC1_Pregnant_Labour_and_Delivery, 'HTS_TST_PMTCT_PostANC1_Pregnant_Labour_and_Delivery')
HTS_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_Negative_pivot = pivot_data(HTS_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_Negative, 'HTS_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_Negative')
HTS_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_Positive_pivot = pivot_data(HTS_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_Positive, 'HTS_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_Positive')
HTS_TST_SNS_pivot = pivot_data(HTS_TST_SNS, 'HTS_TST_SNS')
HTS_SNS_Negative_pivot = pivot_data(HTS_SNS_Negative, 'HTS_SNS_Negative')
HTS_SNS_Positive_pivot = pivot_data(HTS_SNS_Positive, 'HTS_SNS_Positive')
HTS_TST_SNSMod_pivot = pivot_data(HTS_TST_SNSMod, 'HTS_TST_SNSMod')
HTS_SNSMod_Negative_pivot = pivot_data(HTS_SNSMod_Negative, 'HTS_SNSMod_Negative')
HTS_SNSMod_Positive_pivot = pivot_data(HTS_SNSMod_Positive, 'HTS_SNSMod_Positive')
HTS_TST_STI_pivot = pivot_data(HTS_TST_STI, 'HTS_TST_STI')
HTS_STI_Negative_pivot = pivot_data(HTS_STI_Negative, 'HTS_STI_Negative')
HTS_STI_Positive_pivot = pivot_data(HTS_STI_Positive, 'HTS_STI_Positive')
HTS_TST_TB_pivot = pivot_data(HTS_TST_TB, 'HTS_TST_TB')
HTS_TB_Negative_pivot = pivot_data(HTS_TB_Negative, 'HTS_TB_Negative')
HTS_TB_Positive_pivot = pivot_data(HTS_TB_Positive, 'HTS_TB_Positive')
HTS_TST_PrEP_pivot = pivot_data(HTS_TST_TB, 'HTS_TST_TB')
HTS_PrEP_Negative_pivot = pivot_data(HTS_TB_Negative, 'HTS_TB_Negative')
HTS_PrEP_Positive_pivot = pivot_data(HTS_TB_Positive, 'HTS_TB_Positive')
HTS_TST_VCTMod_pivot = pivot_data(HTS_TST_VCTMod, 'HTS_TST_VCTMod')
HTS_VCTMod_Negative_pivot = pivot_data(HTS_VCTMod_Negative, 'HTS_VCTMod_Negative')
HTS_VCTMod_Positive_pivot = pivot_data(HTS_VCTMod_Positive, 'HTS_VCTMod_Positive')
HTS_TST_VCT_pivot = pivot_data(HTS_TST_VCT, 'HTS_TST_VCT')
HTS_VCT_Negative_pivot = pivot_data(HTS_VCT_Negative, 'HTS_VCT_Negative')
HTS_VCT_Positive_pivot = pivot_data(HTS_VCT_Positive, 'HTS_VCT_Positive')
Setting_no_Modality_pivot = pivot_data(Setting_no_Modality, 'Setting_no_Modality')
Blank_Entry_Point_pivot = pivot_data(Blank_Entry_Point, 'Blank_Entry_Point')


# List of all pivots to be merged into the master summary DataFrame
all_pivots_for_summary = [
    HTS_TST_Emergency_pivot,
    HTS_Emergency_Negative_pivot,
    HTS_Emergency_Positive_pivot,
    HTS_TST_Index_pivot,
    HTS_Index_Negative_pivot,
    HTS_Index_Positive_pivot,
    HTS_TST_Inpatient_pivot,
    HTS_Inpatient_Negative_pivot,
    HTS_Inpatient_Positive_pivot,
    HTS_TST_Malnutrition_pivot,
    HTS_Malnutrition_Negative_pivot,
    HTS_Malnutrition_Positive_pivot,
    HTS_TST_MobileMod_pivot,
    HTS_MobileMod_Negative_pivot,
    HTS_MobileMod_Positive_pivot,
    HTS_TST_OtherMod_pivot,
    HTS_OtherMod_Negative_pivot,
    HTS_OtherMod_Positive_pivot,
    HTS_TST_OtherPITC_pivot,
    HTS_OtherPITC_Negative_pivot,
    HTS_OtherPITC_Positive_pivot,
    HTS_TST_Pediatric_pivot,
    HTS_Pediatric_Negative_pivot,
    HTS_Pediatric_Positive_pivot,
    PMTCT_ANC_pivot,
    PMTCT_ANC_Negative_pivot,
    PMTCT_ANC_Positive_pivot,
    HTS_TST_PMTCT_Post_ANC1_Breastfeeding_pivot,
    HTS_PMTCT_Post_ANC1_Breastfeeding_Negative_pivot,
    HTS_PMTCT_Post_ANC1_Breastfeeding_Positive_pivot,
    HTS_TST_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_pivot,
    HTS_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_Negative_pivot,
    HTS_PMTCT_PostANC1_Pregnant_Labour_and_Delivery_Positive_pivot,
    HTS_TST_SNS_pivot ,
    HTS_SNS_Negative_pivot,
    HTS_SNS_Positive_pivot,
    HTS_TST_SNSMod_pivot,
    HTS_SNSMod_Negative_pivot,
    HTS_SNSMod_Positive_pivot,
    HTS_TST_STI_pivot,
    HTS_STI_Negative_pivot,
    HTS_STI_Positive_pivot,
    HTS_TST_TB_pivot,
    HTS_TB_Negative_pivot,
    HTS_TB_Positive_pivot,
    HTS_TST_PrEP_pivot,
    HTS_PrEP_Negative_pivot,
    HTS_PrEP_Positive_pivot,
    HTS_TST_VCTMod_pivot,
    HTS_VCTMod_Negative_pivot,
    HTS_VCTMod_Positive_pivot,
    HTS_TST_VCT_pivot,
    HTS_VCT_Negative_pivot,
    HTS_VCT_Positive_pivot,
    Setting_no_Modality_pivot,
    Blank_Entry_Point_pivot   
    ] 

# Get all unique combinations of ProjectName and Facility from the original combined data
# This ensures all facilities are represented, even if they have no metrics for a pivot
if not combined_data.empty:
    unique_project_facility_combinations = combined_data[['ProjectName', 'Facility','Facility Id (Datim)']].drop_duplicates()
else:
    unique_project_facility_combinations = pd.DataFrame(columns=['ProjectName', 'Facility', 'Facility Id (Datim)'])

# Initialize a master DataFrame that will contain all aggregated data (by ProjectName and Facility)
master_aggregated_df = unique_project_facility_combinations.copy()

# Merge all calculated pivots onto the master aggregated DataFrame
for pivot_df in all_pivots_for_summary:
    # Ensure pivot_df is not empty and has the keys before merging
    if not pivot_df.empty and 'ProjectName' in pivot_df.columns and 'Facility' in pivot_df.columns:
        master_aggregated_df = pd.merge(master_aggregated_df, pivot_df, 
                                        on=['ProjectName', 'Facility', 'Facility Id (Datim)'], how='left')
    else:
        print(f"Warning: An aggregated pivot is empty or missing key columns (ProjectName/Facility/Datim Id) and will not be merged.")
        
# Create the Excel writer object
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    if not master_aggregated_df.empty:
        # Sort the DataFrame for better readability in the final output
        master_aggregated_df.sort_values(by=['ProjectName', 'Facility','Facility Id (Datim)'], inplace=True)
        
        # Save the entire master aggregated data to a single sheet
        sheet_name = 'Facility_Aggregates'
        master_aggregated_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"All aggregated data saved to a single sheet '{sheet_name}'.")
    else:
        print("Master aggregated DataFrame is empty. No data will be saved.")

print(f"Analysis complete. Results saved to: {output_file_path}")



hts_columns = ['Entry Point','Testing Setting', 'Modality', 'Date Of HIV Testing (yyyy-mm-dd)', 'Facility']
if all(col in combined_data.columns for col in hts_columns):
    cutoff_date = Start_of_quarter   #pd.to_datetime('2025-06-30')
    filtered_data = combined_data[combined_data['Date Of HIV Testing (yyyy-mm-dd)'] >= cutoff_date].copy()
    hts_setting_df = filtered_data[hts_columns].copy()
    distinct_hts_setting_df = hts_setting_df.drop_duplicates()

    try:
        distinct_hts_setting_df.to_excel(hts_setting_output_path, index=False)
        print(f"\nSuccessfully saved distinct hts setting data to: {hts_setting_output_path}")
    except Exception as e:
        print(f"\nError saving hts setting data: {e}")
else:
    print("\nWarning: 'Testing Setting' or 'Modality' column not found. Skipping hts setting data export.")
