import os
import pandas as pd
from datetime import datetime

folder_path ='C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/FY26Q1_RADET/NEW'

output_base_dir = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/PMTCT_Project_Export_Quality_Check'

#Defining Periods
filter_date = datetime(2024, 10, 1)

os.makedirs(output_base_dir, exist_ok=True)

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

combined_data['Project_name'] = combined_data['Filename'].str.split('_').str[0]

#convert each date column to date
for col in ['Date Tested for HIV', "Mother''s ART Start Date",  
             'Viral Load Sample Collection Date', 'Date Of Maternal Retesting']:
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

#iterate through each project
for project_name in combined_data['Project_name'].unique():
    project_data = combined_data[combined_data['Project_name']==project_name].copy()

    project_dir = os.path.join(output_base_dir, project_name)
    os.makedirs(project_dir, exist_ok=True)

    project_issues = {}
    all_line_lists_data = []

    
    #blank Marital Status
    blank_Marital_Status = project_data[
        (project_data['Marital Status'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Marital_Status'] = blank_Marital_Status.shape[0]
    blank_Marital_Status['Quality_Issue'] = 'blank Marital Status'
    all_line_lists_data.append(blank_Marital_Status)

    
    #blank Point of Entry
    blank_Point_of_Entry = project_data[
        (project_data['Point of Entry'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Point_of_Entry'] = blank_Point_of_Entry.shape[0]
    blank_Point_of_Entry['Quality_Issue'] = 'blank Point of Entry'
    all_line_lists_data.append(blank_Point_of_Entry)

    #blank Modality
    blank_Modality = project_data[
        (project_data['Modality'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Modality'] = blank_Modality.shape[0]
    blank_Modality['Quality_Issue'] = 'blank Modality'
    all_line_lists_data.append(blank_Modality)

    #blank Gestational Age (Weeks) @ First ANC visit
    blank_Gestational_Age_at_First_ANC_visit = project_data[
        (project_data['Gestational Age (Weeks) @ First ANC visit'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Gestational_Age_at_First_ANC_visit'] = blank_Gestational_Age_at_First_ANC_visit.shape[0]
    blank_Gestational_Age_at_First_ANC_visit['Quality_Issue'] = 'blank Gestational Age (Weeks) @ First ANC visit'
    all_line_lists_data.append(blank_Gestational_Age_at_First_ANC_visit)

    #blank Garvida
    blank_Garvida = project_data[
        (project_data['Gravida'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Garvida'] = blank_Garvida.shape[0]
    blank_Garvida['Quality_Issue'] = 'blank Garvida'
    all_line_lists_data.append(blank_Garvida)

    #blank Parity
    blank_Parity = project_data[
        (project_data['Parity'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Parity'] = blank_Parity.shape[0]
    blank_Parity['Quality_Issue'] = 'blank Parity'
    all_line_lists_data.append(blank_Parity)

    #blank Date Tested for HIV
    blank_Date_Tested_for_HIV = project_data[
        (project_data['Date Tested for HIV'].isna()) #&
        #(project_data['Date Tested for HIV'] > filter_date)
    ]
    project_issues['blank_Date_Tested_for_HIV'] = blank_Date_Tested_for_HIV.shape[0]
    blank_Date_Tested_for_HIV['Quality_Issue'] = 'blank Date Tested for HIV'
    all_line_lists_data.append(blank_Date_Tested_for_HIV)

    #blank HIV Test Result
    blank_HIV_Test_Result = project_data[
        (project_data['HIV Test Result'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_HIV_Test_Result'] = blank_HIV_Test_Result.shape[0]
    blank_HIV_Test_Result['Quality_Issue'] = 'blank HIV Test Result'
    all_line_lists_data.append(blank_HIV_Test_Result)

    #blank Recency Test Type where If Recency Testing Opt In is True
    blank_Recency_Test_Type_where_If_Recency_Testing_Opt_In_is_True = project_data[
        (project_data['If Recency Testing Opt In'] =='True') &
        (project_data['Recency Test Type'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Recency_Test_Type_where_If_Recency_Testing_Opt_In_is_True'] = blank_Recency_Test_Type_where_If_Recency_Testing_Opt_In_is_True.shape[0]
    blank_Recency_Test_Type_where_If_Recency_Testing_Opt_In_is_True['Quality_Issue'] = 'blank Recency Test Type where If Recency Testing Opt In is True'
    all_line_lists_data.append(blank_Recency_Test_Type_where_If_Recency_Testing_Opt_In_is_True)

    
    #blank Recency Interpretation where If Recency Testing Opt In is True
    blank_Recency_Interpretation_where_If_Recency_Testing_Opt_In_is_True = project_data[
        (project_data['If Recency Testing Opt In'] =='True') &
        (project_data['Recency Interpretation'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Recency_Interpretation_where_If_Recency_Testing_Opt_In_is_True'] = blank_Recency_Interpretation_where_If_Recency_Testing_Opt_In_is_True.shape[0]
    blank_Recency_Interpretation_where_If_Recency_Testing_Opt_In_is_True['Quality_Issue'] = 'blank Recency Interpretation where If Recency Testing Opt In is True'
    all_line_lists_data.append(blank_Recency_Interpretation_where_If_Recency_Testing_Opt_In_is_True)

    #blank Final Recency Result where VL Confirmation Result is not blank
    blank_Final_Recency_Result_where_VL_Confirmation_Result_is_not_blank = project_data[
        (~project_data['Viral Load Confirmation Result'].isna()) &
        (project_data['Final Recency Result'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Final_Recency_Result_where_VL_Confirmation_Result_is_not_blank'] = blank_Final_Recency_Result_where_VL_Confirmation_Result_is_not_blank.shape[0]
    blank_Final_Recency_Result_where_VL_Confirmation_Result_is_not_blank['Quality_Issue'] = 'blank Final Recency Result where VL Confirmation Result is not blank'
    all_line_lists_data.append(blank_Final_Recency_Result_where_VL_Confirmation_Result_is_not_blank)

    
    #blank HIV Test Result where Previously Known Hiv Status is negative
    blank_HIV_Test_Result_where_Previously_Known_Hiv_Status_is_negative = project_data[
        (project_data['Previously Known HIV Status'] == 'Negative') &
        (project_data['HIV Test Result'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_HIV_Test_Result_where_Previously_Known_Hiv_Status_is_negative'] = blank_HIV_Test_Result_where_Previously_Known_Hiv_Status_is_negative.shape[0]
    blank_HIV_Test_Result_where_Previously_Known_Hiv_Status_is_negative['Quality_Issue'] = 'blank HIV Test Result where Previously Known Hiv Status is negative'
    all_line_lists_data.append(blank_HIV_Test_Result_where_Previously_Known_Hiv_Status_is_negative)

    #blank Previously Known Hiv Status where Date Of Maternal Retesting is not blank
    blank_Previously_Known_Hiv_Status_where_Date_Of_Maternal_Retesting_is_not_blank = project_data[
        (project_data['Previously Known HIV Status'].isna()) &
        (~project_data['Date Of Maternal Retesting'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Previously_Known_Hiv_Status_where_Date_Of_Maternal_Retesting_is_not_blank'] = blank_Previously_Known_Hiv_Status_where_Date_Of_Maternal_Retesting_is_not_blank.shape[0]
    blank_Previously_Known_Hiv_Status_where_Date_Of_Maternal_Retesting_is_not_blank['Quality_Issue'] = 'blank Previously Known Hiv Status where Date Of Maternal Retesting is not blank'
    all_line_lists_data.append(blank_Previously_Known_Hiv_Status_where_Date_Of_Maternal_Retesting_is_not_blank)

    #blank Mother Unique ID where Mother ART Start Date is not blank
    blank_Mother_Unique_ID_where_Mother_ART_Start_Date_is_not_blank = project_data[
        (project_data["Mother''s Unique ID"].isna()) &
        (~project_data["Mother''s ART Start Date"].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Mother_Unique_ID_where_Mother_ART_Start_Date_is_not_blank'] = blank_Mother_Unique_ID_where_Mother_ART_Start_Date_is_not_blank.shape[0]
    blank_Mother_Unique_ID_where_Mother_ART_Start_Date_is_not_blank['Quality_Issue'] = 'blank Mother Unique ID where Mother ART Start Date is not blank'
    all_line_lists_data.append(blank_Mother_Unique_ID_where_Mother_ART_Start_Date_is_not_blank)

    
    #blank Hepatitis B Test Result where Date Tested for Hepatitis B is not blank
    blank_Hepatitis_B_Test_Result_where_Date_Tested_for_Hepatitis_B_is_not_blank = project_data[
        (project_data['Hepatitis B Test Result'].isna()) &
        (~project_data['Date Tested for Hepatitis B'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Hepatitis_B_Test_Result_where_Date_Tested_for_Hepatitis_B_is_not_blank'] = blank_Hepatitis_B_Test_Result_where_Date_Tested_for_Hepatitis_B_is_not_blank.shape[0]
    blank_Hepatitis_B_Test_Result_where_Date_Tested_for_Hepatitis_B_is_not_blank['Quality_Issue'] = 'blank Hepatitis B Test Result where Date Tested for Hepatitis B is not blank'
    all_line_lists_data.append(blank_Hepatitis_B_Test_Result_where_Date_Tested_for_Hepatitis_B_is_not_blank)

    #blank Hepatitis C Test Result where Date Tested for Hepatitis C is not blank
    blank_Hepatitis_C_Test_Result_where_Date_Tested_for_Hepatitis_C_is_not_blank = project_data[
        (project_data['Hepatitis C Test Result'].isna()) &
        (~project_data['Date Tested for Hepatitis C'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Hepatitis_C_Test_Result_where_Date_Tested_for_Hepatitis_C_is_not_blank'] = blank_Hepatitis_C_Test_Result_where_Date_Tested_for_Hepatitis_C_is_not_blank.shape[0]
    blank_Hepatitis_C_Test_Result_where_Date_Tested_for_Hepatitis_C_is_not_blank['Quality_Issue'] = 'blank Hepatitis C Test Result where Date Tested for Hepatitis C is not blank'
    all_line_lists_data.append(blank_Hepatitis_C_Test_Result_where_Date_Tested_for_Hepatitis_C_is_not_blank)

    
    
    #blank Viral Load Sample Collection Date where HIV Test Result is positive and Recency Interpretation is RTRI Recent
    blank_VL_Sample_Collection_Date_where_HIV_Test_Result_is_positive_and_Recency_Interpretation_is_RTRI_Recent = project_data[
        (project_data['Viral Load Sample Collection Date'].isna()) &
        (project_data['Recency Interpretation'] == 'RTRI Recent') &
        (project_data['HIV Test Result'] == 'Positive') &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_VL_Sample_Collection_Date_where_HIV_Test_Result_is_positive_and_Recency_Interpretation_is_RTRI_Recent'] = blank_VL_Sample_Collection_Date_where_HIV_Test_Result_is_positive_and_Recency_Interpretation_is_RTRI_Recent.shape[0]
    blank_VL_Sample_Collection_Date_where_HIV_Test_Result_is_positive_and_Recency_Interpretation_is_RTRI_Recent['Quality_Issue'] = 'blank Viral Load Sample Collection Date where HIV Test Result is positive and Recency Interpretation is RTRI Recent'
    all_line_lists_data.append(blank_VL_Sample_Collection_Date_where_HIV_Test_Result_is_positive_and_Recency_Interpretation_is_RTRI_Recent)


    #blank Maternal Retesting Result where Date Of Maternal Retesting is not blank
    blank_Maternal_Retesting_Result_where_Date_Of_Maternal_Retesting_is_not_blank = project_data[
        (project_data['Maternal Retesting Result'].isna()) &
        (~project_data['Date Of Maternal Retesting'].isna()) &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Maternal_Retesting_Result_where_Date_Of_Maternal_Retesting_is_not_blank'] = blank_Maternal_Retesting_Result_where_Date_Of_Maternal_Retesting_is_not_blank.shape[0]
    blank_Maternal_Retesting_Result_where_Date_Of_Maternal_Retesting_is_not_blank['Quality_Issue'] = 'blank Maternal Retesting Result where Date Of Maternal Retesting is not blank'
    all_line_lists_data.append(blank_Maternal_Retesting_Result_where_Date_Of_Maternal_Retesting_is_not_blank)


    #blank Mother ART Start Date where HIV Test Result and Maternal Retesting Result is positive 
    blank_Mother_ART_Start_Date_where_HIV_Test_Result_and_Maternal_Retesting_Result_is_positive = project_data[
        (project_data["Mother''s ART Start Date"].isna()) &
        (project_data['Maternal Retesting Result'] == 'Positive') &
        (project_data['HIV Test Result'] == 'Positive') &
        (project_data['Date Tested for HIV'] >= filter_date)
    ]
    project_issues['blank_Mother_ART_Start_Date_where_HIV_Test_Result_and_Maternal_Retesting_Result_is_positive'] = blank_Mother_ART_Start_Date_where_HIV_Test_Result_and_Maternal_Retesting_Result_is_positive.shape[0]
    blank_Mother_ART_Start_Date_where_HIV_Test_Result_and_Maternal_Retesting_Result_is_positive['Quality_Issue'] = 'blank Mother ART Start Date where HIV Test Result and Maternal Retesting Result is positive'
    all_line_lists_data.append(blank_Mother_ART_Start_Date_where_HIV_Test_Result_and_Maternal_Retesting_Result_is_positive)
    



    # Create DataFrame for transposed quality issues summary for the current project
    quality_issues_df_transposed = pd.DataFrame.from_dict(project_issues, orient='index', columns= ['Number of Records']).reset_index()
    quality_issues_df_transposed.columns = ['Quality Issue', 'Number of Records']

    # Combine all line list data for the current project into one DataFrame
    all_line_lists_df = pd.concat(all_line_lists_data, ignore_index=True)

    # Define the output file path for the current project
    output_file_path =os.path.join(project_dir, f"{project_name}_Quality Checks.xlsx")

    # Save the transposed quality issues and the combined line list to one Excel file with two sheets
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        quality_issues_df_transposed.to_excel(writer, sheet_name="Quality Issues Summary", index=False)
        all_line_lists_df.to_excel(writer, sheet_name= 'Quality Issues Linelist', index=False)

    print(f"Quality check for project '{project_name}' saved to: {output_file_path}")
