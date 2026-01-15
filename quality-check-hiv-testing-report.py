import os
import pandas as pd
from datetime import datetime

folder_path ='C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/FY26Q1_RADET/NEW'

output_base_dir = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/HTS_Project_Export_Quality_Check'

#Defining Periods
filter_date = datetime(1900, 1, 1)

os.makedirs(output_base_dir, exist_ok=True)

all_files = [os.path.join(folder_path,f) for f in os.listdir(folder_path) if f.endswith(('csv','xlsx','.xls'))]



# Combine all files into one DataFrame
combined_data = pd.DataFrame()

for file in all_files:
    try:
        if file.endswith('csv'):
            data = pd.read_csv(file, encoding = 'latin1', on_bad_lines='skip')
        elif file.endswith('xlsx'):
            data = pd.read_excel(file, engine='openpyxl')
        elif file.endswith('xls'):
            data = pd.read_excel(file)
        else:
            continue

        data['Filename'] = os.path.basename(file)
        combined_data = pd.concat([combined_data, data], ignore_index=True)
    except Exception as e:
        print(f"Error processing file {file}: {e}")

if combined_data.empty:
    print("No valid files found or data could not be combined.")
    exit()

# Extract project name from the 'Filename' column
combined_data['ProjectName'] = combined_data['Filename'].str.split('_').str[0]

#convert each date column to date
for col in ['Date Of Birth (yyyy-mm-dd)', 'Date of Visit (yyyy-mm-dd)', 'Date Of HIV Testing (yyyy-mm-dd)', 
            'Recency Test Date (yyyy-mm-dd)', 'Recency Viral Load Sample Collection Date', 
            'Recency Test Date (yyyy-mm-dd)', 'Recency Viral Load Result Received Date (yyyy-mm-dd)']: #'Recency Viral Load Result Received Date (yyyy-mm-dd)'
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

quality_issue_counts = []
line_lists = {} 

#iterate through each project
for project_name in combined_data['ProjectName'].unique():
    project_data = combined_data[combined_data['ProjectName'] == project_name].copy()

    # Create directory for the current project
    project_dir = os.path.join(output_base_dir, project_name)
    os.makedirs(project_dir, exist_ok=True)

    project_issues = {}
    all_line_lists_data = []

    #Index client without index type
    Index_client_without_index_type = project_data[
        (project_data['Index Client']=='Yes') &
        (project_data['Index Type'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date) 
    ]
    project_issues['Index_client_without_index_type'] = Index_client_without_index_type.shape[0]
    Index_client_without_index_type['Quality_Issue'] = 'Index client without index type'
    all_line_lists_data.append(Index_client_without_index_type)

    #blank Date of Birth
    blank_date_of_birth = project_data[
        (project_data['Date Of Birth (yyyy-mm-dd)'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_date_of_birth'] = blank_date_of_birth.shape[0]
    blank_date_of_birth['Quality_Issue'] = 'blank date of birth'
    all_line_lists_data.append(blank_date_of_birth)

    #blank Sex
    blank_Sex = project_data[
        (project_data['Sex'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_Sex'] = blank_Sex.shape[0]
    blank_Sex['Quality_Issue'] = 'blank Sex'
    all_line_lists_data.append(blank_Sex)

    #Unisex Sex
    Unisex_Sex = project_data[
        (~project_data['Sex'].isin(["Female", "Male"])) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['Unisex_Sex'] = Unisex_Sex.shape[0]
    Unisex_Sex['Quality_Issue'] = 'Unisex Sex'
    all_line_lists_data.append(Unisex_Sex)

    #blank Date of visit
    blank_date_of_visit = project_data[
        (project_data['Date of Visit (yyyy-mm-dd)'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_date_of_visit'] = blank_date_of_visit.shape[0]
    blank_date_of_visit['Quality_Issue'] = 'blank date of visit'
    all_line_lists_data.append(blank_date_of_visit)

    #blank First Time Visit
    blank_First_Time_Visit = project_data[
        (project_data['First Time Visit'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_First_Time_Visit'] = blank_First_Time_Visit.shape[0]
    blank_First_Time_Visit['Quality_Issue'] = 'blank First Time Visit'
    all_line_lists_data.append(blank_First_Time_Visit)

    #blank Index Client
    blank_Index_Client = project_data[
        (project_data['Index Client'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_Index_Client'] = blank_Index_Client.shape[0]
    blank_Index_Client['Quality_Issue'] = 'blank Index Client'
    all_line_lists_data.append(blank_Index_Client)

    #blank Previously Tested
    blank_Previously_Tested = project_data[
        (project_data['Previously Tested'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_Previously_Tested'] = blank_Previously_Tested.shape[0]
    blank_Previously_Tested['Quality_Issue'] = 'blank Previously Tested'
    all_line_lists_data.append(blank_Previously_Tested)

    #blank Referred From
    blank_Referred_From = project_data[
        (project_data['Referred From'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_Referred_From'] = blank_Referred_From.shape[0]
    blank_Referred_From['Quality_Issue'] = 'blank Referred From'
    all_line_lists_data.append(blank_Referred_From)

    #blank Testing Setting
    blank_Testing_Setting = project_data[
        (project_data['Testing Setting'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_Testing_Setting'] = blank_Testing_Setting.shape[0]
    blank_Testing_Setting['Quality_Issue'] = 'blank Testing Setting'
    all_line_lists_data.append(blank_Testing_Setting)

    #blank Counseling Type
    blank_Counseling_Type = project_data[
        (project_data['Counseling Type'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_Counseling_Type'] = blank_Counseling_Type.shape[0]
    blank_Counseling_Type['Quality_Issue'] = 'blank Counseling Type'
    all_line_lists_data.append(blank_Counseling_Type)

    

    #blank Syphilis Test Result for pregnant women
    blank_Syphilis_Test_Result_for_pregnant_women = project_data[
        (project_data['Syphilis Test Result'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Sex'] == 'Female') &
        (project_data['Age'] >=15) &
        (project_data['Pregnancy/Breastfeeding Status'] == 'Pregnant')
    ]
    project_issues['blank_Syphilis_Test_Result_for_pregnant_women'] = blank_Syphilis_Test_Result_for_pregnant_women.shape[0]
    blank_Syphilis_Test_Result_for_pregnant_women['Quality_Issue'] = 'blank Syphilis Test Result for pregnant women'
    all_line_lists_data.append(blank_Syphilis_Test_Result_for_pregnant_women)

    #blank CD4 Test Result where CD4 Type is not blank
    blank_CD4_Test_Result_where_CD4_Type_is_not_blank = project_data[
        (~project_data['CD4 Type'].isna()) &
        (project_data['CD4 Test Result'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_CD4_Test_Result_where_CD4_Type_is_not_blank'] = blank_CD4_Test_Result_where_CD4_Type_is_not_blank.shape[0]
    blank_CD4_Test_Result_where_CD4_Type_is_not_blank['Quality_Issue'] = 'blank CD4 Test Result where CD4 Type is not blank'
    all_line_lists_data.append(blank_CD4_Test_Result_where_CD4_Type_is_not_blank)

    #CD4 Result for HIV Negative Result
    CD4_Result_for_HIV_Negative_Result = project_data[
        (project_data['Final HIV Test Result'] == 'Negative') &
        (~project_data['CD4 Test Result'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['CD4_Result_for_HIV_Negative_Result'] = CD4_Result_for_HIV_Negative_Result.shape[0]
    CD4_Result_for_HIV_Negative_Result['Quality_Issue'] = 'CD4 Result for HIV Negative Result'
    all_line_lists_data.append(CD4_Result_for_HIV_Negative_Result)

    #blank Final HIV Test Result
    blank_Final_HIV_Test_Result = project_data[
        (project_data['Final HIV Test Result'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_Final_HIV_Test_Result'] = blank_Final_HIV_Test_Result.shape[0]
    blank_Final_HIV_Test_Result['Quality_Issue'] = 'blank Final HIV Test Result'
    all_line_lists_data.append(blank_Final_HIV_Test_Result)

    #blank Date Of HIV Testing
    blank_Date_Of_HIV_Testing = project_data[
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'].isna())
    ]
    project_issues['blank_Date_Of_HIV_Testing'] = blank_Date_Of_HIV_Testing.shape[0]
    blank_Date_Of_HIV_Testing['Quality_Issue'] = 'blank Date Of HIV Testing'
    all_line_lists_data.append(blank_Date_Of_HIV_Testing)

    #Date of visit < Date of Birth
    Date_of_visit_is_earlier_than_Date_of_Birth = project_data[
        (project_data['Date of Visit (yyyy-mm-dd)'] < project_data['Date Of Birth (yyyy-mm-dd)']) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['Date_of_visit_is_earlier_than_Date_of_Birth'] = Date_of_visit_is_earlier_than_Date_of_Birth.shape[0]
    Date_of_visit_is_earlier_than_Date_of_Birth['Quality_Issue'] = 'Date of visit is earlier than Date of Birth'
    all_line_lists_data.append(Date_of_visit_is_earlier_than_Date_of_Birth)


    #Viral Load Confirmation Date < Viral Load Sample Collection Date
    Viral_Load_Confirmation_Date_is_earlier_than_Viral_Load_Sample_Collection_Date = project_data[
        (project_data['Recency Viral Load Result Received Date (yyyy-mm-dd)'] < project_data['Recency Viral Load Sample Collection Date']) & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['Viral_Load_Confirmation_Date_is_earlier_than_Recency_Viral_Load_Result_Received_Date'] = Viral_Load_Confirmation_Date_is_earlier_than_Viral_Load_Sample_Collection_Date.shape[0]
    Viral_Load_Confirmation_Date_is_earlier_than_Viral_Load_Sample_Collection_Date['Quality_Issue'] = 'Viral Load Confirmation Date is earlier than Viral Load Sample Collection Date'
    all_line_lists_data.append(Viral_Load_Confirmation_Date_is_earlier_than_Viral_Load_Sample_Collection_Date)

    
    #Date of HIV Testing is greater than Viral Load Confirmation Date
    Date_Of_HIV_Testing_is_greater_than_Viral_Load_Confirmation_Date = project_data[
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] > project_data['Recency Viral Load Result Received Date (yyyy-mm-dd)']) & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['Date_Of_HIV_Testing_is_greater_than_Viral_Load_Confirmation_Date'] = Date_Of_HIV_Testing_is_greater_than_Viral_Load_Confirmation_Date.shape[0]
    Date_Of_HIV_Testing_is_greater_than_Viral_Load_Confirmation_Date['Quality_Issue'] = 'Date_Of_HIV_Testing_is_greater_than_Viral_Load_Confirmation_Date'
    all_line_lists_data.append(Date_Of_HIV_Testing_is_greater_than_Viral_Load_Confirmation_Date)


    #Date of HIV Testing is greater than Recency Test Date
    Date_Of_HIV_Testing_is_greater_than_Recency_Test_Date = project_data[
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] > project_data['Recency Test Date (yyyy-mm-dd)']) & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['Date_Of_HIV_Testing_is_greater_than_Recency_Test_Date'] = Date_Of_HIV_Testing_is_greater_than_Recency_Test_Date.shape[0]
    Date_Of_HIV_Testing_is_greater_than_Recency_Test_Date['Quality_Issue'] = 'Date of HIV Testing is greater than Recency Test Date'
    all_line_lists_data.append(Date_Of_HIV_Testing_is_greater_than_Recency_Test_Date)


    #Recency Interpretation equals RTRI Recent and Recency Viral Load Sample Collection Date or Recency Viral Load Confirmation Result or Recency Viral Load Result Received Date (yyyy-mm-dd) is blank
    Recency_Interpretation_equals_RTRI_Recent_and_VLSCDate_or_VLConfirmationDate_or_VLRRDate_is_blank = project_data[
        (project_data['Recency Interpretation'] == "RTRI Recent") &
        ((project_data['Recency Viral Load Sample Collection Date'].isna()) |(project_data['Recency Viral Load Confirmation Result'].isna()) | (project_data['Recency Viral Load Result Received Date (yyyy-mm-dd)'].isna())) & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['Recency_Interpretation_equals_RTRI_Recent_and_VLSCDate_or_VLConfirmationDate_or_VLRRDate_is_blank'] = Recency_Interpretation_equals_RTRI_Recent_and_VLSCDate_or_VLConfirmationDate_or_VLRRDate_is_blank.shape[0]
    Recency_Interpretation_equals_RTRI_Recent_and_VLSCDate_or_VLConfirmationDate_or_VLRRDate_is_blank['Quality_Issue'] = 'Recency Interpretation equals RTRI Recent and Recency Viral Load Sample Collection Date or Recency Viral Load Confirmation Result or Viral Load Confirmation Date is blank'
    all_line_lists_data.append(Recency_Interpretation_equals_RTRI_Recent_and_VLSCDate_or_VLConfirmationDate_or_VLRRDate_is_blank)


    #blank Prep Offered for Negative HIV Result
    blank_Prep_Offered_for_Negative_HIV_Result = project_data[
        (project_data['Prep Offered'].isna()) &
        (project_data['Final HIV Test Result'] =='Negative') &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_Prep_Offered_for_Negative_HIV_Result'] = blank_Prep_Offered_for_Negative_HIV_Result.shape[0]
    blank_Prep_Offered_for_Negative_HIV_Result['Quality_Issue'] = 'blank Prep Offered for Negative HIV Result'
    all_line_lists_data.append(blank_Prep_Offered_for_Negative_HIV_Result)

    #Prep Offered for Positive HIV Result
    Prep_Offered_for_Positive_HIV_Result = project_data[
        (project_data['Prep Offered'] =='Yes') &
        (project_data['Final HIV Test Result'] =='Positive') &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['Prep_Offered_for_Positive_HIV_Result'] = Prep_Offered_for_Positive_HIV_Result.shape[0]
    Prep_Offered_for_Positive_HIV_Result['Quality_Issue'] = 'Prep Offered for Positive HIV Result'
    all_line_lists_data.append(Prep_Offered_for_Positive_HIV_Result)

    #blank PrEP Accepted where PrEP Offered is Yes
    blank_PrEP_Accepted_where_PrEP_Offered_is_Yes = project_data[
        (project_data['Prep Offered'] == 'Yes') &
        (project_data['Prep Accepted'].isna()) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_PrEP_Accepted_where_PrEP_Offered_is_Yes'] = blank_PrEP_Accepted_where_PrEP_Offered_is_Yes.shape[0]
    blank_PrEP_Accepted_where_PrEP_Offered_is_Yes['Quality_Issue'] = 'blank PrEP Accepted where PrEP Offered is Yes'
    all_line_lists_data.append(blank_PrEP_Accepted_where_PrEP_Offered_is_Yes)

    #Previously Tested is No where First Time Visit is No
    Previously_Tested_is_No_where_First_Time_Visit_is_No = project_data[
        (project_data['First Time Visit'] == 'No') &
        (project_data['Previously Tested'] == 'No') &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['Previously_Tested_is_No_where_First_Time_Visit_is_No'] = Previously_Tested_is_No_where_First_Time_Visit_is_No.shape[0]
    Previously_Tested_is_No_where_First_Time_Visit_is_No['Quality_Issue'] = 'Previously Tested is No where First Time Visit is No'
    all_line_lists_data.append(Previously_Tested_is_No_where_First_Time_Visit_is_No)

    #Previously Tested is Yes where First Time Visit is Yes
    Previously_Tested_is_Yes_where_First_Time_Visit_is_Yes = project_data[
        (project_data['First Time Visit'] == 'Yes') &
        (project_data['Previously Tested'] == 'Yes') &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['Previously_Tested_is_Yes_where_First_Time_Visit_is_Yes'] = Previously_Tested_is_Yes_where_First_Time_Visit_is_Yes.shape[0]
    Previously_Tested_is_Yes_where_First_Time_Visit_is_Yes['Quality_Issue'] = 'Previously Tested is Yes where First Time Visit is Yes'
    all_line_lists_data.append(Previously_Tested_is_Yes_where_First_Time_Visit_is_Yes)

    #Recency Test Date less than Date of HIV Testing
    Recency_Test_Date_less_than_Date_of_HIV_Testing = project_data[
        (project_data['Recency Test Date (yyyy-mm-dd)'] < project_data['Date Of HIV Testing (yyyy-mm-dd)']) &
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['Recency_Test_Date_less_than_Date_of_HIV_Testing'] = Recency_Test_Date_less_than_Date_of_HIV_Testing.shape[0]
    Recency_Test_Date_less_than_Date_of_HIV_Testing['Quality_Issue'] = 'Recency Test Date less than Date of HIV Testing'
    all_line_lists_data.append(Recency_Test_Date_less_than_Date_of_HIV_Testing)

    #blank Pregnancy Status for childbearing Age women
    blank_Pregnancy_Status_for_childbearing_Age_women = project_data[
        (project_data['Age'] >=15) &
        (project_data['Sex'] == 'Female') &
        (project_data['Pregnancy/Breastfeeding Status'].isna()) & #Pregnancy/Breastfeeding Status
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['blank_Pregnancy_Status_for_childbearing_Age_women'] = blank_Pregnancy_Status_for_childbearing_Age_women.shape[0]
    blank_Pregnancy_Status_for_childbearing_Age_women['Quality_Issue'] = 'blank Pregnancy Status for childbearing Age women'
    all_line_lists_data.append(blank_Pregnancy_Status_for_childbearing_Age_women)


    #wrong modality for Congergational setting
    wrong_modality_for_Congergational_setting = project_data[
        #(project_data['Entry Point'] = "Community") &
        (project_data['Testing Setting'] == 'Congregational setting') &
        (project_data['Modality'] != "PMTCT (ANC1 Only)") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Congergational_setting'] = wrong_modality_for_Congergational_setting.shape[0]
    wrong_modality_for_Congergational_setting['Quality_Issue'] = 'wrong modality for Congergational setting'
    all_line_lists_data.append(wrong_modality_for_Congergational_setting)


    #wrong modality for Bloodblank setting
    wrong_modality_for_bloodbank_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'Blood Bank') &
        (project_data['Modality'] != "Other PITC") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_bloodbank_setting'] = wrong_modality_for_bloodbank_setting.shape[0]
    wrong_modality_for_Congergational_setting['Quality_Issue'] = 'wrong modality for Bloodbank testing setting'
    all_line_lists_data.append(wrong_modality_for_bloodbank_setting)


    #wrong modality for ANC setting
    wrong_modality_for_ANC_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'ANC') &
        (project_data['Modality'] != "PMTCT (ANC1 Only)") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_ANC_setting'] = wrong_modality_for_ANC_setting.shape[0]
    wrong_modality_for_ANC_setting['Quality_Issue'] = 'wrong modality for ANC testing setting'
    all_line_lists_data.append(wrong_modality_for_ANC_setting)

    #wrong modality for CT setting
    wrong_modality_for_CT_setting = project_data[
        #(project_data['Entry Point'].isin(["Facility", "Community"])) &
        (project_data['Testing Setting'] == 'CT') &
        (project_data['Modality'] != "VCT") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_CT_setting'] = wrong_modality_for_CT_setting.shape[0]
    wrong_modality_for_CT_setting['Quality_Issue'] = 'wrong modality for CT testing setting'
    all_line_lists_data.append(wrong_modality_for_CT_setting)


    #wrong modality for Delivery homes setting
    wrong_modality_for_Deliveryhomes_setting = project_data[
        #(project_data['Entry Point'] == "Community") &
        (project_data['Testing Setting'] == 'Delivery homes') &
        (project_data['Modality'] != "PMTCT (ANC1 Only)") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Deliveryhomes_setting'] = wrong_modality_for_Deliveryhomes_setting.shape[0]
    wrong_modality_for_Deliveryhomes_setting['Quality_Issue'] = 'wrong modality for Delivery homes testing setting'
    all_line_lists_data.append(wrong_modality_for_Deliveryhomes_setting)


    #wrong modality for Emergency setting
    wrong_modality_for_Emergency_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'Emergency') &
        (project_data['Modality'] != "Emergency") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Emergency_setting'] = wrong_modality_for_Emergency_setting.shape[0]
    wrong_modality_for_Emergency_setting['Quality_Issue'] = 'wrong modality for Emergency testing setting'
    all_line_lists_data.append(wrong_modality_for_Emergency_setting)


    #wrong modality for FP setting
    wrong_modality_for_FP_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'FP') &
        (project_data['Modality'] != "Other PITC") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_FP_setting'] = wrong_modality_for_FP_setting.shape[0]
    wrong_modality_for_FP_setting['Quality_Issue'] = 'wrong modality for FP testing setting'
    all_line_lists_data.append(wrong_modality_for_FP_setting)


    #wrong modality for Index setting
    wrong_modality_for_Index_setting = project_data[
        #(project_data['Entry Point'].isin(["Facility", "Community"])) &
        (project_data['Testing Setting'] == 'Index') &
        (project_data['Modality'] != "Index") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Index_setting'] = wrong_modality_for_Index_setting.shape[0]
    wrong_modality_for_Index_setting['Quality_Issue'] = 'wrong modality for Index testing setting'
    all_line_lists_data.append(wrong_modality_for_Index_setting)


    #wrong modality for Inpatient setting
    wrong_modality_for_Inpatient_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'Inpatient') &
        (project_data['Modality'] != "Inpatient") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Inpatient_setting'] = wrong_modality_for_Inpatient_setting.shape[0]
    wrong_modality_for_Inpatient_setting['Quality_Issue'] = 'wrong modality for Inpatient testing setting'
    all_line_lists_data.append(wrong_modality_for_Inpatient_setting)


    #wrong modality for L&D setting
    wrong_modality_for_Labour_and_Delivery_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'L&D') &
        (project_data['Modality'] != "PMTCT (Post ANC1: Pregnancy/L&D)") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Labour_and_Delivery_setting'] = wrong_modality_for_Labour_and_Delivery_setting.shape[0]
    wrong_modality_for_Labour_and_Delivery_setting['Quality_Issue'] = 'wrong modality for L&D testing setting'
    all_line_lists_data.append(wrong_modality_for_Labour_and_Delivery_setting)


    #wrong modality for Malnutrition setting
    wrong_modality_for_Malnutrition_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'Malnutrition') &
        (project_data['Modality'] != "Malnutrition") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Malnutrition_setting'] = wrong_modality_for_Malnutrition_setting.shape[0]
    wrong_modality_for_Malnutrition_setting['Quality_Issue'] = 'wrong modality for Malnutrition testing setting'
    all_line_lists_data.append(wrong_modality_for_Malnutrition_setting)


    #wrong modality for Others setting
    wrong_modality_for_Others_setting = project_data[
        #(project_data['Entry Point'].isin(["Facility", "Community"])) &
        (project_data['Testing Setting'] == 'Others') &
        (~project_data['Modality'].isin(["Other PITC", "Other Community Platforms"])) & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Others_setting'] = wrong_modality_for_Others_setting.shape[0]
    wrong_modality_for_Others_setting['Quality_Issue'] = 'wrong modality for Others testing setting'
    all_line_lists_data.append(wrong_modality_for_Others_setting)


    #wrong modality for Outreach setting
    wrong_modality_for_Outreach_setting = project_data[
        #(project_data['Entry Point'] == "Community") &
        (project_data['Testing Setting'] == 'Outreach') &
        (~project_data['Modality'].isin(["Mobile", "Other Community Platforms"])) & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Outreach_setting'] = wrong_modality_for_Outreach_setting.shape[0]
    wrong_modality_for_Outreach_setting['Quality_Issue'] = 'wrong modality for Outreach testing setting'
    all_line_lists_data.append(wrong_modality_for_Outreach_setting)


    #wrong modality for OVC setting
    wrong_modality_for_OVC_setting = project_data[
        #(project_data['Entry Point'] == "Community") &
        (project_data['Testing Setting'] == 'OVC') &
        (project_data['Modality'] != "Other Community Platforms") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_OVC_setting'] = wrong_modality_for_OVC_setting.shape[0]
    wrong_modality_for_OVC_setting['Quality_Issue'] = 'wrong modality for OVC testing setting'
    all_line_lists_data.append(wrong_modality_for_OVC_setting)


    #wrong modality for Pediatric setting
    wrong_modality_for_Pediatric_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'Pediatric') &
        (project_data['Modality'] != "Pediatric") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Pediatric_setting'] = wrong_modality_for_Pediatric_setting.shape[0]
    wrong_modality_for_Pediatric_setting['Quality_Issue'] = 'wrong modality for Pediatric testing setting'
    all_line_lists_data.append(wrong_modality_for_Pediatric_setting)


    #wrong modality for Post Natal Ward/Breastfeeding setting
    wrong_modality_for_Post_Natal_Ward_or_Breastfeeding_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'Post Natal Ward/Breastfeeding') &
        (project_data['Modality'] != "PMTCT (Post ANC1: Breastfeeding)") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Post_Natal_Ward_or_Breastfeeding_setting'] = wrong_modality_for_Post_Natal_Ward_or_Breastfeeding_setting.shape[0]
    wrong_modality_for_Post_Natal_Ward_or_Breastfeeding_setting['Quality_Issue'] = 'wrong modality for Post Natal Ward or Breastfeeding testing setting'
    all_line_lists_data.append(wrong_modality_for_Post_Natal_Ward_or_Breastfeeding_setting)


    #wrong modality for PrEP Testing setting
    wrong_modality_for_PrEP_Testing_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'PrEP Testing') &
        (project_data['Modality'] != "PrEP_CT HTS") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_PrEP_Testing_setting'] = wrong_modality_for_PrEP_Testing_setting.shape[0]
    wrong_modality_for_PrEP_Testing_setting['Quality_Issue'] = 'wrong modality for PrEP Testing testing setting'
    all_line_lists_data.append(wrong_modality_for_PrEP_Testing_setting)


    #wrong modality for Retesting setting
    wrong_modality_for_Retesting_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'Retesting') &
        (project_data['Modality'] != "PMTCT (Post ANC1: Pregnancy/L&D)") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Retesting_setting'] = wrong_modality_for_Retesting_setting.shape[0]
    wrong_modality_for_Retesting_setting['Quality_Issue'] = 'wrong modality for Retesting testing setting'
    all_line_lists_data.append(wrong_modality_for_Retesting_setting)


    #wrong modality for SNS setting
    wrong_modality_for_SNS_setting = project_data[
        #(project_data['Entry Point'].isin(["Facility", "Community"])) &
        (project_data['Testing Setting'] == 'SNS') &
        (project_data['Modality'] != "SNS") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_SNS_setting'] = wrong_modality_for_SNS_setting.shape[0]
    wrong_modality_for_SNS_setting['Quality_Issue'] = 'wrong modality for SNS testing setting'
    all_line_lists_data.append(wrong_modality_for_SNS_setting)


    #wrong modality for Spoke health facility setting
    wrong_modality_for_Spoke_health_facility_setting = project_data[
        #(project_data['Entry Point'].isin(["Facility", "Community"])) &
        (project_data['Testing Setting'] == 'Spoke health facility') &
        (project_data['Modality'] != "PMTCT (ANC1 Only)") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Spoke_health_facility_setting'] = wrong_modality_for_Spoke_health_facility_setting.shape[0]
    wrong_modality_for_Spoke_health_facility_setting['Quality_Issue'] = 'wrong modality for Spoke health facility testing setting'
    all_line_lists_data.append(wrong_modality_for_Spoke_health_facility_setting)


    #wrong modality for Standalone HTS setting
    wrong_modality_for_Standalone_HTS_setting = project_data[
        #(project_data['Entry Point'].isin(["Facility", "Community"])) &
        (project_data['Testing Setting'] == 'Standalone HTS') &
        (~project_data['Modality'].isin(["Other PITC", "Other Community Platforms"])) & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Standalone_HTS_setting'] = wrong_modality_for_Standalone_HTS_setting.shape[0]
    wrong_modality_for_Standalone_HTS_setting['Quality_Issue'] = 'wrong modality for Standalone HTS testing setting'
    all_line_lists_data.append(wrong_modality_for_Standalone_HTS_setting)


    #wrong modality for STI setting
    wrong_modality_for_STI_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'STI') &
        (project_data['Modality'] != "STI") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_STI_setting'] = wrong_modality_for_STI_setting.shape[0]
    wrong_modality_for_STI_setting['Quality_Issue'] = 'wrong modality for STI testing setting'
    all_line_lists_data.append(wrong_modality_for_STI_setting)


    #wrong modality for TB setting
    wrong_modality_for_TB_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'TB') &
        (project_data['Modality'] != "TB") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_TB_setting'] = wrong_modality_for_TB_setting.shape[0]
    wrong_modality_for_TB_setting['Quality_Issue'] = 'wrong modality for TB testing setting'
    all_line_lists_data.append(wrong_modality_for_TB_setting)


    #wrong modality for TBA Orthodox setting
    wrong_modality_for_TBA_Orthodox_setting = project_data[
        #(project_data['Entry Point'] == "Community") &
        (project_data['Testing Setting'].isin(['TBA Orthodox', 'TBA Orthodx'])) &
        (project_data['Modality'] != "PMTCT (ANC1 Only)") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_TBA_Orthodox_setting'] = wrong_modality_for_TBA_Orthodox_setting.shape[0]
    wrong_modality_for_TBA_Orthodox_setting['Quality_Issue'] = 'wrong modality for TBA Orthodox testing setting'
    all_line_lists_data.append(wrong_modality_for_TBA_Orthodox_setting)


    #wrong modality for TBA rt-HCW setting
    wrong_modality_for_TBA_rt_HCW_setting = project_data[
        #(project_data['Entry Point'] == "Community") &
        (project_data['Testing Setting'] == 'TBA rt-HCW') &
        (project_data['Modality'] != "PMTCT (ANC1 Only)") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_TBA_rt_HCW_setting'] = wrong_modality_for_TBA_rt_HCW_setting.shape[0]
    wrong_modality_for_TBA_rt_HCW_setting['Quality_Issue'] = 'wrong modality for TBA rt-HCW testing setting'
    all_line_lists_data.append(wrong_modality_for_TBA_rt_HCW_setting)


    #wrong modality for TB setting
    wrong_modality_for_Ward_or_Inpatient_setting = project_data[
        #(project_data['Entry Point'] == "Facility") &
        (project_data['Testing Setting'] == 'Ward/Inpatient') &
        (project_data['Modality'] != "Inpatient") & 
        (project_data['Date Of HIV Testing (yyyy-mm-dd)'] >= filter_date)
    ]
    project_issues['wrong_modality_for_Ward_or_Inpatient_setting'] = wrong_modality_for_Ward_or_Inpatient_setting.shape[0]
    wrong_modality_for_Ward_or_Inpatient_setting['Quality_Issue'] = 'wrong modality for Ward/Inpatient testing setting'
    all_line_lists_data.append(wrong_modality_for_Ward_or_Inpatient_setting)




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
