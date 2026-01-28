import pandas as pd
import re
import os

def load_data(file_path):
    """Loads data from CSV or Excel, handling encoding issues."""
    ext = os.path.splitext(file_path)[-1].lower()
    if ext in ['.xlsx', '.xls']:
        return pd.read_excel(file_path)
    elif ext == '.csv':
        encodings = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252']
        for enc in encodings:
            try:
                return pd.read_csv(file_path, encoding=enc)
            except (UnicodeDecodeError, TypeError):
                continue
        raise ValueError(f"Could not decode CSV file {file_path}")
    return None

def extract_age_sex_clinical_safe(val):
    """Extracts Age and Sex, ignoring clinical markers like CD4 or Viral Load."""
    if pd.isna(val) or str(val).lower() == 'default':
        return "", ""
    
    parts = [p.strip() for p in str(val).split(',')]
    sex, age = "", ""
    
    sex_terms = ['male', 'female']
    # Terms that should NEVER be treated as 'Age' even if they contain numbers
    ignore_terms = [
        'positive', 'negative', 'status', 'cd4', 'viral load', 'vl', 
        'suppressed', 'unsuppressed', 'arv dispensing quantity', 'died', 'outcome'
    ]

    for part in parts:
        p_lower = part.lower()
        
        # 1. Identify Sex
        if p_lower in sex_terms:
            sex = part
            continue
            
        # 2. Identify Age (contains digits/keywords and NOT in ignore list)
        has_age_indicator = re.search(r'\d', part) or any(k in p_lower for k in ['month', 'year', 'age'])
        is_clinical_term = any(it in p_lower for it in ignore_terms)

        if has_age_indicator and not is_clinical_term:
            age = part
                
    return age, sex

def update_indicator_names(row, cat_col):
    """Updates indicator names based on categoryoptioncombo content."""
    indicator = str(row['indicator'])
    cat_val = str(row[cat_col]).lower()


    # Logic for "Newly Tested Positive" indicators
    if "newly tested positives" in cat_val:
        if indicator.startswith("TB_STAT"):
            return "TB_STAT_POS_Newly_Tested"
        if indicator.startswith("PMTCT_STAT"):
            return "PMTCT_STAT_POS_Newly_Tested"
        if indicator.startswith("HTS_INDEX"):
            return "HTS_INDEX_POS_Newly_Tested"
        
    # Logic for "Known Positive" indicators
    if "known positive" in cat_val:
        if indicator.startswith("TB_STAT"):
            return "TB_STAT_POS_Known_positive"
        if indicator.startswith("PMTCT_STAT"):
            return "PMTCT_STAT_POS_Known_positive"
        if indicator.startswith("HTS_INDEX"):
            return "HTS_INDEX_POS_Known_positive"
    
    # Logic for "Positive" indicators
    if "positive" in cat_val:
        if indicator.startswith("HTS_TST"):
            return "HTS_TST_POS"
        if indicator.startswith("HTS_INDEX"):
            return "HTS_INDEX_POS"
        if indicator.startswith("TB_STAT"):
            return "TB_STAT_POS"
        if indicator.startswith("PMTCT_STAT"):
            return "PMTCT_STAT_POS"
    
    # Logic for "New Positive" indicators
    if "new" in cat_val:
        if indicator.startswith("PMTCT_ART"):
            return "PMTCT_ART_New_positive"
    
    # Logic for "Known Positive" indicators
    if "already" in cat_val:
        if indicator.startswith("PMTCT_ART"):
            return "PMTCT_ART_Known_positive"
    
    
            
    # Logic for "TX_ML Died"
    if indicator.startswith("TX_ML") and "died" in cat_val:
        return "TX_ML_Died"
    # Logic for "TX_ML Stopped"
    if indicator.startswith("TX_ML") and "stopped" in cat_val:
        return "TX_ML_Stopped_Treatment"
    # Logic for "TX_ML IIT"
    if indicator.startswith("TX_ML") and "interruption" in cat_val:
        return "TX_ML_IIT"
    # Logic for "TX_ML Transferred_out"
    if indicator.startswith("TX_ML") and "transfer" in cat_val:
        return "TX_ML_Transferred_out"
        
    return indicator

def transpose_hiv_data(input_path, output_path):
    # Load data
    df = load_data(input_path)
    if df is None:
        return

    # Identify descriptive column (handles duplicate headers from CSV exports)
    cat_col = 'categoryoptioncombo.1' if 'categoryoptioncombo.1' in df.columns else 'categoryoptioncombo'
    
    # Update indicator names based on clinical status in categoryoptioncombo
    # This creates the specific HTS_TST_POS, TX_ML_Died, etc. columns during pivoting
    df['indicator'] = df.apply(lambda row: update_indicator_names(row, cat_col), axis=1)
    
    # Extract clean Age and Sex
    extracted = df[cat_col].apply(lambda x: pd.Series(extract_age_sex_clinical_safe(x)))
    df['Age'] = extracted[0]
    df['Sex'] = extracted[1]
    
    # Handle casing variations for Orgunit
    if 'Orgunit' in df.columns and 'orgUnit' not in df.columns:
        df.rename(columns={'Orgunit': 'orgUnit'}, inplace=True)

    # Pivot dimensions
    index_cols = ['IP', 'orgUnit', 'Period', 'Age', 'Sex']
    
    # Create the transposed table
    pivot_df = df.pivot_table(
        index=index_cols,
        columns='indicator',
        values='Value',
        aggfunc='sum'
    ).reset_index()
    
    pivot_df.columns.name = None
    
    # Save output (CSV or Excel)
    if output_path.endswith('.xlsx'):
        pivot_df.to_excel(output_path, index=False)
    else:
        pivot_df.to_csv(output_path, index=False)
    
    print(f"File saved successfully to: {output_path}")

# Run the transformation
input_file = 'C:/Users/DELL/Documents/DataFi/CS Flatfile Aggregation/FY26Q1_Flatfile with IP.xlsx' #- Sheet1.csv'
output_file = 'C:/Users/DELL/Documents/DataFi/CS Flatfile Aggregation/transposed/transposed_flexible_outputt_final.csv'
transpose_hiv_data(input_file, output_file)