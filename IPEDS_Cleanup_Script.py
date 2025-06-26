import pandas as pd
import re
import os

def load_data(path):
    """Load merged Excel file into DataFrame."""
    return pd.read_excel(path)

def drop_unwanted_columns(df):
    """Drop columns like duplicate institution names and internal unitid."""
    drop_cols = [col for col in df.columns if 'institution name.1' in col or 'hd2023.unitid' in col]
    return df.drop(columns=drop_cols, errors='ignore')

def filter_uc_schools(df):
    """Filter for UC campuses and exclude UCSF and UC Law SF."""
    uc_df = df[df['institution name'].str.contains('University of California', case=False, na=False)]
    exclude = [
        'University of California College of the Law-San Francisco',
        'University of California-San Francisco'
    ]
    return uc_df[~uc_df['institution name'].isin(exclude)]

def drop_duplicate_graduation_columns(df):
    """Keep only the first graduation rate column."""
    grad_cols = [col for col in df.columns if "graduation rate" in col.lower()]
    return df.drop(columns=grad_cols[1:], errors='ignore')

def convert_percent_fields(df):
    """Convert percent-like fields (rates) from strings to float."""
    percent_cols = [col for col in df.columns if 'rate' in col.lower()]
    for col in percent_cols:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.replace('%', '').str.strip()
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def add_derived_columns(df):
    """Add calculated fields like admit_rate and place it after admissions_total."""
    if {'admissions_total', 'applicants_total'}.issubset(df.columns):
        df['admit_rate'] = df['admissions_total'] / df['applicants_total']

        # Move admit_rate after admissions_total
        cols = df.columns.tolist()
        if 'admit_rate' in cols and 'admissions_total' in cols:
            cols.insert(cols.index('admissions_total') + 1, cols.pop(cols.index('admit_rate')))
            df = df[cols]

    return df

def drop_sector_column(df):
    """Drop any column related to sector of institution."""
    sector_cols = [col for col in df.columns if 'sector' in col.lower()]
    return df.drop(columns=sector_cols, errors='ignore')

def standardize_column_names(df):
    """Convert column names to snake_case and strip prefixes."""
    def clean_col(col):
        col = col.lower()
        col = re.sub(r'adm2023\.|drvf2023\.|drvgr2023\.|ef2023d\.|drvef2023\.', '', col)
        col = re.sub(r'\(gasb\)', '', col, flags=re.IGNORECASE)
        col = col.replace('\n', ' ').replace('%', 'percent').replace('-', '')
        col = col.replace(',', '').replace('(', '').replace(')', '')
        col = col.replace('/', '_').strip()
        col = re.sub(r'\s+', '_', col)
        return col

    df.columns = [clean_col(col) for col in df.columns]
    return df

def save_cleaned_data(df, output_path):
    """Save cleaned DataFrame to Excel."""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df.to_excel(output_path, index=False)

def main():
    input_path = "/Users/jamiepanchak/Desktop/portfolio/Education Datasets/data cleaned/IPEDS_2023_Merged.xlsx"
    output_path = "/Users/jamiepanchak/Desktop/portfolio/Education Datasets/data cleaned/IPEDS_2023_Final_Clean.xlsx"

    df = load_data(input_path)
    df = drop_unwanted_columns(df)
    df = filter_uc_schools(df)
    df = drop_duplicate_graduation_columns(df)
    df = convert_percent_fields(df)
    df = drop_sector_column(df)
    df = standardize_column_names(df)
    df = add_derived_columns(df)
    save_cleaned_data(df, output_path)

if __name__ == "__main__":
    main()
