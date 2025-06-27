import os
import pandas as pd

# Set folder path
folder_path = "/Users/jamiepanchak/Desktop/portfolio/Education Datasets"

# File names
files = {
    'admissions': 'Admissions.csv',
    'finances': 'Finances.csv',
    'graduation': 'Graduation Rates.csv',
    'institution': 'Institutional Characteristics.csv',
    'retention': 'Retention Rate.csv',
    'facultyratio': 'Student Faculty Ratio.csv',
    'enrollment': 'UG Enrollment.csv'
}

# Load and clean column names
dfs = {}
for name, filename in files.items():
    df = pd.read_csv(os.path.join(folder_path, filename))
    df.columns = df.columns.str.strip().str.lower()  # Normalize column names
    dfs[name] = df

# Start with Institutional Characteristics as base
merged = dfs['institution']

# Merge other datasets, avoiding duplicate columns
for name, df in dfs.items():
    if name != 'institution':
        # Drop overlapping columns except 'unitid'
        df = df.loc[:, ~df.columns.isin(merged.columns.difference(['unitid']))]
        merged = pd.merge(merged, df, on='unitid', how='left')

# Drop duplicate rows by unitid
merged = merged.drop_duplicates(subset='unitid')

# Create output folder if it doesn't exist
output_folder = os.path.join(folder_path, 'data cleaned')
os.makedirs(output_folder, exist_ok=True)

# Save to Excel
output_path = os.path.join(output_folder, 'IPEDS_2023_Merged.xlsx')
merged.to_excel(output_path, index=False)

print(f"saved to: {output_path}")
