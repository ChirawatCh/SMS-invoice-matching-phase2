import pandas as pd
import os

# Folder path where CSV files are stored
folder_path = 'outputs'

# Define the specific CSV file names
csv_files = [
    'cpfm_b2b_diff_CPAxtra.csv',
    'cpfm_b2b_diff_CPRam.csv',
    'cpfm_b2b_diff_PerfectCompanion.csv',
    'cpfm_b2b_diff_President.csv',
    'cpfm_b2b_diff_ThaiNamThip.csv'
]

# Mapping of file names to vendor names
vendor_mapping = {
    'cpfm_b2b_diff_CPAxtra.csv': 'CP Axtra',
    'cpfm_b2b_diff_CPRam.csv': 'CP Ram',
    'cpfm_b2b_diff_PerfectCompanion.csv': 'Perfect Companion',
    'cpfm_b2b_diff_President.csv': 'President',
    'cpfm_b2b_diff_ThaiNamThip.csv': 'ThaiNamThip'
}

# Read each specific CSV file into a DataFrame and store them in a list
dfs = []
vendors_with_data = set()  # To track vendors with data

for file in csv_files:
    file_path = os.path.join(folder_path, file)
    if os.path.exists(file_path):
        df = pd.read_csv(file_path, dtype={"rDocNumber": str})
        df["rDocNumber"] = df["rDocNumber"].astype(str).str.zfill(15)
        if not df.empty:  # Check if DataFrame is empty
            vendors_with_data.add(vendor_mapping[file])
            df['vendors'] = vendor_mapping[file]  # Map file name to vendor name
            dfs.append(df)
        else:
            print()
            print(f"File {file} is empty.")
    else:
        print()
        print(f"File {file} not found.")

# Concatenate all DataFrames in the list along rows (axis=0)
result_df = pd.concat(dfs, ignore_index=True)

# Assuming 'result_df' is your DataFrame and these are the columns in the desired order
desired_order = ['vendors', 'rDocNumber', "rOperationCode", 'INV.no Check', 'Inv. Date Check', 'ExcludeVAT_diff', 'VAT_diff', 'IncludeVAT_diff']
# Rearrange columns in the DataFrame
result_df = result_df[desired_order]

# Fill NaN values in the DataFrame with 'N/A'
result_df.fillna('N/A', inplace=True)

# Count number of rows for each vendor
vendor_counts = result_df['vendors'].value_counts()

# Include vendors with no data
for vendor in vendor_mapping.values():
    if vendor not in vendors_with_data:
        vendor_counts[vendor] = 0

# Print the number of rows for each vendor
print()
print("Number of rows for each vendor (diff):")
print(vendor_counts)

# Write the concatenated DataFrame to a new CSV file
output_file = 'concatenated_output.csv'
output_path = os.path.join(folder_path, output_file)
result_df.to_csv(output_path, index=False)

print(f"Concatenated CSV file saved to: {output_path}")
print()
