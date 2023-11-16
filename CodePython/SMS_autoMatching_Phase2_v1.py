# Chang release at 23/05/2023
# 1. Accept column name from "rSub_operation_name" to "rOperation_code" and "System" to "system"
#    from CPFT-master-data.xlsx
# 2. User "get" method retrieves the value for a given key if it exists in line 71 and 147

import pandas as pd

# pd.set_option('display.expand_frame_repr', False)

# Only Deploy version
# CPFT_PATH = 'CPFT-master-data.xlsx'
Base_PATH = 'makro Invoice 01-17 July 2023.xlsx'
B2B_PATH = 'B2B.csv'
CPFM_PATH = 'CPFM.csv'

df_CPFT = pd.read_excel(Base_PATH, skiprows=[0])
df_CPFT_B2B = df_CPFT.loc[df_CPFT['system'] == 'B2B']
df_CPFT_CPFM = df_CPFT.loc[df_CPFT['system'] == 'CPFM']

################################################################################################################
# Start B2B (Only CPF,Makro)
# Rename columns
df_B2B = pd.read_csv(B2B_PATH, skiprows=[0])
df_CPFT_B2B = df_CPFT_B2B.rename(columns={'rRef_doc_number': 'PO'})
df_CPFT_B2B = df_CPFT_B2B.rename(columns={'rDoc_number': 'Invoice_Number'})
df_B2B = df_B2B.rename(columns={'rInput_doc_number': 'PO'})
df_B2B = df_B2B.rename(columns={'rRef_doc_number': 'Invoice_Number'})
df_B2B = df_B2B.rename(columns={'rNett_amt_h': 'rNett_amt'})

# Convert data to string
df_CPFT_B2B['Invoice_Number'] = df_CPFT_B2B['Invoice_Number'].astype(str)
df_B2B['Invoice_Number'] = df_B2B['Invoice_Number'].astype(str)
df_CPFT_B2B['PO'] = pd.to_numeric(df_CPFT_B2B['PO'], errors='coerce')
df_B2B['PO'] = pd.to_numeric(df_B2B['PO'], errors='coerce')

# Merge the dataframes with suffixes added to duplicate column names
df_merge = pd.merge(df_CPFT_B2B, df_B2B, on=['PO','Invoice_Number'], how='outer', suffixes=('_CPFT', '_B2B'))
df_merge['rTax_amt_difference'] = df_merge['rTax_amt_CPFT'] - df_merge['rTax_amt_B2B']
df_merge['rNett_amt_difference'] = df_merge['rNett_amt_CPFT'] - df_merge['rNett_amt_B2B']

# create a boolean mask where both columns have values and the values are different
mask = (df_merge['Invoice_Number_CPFT'].notna()
        & df_merge['Invoice_Number_B2B'].notna()
        & (df_merge['Invoice_Number_CPFT'] != df_merge['Invoice_Number_B2B']))

# create a new dataframe with only the columns you specified and the rows where the mask is True
diffINV_B2B = df_merge.loc[mask, ['rOperationCode', 'rDocNumber', 'Invoice_Number_CPFT', 'Invoice_Number_B2B']]

cols = ['rDate_start', 'rDate_end', 'rOperation_code', 'rSub_operation',
        'rWarehouse_code', 'rWarehouse_name', 'rDoc_date', 'rCV_code', 'rCV_name_CPFT', 'PO', 'rNett_amt_without',
        'system', 'rOperationCode', 'rOperation_name', 'rDocNumber', 'rOrderDate', 'rCVCode',
        'rCV_name_B2B', 'rRef_doc_date_show', 'rSumNett', 'rDoc_type_name', 'rTrn', 'rTRN_name',
        'Invoice_Number_CPFT', 'Invoice_Number_B2B', 'rTax_amt_CPFT', 'rTax_amt_B2B', 'rTax_amt_difference',
        'rNett_amt_CPFT', 'rNett_amt_B2B', 'rNett_amt_difference']

df_merge = df_merge[cols]
df_merge.sort_values(by=['rTax_amt_difference'], ascending=True, inplace=True)

# Create a new column with CPFT_Null or B2B_Null depending on the values of rTax_amt_CPFT and rTax_amt_B2B
df_merge['null_report'] = ''
df_merge.loc[df_merge['rTax_amt_CPFT'].isnull(), 'null_report'] = 'CPFT_Null'
df_merge.loc[df_merge['rTax_amt_B2B'].isnull(), 'null_report'] = 'B2B_Null'

# Count null
value_counts = df_merge['null_report'].value_counts()

# Create a dataframe to store the counts
counts_df = pd.DataFrame({'Type': ['CPFT_Null', 'B2B_Null', 'Matching'],
                          'Count': [value_counts.get('CPFT_Null', 0), value_counts.get('B2B_Null', 0),
                                    value_counts.get('', 0)]})

# Add the sum of the "Count" column to the last row
counts_df.loc["Total"] = ["Total", counts_df["Count"].sum()]

# Create null dataframe
df_CPFT_null = df_merge[df_merge['null_report'] == 'CPFT_Null']
df_B2B_null = df_merge[df_merge['null_report'] == 'B2B_Null']

# Create dataframes only different values
df_B2B_diffrNett = df_merge[(df_merge['rNett_amt_difference'].notnull()) & (df_merge['rNett_amt_difference'] != 0)]
df_B2B_diffrTax = df_merge[(df_merge['rTax_amt_difference'].notnull()) & (df_merge['rTax_amt_difference'] != 0)]
df_B2B_diff = pd.concat([df_B2B_diffrTax, df_B2B_diffrNett], ignore_index=True)

df_B2B_diff_select = df_B2B_diff[['rOperationCode', 'rDocNumber', 'Invoice_Number_CPFT', 'Invoice_Number_B2B',
                                  'rTax_amt_CPFT', 'rTax_amt_B2B', 'rTax_amt_difference', 'rNett_amt_CPFT',
                                  'rNett_amt_B2B', 'rNett_amt_difference']]

# Write the reconciled data_v1 and the counts to a single Excel file
# with pd.ExcelWriter("reconciled_data_B2B.xlsx") as writer:
#     df_merge.to_excel(writer, index=False, sheet_name='Reconciled Data')
#     df_B2B_diff.to_excel(writer, index=False, startrow=0, sheet_name='B2B_diff')
#     df_CPFT_null.to_excel(writer, index=False, startrow=0, sheet_name='CPFT_null')
#     df_B2B_null.to_excel(writer, index=False, startrow=0, sheet_name='B2B_null')
#     counts_df.to_excel(writer, index=False, startrow=0, header=False, sheet_name='null_report')

################################################################################################################
# Start CPFM
# Rename columns
df_CPFM = pd.read_csv(CPFM_PATH, skiprows=[0])
df_CPFT_CPFM = df_CPFT_CPFM.rename(columns={'rRef_doc_number': 'PO'})
df_CPFT_CPFM = df_CPFT_CPFM.rename(columns={'rDoc_number': 'Invoice_Number'})
df_CPFM = df_CPFM.rename(columns={'rInput_doc_number': 'PO'})
df_CPFM = df_CPFM.rename(columns={'rRef_doc_number': 'Invoice_Number'})
df_CPFM = df_CPFM.rename(columns={'rNett_amt_h': 'rNett_amt'})

# Convert data to string
df_CPFT_CPFM['Invoice_Number'] = df_CPFT_CPFM['Invoice_Number'].astype(str)
df_CPFM['Invoice_Number'] = df_CPFM['Invoice_Number'].astype(str)
df_CPFT_CPFM['PO'] = pd.to_numeric(df_CPFT_CPFM['PO'], errors='coerce')
df_CPFM['PO'] = pd.to_numeric(df_CPFM['PO'], errors='coerce')

# Merge the dataframes with suffixes added to duplicate column names
df_merge = pd.merge(df_CPFT_CPFM, df_CPFM, on='PO', how='outer', suffixes=('_CPFT', '_CPFM'))
df_merge['rTax_amt_difference'] = df_merge['rTax_amt_CPFT'] - df_merge['rTax_amt_CPFM']
df_merge['rNett_amt_difference'] = df_merge['rNett_amt_CPFT'] - df_merge['rNett_amt_CPFM']

# create a boolean mask where both columns have values and the values are different
mask = (df_merge['Invoice_Number_CPFT'].notna()
        & df_merge['Invoice_Number_CPFM'].notna()
        & (df_merge['Invoice_Number_CPFT'] != df_merge['Invoice_Number_CPFM']))

# create a new dataframe with only the columns you specified and the rows where the mask is True
diffINV_CPFM = df_merge.loc[mask, ['rOperationCode', 'rDocNumber', 'Invoice_Number_CPFT', 'Invoice_Number_CPFM']]

cols = ['rDate_start', 'rDate_end', 'rOperation_code', 'rSub_operation',
        'rWarehouse_code', 'rWarehouse_name', 'rDoc_date', 'rCV_code', 'rCV_name_CPFT', 'PO', 'rNett_amt_without',
        'system', 'rOperationCode', 'rOperation_name', 'rDocNumber', 'rOrderDate', 'rCVCode',
        'rCV_name_CPFM', 'rRef_doc_date_show', 'rSumNett', 'rDoc_type_name', 'rTrn', 'rTRN_name',
        'Invoice_Number_CPFT', 'Invoice_Number_CPFM', 'rTax_amt_CPFT', 'rTax_amt_CPFM', 'rTax_amt_difference',
        'rNett_amt_CPFT', 'rNett_amt_CPFM', 'rNett_amt_difference']

df_merge = df_merge[cols]
df_merge.sort_values(by=['rTax_amt_difference'], ascending=True, inplace=True)

# Create a new column with CPFT_Null or CPFM_Null depending on the values of rTax_amt_CPFT and rTax_amt_CPFM
df_merge['null_report'] = ''
df_merge.loc[df_merge['rTax_amt_CPFT'].isnull(), 'null_report'] = 'CPFT_Null'
df_merge.loc[df_merge['rTax_amt_CPFM'].isnull(), 'null_report'] = 'CPFM_Null'

# Count null
value_counts = df_merge['null_report'].value_counts()

# Create a dataframe to store the counts
counts_df = pd.DataFrame({'Type': ['CPFT_Null', 'CPFM_Null', 'Matching'],
                          'Count': [value_counts.get('CPFT_Null', 0), value_counts.get('CPFM_Null', 0),
                                    value_counts.get('', 0)]})

# Add the sum of the "Count" column to the last row
counts_df.loc["Total"] = ["Total", counts_df["Count"].sum()]

# Create null dataframe
df_CPFT_null = df_merge[df_merge['null_report'] == 'CPFT_Null']
df_CPFM_null = df_merge[df_merge['null_report'] == 'CPFM_Null']

# Create dataframes only different values
df_CPFM_diffrNett = df_merge[(df_merge['rNett_amt_difference'].notnull()) & (df_merge['rNett_amt_difference'] != 0)]
df_CPFM_diffrTax = df_merge[(df_merge['rTax_amt_difference'].notnull()) & (df_merge['rTax_amt_difference'] != 0)]
df_CPFM_diff = pd.concat([df_CPFM_diffrTax, df_CPFM_diffrNett], ignore_index=True)

df_CPFM_diff_select = df_CPFM_diff[['rOperationCode', 'rDocNumber', 'Invoice_Number_CPFT', 'Invoice_Number_CPFM',
                                    'rTax_amt_CPFT', 'rTax_amt_CPFM', 'rTax_amt_difference', 'rNett_amt_CPFT',
                                    'rNett_amt_CPFM', 'rNett_amt_difference']]

# Write the reconciled data_v1 and the counts to a single Excel file
# with pd.ExcelWriter("reconciled_data_CPFM.xlsx") as writer:
#     df_merge.to_excel(writer, index=False, sheet_name='Reconciled Data')
#     df_CPFM_diff.to_excel(writer, index=False, startrow=0, sheet_name='CPFM_diff')
#     df_CPFT_null.to_excel(writer, index=False, startrow=0, sheet_name='CPFT_null')
#     df_CPFM_null.to_excel(writer, index=False, startrow=0, sheet_name='CPFM_null')
#     counts_df.to_excel(writer, index=False, startrow=0, header=False, sheet_name='null_report')

################################################################################################################
df_combined_inv = pd.concat([diffINV_CPFM, diffINV_B2B], ignore_index=True)
df_combined_inv['rOperationCode'] = pd.to_numeric(df_combined_inv['rOperationCode'], errors='coerce').astype('Int64')
df_combined_inv['rDocNumber'] = pd.to_numeric(df_combined_inv['rDocNumber'], errors='coerce').astype('Int64')
df_combined_inv['rDocNumber'] = df_combined_inv['rDocNumber'].astype(str).str.zfill(15)

df_combined_diff = pd.concat([df_CPFM_diff_select, df_B2B_diff_select], ignore_index=True)
df_combined_diff['rOperationCode'] = pd.to_numeric(df_combined_diff['rOperationCode'], errors='coerce').astype('Int64')
df_combined_diff['rDocNumber'] = pd.to_numeric(df_combined_diff['rDocNumber'], errors='coerce').astype('Int64')
df_combined_diff['rDocNumber'] = df_combined_diff['rDocNumber'].astype(str).str.zfill(15)

df_combined = pd.concat([df_CPFM_diff_select, df_B2B_diff_select, diffINV_CPFM, diffINV_B2B], ignore_index=True)
df_combined['rOperationCode'] = pd.to_numeric(df_combined['rOperationCode'], errors='coerce').astype('Int64')
df_combined['rDocNumber'] = pd.to_numeric(df_combined['rDocNumber'], errors='coerce').astype('Int64')
df_combined['rDocNumber'] = df_combined['rDocNumber'].astype(str).str.zfill(15)

print("B2B_Diff: ", df_B2B_diff.shape[0])
print("CPFM_Diff: ", df_CPFM_diff.shape[0])
# Print the removed rows
print("COMBINED_Diff: ", df_combined.shape[0])

# # find the duplicate rows based on 'rDocNumber'
# duplicates = df_combined[df_combined.duplicated(subset='rDocNumber', keep='first')]
# # remove the duplicate rows from the original DataFrame
# df_combined = df_combined.drop_duplicates(subset='rDocNumber', keep='first')
# print(f"Removed {len(duplicates)} rows with duplicate 'rDocNumber' values")
# print("COMBINED_Diff_noDup: ", df_combined.shape[0])
# # print(f"Removed {len(duplicates)} rows with duplicate 'rDocNumber' values:\n{duplicates}")

df_combined_inv.to_csv('cpfm_b2b_invdiff.csv', index=False)
df_combined_diff.to_csv('cpfm_b2b_diff.csv', index=False)
df_combined.to_csv('combined_diff.csv', index=False)
