## Makro at 26/07/2023 -> First Draft
## Add Condition check with threshold -> 20/09/2023
## Add Calculate Vat = Total - Ext Vat -> 20/09/2023
## Add Matching with Invoice Date -> 28/09/2023

import pandas as pd
import os
import locale
from datetime import datetime as dt

## Check Current Path ##
current_directory = os.getcwd() 
print(current_directory)
# pd.set_option('display.expand_frame_repr', False)

## Only Deploy version ##
# CPFT_PATH = 'CPFT-master-data.xlsx'
Base_PATH = 'makro Invoice.xlsx'
B2B_PATH = 'B2B.csv'
# CPFM_PATH = 'CPFM.csv'

df_Base = pd.read_excel(Base_PATH,converters={'วันที่':str,'เลขที่เอกสาร':str,'รหัสลูกค้า':str,'รหัส Store':str,'Branch':str}, skiprows=2)
df_Base = df_Base.loc[df_Base['Branch'] != '0']
# df_CPFT_CPFM = df_CPFT.loc[df_CPFT['system'] == 'CPFM']

################################################################################################################
## Start B2B (Only CPF,Makro) ##
## Read and Filter only Makro ##
df_B2B = pd.read_csv(B2B_PATH, skiprows=[0])
df_B2B = df_B2B[df_B2B['rCV_name'].str.contains("ซีพี แอ็กซ์ตร้า")]
print(df_B2B.head())

## Rename columns ##
# df_Base = df_Base.rename(columns={'rRef_doc_number': 'PO'}) # Makro hasn't Tax then set Tax = null
df_Base = df_Base.rename(columns={'เลขที่เอกสาร': 'Invoice_No'})
df_Base = df_Base.rename(columns={'Branch': 'Store_No'})
df_Base = df_Base.rename(columns={'store name': 'Store_Name'})
df_Base = df_Base.rename(columns={'วันที่': 'Invoice_Date'})
df_Base = df_Base.rename(columns={'Invoice Price Ex. Vat': 'Exc_Vat'})
# df_Base = df_Base.rename(columns={'เลขที่เอกสาร': 'Tax'}) ## Makro hasn't Tax then set Tax = 0
df_Base = df_Base.rename(columns={'จำนวน': 'Total_Amt'})
df_Base['Tax'] = 0
df_Base['PO'] = ""

## Adjust Base Data
df_Base['Invoice_Date'] = pd.to_datetime(df_Base.Invoice_Date,format='%d/%m/%Y')
print(df_Base.head())

df_B2B = df_B2B.rename(columns={'rInput_doc_number': 'PO'})
df_B2B = df_B2B.rename(columns={'rRef_doc_number': 'Invoice_No'})
df_B2B = df_B2B.rename(columns={'rRef_doc_date_show': 'Invoice_Date'})
df_B2B = df_B2B.rename(columns={'rNett_amt_h': 'Total_Amt'})
df_B2B = df_B2B.rename(columns={'rTax_amt': 'Tax'})

## Adjust B2B Data
# locale.setlocale(locale.LC_TIME, "th_TH")
df_B2B['Invoice_Date'] = pd.to_datetime(df_B2B.Invoice_Date,format='%d/%m/%Y-543')


## Using drop() function to delete last row ##
df_Base.drop(index=df_Base.index[-1],axis=0,inplace=True)

## Convert data to string ##
df_Base['Invoice_No'] = df_Base['Invoice_No'].astype(str)
df_B2B['Invoice_No'] = df_B2B['Invoice_No'].astype(str)
df_B2B['rOperationCode'] = df_B2B['rOperationCode'].astype(str)
df_B2B['rDocNumber'] = df_B2B['rDocNumber'].astype(str)
# df_CPFT_B2B['PO'] = pd.to_numeric(df_CPFT_B2B['PO'], errors='coerce')
# df_B2B['PO'] = pd.to_numeric(df_B2B['PO'], errors='coerce')

## Check Leng of invoice number ##
for index, row in df_Base.iterrows():        
        if(len((row['Invoice_No'])) > 12):
                print(index,row['Invoice_No'], row['Invoice_No'])
                new_invoiceID = row['Invoice_No'][2:]
                print(new_invoiceID)
                df_Base.at[index,'Invoice_No'] = new_invoiceID
        
print(df_Base['Invoice_No'])

## Merge the dataframes with suffixes added to duplicate column names ##
df_merge = pd.merge(df_Base, df_B2B, on=['Invoice_No','Invoice_Date'], how='outer', suffixes=('_BASE', '_B2B'))
df_merge['Tax_BASE'] = round(df_merge['Total_Amt_BASE'] - df_merge['Exc_Vat'],2) ## Calculate Vat for makro
df_merge['Exc_Vat_difference'] = round(df_merge['Exc_Vat'] - df_merge['rSumNett'],2)
df_merge['Tax_difference'] = round(df_merge['Tax_BASE'] - df_merge['Tax_B2B'],2)
df_merge['Total_Amt_difference'] = round(df_merge['Total_Amt_BASE'] - df_merge['Total_Amt_B2B'],2)


print(df_merge.head())
print(df_merge.shape)
print(list(df_merge))

## create a boolean mask where both columns have values and the values are different ##
## Makro cannot match Diff Invoice ## 
# mask = (df_merge['Invoice_No'].notna()
        # &df_merge['rTrn'].notna())

## create a new dataframe with only the columns you specified and the rows where the mask is True ##
# diffINV_B2B = df_merge.loc[mask, ['rOperationCode', 'rDocNumber', 'Invoice_Number_CPFT', 'Invoice_Number_B2B']]

cols = ['rOperationCode','rDocNumber','Store_No', 'Store_Name', 'Invoice_No', 'Invoice_Date',
        'Exc_Vat', 'rSumNett', 'Exc_Vat_difference','Tax_BASE', 'Tax_B2B', 'Tax_difference', 'Total_Amt_BASE', 'Total_Amt_B2B', 'Total_Amt_difference',
        'PO_BASE']


df_merge = df_merge[cols]
df_merge.sort_values(by=['Total_Amt_difference'], ascending=True, inplace=True)

df_merge = df_merge.rename(columns={'Exc_Vat': 'Exc_Vat_BASE'})
df_merge = df_merge.rename(columns={'rSumNett': 'Exc_Vat_B2B'})

print(df_merge.head())
print(df_merge.shape)

## Create a new column with CPFT_Null or B2B_Null depending on the values of rTax_amt_CPFT and rTax_amt_B2B
df_merge['null_report'] = ''
df_merge.loc[df_merge['Total_Amt_BASE'].isnull(), 'null_report'] = 'BASE_Null'
df_merge.loc[df_merge['Total_Amt_B2B'].isnull(), 'null_report'] = 'B2B_Null'

## Count null
value_counts = df_merge['null_report'].value_counts()

## Create a dataframe to store the counts
counts_df = pd.DataFrame({'Type': ['BASE_Null', 'B2B_Null', 'Matching'],
                          'Count': [value_counts.get('BASE_Null', 0), value_counts.get('B2B_Null', 0),
                                    value_counts.get('', 0)]})

## Add the sum of the "Count" column to the last row
counts_df.loc["Total"] = ["Total", counts_df["Count"].sum()]

## Create null dataframe
df_CPFT_null = df_merge[df_merge['null_report'] == 'BASE_Null']
df_B2B_null = df_merge[df_merge['null_report'] == 'B2B_Null']

## Create dataframes only different values
## Set Threshold for reject invoice transactions (Now Setting = tax&vat not between -0.01 to 0.01)
df_B2B_diffrNett = df_merge[(df_merge['Total_Amt_difference'].notnull()) & ((df_merge['Total_Amt_difference'] < -0.01)|(df_merge['Total_Amt_difference'] > 0.01)) & (df_merge['null_report'] == '')]
df_B2B_diffrTax = df_merge[(df_merge['Tax_difference'].notnull()) & ((df_merge['Tax_difference'] < -0.01)|(df_merge['Tax_difference'] > 0.01)) & (df_merge['null_report'] == '')]
df_B2B_diff = pd.concat([df_B2B_diffrTax, df_B2B_diffrNett], ignore_index=True).drop_duplicates().reset_index(drop=True)

cols_select = ['Store_No', 'Store_Name', 'Invoice_No', 'Invoice_Date',
        'Exc_Vat_BASE', 'Exc_Vat_B2B', 'Exc_Vat_difference','Tax_BASE', 'Tax_B2B', 'Tax_difference', 'Total_Amt_BASE', 'Total_Amt_B2B', 'Total_Amt_difference',
        'PO_BASE']

df_merge_select = df_merge[cols_select]

df_B2B_diff_select = df_B2B_diff[['rOperationCode', 'rDocNumber', 'Invoice_No',
                                  'Tax_BASE', 'Tax_B2B', 'Tax_difference', 'Total_Amt_BASE',
                                  'Total_Amt_B2B', 'Total_Amt_difference']]

## Write the reconciled data_v1 and the counts to a single Excel file ##
with pd.ExcelWriter("reconciled_data_Makro.xlsx") as writer:
    df_merge_select.to_excel(writer, index=False, sheet_name='Reconciled Data')
    df_B2B_diff.to_excel(writer, index=False, startrow=0, sheet_name='B2B_diff')
    df_CPFT_null.to_excel(writer, index=False, startrow=0, sheet_name='BASE_null')
    df_B2B_null.to_excel(writer, index=False, startrow=0, sheet_name='B2B_null')
    counts_df.to_excel(writer, index=False, startrow=0, header=False, sheet_name='null_report')

# ################################################################################################################
# # Start CPFM
# # Rename columns
# # df_CPFM = pd.read_csv(CPFM_PATH, skiprows=[0])
# # df_CPFT_CPFM = df_CPFT_CPFM.rename(columns={'rRef_doc_number': 'PO'})
# # df_CPFT_CPFM = df_CPFT_CPFM.rename(columns={'rDoc_number': 'Invoice_Number'})
# # df_CPFM = df_CPFM.rename(columns={'rInput_doc_number': 'PO'})
# # df_CPFM = df_CPFM.rename(columns={'rRef_doc_number': 'Invoice_Number'})
# # df_CPFM = df_CPFM.rename(columns={'rNett_amt_h': 'rNett_amt'})

# # # Convert data to string
# # df_CPFT_CPFM['Invoice_Number'] = df_CPFT_CPFM['Invoice_Number'].astype(str)
# # df_CPFM['Invoice_Number'] = df_CPFM['Invoice_Number'].astype(str)
# # df_CPFT_CPFM['PO'] = pd.to_numeric(df_CPFT_CPFM['PO'], errors='coerce')
# # df_CPFM['PO'] = pd.to_numeric(df_CPFM['PO'], errors='coerce')

# # # Merge the dataframes with suffixes added to duplicate column names
# # df_merge = pd.merge(df_CPFT_CPFM, df_CPFM, on='PO', how='outer', suffixes=('_CPFT', '_CPFM'))
# # df_merge['rTax_amt_difference'] = df_merge['rTax_amt_CPFT'] - df_merge['rTax_amt_CPFM']
# # df_merge['rNett_amt_difference'] = df_merge['rNett_amt_CPFT'] - df_merge['rNett_amt_CPFM']

# # # create a boolean mask where both columns have values and the values are different
# # mask = (df_merge['Invoice_Number_CPFT'].notna()
        # # & df_merge['Invoice_Number_CPFM'].notna()
        # # & (df_merge['Invoice_Number_CPFT'] != df_merge['Invoice_Number_CPFM']))

# # # create a new dataframe with only the columns you specified and the rows where the mask is True
# # diffINV_CPFM = df_merge.loc[mask, ['rOperationCode', 'rDocNumber', 'Invoice_Number_CPFT', 'Invoice_Number_CPFM']]

# # cols = ['rDate_start', 'rDate_end', 'rOperation_code', 'rSub_operation',
        # # 'rWarehouse_code', 'rWarehouse_name', 'rDoc_date', 'rCV_code', 'rCV_name_CPFT', 'PO', 'rNett_amt_without',
        # # 'system', 'rOperationCode', 'rOperation_name', 'rDocNumber', 'rOrderDate', 'rCVCode',
        # # 'rCV_name_CPFM', 'rRef_doc_date_show', 'rSumNett', 'rDoc_type_name', 'rTrn', 'rTRN_name',
        # # 'Invoice_Number_CPFT', 'Invoice_Number_CPFM', 'rTax_amt_CPFT', 'rTax_amt_CPFM', 'rTax_amt_difference',
        # # 'rNett_amt_CPFT', 'rNett_amt_CPFM', 'rNett_amt_difference']

# # df_merge = df_merge[cols]
# # df_merge.sort_values(by=['rTax_amt_difference'], ascending=True, inplace=True)

# # # Create a new column with CPFT_Null or CPFM_Null depending on the values of rTax_amt_CPFT and rTax_amt_CPFM
# # df_merge['null_report'] = ''
# # df_merge.loc[df_merge['rTax_amt_CPFT'].isnull(), 'null_report'] = 'CPFT_Null'
# # df_merge.loc[df_merge['rTax_amt_CPFM'].isnull(), 'null_report'] = 'CPFM_Null'

# # # Count null
# # value_counts = df_merge['null_report'].value_counts()

# # # Create a dataframe to store the counts
# # counts_df = pd.DataFrame({'Type': ['CPFT_Null', 'CPFM_Null', 'Matching'],
                        # #   'Count': [value_counts.get('CPFT_Null', 0), value_counts.get('CPFM_Null', 0),
                                # #     value_counts.get('', 0)]})

# # # Add the sum of the "Count" column to the last row
# # counts_df.loc["Total"] = ["Total", counts_df["Count"].sum()]

# # # Create null dataframe
# # df_CPFT_null = df_merge[df_merge['null_report'] == 'CPFT_Null']
# # df_CPFM_null = df_merge[df_merge['null_report'] == 'CPFM_Null']

# # # Create dataframes only different values
# # df_CPFM_diffrNett = df_merge[(df_merge['rNett_amt_difference'].notnull()) & (df_merge['rNett_amt_difference'] != 0)]
# # df_CPFM_diffrTax = df_merge[(df_merge['rTax_amt_difference'].notnull()) & (df_merge['rTax_amt_difference'] != 0)]
# # df_CPFM_diff = pd.concat([df_CPFM_diffrTax, df_CPFM_diffrNett], ignore_index=True)

# # df_CPFM_diff_select = df_CPFM_diff[['rOperationCode', 'rDocNumber', 'Invoice_Number_CPFT', 'Invoice_Number_CPFM',
                                # #     'rTax_amt_CPFT', 'rTax_amt_CPFM', 'rTax_amt_difference', 'rNett_amt_CPFT',
                                # #     'rNett_amt_CPFM', 'rNett_amt_difference']]

# # Write the reconciled data_v1 and the counts to a single Excel file
# # with pd.ExcelWriter("reconciled_data_CPFM.xlsx") as writer:
# #     df_merge.to_excel(writer, index=False, sheet_name='Reconciled Data')
# #     df_CPFM_diff.to_excel(writer, index=False, startrow=0, sheet_name='CPFM_diff')
# #     df_CPFT_null.to_excel(writer, index=False, startrow=0, sheet_name='CPFT_null')
# #     df_CPFM_null.to_excel(writer, index=False, startrow=0, sheet_name='CPFM_null')
# #     counts_df.to_excel(writer, index=False, startrow=0, header=False, sheet_name='null_report')

# ################################################################################################################
# # df_combined_inv = pd.concat([diffINV_CPFM, diffINV_B2B], ignore_index=True)
# df_combined_inv = diffINV_B2B
# df_combined_inv['rOperationCode'] = pd.to_numeric(df_combined_inv['rOperationCode'], errors='coerce').astype('Int64')
# df_combined_inv['rDocNumber'] = pd.to_numeric(df_combined_inv['rDocNumber'], errors='coerce').astype('Int64')
# df_combined_inv['rDocNumber'] = df_combined_inv['rDocNumber'].astype(str).str.zfill(15)

# # df_combined_diff = pd.concat([df_CPFM_diff_select, df_B2B_diff_select], ignore_index=True)
# df_combined_diff = df_B2B_diff_select
# df_combined_diff['rOperationCode'] = pd.to_numeric(df_combined_diff['rOperationCode'], errors='coerce').astype('Int64')
# df_combined_diff['rDocNumber'] = pd.to_numeric(df_combined_diff['rDocNumber'], errors='coerce').astype('Int64')
# df_combined_diff['rDocNumber'] = df_combined_diff['rDocNumber'].astype(str).str.zfill(15)

# # df_combined = pd.concat([df_CPFM_diff_select, df_B2B_diff_select, diffINV_CPFM, diffINV_B2B], ignore_index=True)
# df_combined = pd.concat([df_B2B_diff_select, diffINV_B2B], ignore_index=True)
# df_combined['rOperationCode'] = pd.to_numeric(df_combined['rOperationCode'], errors='coerce').astype('Int64')
# df_combined['rDocNumber'] = pd.to_numeric(df_combined['rDocNumber'], errors='coerce').astype('Int64')
# df_combined['rDocNumber'] = df_combined['rDocNumber'].astype(str).str.zfill(15)

# print("B2B_Diff: ", df_B2B_diff.shape[0])
# # print("CPFM_Diff: ", df_CPFM_diff.shape[0])
# # Print the removed rows
# print("COMBINED_Diff: ", df_combined.shape[0])

# # # find the duplicate rows based on 'rDocNumber'
# # duplicates = df_combined[df_combined.duplicated(subset='rDocNumber', keep='first')]
# # # remove the duplicate rows from the original DataFrame
# # df_combined = df_combined.drop_duplicates(subset='rDocNumber', keep='first')
# # print(f"Removed {len(duplicates)} rows with duplicate 'rDocNumber' values")
# # print("COMBINED_Diff_noDup: ", df_combined.shape[0])
# # # print(f"Removed {len(duplicates)} rows with duplicate 'rDocNumber' values:\n{duplicates}")

# df_combined_inv.to_csv('cpfm_b2b_invdiff.csv', index=False)
df_B2B_diff_select.to_csv('cpfm_b2b_diff_Makro.csv', index=False)
# df_combined.to_csv('combined_diff.csv', index=False)
