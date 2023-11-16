## President Bakery at 26/07/2023 -> First Draft
## Add Condition check with threshold -> 22/09/2023
## Change encoding of TXT files to "cp874" -> 17/10/2023

import pandas as pd
import os
import csv
import glob
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta

## Function Read Text file ##
def parse_txt_file_and_save_to_csv(file_path, output_csv):
    # Create a list of headers in the order specified by the dictionary keys
    column_headers = ['Blank','Header','Branch','InvoiceDate','InvoiceNo','Total','Vat','ExcVat','Code1','Code2','Code3','Name','Code4','Province','ConpanyCode','BranchName','Code5','Code6','Code7']

    with open(file_path, 'r', encoding='cp874') as file:
        for line in file:
            # Extract data for the header record and create a dictionary
            dataSpilt = line.split('|')             
        #     dS = list(map(lambda x: x.replace('"', ''), dataSpilt))
        #     header_data = dict(zip(column_headers,dS))
            header_data = dict(zip(column_headers,(dS.replace('"','') for dS in dataSpilt) ))
            if(dataSpilt[1] == 'H'):
                result.append(header_data)

    # Write the extracted header data to a CSV file
    with open(output_csv, 'w', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=column_headers)
        writer.writeheader()
        writer.writerows(result)

## Read Input Data ##
# txt_data = "President Invoice.txt"
path = r'*.TXT'
files = glob.glob(path)
 # Lists to store data for each column
result = []
for txt_data in files:
        parse_txt_file_and_save_to_csv(txt_data, 'result_InputPresident.csv')

## Check Current Path ##
current_directory = os.getcwd() 
print(current_directory)
# pd.set_option('display.expand_frame_repr', False)

## Only Deploy version ##
# CPFT_PATH = 'CPFT-master-data.xlsx'
Base_PATH = 'result_InputPresident.csv'
# B2B_PATH = 'B2B.csv'
CPFM_PATH = 'CPFM.csv'

#df_Base = pd.read_excel(Base_PATH,converters={'วันที่':str,'เลขที่เอกสาร':str,'รหัสลูกค้า':str,'รหัส Store':str,'Branch':str}, skiprows=2)
df_Base = pd.read_csv(Base_PATH)

## Start CPFM ##
## Rename columns ##
df_CPFM = pd.read_csv(CPFM_PATH, skiprows=[0])
df_CPFM = df_CPFM[df_CPFM['rCV_name'].str.contains("เพรซิเดนท์ เบเกอรี่")]

# df_Base = df_Base.rename(columns={'rRef_doc_number': 'PO'}) # President hasn't Tax then set Tax = null
df_Base = df_Base.rename(columns={'InvoiceNo': 'Invoice_No'})
df_Base = df_Base.rename(columns={'Branch': 'Store_No'})
df_Base = df_Base.rename(columns={'BranchName': 'Store_Name'})
df_Base = df_Base.rename(columns={'InvoiceDate': 'Invoice_Date'})
df_Base = df_Base.rename(columns={'ExcVat': 'Exc_Vat'})
df_Base = df_Base.rename(columns={'Vat': 'Tax'}) 
df_Base = df_Base.rename(columns={'Total': 'Total_Amt'})
# df_Base['Tax'] = 0
df_Base['PO'] = ""

print(df_Base.head())

df_CPFM = df_CPFM.rename(columns={'rInput_doc_number': 'PO'})
df_CPFM = df_CPFM.rename(columns={'rRef_doc_number': 'Invoice_No'})
df_CPFM = df_CPFM.rename(columns={'rNett_amt_h': 'Total_Amt'})
df_CPFM = df_CPFM.rename(columns={'rTax_amt': 'Tax'})

# ## Convert data to string ##
df_Base['Invoice_No'] = df_Base['Invoice_No'].astype(str)
df_CPFM['Invoice_No'] = df_CPFM['Invoice_No'].astype(str)
df_CPFM['rOperationCode'] = df_CPFM['rOperationCode'].astype(str)
df_CPFM['rDocNumber'] = df_CPFM['rDocNumber'].astype(str)
df_Base['Exc_Vat'] = pd.to_numeric(df_Base['Exc_Vat'], errors='coerce')
df_Base['Tax'] = pd.to_numeric(df_Base['Tax'], errors='coerce')
df_Base['Total_Amt'] = pd.to_numeric(df_Base['Total_Amt'], errors='coerce')

print(df_CPFM.head())

## Merge the dataframes with suffixes added to duplicate column names ##
df_merge = pd.merge(df_Base, df_CPFM, on=['Invoice_No','Invoice_No'], how='outer', suffixes=('_BASE', '_CPFM'))
df_merge['Exc_Vat_difference'] = round(df_merge['Exc_Vat'] - df_merge['rSumNett'],2)
df_merge['Tax_difference'] = round(df_merge['Tax_BASE'] - df_merge['Tax_CPFM'],2)
df_merge['Total_Amt_difference'] = round(df_merge['Total_Amt_BASE'] - df_merge['Total_Amt_CPFM'],2)

print(df_merge.head())
print(df_merge.shape)
print(list(df_merge))


cols = ['rOperationCode','rDocNumber','Store_No', 'Store_Name', 'Invoice_No', 'Invoice_Date',
        'Exc_Vat', 'rSumNett', 'Exc_Vat_difference','Tax_BASE', 'Tax_CPFM', 'Tax_difference', 'Total_Amt_BASE', 'Total_Amt_CPFM', 'Total_Amt_difference',
        'PO_BASE']

df_merge = df_merge[cols]
df_merge.sort_values(by=['Total_Amt_difference'], ascending=True, inplace=True)

df_merge = df_merge.rename(columns={'Exc_Vat': 'Exc_Vat_BASE'})
df_merge = df_merge.rename(columns={'rSumNett': 'Exc_Vat_CPFM'})

print(df_merge.head())
print(df_merge.shape)

## Create a new column with CPFT_Null or CPFM_Null depending on the values of rTax_amt_CPFT and rTax_amt_CPFM ##
df_merge['null_report'] = ''
df_merge.loc[df_merge['Total_Amt_BASE'].isnull(), 'null_report'] = 'BASE_Null'
df_merge.loc[df_merge['Total_Amt_CPFM'].isnull(), 'null_report'] = 'CPFM_Null'

## Count null ##
value_counts = df_merge['null_report'].value_counts()

## Create a dataframe to store the counts ##
counts_df = pd.DataFrame({'Type': ['BASE_Null', 'CPFM_Null', 'Matching'],
                          'Count': [value_counts.get('BASE_Null', 0), value_counts.get('CPFM_Null', 0),
                                    value_counts.get('', 0)]})

## Add the sum of the "Count" column to the last row ##
counts_df.loc["Total"] = ["Total", counts_df["Count"].sum()]

## Create null dataframe ##
df_BASE_null = df_merge[df_merge['null_report'] == 'BASE_Null']
df_CPFM_null = df_merge[df_merge['null_report'] == 'CPFM_Null']

## Create dataframes only different values ##
df_CPFM_diffrNett = df_merge[(df_merge['Total_Amt_difference'].notnull()) & ((df_merge['Total_Amt_difference'] < -0.01)|(df_merge['Total_Amt_difference'] > 0.01)) & (df_merge['null_report'] == '')]
df_CPFM_diffrTax = df_merge[(df_merge['Tax_difference'].notnull()) & ((df_merge['Tax_difference'] < -0.01)|(df_merge['Tax_difference'] > 0.01)) & (df_merge['null_report'] == '')]
df_CPFM_diff = pd.concat([df_CPFM_diffrTax, df_CPFM_diffrNett], ignore_index=True).drop_duplicates().reset_index(drop=True)

cols_select = ['Store_No', 'Store_Name', 'Invoice_No', 'Invoice_Date',
        'Exc_Vat_BASE', 'Exc_Vat_CPFM', 'Exc_Vat_difference','Tax_BASE', 'Tax_CPFM', 'Tax_difference', 'Total_Amt_BASE', 'Total_Amt_CPFM', 'Total_Amt_difference',
        'PO_BASE']

df_merge_select = df_merge[cols_select]

df_CPFM_diff_select = df_CPFM_diff[['rOperationCode', 'rDocNumber', 'Invoice_No',
                                  'Tax_BASE', 'Tax_CPFM', 'Tax_difference', 'Total_Amt_BASE',
                                  'Total_Amt_CPFM', 'Total_Amt_difference']]

## Write the reconciled data_v1 and the counts to a single Excel file ##
with pd.ExcelWriter("outputs/reconciled_data_President.xlsx") as writer:
    df_merge_select.to_excel(writer, index=False, sheet_name='Reconciled Data')
    df_CPFM_diff.to_excel(writer, index=False, startrow=0, sheet_name='CPFM_diff')
    df_BASE_null.to_excel(writer, index=False, startrow=0, sheet_name='BASE_null')
    df_CPFM_null.to_excel(writer, index=False, startrow=0, sheet_name='CPFM_null')
    counts_df.to_excel(writer, index=False, startrow=0, header=False, sheet_name='null_report')

df_CPFM_diff_select.to_csv('outputs/cpfm_b2b_diff_President.csv', index=False)
