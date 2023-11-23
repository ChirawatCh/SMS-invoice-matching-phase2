import pandas as pd
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta

# File paths for input and output
Base_PATH = r"inputs/PerfectCompanion Invoice.xlsx"
CPFM_PATH = r"inputs/CPFM.csv"
OUTPUT_CSV_PATH = r"outputs/cpfm_diff_PerfectCompanion.csv"
OUTPUT_EXCEL_PATH = r"outputs/reconciled_data_PerfectCompanion.xlsx"

######################################### Session 1 Data cleansing and reading ###############################################

# Reading data from Excel and CSV files into dataframes
df_BASE = pd.read_excel(
    Base_PATH,
    converters={"เลขที่ PO": str, "เลขที่ใบแจ้งหนี้": str, "สาขาที่ออกใบกำกับภาษี": str},
    skiprows=6,
)
df_BASE = df_BASE.iloc[:, 1:]
# Drop the last n rows ("Grand total")
df_BASE = df_BASE.drop(df_BASE.tail(1).index)

df_CPFM = pd.read_csv(
    CPFM_PATH, dtype={"rRef_doc_number": str, "rInput_doc_number": str, "rDocNumber": str}, skiprows=[0]
)
df_CPFM = df_CPFM[df_CPFM["rCV_name"].str.contains("เพอร์เฟค คอมพาเนียน")]

# Renaming columns for consistency
column_mappings_base = {
    "เลขที่ PO": "PO",
    "เลขที่ใบแจ้งหนี้": "Invoice_No",
    "สถานที่ส่งสินค้า": "Store_Name",
    "วันที่ใบแจ้งหนี้": "Invoice_Date",
    "จำนวนเงิน\n(ก่อนVAT)": "Exc_Vat",
    "VAT 7%": "Tax",
    "จำนวนเงิน\n(รวมVAT)": "Total_Amt",
}
column_mappings_cpfm = {
    "rInput_doc_number": "PO",
    "rRef_doc_number": "Invoice_No",
    "rRef_doc_date_show": "Invoice_Date",
    "rNett_amt_h": "Total_Amt",
    "rTax_amt": "Tax",
}
df_BASE = df_BASE.rename(columns=column_mappings_base)
df_CPFM = df_CPFM.rename(columns=column_mappings_cpfm)

# Converting specific columns to string and numeric types
for col in ["PO", "Invoice_No"]:
    df_BASE[col] = df_BASE[col].astype(str)
    df_CPFM[col] = df_CPFM[col].astype(str)

for col in ["Tax", "Total_Amt", "Exc_Vat"]:
    df_BASE[col] = pd.to_numeric(df_BASE[col], errors="coerce")

for col in ["Tax", "Total_Amt", "rSumNett"]:
    df_CPFM[col] = pd.to_numeric(df_CPFM[col], errors="coerce")
    

# Function to convert dates from Buddhist Era (BE) to Anno Domini (AD)
def convert_date(date_str):
    date_obj = dt.strptime(date_str, "%d/%m/%Y") - relativedelta(years=543)
    return date_obj

# Applying date conversion function to the 'Invoice_Date' column in df_CPFM
df_CPFM["Invoice_Date"] = df_CPFM["Invoice_Date"].apply(convert_date)

# Check for NaT values and convert others to the desired format
# Convert the 'Invoice_Date' column to datetime
df_BASE['Invoice_Date'] = pd.to_datetime(df_BASE['Invoice_Date'], errors='coerce')
df_CPFM['Invoice_Date'] = pd.to_datetime(df_CPFM['Invoice_Date'], errors='coerce')

# Convert to datetime type into this format "01/11/2023"
df_BASE[['Invoice_Date']] = df_BASE[['Invoice_Date']].apply(lambda col: col.dt.strftime('%d/%m/%Y') if col.dtype == 'datetime64[ns]' else col)
df_CPFM[['Invoice_Date']] = df_CPFM[['Invoice_Date']].apply(lambda col: col.dt.strftime('%d/%m/%Y') if col.dtype == 'datetime64[ns]' else col)

# Check and remove some columns
print("Vender columns name:")
print(df_BASE.columns)
print()
print("CPFM columns name:")
print(df_CPFM.columns)
print()

# Remove multiple columns from BASE
base_columns_to_remove = ['รหัสลูกค้า', 'ปริมาณ', 'น้ำหนัก']  
df_BASE.drop(columns=base_columns_to_remove, inplace=True)
# Remove multiple columns by CPFM
cpfm_columns_to_remove = ['rDoc_type_name', 'rTrn', 'rTRN_name', 'rCVCode']  
df_CPFM.drop(columns=cpfm_columns_to_remove, inplace=True)

# Padding 'สาขาที่ออกใบกำกับภาษี' convert to strnumber
df_BASE['สาขาที่ออกใบกำกับภาษี'] = df_BASE['สาขาที่ออกใบกำกับภาษี'].astype(str).str.zfill(5)

############################################# Session 2 CSV file ####################################################

# Merging dataframes based on common columns and creating comparison columns
df_CPFM_diff = pd.merge(df_BASE, df_CPFM, on=["PO", "PO"], how="inner", suffixes=("_BASE", "_CPFM"))

df_CPFM_diff["INV.no Check"] = df_CPFM_diff["Invoice_No_BASE"] == df_CPFM_diff["Invoice_No_CPFM"]
df_CPFM_diff["Inv. Date Check"] = (df_CPFM_diff["Invoice_Date_BASE"] == df_CPFM_diff["Invoice_Date_CPFM"])
df_CPFM_diff["ExcludeVAT_diff"] = round(df_CPFM_diff["Exc_Vat"] - df_CPFM_diff["rSumNett"], 2)
df_CPFM_diff["VAT_diff"] = round(df_CPFM_diff["Tax_BASE"] - df_CPFM_diff["Tax_CPFM"], 2)
df_CPFM_diff["IncludeVAT_diff"] = round(df_CPFM_diff["Total_Amt_BASE"] - df_CPFM_diff["Total_Amt_CPFM"], 2)

## Convert data to string ##
df_CPFM_diff["PO"] = df_CPFM_diff["PO"].astype(str)

# Round a specific column to 2 decimal points
columns_to_round = ['Tax_BASE', 'Exc_Vat', 'Total_Amt_BASE', 'rSumNett', 'Tax_CPFM', 'Total_Amt_CPFM']
for col in columns_to_round:
    df_CPFM_diff[col] = df_CPFM_diff[col].round(2)

# Filtering rows based on specific conditions
df_CPFM_diff = df_CPFM_diff[
    ((df_CPFM_diff["ExcludeVAT_diff"] != 0) & df_CPFM_diff["ExcludeVAT_diff"].notnull())
    | ((df_CPFM_diff["VAT_diff"] != 0) & df_CPFM_diff["VAT_diff"].notnull())
    | ((df_CPFM_diff["IncludeVAT_diff"] != 0) & df_CPFM_diff["IncludeVAT_diff"].notnull())
    | (df_CPFM_diff["INV.no Check"] == False)
    | (df_CPFM_diff["Inv. Date Check"] == False)
]

# df_CPFM_diff.to_csv(OUTPUT_CSV_PATH, index=False, encoding='utf-8-sig')

# Select columns
cols_to_export = [
    "rDocNumber",
    # "PO",
    # "Invoice_No_BASE",
    # "Invoice_No_CPFM",
    "INV.no Check",
    # "Invoice_Date_BASE",
    # "Invoice_Date_CPFM",
    "Inv. Date Check",
    # 'Exc_Vat',
    # 'rSumNett',
    "ExcludeVAT_diff",
    # 'Tax_BASE',
    # 'Tax_CPFM',
    "VAT_diff",
    # 'Total_Amt_BASE',
    # 'Total_Amt_CPFM',
    "IncludeVAT_diff",
]
filtered_df = df_CPFM_diff[cols_to_export]

# Displaying the resulting dataframe and the number of differing rows
print(filtered_df)
# filtered_df.to_html("perfect-companion/output/cpfm_diff_CPAxtra.html")
filtered_df.to_csv(OUTPUT_CSV_PATH, index=False, encoding='utf-8-sig')
print("NO. of diff rows:", filtered_df.shape[0])

############################################# Session 3 Excel file ####################################################
# Merging dataframes and creating comparison columns
df_merge_excel = pd.merge(df_BASE, df_CPFM, on=["PO", "PO"], how="outer", suffixes=("_BASE", "_CPFM"))

df_merge_excel["INV.no Check"] = df_merge_excel["Invoice_No_BASE"] == df_merge_excel["Invoice_No_CPFM"]
df_merge_excel["Inv. Date Check"] = (df_merge_excel["Invoice_Date_BASE"] == df_merge_excel["Invoice_Date_CPFM"])
df_merge_excel["ExcludeVAT_diff"] = round(df_merge_excel["Exc_Vat"] - df_merge_excel["rSumNett"], 2)
df_merge_excel["VAT_diff"] = round(df_merge_excel["Tax_BASE"] - df_merge_excel["Tax_CPFM"], 2)
df_merge_excel["IncludeVAT_diff"] = round(df_merge_excel["Total_Amt_BASE"] - df_merge_excel["Total_Amt_CPFM"], 2)

## Convert data to string ##
df_CPFM_diff["PO"] = df_CPFM_diff["PO"].astype(str)

# Round a specific column to 2 decimal points
columns_to_round = ['Tax_BASE', 'Exc_Vat', 'Total_Amt_BASE', 'rSumNett', 'Tax_CPFM', 'Total_Amt_CPFM']
for col in columns_to_round:
    df_merge_excel[col] = df_merge_excel[col].round(2)

## Create a new column with CPFT_Null or CPFM_Null depending on the values of rTax_amt_CPFT and rTax_amt_CPFM ##
df_merge_excel['null_report'] = ''
df_merge_excel.loc[df_merge_excel['Total_Amt_BASE'].isnull(), 'null_report'] = 'BASE_Null'
df_merge_excel.loc[df_merge_excel['Total_Amt_CPFM'].isnull(), 'null_report'] = 'CPFM_Null'

## Count null ##
value_counts = df_merge_excel['null_report'].value_counts()

## Create a dataframe to store the counts ##
counts_df = pd.DataFrame({'Type': ['BASE_Null', 'B2B_Null', 'Matching'],
                          'Count': [value_counts.get('BASE_Null', 0), 
                                    value_counts.get('CPFM_Null', 0),
                                    value_counts.get('', 0)]})

## Add the sum of the "Count" column to the last row ##
counts_df.loc["Total"] = ["Total", counts_df["Count"].sum()]

## Create null dataframe ##
df_BASE_null = df_merge_excel[df_merge_excel['null_report'] == 'BASE_Null']
df_CPFM_null = df_merge_excel[df_merge_excel['null_report'] == 'CPFM_Null']

## Write the reconciled data_v1 and the counts to a single Excel file ##
with pd.ExcelWriter(OUTPUT_EXCEL_PATH) as writer:
    df_merge_excel.to_excel(writer, index=False, sheet_name='Reconciled Data')
    df_CPFM_diff.to_excel(writer, index=False, startrow=0, sheet_name='CPFM_diff')
    df_BASE_null.to_excel(writer, index=False, startrow=0, sheet_name='BASE_null')
    df_CPFM_null.to_excel(writer, index=False, startrow=0, sheet_name='CPFM_null')
    counts_df.to_excel(writer, index=False, startrow=0, header=False, sheet_name='report')
    
