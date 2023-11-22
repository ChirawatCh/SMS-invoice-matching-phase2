import pandas as pd
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
pd.set_option('display.precision', 2)

# File paths for input and output
BASE_PATH = r"inputs/ThaiNamthip Invoice.xlsx"
CPFM_PATH = r"inputs/CPFM.csv"
OUTPUT_CSV_PATH = r"outputs/cpfm_diff_ThaiNamThip.csv"
OUTPUT_EXCEL_PATH = r"outputs/reconciled_data_ThaiNamThip.xlsx"

######################################### Session 1 Data cleansing and reading ###############################################

df_BASE = pd.read_excel(
    BASE_PATH,
    converters={
        "Branch account": str,
        "Store No.": str,
        "Reference": str,
        "Payment reference": str,
        # "Doc. Date": str,
        "Business Place": str,
        # "Net due date": str,
    },
    skiprows=0,
)

# Trim column names
def trim_column_names(df):
    df.columns = df.columns.str.strip()
    return df

df_BASE = trim_column_names(df_BASE)
# print(df_BASE.columns)

## Read Data
df_CPFM = pd.read_csv(CPFM_PATH, dtype={"rRef_doc_number": str, "rInput_doc_number": str, "rRef_doc_date_show": str, "rDocNumber": str}, skiprows=[0])
df_CPFM = df_CPFM[df_CPFM["rCV_name"].str.contains("ไทยน้ำทิพย์")]

## Rename columns ##
df_BASE = df_BASE.rename(columns={"Payment reference": "PO"})
df_BASE = df_BASE.rename(columns={"Reference": "Invoice_No"})
df_BASE = df_BASE.rename(columns={"Store No.": "Store_No"})
df_BASE = df_BASE.rename(columns={"Store name": "Store_Name"})
df_BASE = df_BASE.rename(columns={"Doc. Date": "Invoice_Date"})
df_BASE = df_BASE.rename(columns={"Exc.Vat": "Exc_Vat"})
df_BASE = df_BASE.rename(columns={"Vat": "Tax"})
df_BASE = df_BASE.rename(columns={"LC amnt": "Total_Amt"})

## Adjust Base Data
df_BASE.PO = df_BASE.PO.str.zfill(15)

df_CPFM = df_CPFM.rename(columns={"rInput_doc_number": "PO"})
df_CPFM = df_CPFM.rename(columns={"rRef_doc_number": "Invoice_No"})
df_CPFM = df_CPFM.rename(columns={"rRef_doc_date_show": "Invoice_Date"})
df_CPFM = df_CPFM.rename(columns={"rNett_amt_h": "Total_Amt"})
df_CPFM = df_CPFM.rename(columns={"rTax_amt": "Tax"})

## Convert data to string ##
df_BASE["PO"] = df_BASE["PO"].astype(str)
df_CPFM["PO"] = df_CPFM["PO"].astype(str)
df_BASE["Invoice_No"] = df_BASE["Invoice_No"].astype(str)
df_CPFM["Invoice_No"] = df_CPFM["Invoice_No"].astype(str)
df_CPFM["rOperationCode"] = df_CPFM["rOperationCode"].astype(str)
df_CPFM["rDocNumber"] = df_CPFM["rDocNumber"].astype(str)
df_BASE["Exc_Vat"] = pd.to_numeric(df_BASE["Exc_Vat"], errors="coerce")
df_BASE["Tax"] = pd.to_numeric(df_BASE["Tax"], errors="coerce")
df_BASE["Total_Amt"] = pd.to_numeric(df_BASE["Total_Amt"], errors="coerce")
df_CPFM["rSumNett"] = pd.to_numeric(df_CPFM["rSumNett"], errors="coerce")
df_CPFM["Tax"] = pd.to_numeric(df_CPFM["Tax"], errors="coerce")
df_CPFM["Total_Amt"] = pd.to_numeric(df_CPFM["Total_Amt"], errors="coerce")

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

# Apply formatting to valid timestamps, keeping NaT for invalid/missing values
df_BASE['Invoice_Date'] = df_BASE['Invoice_Date'].apply(lambda x: x.strftime('%Y-%m-%d') if not pd.isnull(x) else pd.NaT)
df_CPFM['Invoice_Date'] = df_CPFM['Invoice_Date'].apply(lambda x: x.strftime('%Y-%m-%d') if not pd.isnull(x) else pd.NaT)

# Check and remove some columns
print("Vender columns name:")
print(df_BASE.columns)
print()
print("CPFM columns name:")
print(df_CPFM.columns)
print()
# Assuming df is your DataFrame and 'columns_to_remove' contains the names of columns you want to remove
columns_to_remove = ['rDoc_type_name', 'rTrn', 'rTRN_name', 'rCVCode']  # Replace these with your column names
# Remove multiple columns by names
df_CPFM.drop(columns=columns_to_remove, inplace=True)

############################################# Session 2 CSV file ####################################################

## Merge the dataframes with suffixes added to duplicate column names ##
df_CPFM_diff = pd.merge(df_BASE, df_CPFM, on=["PO", "PO"], how="inner", suffixes=("_BASE", "_CPFM"))

df_CPFM_diff["INV.no Check"] = df_CPFM_diff["Invoice_No_BASE"] == df_CPFM_diff["Invoice_No_CPFM"]
df_CPFM_diff["Inv. Date Check"] = (df_CPFM_diff["Invoice_Date_BASE"] == df_CPFM_diff["Invoice_Date_CPFM"])
df_CPFM_diff["ExcludeVAT_diff"] = round(df_CPFM_diff["Exc_Vat"] - df_CPFM_diff["rSumNett"], 2)
df_CPFM_diff["VAT_diff"] = round(df_CPFM_diff["Tax_BASE"] - df_CPFM_diff["Tax_CPFM"], 2)
df_CPFM_diff["IncludeVAT_diff"] = round(df_CPFM_diff["Total_Amt_BASE"] - df_CPFM_diff["Total_Amt_CPFM"], 2)

## Convert data to string ##
df_CPFM_diff["PO"] = df_CPFM_diff["PO"].astype(str)
# Round a specific column ('Tax_BASE') to 2 decimal points
df_CPFM_diff['Tax_BASE'] = df_CPFM_diff['Tax_BASE'].round(2)
df_CPFM_diff['Exc_Vat'] = df_CPFM_diff['Exc_Vat'].round(2)
df_CPFM_diff['Total_Amt_BASE'] = df_CPFM_diff['Total_Amt_BASE'].round(2)

# Filtering rows based on specific conditions
df_CPFM_diff = df_CPFM_diff[
    ((df_CPFM_diff["ExcludeVAT_diff"] != 0) & df_CPFM_diff["ExcludeVAT_diff"].notnull())
    | ((df_CPFM_diff["VAT_diff"] != 0) & df_CPFM_diff["VAT_diff"].notnull())
    | ((df_CPFM_diff["IncludeVAT_diff"] != 0) & df_CPFM_diff["IncludeVAT_diff"].notnull())
    | (df_CPFM_diff["INV.no Check"] == False)
    | (df_CPFM_diff["Inv. Date Check"] == False)
]

df_CPFM_diff.to_csv(OUTPUT_CSV_PATH, index=False, encoding='utf-8-sig')

# Selecting columns for export and writing the result to a CSV file

cols_to_export = [
    "Store_Name",
    "PO",
    "Invoice_No_BASE",
    "Invoice_No_CPFM",
    "INV.no Check",
    "Invoice_Date_BASE",
    "Invoice_Date_CPFM",
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
# filtered_df.to_html("thainamthip/output/cpfm_diff_ThaiNamThip.html")
print("NO. of diff rows:", filtered_df.shape[0])

############################################# Session 3 Excel file ####################################################
# Merging dataframes and creating comparison columns
df_merge_excel = pd.merge(df_BASE, df_CPFM, on=["PO", "PO"], how="outer", suffixes=("_BASE", "_CPFM"))

## Create a new column with CPFT_Null or CPFM_Null depending on the values of rTax_amt_CPFT and rTax_amt_CPFM ##
df_merge_excel['null_report'] = ''
df_merge_excel.loc[df_merge_excel['Total_Amt_BASE'].isnull(), 'null_report'] = 'BASE_Null'
df_merge_excel.loc[df_merge_excel['Total_Amt_CPFM'].isnull(), 'null_report'] = 'CPFM_Null'

## Count null ##
value_counts = df_merge_excel['null_report'].value_counts()

## Create a dataframe to store the counts ##
counts_df = pd.DataFrame({'Type': ['BASE_Null', 'B2B_Null', "DIFF", 'Matching'],
                          'Count': [value_counts.get('BASE_Null', 0), 
                                    value_counts.get('CPFM_Null', 0),
                                    df_CPFM_diff.shape[0],
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
    
