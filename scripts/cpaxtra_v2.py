import pandas as pd
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta

# Paths
BASE_PATH = r"inputs/CPAxtra Invoice.xlsx"
B2B_PATH = r"inputs/B2B.csv"
OUTPUT_CSV_PATH = r"outputs/b2b_diff_CPAxtra.csv"
OUTPUT_EXCEL_PATH = r"outputs/reconciled_data_CPAxtra.xlsx"

######################################### Session 1 Data cleansing and reading ###############################################


# Read data
df_BASE = pd.read_excel(
    BASE_PATH,
    converters={
        "วันที่": str,
        "เลขที่เอกสาร": str,
        "รหัสลูกค้า": str,
        "รหัส Store": str,
        "Branch": str,
        " Tax Branch": str,
    },
    skiprows=2,
)
df_BASE = df_BASE[df_BASE["Branch"] != "0"]

# Trim column names
def trim_column_names(df):
    df.columns = df.columns.str.strip()
    return df

df_BASE = trim_column_names(df_BASE)
# print(df_BASE.columns)

df_B2B = pd.read_csv(B2B_PATH, dtype={"rRef_doc_number": str, "rInput_doc_number": str, "rDocNumber": str}, skiprows=[0])
df_B2B = df_B2B[df_B2B['rCV_name'].str.contains("ซีพี แอ็กซ์ตร้า")]



# Rename columns and prepare data
df_BASE = df_BASE.rename(
    columns={
        "เลขที่เอกสาร": "Invoice_No",
        "Branch": "Store_No",
        "store name": "Store_Name",
        "วันที่": "Invoice_Date",
        "Invoice Price Ex. Vat": "Exc_Vat",
        "จำนวน": "Total_Amt",
    }
)
df_BASE["Tax"] = 0
df_BASE["PO"] = ""
df_BASE["Invoice_No"] = df_BASE["Invoice_No"].astype(
    str
)  # Ensure 'Invoice_No' column is string type

df_B2B = df_B2B.rename(
    columns={
        "rInput_doc_number": "PO",
        "rRef_doc_number": "Invoice_No",
        "rRef_doc_date_show": "Invoice_Date",
        "rNett_amt_h": "Total_Amt",
        "rTax_amt": "Tax",
    }
)

# Convert date function
def convert_date(date_str):
    date_obj = dt.strptime(date_str, "%d/%m/%Y") - relativedelta(years=543)
    return date_obj

df_B2B["Invoice_Date"] = df_B2B["Invoice_Date"].apply(convert_date)
df_BASE["Invoice_Date"] = pd.to_datetime(df_BASE.Invoice_Date, format="%d/%m/%Y")

# Modify invoice numbers
df_BASE["Invoice_No"] = df_BASE["Invoice_No"].apply(
    lambda x: x[2:] if isinstance(x, str) and len(x) > 12 else x
)

# Check for NaT values and convert others to the desired format
# Convert the 'Invoice_Date' column to datetime
df_BASE['Invoice_Date'] = pd.to_datetime(df_BASE['Invoice_Date'], errors='coerce')
df_B2B['Invoice_Date'] = pd.to_datetime(df_B2B['Invoice_Date'], errors='coerce')

# Convert to datetime type into this format "01/11/2023"
df_BASE[['Invoice_Date']] = df_BASE[['Invoice_Date']].apply(lambda col: col.dt.strftime('%d/%m/%Y') if col.dtype == 'datetime64[ns]' else col)
df_B2B[['Invoice_Date']] = df_B2B[['Invoice_Date']].apply(lambda col: col.dt.strftime('%d/%m/%Y') if col.dtype == 'datetime64[ns]' else col)

# Check and remove some columns
print("Vender columns name:")
print(df_BASE.columns)
print()
print("B2B columns name:")
print(df_B2B.columns)
print()

# Assuming df is your DataFrame and 'columns_to_remove' contains the names of columns you want to remove
columns_to_remove = ['rDoc_type_name', 'rTrn', 'rTRN_name', 'rCVCode']  # Replace these with your column names
# Remove multiple columns by names
df_B2B.drop(columns=columns_to_remove, inplace=True)

# Padding "Tax Branch" convert to strnumber
df_BASE["Tax Branch"] = df_BASE["Tax Branch"].astype(str).str.zfill(5)
df_BASE["Store_No"] = df_BASE["Store_No"].astype(str).str.zfill(5)

############################################# Session 2 CSV file ####################################################

## Merge the dataframes with suffixes added to duplicate column names ##
df_B2B_diff = pd.merge(df_BASE, df_B2B, on=['Invoice_No','Invoice_No'], how='inner', suffixes=('_BASE', '_B2B'))

# Calculate differences
df_B2B_diff["Inv. Date Check"] = (df_B2B_diff["Invoice_Date_BASE"] == df_B2B_diff["Invoice_Date_B2B"])
df_B2B_diff["Tax_BASE"] = round(df_B2B_diff["Total_Amt_BASE"] - df_B2B_diff["Exc_Vat"], 2)
df_B2B_diff["ExcludeVAT_diff"] = round(df_B2B_diff["Exc_Vat"] - df_B2B_diff["rSumNett"], 2)
df_B2B_diff["VAT_diff"] = round(df_B2B_diff["Tax_BASE"] - df_B2B_diff["Tax_B2B"], 2)
df_B2B_diff["IncludeVAT_diff"] = round(df_B2B_diff["Total_Amt_BASE"] - df_B2B_diff["Total_Amt_B2B"], 2)

# Round a specific column to 2 decimal points
columns_to_round = ['Tax_BASE', 'Exc_Vat', 'Total_Amt_BASE', 'rSumNett', 'Tax_B2B', 'Total_Amt_B2B']
for col in columns_to_round:
    df_B2B_diff[col] = df_B2B_diff[col].round(2)

# Filtered data
df_B2B_diff = df_B2B_diff[
    ((df_B2B_diff["ExcludeVAT_diff"] != 0) & df_B2B_diff["ExcludeVAT_diff"].notnull())
    | ((df_B2B_diff["VAT_diff"] != 0) & df_B2B_diff["VAT_diff"].notnull())
    | ((df_B2B_diff["IncludeVAT_diff"] != 0) & df_B2B_diff["IncludeVAT_diff"].notnull())
    | (df_B2B_diff["Inv. Date Check"] == False)
]

# df_B2B_diff.to_csv(OUTPUT_CSV_PATH, index=False, encoding='utf-8-sig')

# Select columns
cols_to_export = [
    "rDocNumber",
    # "PO",
    # "Invoice_No_BASE",
    # "Invoice_No_CPFM",
    # "INV.no Check",
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
filtered_df = df_B2B_diff[cols_to_export]

print(filtered_df)
filtered_df.to_csv(OUTPUT_CSV_PATH, index=False, encoding='utf-8-sig')
print("NO. of diff rows:", filtered_df.shape[0])

############################################# Session 3 Excel file ####################################################
# Merging dataframes and creating comparison columns
df_merge_excel = pd.merge(df_BASE, df_B2B, on=['Invoice_No','Invoice_No'], how='outer', suffixes=('_BASE', '_B2B'))

# Calculate differences
df_merge_excel["Inv. Date Check"] = (df_merge_excel["Invoice_Date_BASE"] == df_merge_excel["Invoice_Date_B2B"])
df_merge_excel["Tax_BASE"] = round(df_merge_excel["Total_Amt_BASE"] - df_merge_excel["Exc_Vat"], 2)
df_merge_excel["ExcludeVAT_diff"] = round(df_merge_excel["Exc_Vat"] - df_merge_excel["rSumNett"], 2)
df_merge_excel["VAT_diff"] = round(df_merge_excel["Tax_BASE"] - df_merge_excel["Tax_B2B"], 2)
df_merge_excel["IncludeVAT_diff"] = round(df_merge_excel["Total_Amt_BASE"] - df_B2B_diff["Total_Amt_B2B"], 2)

# Round a specific column to 2 decimal points
columns_to_round = ['Tax_BASE', 'Exc_Vat', 'Total_Amt_BASE', 'rSumNett', 'Tax_B2B', 'Total_Amt_B2B']
for col in columns_to_round:
    df_B2B_diff[col] = df_B2B_diff[col].round(2)

## Create a new column with CPFT_Null or B2B_Null depending on the values of rTax_amt_CPFT and rTax_amt_B2B ##
df_merge_excel['null_report'] = ''
df_merge_excel.loc[df_merge_excel['Total_Amt_BASE'].isnull(), 'null_report'] = 'BASE_Null'
df_merge_excel.loc[df_merge_excel['Total_Amt_B2B'].isnull(), 'null_report'] = 'B2B_Null'

## Count null ##
value_counts = df_merge_excel['null_report'].value_counts()

## Create a dataframe to store the counts ##
counts_df = pd.DataFrame({'Type': ['BASE_Null', 'B2B_Null', 'Matching'],
                          'Count': [value_counts.get('BASE_Null', 0), 
                                    value_counts.get('B2B_Null', 0),
                                    value_counts.get('', 0)]})

## Add the sum of the "Count" column to the last row ##
counts_df.loc["Total"] = ["Total", counts_df["Count"].sum()]

## Create null dataframe ##
df_BASE_null = df_merge_excel[df_merge_excel['null_report'] == 'BASE_Null']
df_B2B_null = df_merge_excel[df_merge_excel['null_report'] == 'B2B_Null']

## Write the reconciled data_v1 and the counts to a single Excel file ##
with pd.ExcelWriter(OUTPUT_EXCEL_PATH) as writer:
    df_merge_excel.to_excel(writer, index=False, sheet_name='Reconciled Data')
    df_B2B_diff.to_excel(writer, index=False, startrow=0, sheet_name='B2B_diff')
    df_BASE_null.to_excel(writer, index=False, startrow=0, sheet_name='BASE_null')
    df_B2B_null.to_excel(writer, index=False, startrow=0, sheet_name='B2B_null')
    counts_df.to_excel(writer, index=False, startrow=0, header=False, sheet_name='report')

