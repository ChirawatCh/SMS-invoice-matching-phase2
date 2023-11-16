import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

B2B_PATH = r'B2B.csv'
df_B2B = pd.read_csv(B2B_PATH, skiprows=[0])
df_B2B = df_B2B.rename(columns={'rRef_doc_date_show': 'Invoice_Date'})

def convert_date(date_str):
  date_obj = datetime.strptime(date_str, "%d/%m/%Y") - relativedelta(years=543)
  return date_obj

# Apply function to column 
df_B2B['Invoice_Date'] = df_B2B['Invoice_Date'].apply(convert_date)

for i, date_obj in df_B2B['Invoice_Date'].items():
  print(i, date_obj, type(date_obj))