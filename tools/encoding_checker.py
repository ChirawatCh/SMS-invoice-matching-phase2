import pandas as pd
import os
import csv
import glob
import re

# Trying different encodings suitable for Thai text on Mac
possible_encodings = ['cp874', 'TIS-620', 'utf-8', 'latin-1']

for encoding in possible_encodings:
    try:
        # df_Base = pd.read_csv("CPF231003.TXT", encoding=encoding)
        df_Base = pd.read_csv("result_InputPresident.csv", encoding=encoding)
        print(f"File read successfully with encoding: {encoding}")
        break  # Stop trying encodings if successful
    except Exception as e:
        print(f"Failed with encoding {encoding}: {e}")
