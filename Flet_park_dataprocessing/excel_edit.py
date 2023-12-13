# pip install pandas openpyxl

import pandas as pd

df = pd.read_excel('excel_edit.xlsx')

rows_to_delete = [0,1,2,3,4,5,6,7,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31]
df.drop(index=rows_to_delete, inplace = True)

df.columns = df.iloc[0]
df = df[1:]
