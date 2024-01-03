import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
import re

excel_file_path = "excel_edit.xlsx"
xls = pd.ExcelFile(excel_file_path)
dfs = {xls.sheet_names[0]: pd.read_excel(excel_file_path, xls.sheet_names[0])}
rows_to_delete = [0, 1, 2, 3, 4, 5, 6, 7, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]
df01.drop(index=rows_to_delete, inplace=True)
df01 = df01.reset_index(drop = True)

df01.iloc[:, 0]
df01.iloc[:, 1]

df01.iloc[:, 3]
df01.iloc[:, 4]

df01.iloc[:, 6]
df01.iloc[:, 7]

df_merge01 = pd.concat([df01.iloc[:, 0],df01.iloc[:, 1]], axis = 1).T
df_merge01 = df_merge01.reset_index(drop = True)
df_merge34 = pd.concat([df01.iloc[:, 3],df01.iloc[:, 4]], axis = 1).T
df_merge34 = df_merge34.reset_index(drop = True)
df_merge67 = pd.concat([df01.iloc[:, 6],df01.iloc[:, 7]], axis = 1).T
df_merge67 = df_merge67.reset_index(drop = True)

df_merge012 = df_merge01.rename(columns = df_merge01.iloc[0,:])
df_merge012 = df_merge012.drop([df_merge012.index[0]])
df_merge012 = df_merge012.reset_index(drop = True)

df_merge342 = df_merge34.rename(columns = df_merge34.iloc[0,:])
df_merge342 = df_merge342.drop([df_merge342.index[0]])
df_merge342 = df_merge342.reset_index(drop = True)

df_merge672 = df_merge67.rename(columns = df_merge67.iloc[0,:])
df_merge672 = df_merge672.drop([df_merge672.index[0]])
df_merge672 = df_merge672.reset_index(drop = True)

df_merge_final = pd.concat([df_merge012,df_merge342], axis = 1)
df_merge_final = pd.concat([df_merge_final,df_merge672], axis = 1)
df_merge_final
