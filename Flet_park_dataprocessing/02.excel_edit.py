


###################### 엑셀파일의 위 아래 row column 잘라버리기 ###################################

# pip install pandas openpyxl


import pandas as pd

df = pd.read_excel('excel_edit.xlsx')

rows_to_delete = [0,1,2,3,4,5,6,7,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31]
df.drop(index=rows_to_delete, inplace = True)

df.columns = df.iloc[0]
df = df[1:]


import pandas as pd



# ## df1 => 계약번호, 차명, 지점, 대리점, 주행거리, 불량내용
# ## df2 => 생산번호, 상호, 출하일, 부위구분, Nan
# ## df3 => 차대번호, 성명, 발생일, 불량구분, Nan



df1 = df.iloc[:,:2].T.reset_index()
df2 = df.iloc[:,3:5].T.reset_index()
df3 = df.iloc[:,6:8].T.reset_index()



df11 = pd.DataFrame(columns=df1.iloc[0, :], data=[df1.iloc[1, :].values])
df22 = pd.DataFrame(columns=df2.iloc[0, :], data=[df2.iloc[1, :].values])
df33 = pd.DataFrame(columns=df3.iloc[0, :], data=[df3.iloc[1, :].values])
df33


df_final = pd.concat([df11,df22.iloc[:,:4]], axis = 1)
df_final = pd.concat([df_final,df33.iloc[:,:4]], axis =1)
df_final
