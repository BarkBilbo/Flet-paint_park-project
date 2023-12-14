
# 엑셀 파일 경로 
excel_file_path = 'excel_edit.xlsx'


# 모든 시트를 담을 빈 딕셔너리 만들기
dfs = {}

# 엑셀 파일 불러오기
xls = pd.ExcelFile(excel_file_path)

# 모든 시트에 대해서 반복해주기
for sheet_name in xls.sheet_names:
    # 각 시트를 데이터프레임으로 읽어와서 딕셔너리에 추가하기
    dfs[sheet_name] = pd.read_excel(excel_file_path, sheet_name)
    
# sheet1, sheet2, ,,,,,,,,, sheetx 데이터 프레임 만들기
for i, (sheet_name, df) in enumerate(dfs.items(), start=1):
    globals()[f'sheet{i}'] = df
