import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
import re

nltk.download('punkt')

def process_multiple_excel_files(file_paths):
    final_result_df = pd.DataFrame()

    for excel_file_path in file_paths:
        xls = pd.ExcelFile(excel_file_path)
        dfs = {sheet_name: pd.read_excel(excel_file_path, sheet_name) for sheet_name in xls.sheet_names}

        for sheet_name, df in dfs.items():
            rows_to_delete = [0, 1, 2, 3, 4, 5, 6, 7, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]
            df.drop(index=rows_to_delete, inplace=True)
            df.columns = df.iloc[0]
            df = df[1:]
            dfs[sheet_name] = df

        result_dfs = []

        for sheet_name, df in dfs.items():
            df1 = df.iloc[:, :2].T.reset_index()
            df2 = df.iloc[:, 3:5].T.reset_index()
            df3 = df.iloc[:, 6:8].T.reset_index()
            df4 = df.iloc[:, 7:9].T.reset_index()

            df11 = pd.DataFrame(columns=df1.iloc[0, :], data=[df1.iloc[1, :].values])
            df22 = pd.DataFrame(columns=df2.iloc[0, :], data=[df2.iloc[1, :].values])
            df33 = pd.DataFrame(columns=df3.iloc[0, :], data=[df3.iloc[1, :].values])
            df44 = pd.DataFrame(columns=df4.iloc[0, :], data=[df4.iloc[1, :].values])

            df_final = pd.concat([df11, df22.iloc[:, :4]], axis=1, ignore_index=True)
            df_final = pd.concat([df_final, df33.iloc[:, :4]], axis=1, ignore_index=True)
            df_final['도장결함'] = ""

            defect = df_final.iloc[0, 4]
            result = re.sub(r"[(),a-zA-Z0-9]", "", defect)
            words = word_tokenize(result)

            dict_collect = {
                "이물": ["이물질", "이물"],
                "이색": ["이색"],
                "얼룩": ["얼룩"],
                "핀홀": ["핀홀", "핀 홀"],
                "칠부족": ["칠부족", "칠 부족"],
                "칠튐": ["칠튐", "칠 튐"],
                "전착불량": ["전착", "전착흐름", "전착 흐름"],
                "실러불량": ["실러", "실라", "실링", "실러불량", "실라불량", "실링불량", "실리콘"],
                "폴리싱자국": ["광택", "폴리싱"]
            }

            for defect_txt in words:
                for key, values in dict_collect.items():
                    if defect_txt in values:
                        df_final["도장결함"] += key + " "

            df_final.reset_index(drop=True, inplace=True)
            result_dfs.append(df_final.iloc[0, :])

        final_df = pd.concat(result_dfs, axis=1, ignore_index=True).T
        final_result_df = pd.concat([final_result_df, final_df], ignore_index=True)

    final_result_df['도장결함'] = final_result_df['도장결함'].apply(lambda x: ' '.join(set(x.split())))

    return final_result_df

file_paths = ['excel_edit.xlsx', '품질문제 기타환입 24년 1월 1차_덕평(15건).xlsx', '품질문제 기타환입 23년 10월 3차_경산(18건).xlsx', '품질문제 기타환입 23년 8월 4차_광명(5건).xlsx', '품질문제 기타환입 24년 1월 2차_광명(5건).xlsx', '품질문제 기타환입 23년 9월 4차_광명(3건).xlsx' ]
result_df = process_multiple_excel_files(file_paths)
result_df.columns = ["계약번호", "차명", "대리점", "주행거리", "불량내용", "생산번호", "상호", "출하일", "부위", "차대번호", "성명", "발생일", "불량종류", "도장결함"]

result_df

