import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
import re
import os  # 추가된 부분

nltk.download('punkt')

def process_multiple_excel_files(file_paths):  # 함수명 수정

    # Initialize an empty DataFrame to store the results
    final_result_df = pd.DataFrame()

    for excel_file_path in file_paths:
        # Step 1: Read all sheets into separate DataFrames
        xls = pd.ExcelFile(excel_file_path)
        dfs = {sheet_name: pd.read_excel(excel_file_path, sheet_name) for sheet_name in xls.sheet_names}

        # Step 2: Delete unnecessary rows and columns
        for sheet_name, df in dfs.items():
            rows_to_delete = [0, 1, 2, 3, 4, 5, 6, 7, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]
            df.drop(index=rows_to_delete, inplace=True)
            df.columns = df.iloc[0]
            df = df[1:]
            dfs[sheet_name] = df

        # Step 3: Process each DataFrame and concatenate the results
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

            dict1, dict2, dict3, dict4, dict5, dict6, dict7, dict8, dict9 = {}, {}, {}, {}, {}, {}, {}, {}, {}
            dict1["이물"] = ["이물질", "이물"]
            dict2["이색"] = ["이색"]
            dict3["얼룩"] = ["얼룩"]
            dict4["핀홀"] = ["핀홀", "핀 홀"]
            dict5["칠부족"] = ["칠부족", "칠 부족"]
            dict6["칠튐"] = ["칠튐", "칠 튐"]
            dict7["전착불량"] = ["전착", "전착흐름", "전착 흐름"]
            dict8["실러불량"] = ["실러", "실라", "실링", "실러불량", "실라불량", "실링불량", "실리콘"]
            dict9["폴리싱자국"] = ["광택", "폴리싱"]

            dict_collect_list = [dict1, dict2, dict3, dict4, dict5, dict6, dict7, dict8, dict9]

            for defect_txt in words:
                for i in range(len(dict_collect_list)):
                    if list(dict_collect_list[i].values()):
                        if defect_txt in list(dict_collect_list[i].values())[0]:
                            df_final["도장결함"] += list(dict_collect_list[i].keys())[0] + " "

            # Resetting index before appending to the result list
            df_final.reset_index(drop=True, inplace=True)
            result_dfs.append(df_final.iloc[0, :])

        # Concatenating all DataFrames in the result list with a new index
        final_df = pd.concat(result_dfs, axis=1, ignore_index=True).T

        # Add the processed DataFrame to the final result DataFrame
        final_result_df = pd.concat([final_result_df, final_df], ignore_index=True)

    return final_result_df

# Example Usage:
file_paths = ['excel_edit.xlsx', '품질문제 기타환입 24년 1월 1차_덕평(15건).xlsx', '품질문제 기타환입 23년 10월 3차_경산(18건).xlsx']
result_df = process_multiple_excel_files(file_paths)
result_df.columns = ["계약번호","차명","대리점","주행거리","불량내용","생산번호","상호","출하일","부위","차대번호","성명","발생일","불량종류","결함내용"]
result_df
