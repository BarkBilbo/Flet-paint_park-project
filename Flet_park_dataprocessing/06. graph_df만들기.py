import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
import re

nltk.download('punkt')


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)


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

# 발생일 기준으로 연월일 만들기
result_df['발생일'] = pd.to_datetime(result_df['발생일'])
result_df['year'] = result_df['발생일'].dt.year
result_df['month'] = result_df['발생일'].dt.month

# result_df





import pandas as pd

# 기존의 result_df 데이터프레임을 그대로 사용
# ...

# graph_df 만들기
graph_df = pd.DataFrame({
    'year': result_df['발생일'].dt.year,
    'month': result_df['발생일'].dt.month
})

# `graph_df`의 인덱스를 MultiIndex로 설정
graph_df = graph_df.set_index(['year', 'month'])

# `result_df`의 인덱스를 `graph_df`와 동일한 MultiIndex로 설정
result_df = result_df.set_index(['year', 'month'])

# `graph_df`에 환입 건수 계산
graph_df['환입_전체'] = result_df.groupby(['year', 'month']).size()

# `graph_df`에서 도장 결함 건수 계산
graph_df['환입_도장'] = result_df[result_df['불량종류'] == '도장'].groupby(['year', 'month']).size()

# 도장 결함 비율 계산
graph_df['도장_결함비율'] = graph_df['환입_도장'] / graph_df['환입_전체'] * 100

# MV_전체결함 / KA4_전체결함 / RJ_전체결함 칼럼 추가
graph_df['MV_전체결함'] = result_df[result_df['차명'] == 'EV9'].groupby(['year', 'month']).size()
graph_df['KA4_전체결함'] = result_df[result_df['차명'] == '카니발'].groupby(['year', 'month']).size()
graph_df['RJ_전체결함'] = result_df[result_df['차명'] == 'K9'].groupby(['year', 'month']).size()

# MV_도장결함 / KA4_도장결함 / RJ_도장결함 칼럼 추가
graph_df['MV_도장결함'] = result_df[(result_df['차명'] == 'EV9') & (result_df['불량종류'] == '도장')].groupby(['year', 'month']).size()
graph_df['KA4_도장결함'] = result_df[(result_df['차명'] == '카니발') & (result_df['불량종류'] == '도장')].groupby(['year', 'month']).size()
graph_df['RJ_도장결함'] = result_df[(result_df['차명'] == 'K9') & (result_df['불량종류'] == '도장')].groupby(['year', 'month']).size()

# 도장결함비율 칼럼 추가
graph_df['도장결함비율'] = (graph_df['MV_도장결함'] + graph_df['KA4_도장결함'] + graph_df['RJ_도장결함']) / (
        graph_df['MV_전체결함'] + graph_df['KA4_전체결함'] + graph_df['RJ_전체결함']) * 100
graph_df['MV_도장결함비율'] = graph_df['MV_도장결함'] / graph_df['MV_전체결함'] * 100
graph_df['KA4_도장결함비율'] = graph_df['KA4_도장결함'] / graph_df['KA4_전체결함'] * 100
graph_df['RJ_도장결함비율'] = graph_df['RJ_도장결함'] / graph_df['RJ_전체결함'] * 100

# ※외관 :  긁힘, 굴곡, 오염, 간단차
# ※기타 : 파손, 기능, 누수, 이종, 미장착, 조립불, 기타
## MV 기능/외관/기타 결함
graph_df['MV_기능결함'] = result_df[(result_df['차명'] == 'EV9') 
                                & (result_df['불량종류'] == '기능')].groupby(['year', 'month']).size()
graph_df['MV_외관결함'] = result_df[((result_df['차명'] == 'EV9')) 
                                 & ((result_df['불량종류'] == '긁힘') 
                                    | (result_df['불량종류'] == '굴곡') 
                                    | (result_df['불량종류'] == '오염') 
                                    | (result_df['불량종류'] == '간단차'))].groupby(['year', 'month']).size()


graph_df['MV_기타결함'] = result_df[((result_df['차명'] == 'EV9')) 
                                 & ((result_df['불량종류'] == '긁힘') 
                                    | (result_df['불량종류'] == '굴곡') 
                                    | (result_df['불량종류'] == '오염') 
                                    | (result_df['불량종류'] == '간단차'))].groupby(['year', 'month']).size()

## KA4 기능/외관/기타 결함
graph_df['KA4_기능결함'] = result_df[(result_df['차명'] == '카니발') 
                                 & (result_df['불량종류'] == '기능')].groupby(['year', 'month']).size()
graph_df['KA4_외관결함'] = result_df[((result_df['차명'] == '카니발')) 
                                 & ((result_df['불량종류'] == '긁힘') 
                                    | (result_df['불량종류'] == '굴곡') 
                                    | (result_df['불량종류'] == '오염') 
                                    | (result_df['불량종류'] == '간단차'))].groupby(['year', 'month']).size()


graph_df['KA4_기타결함'] = result_df[((result_df['차명'] == '카니발')) 
                                 & ((result_df['불량종류'] == '긁힘') 
                                    | (result_df['불량종류'] == '굴곡') 
                                    | (result_df['불량종류'] == '오염') 
                                    | (result_df['불량종류'] == '간단차'))].groupby(['year', 'month']).size()


## RJ 기능/외관/기타 결함
graph_df['K9_기능결함'] = result_df[(result_df['차명'] == 'K9') 
                                 & (result_df['불량종류'] == '기능')].groupby(['year', 'month']).size()
graph_df['K9_외관결함'] = result_df[((result_df['차명'] == 'K9')) 
                                 & ((result_df['불량종류'] == '긁힘') 
                                    | (result_df['불량종류'] == '굴곡') 
                                    | (result_df['불량종류'] == '오염') 
                                    | (result_df['불량종류'] == '간단차'))].groupby(['year', 'month']).size()


graph_df['K9_기타결함'] = result_df[((result_df['차명'] == 'K9')) 
                                 & ((result_df['불량종류'] == '긁힘') 
                                    | (result_df['불량종류'] == '굴곡') 
                                    | (result_df['불량종류'] == '오염') 
                                    | (result_df['불량종류'] == '간단차'))].groupby(['year', 'month']).size()








# graph_df에서 중복된 행 제거
graph_df.drop_duplicates(inplace=True)

# year, month 오름차순 정렬
graph_df = graph_df.sort_values(['month'], ascending=True)
graph_df = graph_df.sort_values(['year'], ascending=True)

# NaN값을 0으로 대체
graph_df['도장결함비율'] = graph_df['도장결함비율'].fillna(0)


# 맨 밑에 합계 row 추가
graph_df.loc[('Total', ''), :] = graph_df.sum()

