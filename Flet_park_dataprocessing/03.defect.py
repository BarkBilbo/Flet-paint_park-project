import pandas as pd
import nltk
from nltk import FreqDist
from nltk.tokenize import word_tokenize
import re
nltk.download('punkt')

def hwanip_df_convert(file_name):
    df = pd.read_excel(file_name)
    # ## df1 => 계약번호, 차명, 지점, 대리점, 주행거리, 불량내용
    # ## df2 => 생산번호, 상호, 출하일, 부위구분, Nan
    # ## df3 => 차대번호, 성명, 발생일, 불량구분, Nan

    df1 = df.iloc[:,:2].T.reset_index()
    df2 = df.iloc[:,2:4].T.reset_index()
    df3 = df.iloc[:,4:6].T.reset_index()
    df11 = pd.DataFrame(columns=df1.iloc[0, :], data=[df1.iloc[1, :].values])
    df22 = pd.DataFrame(columns=df2.iloc[0, :], data=[df2.iloc[1, :].values])
    df33 = pd.DataFrame(columns=df3.iloc[0, :], data=[df3.iloc[1, :].values])
    df_final = pd.concat([df11,df22.iloc[:,:4]], axis = 1)
    df_final = pd.concat([df_final,df33.iloc[:,:4]], axis =1)
    df_final['도장결함'] = ""

    defect = df_final.iloc[0, 4]
    # 특수문자 제거
    result = re.sub(r"[(),a-zA-Z0-9]","", defect)
    # 단어 쪼개기
    words = word_tokenize(result)
    print(words)

    dict1,dict2,dict3,dict4,dict5,dict6,dict7,dict8,dict9 = {},{},{},{},{},{},{},{},{}
    dict1["이물"] = ["이물질", "이물"]
    dict2["이색"] = ["이색"]
    dict3["얼룩"] = ["얼룩"]
    dict4["핀홀"] = ["핀홀", "핀 홀"]
    dict5["칠부족"] = ["칠부족", "칠 부족"]
    dict6["칠튐"] = ["칠튐", "칠 튐"]
    dict7["전착불량"] = ["전착", "전착흐름", "전착 흐름"]
    dict8["실러불량"] = ["실러", "실라", "실링", "실러불량", "실라불량", "실링불량"]
    dict9["폴리싱자국"] = ["광택", "폴리싱"]
    dict_collect_list = [dict1,dict2,dict3,dict4,dict5,dict6,dict7,dict8,dict9]

    # 실제 결함 단어 찾아내서 dataframe "도장결함"열에 연속으로 채워넣기
    for defect_txt in words:
        for i in range(len(dict_collect_list)):
            if defect_txt in list(dict_collect_list[i].values())[0]:
                # print(list(dict_collect_list[i].keys())[0])
                df_final["도장결함"] += list(dict_collect_list[i].keys())[0] + " "
    return df_final

file_name = "hwanip_data.xlsx"
df_final = hwanip_df_convert(file_name)
# df = pd.DataFrame()
df_complete = pd.concat([df_final, df_final])
df_complete

########################################################
# 가변적으로 전역변수이름을 생성하기
for i in range(3):
    globals()["df_comp_{}".format(i)] = i
