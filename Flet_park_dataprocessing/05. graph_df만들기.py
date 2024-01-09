import pandas as pd


# 발생일 기준으로 연월일 만들기
result_df['발생일'] = pd.to_datetime(result_df['발생일'])
result_df['year'] = result_df['발생일'].dt.year
result_df['month'] = result_df['발생일'].dt.month

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
graph_df['환입_전체'] = graph_df.groupby(['year', 'month']).size()

# `graph_df`에서 도장 결함 건수 계산
graph_df['환입_도장'] = graph_df[result_df['불량종류'] == '도장'].groupby(['year', 'month']).size()

# 도장 결함 비율 계산
graph_df['도장_결함비율'] = graph_df['환입_도장'] / graph_df['환입_전체'] * 100

# graph_df에서 중복된 행 제거
graph_df.drop_duplicates(inplace=True)

graph_df
