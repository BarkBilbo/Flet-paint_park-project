#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd


# In[2]:


df = pd.read_excel("hwanip_data.xlsx")
df


# df1 column : 계약번호, 차명, 지점/대리점, 주행거리, 불량내용 (columns)
# df2 column : 생산번호, 상호, 출하일, 부위구분, NaN
# df3 column : 차대번호, 성명, 발생일, 불량구분, Nan
# 
# df4        : df1 + df2 + df3 (row방향으로)   ==> 이렇게 되면 row가 1개가 됨
# 
# 

# In[3]:


df1 = df.iloc[:,:2].T.reset_index()
df2 = df.iloc[:,2:4].T.reset_index()
df3 = df.iloc[:,4:6].T.reset_index()
index_ = [0]


# In[4]:


df1


# In[5]:


df111 = pd.DataFrame(columns=df1.iloc[0, :], index = index_, data=[df1.iloc[1, :].values])
print(df111)


# In[6]:


df11 = pd.DataFrame(columns=df1.iloc[0, :], data=[df1.iloc[1, :].values])
df22 = pd.DataFrame(columns=df2.iloc[0, :], data=[df2.iloc[1, :].values])
df33 = pd.DataFrame(columns=df3.iloc[0, :], data=[df3.iloc[1, :].values])


# In[7]:


df11


# In[8]:


df22


# In[9]:


df33


# In[10]:


df_final = pd.concat([df11,df22.iloc[:,:3]], axis = 1)
df_final = pd.concat([df_final,df33.iloc[:,:3]], axis =1)
df_final


# In[11]:


df_final.info()


# In[ ]:
kw_1 = '이물'   #이물
kw_2 = '이색'   #이색
kw_3 = '얼룩'   #얼룩
kw_4 = '핀홀|핀 홀'   #핀홀
kw_5 = '칠부족|칠 부족'   #칠부족
kw_6 = '흐름'   #칠흐름
kw_7 = '튐'   #칠튐
kw_8 = '전착'   #전착불량
kw_9 = '실러|실라|실링'   #실러불량
kw_10 = '수연'   #수연불량
kw_11 = '폴리싱|광택'   #폴리싱자국




cond1 = df_final['불량내용'].str.contains(kw_1) 
cond2 = df_final['불량내용'].str.contains(kw_2) 
cond3 = df_final['불량내용'].str.contains(kw_3) 
cond4 = df_final['불량내용'].str.contains(kw_4) 
cond5 = df_final['불량내용'].str.contains(kw_5)
cond6 = df_final['불량내용'].str.contains(kw_6)
cond7 =  df_final['불량내용'].str.contains(kw_7) 
cond8 =  df_final['불량내용'].str.contains(kw_8) 
cond9 =  df_final['불량내용'].str.contains(kw_9) 
cond10 =  df_final['불량내용'].str.contains(kw_10) 
cond11 =  df_final['불량내용'].str.contains(kw_11) 


ft1 = df[df_final['불량내용'].str.contains(kw_1)]
print(ft1)
 # ===========> 여기서 불량 발생


df_fianl['도장 결함'] = 
