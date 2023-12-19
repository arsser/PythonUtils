#把出勤excel表按公司拆分成多个文件；

import pandas as pd

# 定义月份变量
month = '2023-11'

# 读取Excel文件
df = pd.read_excel(fr'\\10.12.21.65\share\外包费用\新流程\{month}\Step1-考勤数据\外包钉钉考勤.xlsm')

# 拆分数据
for value in df['公司'].unique():
    df_subset = df[df['公司'] == value]
    df_subset.to_excel(rf"\\10.12.21.65\share\外包费用\新流程\{month}\Step1-考勤数据\chai_{value}_{month}月出勤.xlsx", index=False)
