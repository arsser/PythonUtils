#把报销excel表按公司拆分成多个文件；

import pandas as pd

# 读取Excel文件
df = pd.read_excel(r'C:\Users\Arsser\Desktop\外包费用\出勤和报销\外包报销单据确认-赵思逸1699875477458.xlsx')

# 拆分数据
for value in df['公司'].unique():
    df_subset = df[df['公司'] == value]
    df_subset.to_excel(rf"C:\Users\Arsser\Desktop\外包费用\出勤和报销\分拆\{value}_10月报销.xlsx", index=False)
