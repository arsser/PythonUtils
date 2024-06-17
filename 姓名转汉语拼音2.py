import pandas as pd
from pypinyin import pinyin, Style

def convert_name_to_pinyin(name, employee_id, special_names=None, name_counts=None):
    if special_names is None:
        special_names = {}
    if name_counts is None:
        name_counts = {}
    
    # 检查员工ID是否在特殊列表中
    if employee_id in special_names:
        return special_names[employee_id]
    
    # 如果不在特殊列表中，按原来的方法处理
    surname, given_name = name[:1], name[1:]
    surname_pinyin = pinyin(surname, style=Style.NORMAL)[0][0]
    given_name_pinyin = ''.join([x[0] for x in pinyin(given_name, style=Style.NORMAL)])
    
    full_pinyin = f"{given_name_pinyin}.{surname_pinyin}"
    
    # 检查是否有同名情况并进行处理
    if name in name_counts:
        current_count = name_counts[name]
        full_pinyin += str(current_count)
        name_counts[name] += 1
    else:
        name_counts[name] = 1
    
    return full_pinyin

def process_excel(file_path):
    # 读取Excel文件
    df = pd.read_excel(file_path)
    
    # 假设第一列是中文姓名，第三列是员工ID
    special_names = {}
    name_counts = {}
    # 更新DataFrame，为每一行应用转换函数
    df['Pinyin'] = df.apply(lambda row: convert_name_to_pinyin(row.iloc[0], row.iloc[2], special_names, name_counts), axis=1)
    
    # 将结果写回Excel文件
    df.to_excel(file_path, index=False)

# 使用示例，替换'path_to_your_file.xlsx'为你的文件路径
process_excel("z:/姓名.xlsx")
