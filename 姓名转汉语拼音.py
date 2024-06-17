import pandas as pd
from pypinyin import pinyin, Style

def convert_name_to_pinyin(name):
    # 将姓名分开成姓和名
    surname, given_name = name[:1], name[1:]
    
    # 分别获取姓和名的拼音
    surname_pinyin = pinyin(surname, style=Style.NORMAL)[0][0]
    given_name_pinyin = ''.join([x[0] for x in pinyin(given_name, style=Style.NORMAL)])
    
    # 组合成所需的格式并返回
    return f"{given_name_pinyin}.{surname_pinyin}"


def convert_name_to_pinyin2(name, special_names=None):
    if special_names is None:
        special_names = {}
    
    # 检查名字是否在特殊列表中
    if name in special_names:
        return special_names[name]
    
    # 如果不在特殊列表中，按原来的方法处理
    surname, given_name = name[:1], name[1:]
    surname_pinyin = pinyin(surname, style=Style.NORMAL)[0][0]
    given_name_pinyin = ''.join([x[0] for x in pinyin(given_name, style=Style.NORMAL)])
    
    return f"{given_name_pinyin}.{surname_pinyin}"

# 特殊姓名列表示例
special_names = {
    "张三": "San.Zhang",
    "李四": "Si.Li"
}


def process_excel(file_path):
    # 读取Excel文件
    df = pd.read_excel(file_path)
    
    # 假设第一列是中文姓名，我们将其转换为拼音
    df['Pinyin'] = df.iloc[:, 0].apply(convert_name_to_pinyin2)
    
    # 将结果写回Excel文件
    df.to_excel(file_path, index=False)

# 使用示例，替换'path_to_your_file.xlsx'为你的文件路径
process_excel(f"z:\姓名.xlsx")
