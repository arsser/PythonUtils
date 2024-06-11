# 用来合并windows某个目录下的excel文件的第一个sheet，需要检查列名，
# 如果发现列名不同，不要合并，先输出不一致的说明，否则进行合并，
# 合并后的文件保存在相同目录下，并命名为"合并.xlsx"，
# 如果该文件已存在，则自动命名为"合并1.xlsx"，以此类推

import os
import pandas as pd

def merge_excel_files_old1(directory):
    # 获取目录下所有的Excel文件
    excel_files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]

    if not excel_files:
        print("没有找到任何Excel文件。")
        return

    # 用于存储所有数据的DataFrame
    all_data = []
    column_names = None

    # 遍历所有Excel文件
    for file_name in excel_files:
        file_path = os.path.join(directory, file_name)
        try:
            # 读取第一个Sheet的数据
            data = pd.read_excel(file_path, engine='openpyxl', sheet_name=0)

            # 检查列名是否一致
            if column_names is None:
                column_names = data.columns.tolist()
            elif list(data.columns) != column_names:
                print(f"{file_name} 的列名与之前的文件不一致。")
                continue

            # 将数据添加到all_data列表中
            all_data.append(data)

        except Exception as e:
            print(f"读取 {file_name} 出错：{str(e)}")

    # 如果有数据要合并
    if all_data:
        merged_data = pd.concat(all_data, ignore_index=True)

        # 确定输出文件名
        output_file = "合并.xlsx"
        i = 1
        while os.path.exists(os.path.join(directory, output_file)):
            output_file = f"合并{i}.xlsx"
            i += 1

        # 保存合并后的数据
        merged_data.to_excel(os.path.join(directory, output_file), index=False, engine='openpyxl')
        print(f"文件已合并并保存为 {output_file}")
    else:
        print("没有可合并的文件。")
        
        
def merge_excel_files_old2(directory):
    # 获取目录下所有的Excel文件
    excel_files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]

    if not excel_files:
        print("没有找到任何Excel文件。")
        return

    # 用于存储所有数据的DataFrame
    all_headers = []  # 存储所有文件的首行
    all_data = []
    column_names = None

    # 遍历所有Excel文件
    for file_name in excel_files:
        file_path = os.path.join(directory, file_name)
        try:
            # 读取第一个Sheet的数据
            data = pd.read_excel(file_path, engine='openpyxl', sheet_name=0)

            # 把第一行加入到headers列表中
            header_row = data.iloc[0].to_frame().T
            all_headers.append(header_row)

            # 检查列名是否一致
            if column_names is None:
                column_names = data.columns.tolist()
            elif list(data.columns) != column_names:
                print(f"{file_name} 的列名与之前的文件不一致。")
                continue

            # 从第二行开始读取数据
            data = data.iloc[1:]
            all_data.append(data)

        except Exception as e:
            print(f"读取 {file_name} 出错：{str(e)}")

    # 如果有数据要合并
    if all_data:
        # 合并数据
        merged_data = pd.concat(all_data, ignore_index=True)

        # 合并headers
        headers_df = pd.concat(all_headers, ignore_index=True)

        # 将headers和data合并
        final_data = pd.concat([headers_df, merged_data], ignore_index=True)

        # 确定输出文件名
        output_file = "合并.xlsx"
        i = 1
        while os.path.exists(os.path.join(directory, output_file)):
            output_file = f"合并{i}.xlsx"
            i += 1

        # 保存合并后的数据
        final_data.to_excel(os.path.join(directory, output_file), index=False, engine='openpyxl')
        print(f"文件已合并并保存为 {output_file}")
    else:
        print("没有可合并的文件。")

def merge_excel_files(directory):
    # 获取目录下所有的Excel文件
    excel_files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]

    if not excel_files:
        print("没有找到任何Excel文件。")
        return

    # 用于存储所有数据的DataFrame
    all_data = []
    column_names = None

    # 遍历所有Excel文件
    for file_name in excel_files:
        file_path = os.path.join(directory, file_name)
        try:
            # 读取第一个Sheet的数据，忽略第一行，使用第二行作为标题行
            data = pd.read_excel(file_path, engine='openpyxl', sheet_name=0, skiprows=1, header=0)

            # 检查列名是否一致
            if column_names is None:
                column_names = data.columns.tolist()
            elif list(data.columns) != column_names:
                print(f"{file_name} 的列名与之前的文件不一致。")
                continue

            all_data.append(data)

        except Exception as e:
            print(f"读取 {file_name} 出错：{str(e)}")

    # 如果有数据要合并
    if all_data:
        merged_data = pd.concat(all_data, ignore_index=True)

        # 确定输出文件名
        output_file = "合并.xlsx"
        i = 1
        while os.path.exists(os.path.join(directory, output_file)):
            output_file = f"合并{i}.xlsx"
            i += 1

        # 保存合并后的数据
        merged_data.to_excel(os.path.join(directory, output_file), index=False, engine='openpyxl')
        print(f"文件已合并并保存为 {output_file}")
    else:
        print("没有可合并的文件。")
        
# 调用函数，传入你的目录路径
merge_excel_files('D:\供应链\Emma\其他\emma yan\间采\供应商基本资料登记详细')