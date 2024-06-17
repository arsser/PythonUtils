import pandas as pd
import os

def merge_excel_files_old1(directory):
    # 获取目录中所有Excel文件
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    all_data = []
    columns_consistent = True
    first_file_columns = None

    # 遍历文件并添加到DataFrame列表中
    for file in files:
        df = pd.read_excel(os.path.join(directory, file), sheet_name=0, header=1)  # 使用第二行作为标题行
        if first_file_columns is None:
            first_file_columns = df.columns
        elif not df.columns.equals(first_file_columns):
            columns_consistent = False
            print(f"列名不一致：文件 '{file}' 的列名与其他文件不同。")
            print(df.columns,first_file_columns)
            break
        all_data.append(df)

    if columns_consistent and all_data:
        # 合并所有DataFrame
        merged_df = pd.concat(all_data, ignore_index=True)

        # 保存到新的Excel文件，避免覆盖已存在的文件
        output_file = '合并.xlsx'
        counter = 1
        while os.path.exists(os.path.join(directory, output_file)):
            output_file = f'合并{counter}.xlsx'
            counter += 1

        merged_df.to_excel(os.path.join(directory, output_file), index=False)
        print(f"合并后的文件已保存为：{output_file}")
    elif not columns_consistent:
        print("由于列名不一致，文件未合并。")

def merge_excel_files(directory):
    # 获取目录中所有Excel文件
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx') and '合并' not in f]
    all_data = []
    columns_consistent = True
    first_file_columns = None
    first_file_name = None  # 用于存储第一个文件的文件名

    # 遍历文件并添加到DataFrame列表中
    for file in files:
        df = pd.read_excel(os.path.join(directory, file), sheet_name=0, header=1)  # 使用第二行作为标题行
        if first_file_columns is None:
            first_file_columns = df.columns
            first_file_name = file  # 记录第一个文件的文件名
        elif not df.columns.equals(first_file_columns):
            columns_consistent = False
            # 打印不一致的列名情况和第一个文件的名称
            different_columns = set(df.columns).symmetric_difference(first_file_columns)
            print(f"列名不一致：文件 '{file}' 的列名与第一个文件 '{first_file_name}' 的列名不同。不一致的列名包括: {different_columns}")
            break
        all_data.append(df)

    if columns_consistent and all_data:
        # 合并所有DataFrame
        merged_df = pd.concat(all_data, ignore_index=True)

        # 保存到新的Excel文件，避免覆盖已存在的文件
        output_file = '合并.xlsx'
        counter = 1
        while os.path.exists(os.path.join(directory, output_file)):
            output_file = f'合并{counter}.xlsx'
            counter += 1

        merged_df.to_excel(os.path.join(directory, output_file), index=False)
        print(f"合并后的文件已保存为：{output_file}")
    elif not columns_consistent:
        print("由于列名不一致，文件未合并。")

# 调用函数，参数为目标目录路径
merge_excel_files('D:\供应链\Emma\其他\emma yan\间采\供应商基本资料登记详细')
