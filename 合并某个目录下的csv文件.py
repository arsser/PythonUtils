#把某个目录下的csv文件合并到一个文件

import pandas as pd
import os

# 设置要合并的 CSV 文件所在的目录
directory = r'Z:\2023刷脸数据'

# 特定字符串开头
prefix = 'export'


# 找到目录中的所有 CSV 文件
#csv_files = [file for file in os.listdir(directory) if file.endswith('.csv')]
# 找到目录中以特定字符串开头的所有 CSV 文件
csv_files = [file for file in os.listdir(directory) if file.endswith('.csv') and file.startswith(prefix)]

# 创建一个空的 DataFrame 用于后续合并数据
combined_csv = pd.DataFrame()

# 遍历并合并每个 CSV 文件
for file in csv_files:
    # 打印当前处理的文件名
    print(f"Processing file: {file}")

    # 读取并追加到 combined_csv DataFrame
    file_path = os.path.join(directory, file)
    current_csv = pd.read_csv(file_path)
    combined_csv = pd.concat([combined_csv, current_csv])

# 设置合并后的 CSV 文件名
output_filename = 'combined_csv.csv'

# 将合并后的数据保存为新的 CSV 文件
combined_csv.to_csv(output_filename, index=False)

print(f'Combined CSV saved as {output_filename}')


exit




# 检查是否找到了符合条件的文件
if not csv_files:
    print(f"No CSV files starting with '{prefix}' found in the directory.")
else:
    # 读取并合并 CSV 文件
    
    combined_csv = pd.concat([pd.read_csv(os.path.join(directory, file)) for file in csv_files])

    # 设置合并后的 CSV 文件名
    output_filename = os.path.join(directory, 'combined_csv.csv')

    # 将合并后的数据保存为新的 CSV 文件
    combined_csv.to_csv(output_filename, index=False)

    print(f'Combined CSV saved as {output_filename}')
