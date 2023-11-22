import chardet



def detect_encoding(file_path):
    with open(file_path, 'rb') as file:
        return chardet.detect(file.read())['encoding']


# 示例使用
file = r'Z:\2023刷脸数据\exported_data_2023-01-01_to_2023-02-01.csv'
encoding = detect_encoding(file)
print(encoding)