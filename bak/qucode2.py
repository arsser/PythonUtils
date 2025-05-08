import pandas as pd
import qrcode
from urllib.parse import urlparse
import os

# 指定XLS文件路径和保存二维码的目录
xls_path = '合作伙伴.xls'
save_dir = 'D:\Git\PythonUtils\合作伙伴二维码'

# 读取XLS文件，假设URL在第一列
df = pd.read_excel(xls_path)
urls = df.iloc[:, 4]  # 调整列的索引，如果URL不在第一列

# 确保保存二维码的目录存在
if not os.path.exists(save_dir):
    os.makedirs(save_dir)

# 遍历所有URL
for url in urls:
    # 解析URL以生成合适的文件名
    parsed_url = urlparse(url)
    filename = os.path.basename(parsed_url.path)
    if filename == '':
        filename = 'index'  # 如果URL以'/'结束，则使用'index'作为文件名
    #filename = filename.replace('.', '_') + '.png'  # 替换文件名中的点，避免扩展名问题
    
    # 创建二维码生成器
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)

    # 生成二维码图像
    img = qr.make_image(fill='black', back_color='white')

    # 保存二维码图像
    img.save(os.path.join(save_dir, filename))

print("二维码生成完毕，保存在：" + save_dir)
