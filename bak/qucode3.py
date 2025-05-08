import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
from urllib.parse import urlparse
import os
import hashlib

# 定义一个函数来生成哈希文件名
def generate_hashed_filename(original_name):
    hash_object = hashlib.sha256(original_name.encode())
    hex_dig = hash_object.hexdigest()
    return hex_dig[:10] + '.png'  # 返回哈希值的前10个字符作为文件名


# 指定XLS文件路径和保存二维码的目录
xls_path = '合作伙伴.xlsx'
save_dir = 'D:\\Git\\PythonUtils\\合作伙伴二维码'

# 读取XLS文件，假设URL在第一列，文本在第二列
df = pd.read_excel(xls_path)
urls = df.iloc[:, 4]  # URL在第一列
texts = df.iloc[:, 2]  # 文本在第二列

# 确保保存二维码的目录存在
if not os.path.exists(save_dir):
    os.makedirs(save_dir)

# 指定颜色
qr_color = '#0B579F'  # 二维码和文本的颜色
# 准备存储新文件名的列
hashed_filenames = []

# 遍历所有URL和文本
for url, text in zip(urls, texts):
    # 解析URL以生成合适的文件名
    parsed_url = urlparse(url)
    filename = os.path.basename(parsed_url.path)
    if filename == '':
        filename = 'index'  # 如果URL以'/'结束，则使用'index'作为文件名
    #filename = filename.replace('.', '_') + '.png'  # 替换文件名中的点，避免扩展名问题
    filename = generate_hashed_filename(filename)  
    
    # 创建二维码生成器
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color=qr_color, back_color='white').convert('RGB')

    # 在二维码下方添加文本
    draw = ImageDraw.Draw(img)
    font = ImageFont.load_default()  # 默认字体
    text_width, text_height = draw.textbbox((0, 0), text, font=font)[2:]
    new_height = img.height + text_height + 10
    new_img = Image.new('RGB', (img.width, new_height), 'white')
    new_img.paste(img, (0, 0))

    text_x = (img.width - text_width) / 2
    text_y = img.height + 5
    draw = ImageDraw.Draw(new_img)
    draw.text((text_x, text_y), text, fill=qr_color, font=font)

    # 保存二维码图像
    new_img.save(os.path.join(save_dir, filename))
    # 将哈希文件名保存到列表
    hashed_filenames.append(filename)

# 将哈希文件名写入DataFrame
df['Hashed Filename'] = hashed_filenames
df.to_excel(xls_path, index=False)
print("二维码生成完毕，保存在：" + save_dir)
