import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
from urllib.parse import urlparse, urlunparse
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
    # 解析URL以获取和替换文件名部分
    parsed_url = urlparse(url)
    path_parts = parsed_url.path.split('/')
    filename = os.path.basename(parsed_url.path)
    if filename == '':
        filename = 'index'  # 如果URL以'/'结束，则使用'index'作为文件名
    hashed_filename = generate_hashed_filename(filename)
    path_parts[-1] = hashed_filename if filename != 'index' else filename
    new_path = '/'.join(path_parts)
    new_url = urlunparse(parsed_url._replace(path=new_path))

    # 创建二维码生成器
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(new_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color=qr_color, back_color='white').convert('RGBA')

    # 将二维码的白色背景转换为透明
    datas = img.getdata()
    new_data = []
    for item in datas:
        if item[0] == 255 and item[1] == 255 and item[2] == 255:  # 白色变透明
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)
    img.putdata(new_data)

    # 在二维码下方添加文本
    draw = ImageDraw.Draw(img)
    #font = ImageFont.load_default()  # 默认字体

    # 加载指定字号的字体，假设使用系统默认的 truetype 字体文件
    font_path = "C:\\Windows\\Fonts\\arial.ttf"  # 替换为您字体文件的路径
    font_size = 30  # 设置字体大小
    font = ImageFont.truetype(font_path, font_size)

    text_width, text_height = draw.textbbox((0, 0), text, font=font)[2:]
    new_height = img.height + text_height + 3
    new_img = Image.new('RGBA', (img.width, new_height), (255, 255, 255, 0))
    new_img.paste(img, (0, 0), img)

    text_x = (img.width - text_width) / 2
    text_y = img.height 
    draw = ImageDraw.Draw(new_img)
    draw.text((text_x, text_y), text, fill=qr_color, font=font)

    # 保存二维码图像
    final_filename = hashed_filename
    new_img.save(os.path.join(save_dir, final_filename))
    hashed_filenames.append(final_filename)

# 将哈希文件名写入DataFrame
df['Hashed Filename'] = hashed_filenames
df.to_excel(xls_path, index=False)

print("二维码生成完毕，保存在：" + save_dir)
