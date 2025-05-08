import qrcode
from PIL import Image

# URL链接
url = 'http://pbi.yitu-inc.com/img/305d4ea8_E815772_a271014e.png'
url = 'https://www.yitutech.com/sites/default/files/partner/xxxx.png'
# 创建二维码生成器
qr = qrcode.QRCode(
    version=1,  # 控制二维码的大小，1表示最小的21x21格子
    error_correction=qrcode.constants.ERROR_CORRECT_L,  # 错误修正能力，L表示约可纠正7%的错误
    box_size=10,  # 每个格子的像素大小
    border=4,  # 边框的格子厚度，默认是4
)

# 添加数据
qr.add_data(url)
qr.make(fit=True)

# 生成二维码图像
img = qr.make_image(fill='black', back_color='white')

# 显示图像（如果在图形界面环境中）
img.show()

# 保存图像
img.save("url_qrcode.png")
