import qrcode
from PIL import Image

# 创建二维码
qr = qrcode.QRCode(
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_L,
    box_size=10,
    border=4
)
qr.add_data('https://www.example.com')
qr.make(fit=True)

# 创建一个带白色背景的二维码图像
img = qr.make_image(fill_color="black", back_color="white")

# 将PIL图像对象转换为RGBA模式（支持透明度）
img = img.convert("RGBA")

# 获取图像数据
datas = img.getdata()

# 修改图像数据中的背景色（白色）为透明
new_data = []
for item in datas:
    # 改变所有白色（也就是背景色）到透明色
    if item[0] == 255 and item[1] == 255 and item[2] == 255:
        new_data.append((255, 255, 255, 0))
    else:
        new_data.append(item)

# 更新图像数据
img.putdata(new_data)

# 保存图像
img.save('transparent_qrcode.png')

# 显示图像
img.show()
