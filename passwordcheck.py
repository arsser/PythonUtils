import argon2

# 输入的明文密码
input_password = "admin887"

# 数据库中存储的哈希密码
stored_hash1 = "$argon2id$v=19$m=19456,t=2,p=1$Pe8/OTVMLtIHJtcCrGUXdw$QQtdLz1dYacWj7K6iAsf0Fdsv5jas4GCGoZfLFKJAwY"
stored_hash = "$argon2id$v=19$m=19456,t=2,p=1$ULaB0nihPc42DXcCzpUEJQ$397qccaxcxpumVI3XNtweZsis8ZGl/4flxvbZQix/kU"

# 创建 Argon2 密码验证器
ph = argon2.PasswordHasher()

# 验证密码
# 验证密码
try:
    is_valid = ph.verify(stored_hash, input_password)
    print(f"密码验证结果: {is_valid}")  # 如果验证成功,打印 True
except argon2.exceptions.VerifyMismatchError:
    print(f"密码验证结果: False")  # 如果验证失败,打印 False
