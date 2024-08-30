import hashlib

def generate_hashed_filename(original_name):
    hash_object = hashlib.sha256(original_name.encode())
    hex_dig = hash_object.hexdigest()
    return hex_dig[:10]  # 返回哈希值的前10个字符作为文件名

# 示例使用
original_filenames = ["NO_MINDV_G2024001", "NO_MINDV_G2024002", "NO_MINDV_G2024003"]
hashed_filenames = [generate_hashed_filename(name) for name in original_filenames]

print(hashed_filenames)
