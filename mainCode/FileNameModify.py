import os

# 更改文件名函数
def rename_files(path):
    # 获取目录下所有文件列表
    file_list = os.listdir(path)
    # 循环处理每个文件
    for file_name in file_list:
        # 如果文件名中包含空格
        if ' ' in file_name:
            # 去掉空格，或者可以替换成其他符号
            new_name = file_name.replace(' ', '')
            # 构造新的文件路径
            old_path = os.path.join(path, file_name)
            new_path = os.path.join(path, new_name)
            # 更改文件名
            os.rename(old_path, new_path)
            print(f'{file_name}已更名为{new_name}')

# 测试
if __name__ == '__main__':
    # 修改该路径为你想要更改文件名的目录
    path = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/xxx'
    rename_files(path)
