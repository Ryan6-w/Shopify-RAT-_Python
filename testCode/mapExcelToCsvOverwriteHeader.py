import pandas as pd

# 读取Excel文件
excel_file = pd.read_excel('/Users/ryanweng/Documents/Cuppowood/Python/output.xlsx')
excel_header = excel_file.columns.tolist()

# 读取CSV文件
csv_file = pd.read_csv('/Users/ryanweng/Documents/Cuppowood/Python/product_template.csv')
csv_header = csv_file.columns.tolist()

# 创建一个字典来存储header之间的映射关系
header_map = dict(zip(excel_header, csv_header))


# 将Excel文件的数据写入CSV文件
excel_file.to_csv('new_csv_file.csv', index=False)
