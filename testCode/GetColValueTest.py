import pandas as pd

# 读取 Excel 文件
excel_file = pd.ExcelFile('/Users/ryanweng/Documents/Cuppowood/website/产品导入/Color info_detail.xlsx')

# 读取 Sheet1 中的数据

T = pd.read_excel(excel_file, sheet_name='T')

# 获取列1和列2的数据
TN = T['Name'].values
TC = T['Code'].values

# 打印列1和列2的数据
print('列1数据：')
print(TN)
print('\n')
print('列2数据：')
print(TC)

# 获取 Col1 和 Col2 列的值，并将它们存储在列表中
TN = T['Name'].tolist()
TC = T['Code'].tolist()

# 将 Col1 和 Col2 列的值循环打印出来
for i in range(len(TN)):
    print(f'Col1: {TN[i]}, Col2: {TC[i]}')
