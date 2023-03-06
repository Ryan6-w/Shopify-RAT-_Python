import pandas as pd

# 读取 Excel 文件
excel_file = pd.ExcelFile('/Users/ryanweng/Documents/Cuppowood/website/产品导入/Color info_detail.xlsx')

# 读取 Sheet1 中的数据
T = pd.read_excel(excel_file, sheet_name='T')

# 将 Col2 列中的数据按照 Col1 列中的值进行分组，并将它们组合成列表
grouped_data = T.groupby('Name')['Code'].apply(list).reset_index(name='Code')

# 将每个大组中的数据打印出来
for index, row in grouped_data.iterrows():
    col1_value = row['Name']
    col2_values = row['Code']
    print(f'{col1_value}: {col2_values}')