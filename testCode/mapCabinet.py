import pandas as pd

mainFilePath = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/Cabinet_detail.xlsx'
updatedFilePath = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/产品信息/CNG_Cabinet_ Data.xlsx'
newFilePath = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/test.xlsx'

# 读取第一个Excel文件并选择第一列和要更新的列
df1 = pd.read_excel(mainFilePath, sheet_name='Demo')

# 读取第二个Excel文件并选择所有列
df2 = pd.read_excel(updatedFilePath,sheet_name='Demo')

# 将需要匹配的列"Cabinet"设置为索引列
df1.set_index('CABINET', inplace=True)
df2.set_index('CABINET', inplace=True)

# 将新Excel文件中的数据更新到旧Excel文件中
df1.update(df2, overwrite=True)

# 将更新后的DataFrame对象保存到Excel文件中，保留原有的Excel文件格式
with pd.ExcelWriter(newFilePath, engine='openpyxl') as writer:
    df1.to_excel(writer, sheet_name='Sheet1')