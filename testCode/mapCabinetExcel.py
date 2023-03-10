import pandas as pd

mainFilePath = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/Cabinet_detail.xlsx'
updatedFilePath = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/产品信息/CNG_Cabinet_ Data.xlsx'
newFilePath = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/test.xlsx'

# 读取第一个Excel文件并选择第一列和要更新的列
df1 = pd.read_excel(mainFilePath, sheet_name='Demo')

# 读取第二个Excel文件并选择所有列
df2 = pd.read_excel(updatedFilePath)


# 方法1： 这个方法只是把更新的列存储到新文件里
# 将第一个文件中需要更新的列与第二个文件中的共同列进行匹配, on =共同列
#merged_df = pd.merge(df1[['CABINET','COMODO_BOX','A','B','C','D','E','F']], df2, on='CABINET', how='left')

# 方法2：索引列不会保留
# # 将需要映射的列设置为索引列
# df1.set_index('CABINET', inplace=True)
# df2.set_index('CABINET', inplace=True)
# # 将新数据更新到第二个Excel文件中
# df1.update(df2, overwrite=True)
# # 将更新后的数据保存为新的Excel文件
# df1.to_excel( newFilePath, index=False)

#方法3：
# 将需要映射的列设置为索引列
df1.set_index('CABINET', inplace=True)
df2.set_index('CABINET', inplace=True)

# 将索引列转为普通列，并添加一个新的索引列
df1.reset_index(inplace=True)
df2.reset_index(inplace=True)
df1.rename(columns={'CABINET': 'SKU'}, inplace=True)
df2.rename(columns={'CABINET': 'SKU'}, inplace=True)

# 将新数据更新到第二个Excel文件中
df1.update(df2, overwrite=True)

# 保存更新后的DataFrame对象到Excel文件中
with pd.ExcelWriter(newFilePath, engine='openpyxl') as writer:
    df1.to_excel(writer, index=False)




