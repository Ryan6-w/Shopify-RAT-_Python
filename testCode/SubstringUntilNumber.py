import re
import pandas as pd
from openpyxl import Workbook


productPath = pd.ExcelFile('/Users/ryanweng/Documents/Cuppowood/website/产品导入/CNG_Cabinet_ Data.xlsx')

products= pd.read_excel(productPath, sheet_name='demo1', usecols=['CABINET','URL','COMODO_BOX','A','B','C','D','E','F'])

# s = "Hello123World"
# result = re.match(r'\D*', s).group()
# print(result)

def substring_until_number(s):
    result = ""
    for i in s:
        if s[0].isdigit():
            return "DB"
        if i.isdigit():
            break
        result += i
    return result

# 指定要获取值的列名列表
productExtract = ['CABINET','URL','COMODO_BOX','A','B','C','D','E','F']
# 创建一个空列表，用于存储提取的值
productList = []
# 遍历每一行，提取指定列的值并添加到列表中；用iterrows 来遍历每一行，index为索引，row 为当前行数
for index, row in products.iterrows():
    values = [row[columnHeader] for columnHeader in productExtract]
    productList.append(values)

insertRow = 2 
price =0
count =0
# productList index: 0=sku, 1= url, 2= box price, 3 = A ,4= B, 5=C, 6=D ,7=E ,8 =F 
# colorsList index: 0=name, 1 =code, 2= price level
for productRow in productList:
    result= substring_until_number(str(productRow[0]))
    print(result)