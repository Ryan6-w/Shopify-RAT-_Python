import pandas as pd
from openpyxl import Workbook
import os
import re


def seperateSKU(s):
    numString = ""
    for i in s:
        if i.isnumeric():
            numString+=i
            
    return numString

def tagFormat(s):
    return s.replace(' ', '_').replace('(','').replace(')','').replace('\"','').lower()


colorPath = pd.ExcelFile('/Users/ryanweng/Documents/Cuppowood/website/产品导入/Adroit Stocked Color info.xlsx')
productPath = pd.ExcelFile('/Users/ryanweng/Documents/Cuppowood/website/产品导入/CNG_Cabinet_ Data.xlsx')
newExcelPath = '/Users/ryanweng/Documents/Cuppowood/Python/Testfiles/output.xlsx'

cabinetURL ="https://s3.us-east-2.amazonaws.com/static.spaice.ca/share/cuppowood/Cabinet/"
# just to get the photo name is could
cabinetPhotoPath ="/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/ConcatPhoto/"
cabinetPhotoName = os.listdir(cabinetPhotoPath)

# remove the excel file and csv file
if os.path.exists(newExcelPath):
    os.remove(newExcelPath)

# 读取第一个 Excel 文件，提取指定列的数据
colors = pd.read_excel(colorPath, usecols=['Color name','Panel Code','Price Level'])
# 读取第二个 Excel 文件，提取指定列的数据
products= pd.read_excel(productPath, sheet_name='demo1', usecols=['CABINET','URL','COMODO_BOX','A','B','C','D','E','F','SKU'])

# 指定要获取值的列名列表
colorsExtract = ['Color name','Panel Code','Price Level']
# 创建一个空列表，用于存储提取的值
colorsList = []
# 遍历每一行，提取指定列的值并添加到列表中；用iterrows 来遍历每一行，index为索引，row 为当前行数
for index, row in colors.iterrows():
    values = [row[columnHeader] for columnHeader in colorsExtract]
    colorsList.append(values)

# 指定要获取值的列名列表
productExtract = ['CABINET','URL','COMODO_BOX','A','B','C','D','E','F','SKU']
# 创建一个空列表，用于存储提取的值
productList = []
# 遍历每一行，提取指定列的值并添加到列表中；用iterrows 来遍历每一行，index为索引，row 为当前行数
for index, row in products.iterrows():
    values = [row[columnHeader] for columnHeader in productExtract]
    productList.append(values)


# 将字典写入到 Excel 文件中,我们使用 openpyxl 库将这个字典写入到一个新的 Excel 文件中，其中第一列包含第一个文件中的值，第二列包含第二个文件中的整个列。
workbook = Workbook()
worksheet = workbook.active

insertRow = 2 
price = count = depth = height = width = 0
pTitle = pTag = pType = pDes = tempSKU= ""


# productList index: 0=sku, 1= url, 2= box price, 3 = A ,4= B, 5=C, 6=D ,7=E ,8 =F , 9= kiSKU
# colorsList index: 0=name, 1 =code, 2= price level
for productRow in productList:
    if not isinstance(productRow[2],(int,float)):
        # print("Product with empty price is: " + str(productRow[0])) # 不要用Remove 因为list 是有序，会自动向上移,会少读一个产品
        productRow =[]
        count +=1
        continue
    
    if not isinstance(productRow[1],(str)):
        # print("Product with price but not photo: " + str(productRow[0])) # 不要用Remove 因为list 是有序，会自动向上移,会少读一个产品
        productRow =[]
        count +=1
        continue
    
    tempSKU = productRow[9]
    if productRow[0][0].isnumeric():
        numString = seperateSKU(productRow[0][1:])
    else:
        numString =seperateSKU(productRow[0])
    
    # 如果要改下面的，还需要改其他2个地方，一个是base cabinet 一个是knee drawer
    width = numString[:2]
    height = numString[2:4]
    depth = numString[4:6]
    tempTitle = f"{width}\"W {height}\"H {depth}\"D ({productRow[0]})"
    tempTag = f"{width}W, {height}H, D{depth}D"
    pDes = "Depth: "+ depth +", Height: "+ height + ", Width: "+ width

    if(productRow[9][3:5] == "EB"):
        pType = "Base Cabinet"
        width = numString[:2]
        height = "34.5"
        depth = "24"    

    for colorRow in colorsList:

        # width大于或者等于24 为doubledoor, 小于24为Singledoor
        photoSKU = productRow[9][3:].replace('-', '')
        tempColor = colorRow[0].replace(' ','').replace('-','')

        # selfSKU-kiSKU--color

        # .+任意字符串
        # (?P<name>pattern) =》以下语法来创建命名捕获组
        # 使用了非捕获组 (?:_SINGLEDOOR)? 和 (?:_DOUBLEDOOR)?，表示它们是可选的，即可能存在也可能不存在
        
        width = int(width)
        for cName in cabinetPhotoName:
            if width < 24 :        
                pattern = re.compile(rf".+-{photoSKU}(?:_SINGLEDOOR)?--{tempColor}")
                if re.match(pattern, cName):
                    print(f"{cName} matches the pattern {pattern} and sku are {productRow[0]}\n")
            if width >= 24 :        
                pattern = re.compile(rf".+-{photoSKU}(?:_DOUBLEDOOR)?--{tempColor}")
                if re.match(pattern, cName):
                    print(f"{cName} matches the pattern {pattern} and sku are {productRow[0]}\n")
 

            