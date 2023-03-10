import pandas as pd
from openpyxl import Workbook
import os
import glob
import re


color = pd.ExcelFile('/Users/ryanweng/Documents/Cuppowood/website/产品导入/Adroit Stocked Color info.xlsx')
sku = pd.ExcelFile('/Users/ryanweng/Documents/Cuppowood/website/产品导入/CNG_Cabinet_ Data.xlsx')
# s3Path= "https://s3.us-east-2.amazonaws.com/static.spaice.ca/share/cuppowood/Cabinet/"

# 读取第一个 Excel 文件，提取指定列的数据
cName = pd.read_excel(color, usecols=['Color name','Panel Code ( BFM )','Price Level'])

# 读取第二个 Excel 文件，提取指定列的数据
pSku= pd.read_excel(sku, sheet_name='demo', usecols=['CABINET','URL','COMODO_BOX','A','B','C','D','E','F'])

# #读取橱柜照片的文件名; 定义文件路径和文件类型
# photoPath = "/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/Cabinet/"
# photoType = "*.jpg"
# CabinetPhotoNames = []

# # 获取符合条件的文件列表
# cabinets = glob.glob(os.path.join(photoPath, photoType))
# # 对获取到的橱柜产品进行循环查找
# for cabinet in cabinets:
#     # 获取橱柜产品照片文件名
#     cabinetPhoto = os.path.basename(cabinet)
#     # 把图片名字存储到准备好的array 里
#     CabinetPhotoNames.append(cabinetPhoto)

colors = []
for i, color in enumerate(cName['Name']):
    colors.append(color)

codes =[]
for i, code in enumerate(cName['Code']) :
    codes.append(code)

skus = []
for i, sku in enumerate(pSku['SKU']):
    skus.append(sku)

tPs =[]
for i, tp in enumerate(pSku['T']):
    tPs.append(tp)

ePs =[]
for i, ep in enumerate(pSku['E']):
    ePs.append(ep)

bPs =[]
for i, bp in enumerate(pSku['B']):
    bPs.append(bp)

a1Ps =[]
for i, a1p in enumerate(pSku['A1']):
    a1Ps.append(a1p)

a2Ps =[]
for i, a2p in enumerate(pSku['A2']):
    a2Ps.append(a2p)

a3Ps =[]
for i, a3p in enumerate(pSku['A3']):
    a3Ps.append(a3p)

urls =[]
for i,url in enumerate(pSku['URL']):
    urls.append(url) 


# 将字典写入到 Excel 文件中,我们使用 openpyxl 库将这个字典写入到一个新的 Excel 文件中，其中第一列包含第一个文件中的值，第二列包含第二个文件中的整个列。
workbook = Workbook()
worksheet = workbook.active
worksheet.cell(row=1, column=1, value='Handle')
worksheet.cell(row=1, column=2, value='Title')
worksheet.cell(row=1, column=3, value='Option1 Name')
worksheet.cell(row=1, column=4, value='Option1 Value') 
worksheet.cell(row=1, column=5, value='Variant SKU') 
worksheet.cell(row=1, column=6, value='Variant Price') 
worksheet.cell(row=1, column=7, value='Status') 
worksheet.cell(row=1, column=8, value='Variant Inventory Policy') 
worksheet.cell(row=1, column=9, value='Variant Fulfillment Service') 
worksheet.cell(row=1, column=10, value='Variant Requires Shipping') 
worksheet.cell(row=1, column=11, value='Variant Taxable') 
worksheet.cell(row=1, column=12, value='Variant Weight Unit') 
worksheet.cell(row=1, column=13, value='Image Src') 
worksheet.cell(row=1, column=14, value='Image Position') 



skuRow =2 
codeRow =2
urlIndex =0
for i,sku in enumerate(skus):
    worksheet.cell(row=skuRow, column=2, value=sku)
    worksheet.cell(row=skuRow,column=7,value="active")
    worksheet.cell(row=skuRow,column=14,value=1) 
    worksheet.cell(row=skuRow,column=13,value=urls[urlIndex]) 
    urlIndex = urlIndex+1
    for j, color in enumerate(colors):
        worksheet.cell(row=skuRow, column=4, value=color)
        worksheet.cell(row=skuRow,column=1,value="Cuppowood-"+ sku)
        worksheet.cell(row=skuRow,column=3,value="Material")
        worksheet.cell(row=skuRow,column=8,value="deny")
        worksheet.cell(row=skuRow,column=9,value="manual")
        worksheet.cell(row=skuRow,column=10,value="TRUE")
        worksheet.cell(row=skuRow,column=11,value="TRUE")
        worksheet.cell(row=skuRow,column=12,value="g")


        skuRow +=1
    for x, code in enumerate(codes):
        worksheet.cell(row=codeRow,column=5,value=sku + "-" + code)
        codeRow +=1

# T 的价格是2-10， E:11-18, B: 19-30 , a1：31-34, a2: 35-38, a3:39-41 ; 下一次SKU就是第一次价格出现+40
tRow = 2
for t, tp in enumerate(tPs):
    for i in range(tRow,tRow+9):
        worksheet.cell(row=i, column=6, value=tp)  
    tRow +=40

eRow = 11
for t, ep in enumerate(ePs):
    for i in range(eRow,eRow+8):
        worksheet.cell(row=i, column=6, value=ep)
    eRow +=40

bRow = 19
for t, bp in enumerate(bPs):
    for i in range(bRow,bRow+12):
        worksheet.cell(row=i, column=6, value=bp)
    bRow += 40

a1Row = 31
for t, a1p in enumerate(a1Ps):
    for i in range(a1Row,a1Row+4):
        worksheet.cell(row=i, column=6, value=a1p)
    a1Row += 40

a2Row = 35
for t, a2p in enumerate(a2Ps):
    for i in range(a2Row,a2Row+4):
        worksheet.cell(row=i, column=6, value=a2p)
    a2Row += 40

a3Row = 39
for t, a3p in enumerate(a3Ps):
    for i in range(a3Row,a3Row+3):
        worksheet.cell(row=i, column=6, value=a3p)
    a3Row += 40


workbook.save('/Users/ryanweng/Documents/Cuppowood/Python/Testfiles/output.xlsx')
