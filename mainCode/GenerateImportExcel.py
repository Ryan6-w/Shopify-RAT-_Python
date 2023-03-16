import pandas as pd
from openpyxl import Workbook
import os

# # result[0] 为字母
# def substring_until_number(s):
#     result = ""
#     for i in s:
#         if s[0].isdigit():
#             return "DB"
#         if i.isdigit():
#             break
#         result += i
#     return result

# except base
def seperateSKU(s):
    numString = ""
    for i in s:
        if i.isnumeric():
            numString+=i
            
    return numString

colorPath = pd.ExcelFile('/Users/ryanweng/Documents/Cuppowood/website/产品导入/Adroit Stocked Color info.xlsx')
productPath = pd.ExcelFile('/Users/ryanweng/Documents/Cuppowood/website/产品导入/CNG_Cabinet_ Data.xlsx')
newExcelPath = '/Users/ryanweng/Documents/Cuppowood/Python/Testfiles/output.xlsx'

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
worksheet.cell(row=1, column=14, value='Image Position')  # not yet
worksheet.cell(row=1, column=15, value='Tags') 
worksheet.cell(row=1, column=16, value='Product Category') 
worksheet.cell(row=1, column=17, value='Type') 
worksheet.cell(row=1, column=18, value='Body (HTML)') 


# 如果没有价格那么价格是String, 有价格会是float 或者int
# for productRow in productList:
    # print(pTag(productRow[2]))


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
    
    width = numString[:2]
    height = numString[2:4]
    depth = numString[4:6]
    tempTitle = f"{width}\"{height}\"{depth}\" ({productRow[0]})"
    tempTag = f"W{width}, H{height}, D{depth}"
    
    pDes = "Depth: "+ depth +", Height: "+ height + ", Width: "+ width

    if tempSKU[3:5] == "EW":
        pType = "Wall Cabinet"
        
        if tempSKU[3:6]== "EWR":
            if int(depth) ==24:
                tempType = "Refrigerator Wall Cabinet"
                pTag = f"{tempType}, {tempTag}" 
                pTitle = f"{tempType} {tempTitle}"
            elif 30<=int(height)<=42:
                tempType = "High Wall Cabinet"
                pTag = f"{height}\" {tempType}, {tempTag}"
                pTitle = f"{height}\" {tempType} {tempTitle}"
            elif int(height)<30:
                tempType = "Standard Hight Wall Cabinet"
                pTag = f"{tempType}, {tempTag}"
                pTitle = f"{tempType} {tempTitle}"
            elif 48<=int(height):
                tempType = "Standing Wall Cabinet"
                pTag = f"{tempType}, {tempTag}"
                pTitle = f"{tempType} {tempTitle}"

        elif tempSKU[3:6] == "EWL":
            #K2,HX, HK
            if tempSKU[-2:] =="K2":
                tempType = "Standard Lift Up Door Wall Cabinet"
                pTag = f"{tempType}, {tempTag}"
                pTitle = f"{tempType} {tempTitle}"
            elif tempSKU[-2:] =="HX":
                tempType = "Lift Up Door Wall Cabinet"
                pTag = f"{tempType}, {tempTag}, HK-XS"
                pTitle = f"{tempType} with HK-XS {tempTitle}"
            elif tempSKU[-2:] =="HK":
                tempType = "Lift Up Door Wall Cabinet"
                pTag = f"{tempType}, {tempTag}, HK-Top"
                pTitle = f"{tempType} with HK-Top {tempTitle}"

        elif tempSKU[3:6] == "EWC":
            #DR, PR
            if tempSKU[-2:] =="DR":
                tempType = "Diagonal Corner Wall Cabinet"
                pTag = f"{tempType}, {tempTag}"
                pTitle = f"{tempType} {tempTitle}"
            elif tempSKU[-2:] =="PR":
                tempType = "Pie-Cut Corner Wall Cabinet"
                pTag = f"{tempType}, {tempTag}"
                pTitle = f"{tempType} {tempTitle}"
                
        elif tempSKU[3:6] == "EWB":
            tempType = "Blind Corner Wall Cabinet"
            pTag = f"{tempType}, {tempTag}"
            pTitle = f"{tempType} {tempTitle}"
        # print(f"tag: {pTag}")
        # print(f"title: {pTitle}")

    elif(productRow[9][3:5] == "EP"):
        pType = "Pantry"

        if tempSKU[-2:]!= "OV":
            tempType = f"{depth}\" Deep Pantry"
            if tempSKU[-2:] == "PT":
                pTag = f"{tempType}, {tempTag}" 
                pTitle = f"{tempType} {tempTitle}"
            elif tempSKU[-2:] == "R3":
                pTag = f"{tempType}, {tempTag}, 3RO" 
                pTitle = f"{tempType} (3RO) {tempTitle}"
            elif tempSKU[-2:] == "R4":
                pTag = f"{tempType}, {tempTag}, 4RO" 
                pTitle = f"{tempType} (4RO) {tempTitle}"
            elif tempSKU[-2:] == "FD":
                pTag = f"{tempType}, {tempTag}, FHD" 
                pTitle = f"{tempType} (FHD) {tempTitle}"
        elif tempSKU[-2:] =="OV":
            tempType = "Oven Pantry"
            pTag = f"{tempType}, {tempTag}" 
            pTitle = f"{tempType} {tempTitle}"
        
        # print(f"tag: {pTag}")
        # print(f"title: {pTitle}")

    elif(productRow[9][3:5] == "EB"):
        pType = "Base Cabinet"
        width = numString[:2]
        height = 34.5
        depth = 24
        tempTitle = f"{width}\" ({productRow[0]})"
        tempTag = f"W{width}"
        if tempSKU[3:6]== "EBD":
            tempType = "Drawer Base Cabinet"
            if tempSKU[-2:] == "W1":
                pTag = f"1 {tempType}, {tempType}, {tempTag}" 
                pTitle = f"1 {tempType} {tempTitle}"
            elif tempSKU[-2:] == "W2":
                pTag = f"2 {tempType}, {tempType}, {tempTag}" 
                pTitle = f"2 {tempType} {tempTitle}"
            elif tempSKU[-2:] == "T1":
                pTag = f"2 {tempType}, {tempType}, {tempTag}, Top 1RO" 
                pTitle = f"2 {tempType} (Top 1RO) {tempTitle}"
            elif tempSKU[-2:] == "W3":
                pTag = f"3 {tempType}, {tempType}, {tempTag}" 
                pTitle = f"3 {tempType} {tempTitle}"
            elif tempSKU[-2:] == "W4":
                pTag = f"4 {tempType}, {tempType}, {tempTag}" 
                pTitle = f"4 {tempType} {tempTitle}"
        elif tempSKU[3:6]== "EBC":
            tempType = "Corner Base Cabinet"
            if tempSKU[-2:] in ("DR","SR"):
                pTag = f"Diagonal {tempType}, {tempTag}" 
                pTitle = f"Diagonal {tempType} {tempTitle}"
            elif tempSKU[-2:] in ("PR","PW","PM"):
                pTag = f"Pie-Cut  {tempType}, {tempTag}" 
                pTitle = f"Pie-Cut  {tempType} {tempTitle}"
        elif tempSKU[3:6]== "EBB":
            if tempSKU[-2:]== "FD":
                tempType = "Blind Base Cabinet (FHD)"
                pTag = f"{tempType}, {tempTag}" 
                pTitle = f"{tempType} {tempTitle}"
            elif tempSKU[-2:]== "BB":
                tempType = "Blind Base Cabinet"
                pTag = f"{tempType}, {tempTag} " 
                pTitle = f"{tempType} (1 Drawer) {tempTitle}"
            elif tempSKU[-2:]== "SR":
                tempType = "Blind Sink Base Cabinet"
                pTag = f"{tempType}, {tempTag} " 
                pTitle = f"{tempType} {tempTitle}"
            elif tempSKU[-2:]== "SF":
                tempType = "Blind Sink Base Cabinet (FHD)"
                pTag = f"{tempType}, {tempTag} " 
                pTitle = f"{tempType} {tempTitle}"
        elif tempSKU[3:6]== "EBS":
            tempType = "Sink Base Cabinet"
            if tempSKU[-2:]== "BS":
                pTag = f"{tempType}, {tempTag}" 
                pTitle = f"{tempType} {tempTitle}"
            elif tempSKU[-4:]== "S-R1":
                pTag = f"{tempType} (1RO), {tempTag}" 
                pTitle = f"{tempType} (1RO) {tempTitle}"
            elif tempSKU[-4:]== "S-R2":
                pTag = f"{tempType} (2RO), {tempTag}" 
                pTitle = f"{tempType} (2RO) {tempTitle}"
            elif tempSKU[-2:]== "TT":
                pTag = f"{tempType} (Tilt Out), {tempTag}" 
                pTitle = f"{tempType} (Tilt Out) {tempTitle}"
            elif tempSKU[-2:]== "FD":
                pTag = f"{tempType} (FHD), {tempTag}" 
                pTitle = f"{tempType} (FHD) {tempTitle}"
            elif tempSKU[-4:]== "FDR1":
                pTag = f"{tempType} (FHD BOT 1RO), {tempTag} " 
                pTitle = f"{tempType} (FHD BOT 1RO) {tempTitle}"
            elif tempSKU[-4:]== "FDR2":
                pTag = f"{tempType} (FHD BOT 2RO), {tempTag} " 
                pTitle = f"{tempType} (FHD BOT 2RO) {tempTitle}"
            elif tempSKU[-2:]== "FS":
                pTag = f"Farm {tempType}, {tempTag} " 
                pTitle = f"Farm {tempType} {tempTitle}"
        elif tempSKU[3:6]== "EBR":
            tempType = "Base Cabinet"
            if tempSKU[-2:]== "BR":
                pTag = f"{tempType}, {tempTag}" 
                pTitle = f"{tempType} {tempTitle}"
            elif tempSKU[-2:]== "R1":
                pTag = f"{tempType} (1RO), {tempTag}" 
                pTitle = f"{tempType} (1RO) {tempTitle}"            
            elif tempSKU[-2:]== "R2":
                pTag = f"{tempType} (2RO), {tempTag}" 
                pTitle = f"{tempType} (2RO) {tempTitle}"
            elif tempSKU[-2:]== "GP":
                pTag = f"Pull-Out Basket {tempType}, {tempTag}" 
                pTitle = f"Pull-Out Basket {tempType} {tempTitle}"            
            elif tempSKU[-2:]== "HM":
                pTag = f"Hamper Basket {tempType}, {tempTag}" 
                pTitle = f"Hamper Basket {tempType} {tempTitle}"
            elif tempSKU[-2:]== "OV":
                pTag = f"Oven {tempType}, {tempTag}" 
                pTitle = f"Oven {tempType} {tempTitle}"            
            elif tempSKU[-2:]== "MW":
                pTag = f"Microwave {tempType}, {tempTag}" 
                pTitle = f"Microwave {tempType} {tempTitle}"
            elif tempSKU[-2:]== "KN":
                pTag = f"Knee Drawer Cabinet, W{width}, H{height}, D{depth}" 
                pTitle = f"Knee Drawer Cabinet {width}\"{height}\"{depth}\" ({productRow[0]})"
        elif tempSKU[3:6]== "EBF":
            tempType = "Base Cabinet"
            if tempSKU[-2:]== "BF":
                pTag = f"{tempType} (FHD), {tempTag}" 
                pTitle = f"{tempType} (FHD) {tempTitle}"
            elif tempSKU[-2:]== "T1":
                pTag = f"{tempType} (FHD Top 1RO), {tempTag}" 
                pTitle = f"{tempType} (FHD Top 1RO) {tempTitle}"   
            elif tempSKU[-2:]== "R1":
                pTag = f"{tempType} (FHD BOT 1RO), {tempTag}" 
                pTitle = f"{tempType} (FHD BOT 1RO) {tempTitle}"         
            elif tempSKU[-2:]== "R2":
                pTag = f"{tempType} (FHD 2RO), {tempTag}" 
                pTitle = f"{tempType} (FHD 2RO) {tempTitle}"
            elif tempSKU[-2:]== "R3":
                pTag = f"{tempType} (FHD 3RO), {tempTag}" 
                pTitle = f"{tempType} (FHD 3RO) {tempTitle}"
            elif tempSKU[-2:]== "GP":
                pTag = f"Pull-Out Basket {tempType} (FHD), {tempTag}" 
                pTitle = f"Pull-Out Basket {tempType} (FHD) {tempTitle}"   
            elif tempSKU[-2:]== "SP":
                pTag = f"Pull-Out Basket {tempType} (FHD Spice), {tempTag}" 
                pTitle = f"Pull-Out Basket {tempType} (FHD Spice) {tempTitle}"   
            elif tempSKU[-2:]== "HM":
                pTag = f"Hamper {tempType} (FHD), {tempTag}" 
                pTitle = f"Hamper {tempType} (FHD) {tempTitle}"       

        # print(f"tag: {pTag}")
        # print(f"title: {pTitle}")



    # worksheet.cell(row=insertRow, column=2, value=pTitle)
    # worksheet.cell(row=insertRow,column=7,value="active")
    # worksheet.cell(row=insertRow,column=13,value=productRow[1]) 
    # worksheet.cell(row=insertRow,column=15,value=pTag) 
    # worksheet.cell(row=insertRow,column=16,value="Furniture > Cabinets & Storage > Kitchen Cabinets") 
    # worksheet.cell(row=insertRow,column=17,value=pType) 
    # worksheet.cell(row=insertRow,column=18,value=pDes) 


#     for colorRow in colorsList:
#         worksheet.cell(row=insertRow, column=4, value=colorRow[0])
#         worksheet.cell(row=insertRow, column=1, value="Cuppowood-"+ str(productRow[0]))
#         worksheet.cell(row=insertRow,column=3,value="Material")

#         worksheet.cell(row=insertRow,column=5,value=str(productRow[0])+"-"+str(colorRow[1]))
#         if(colorRow[2] == 'A'):
#             price = round(productRow[2]+productRow[3],2)
#         elif (colorRow[2] == 'B'):
#             price = round(productRow[2]+productRow[4],2)
#         elif (colorRow[2] == 'C'):
#             price = round(productRow[2]+productRow[5],2)
#         elif (colorRow[2] == 'D'):
#             price = round(productRow[2]+productRow[6],2)
#         elif (colorRow[2] == 'E'):
#             price = round(productRow[2]+productRow[7],2)
#         elif (colorRow[2] == 'F'):
#             price = round(productRow[2]+productRow[8],2)
#         else:
#             price =0
#         worksheet.cell(row=insertRow,column=6,value= price)
#         worksheet.cell(row=insertRow,column=8,value="deny")
#         worksheet.cell(row=insertRow,column=9,value="manual")
#         worksheet.cell(row=insertRow,column=10,value="TRUE")
#         worksheet.cell(row=insertRow,column=11,value="TRUE")
#         worksheet.cell(row=insertRow,column=12,value="g")
#         insertRow +=1

# print("Total removed numbers are: "+ str(count))
# workbook.save(newExcelPath)
