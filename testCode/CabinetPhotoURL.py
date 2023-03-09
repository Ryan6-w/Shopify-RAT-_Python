import os
import glob
from openpyxl import Workbook


#读取橱柜照片的文件名; 定义文件路径和文件类型
photoPath = "/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/Cabinet/"
photoType = "*.jpg"
CabinetPhotoNames = []


s3Path= "https://s3.us-east-2.amazonaws.com/static.spaice.ca/share/cuppowood/Cabinet/"


# 获取符合条件的文件列表
cabinets = glob.glob(os.path.join(photoPath, photoType))
# 对获取到的橱柜产品进行循环查找
for cabinet in cabinets:
    # 获取橱柜产品照片文件名
    cabinetPhoto = os.path.basename(cabinet)
    # 把图片名字存储到准备好的array 里
    CabinetPhotoNames.append(cabinetPhoto)

workbook = Workbook()
worksheet = workbook.active
worksheet.cell(row=1, column=1, value='url')

for i,CabinetPhotoName in enumerate(CabinetPhotoNames):
    worksheet.cell(row=i+2,column=1,value=s3Path+CabinetPhotoName)


workbook.save('/Users/ryanweng/Documents/Cuppowood/Python/Testfiles/CabinetURL.xlsx')
