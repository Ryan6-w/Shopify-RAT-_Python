import cv2
import numpy as np


# 设置两张照片的路径
image1_path = "/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/PhotoConcate/output/1DB-EBDW1.jpg"
image2_path = "/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/PhotoConcate/output/Ida012S.jpg"

# 设置输出图片的路径和名称
output_path = "/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/PhotoConcate/output/output.jpg"


count = 0
taskname ='makeup'

def cv_imread(file_path):
    cv_img = cv2.imdecode(np.fromfile(file_paht,dtype=np.uint8),-1)
    return cv_img


# # 将第二张照片放在第一张照片的右上角
# rows, cols, channels = img2.shape
# roi = img1[0:rows, img1.shape[1]-cols:img1.shape[1]]
# img2gray = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
# ret, mask = cv2.threshold(img2gray, 10, 255, cv2.THRESH_BINARY)
# mask_inv = cv2.bitwise_not(mask)
# img1_bg = cv2.bitwise_and(roi,roi,mask = mask_inv)
# img2_fg = cv2.bitwise_and(img2,img2,mask = mask)
# dst = cv2.add(img1_bg,img2_fg)
# img1[0:rows, img1.shape[1]-cols:img1.shape[1]] = dst

# 保存输出图片
cv2.imwrite('path/to/output.png', img1)

 