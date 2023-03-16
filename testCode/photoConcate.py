import os
import subprocess

# 定义输入和输出路径
cabinetPath = "/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/PhotoConcate/Cabinet"
colorPath = "/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/PhotoConcate/Colors"
concatedPath = "/Users/ryanweng/Documents/Cuppowood/website/产品导入/Shopify/PhotoConcate/output"

# 用os.listdir 把文件全部列出来
cabinetName =[]
for n in os.listdir(cabinetPath):
    if n.endswith('.jpg'):
        cabinetName.append(n)

colorName =[]
for n in os.listdir(colorPath):
    if n.endswith('.jpg'):
        colorName.append(n)

# 获取输入目录中的所有照片文件地址, 用os.path.join 把path 拼接
cabinetFiles = [os.path.join(cabinetPath, f) for f in cabinetName ]
colorFiles = [os.path.join(colorPath, f) for f in colorName]

 

# 遍历每个文件，并拼接它们
for i in range(len(cabinetFiles)):
    # 获取当前两张照片的路径
    cabinet = cabinetName[i]

    # 用os.path.splitext 把extension 去掉
    cabinetNoExtension = os.path.splitext(cabinetName[i])[0]
    for j in range(len(colorFiles)):
        color = colorFiles[j]

        # 定义输出文件路径
        output_file = os.path.join(concatedPath, '{}'.format(cabinetNoExtension+"--"+colorName[j]))

        # 执行 ImageMagick 命令
        cmd = ['convert', cabinet, color, '-geometry', '+200+0', '-composite', output_file]

        print(output_file)
        # 设置 ImageMagick 命令
        # 调用ImageMagick命令行工具，将两张照片融合成一张
        # -gravity参数用于设置水平和垂直方向上的对齐方式
        # -composite参数用于将两张照片融合
        cmd = ["convert", cabinet, color,
            "-gravity", "center", "-composite",
            "-geometry", "+50+50", "-composite",
            concatedPath]    
        subprocess.call(cmd)
