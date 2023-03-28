import csv
import pandas as pd
from openpyxl import load_workbook
import os

tempCSVPath = '/Users/ryanweng/Documents/Cuppowood/website/产品导入/product_template.csv'
excelPath = '/Users/ryanweng/Documents/Cuppowood/Python/Testfiles/FailedItems.xlsx'
newCSVpath = '/Users/ryanweng/Documents/Cuppowood/Python/Testfiles/FailedItems.csv'

if os.path.exists(newCSVpath):
    os.remove(newCSVpath)


# 打开 csv 文件并读取 header
with open(tempCSVPath, 'r') as f:
    reader = csv.reader(f)
    header = next(reader)

# 打开 Excel 文件
wb = load_workbook(excelPath)

# 获取第一个 sheet
ws = wb.active

# 将 Excel 数据读取为 DataFrame
df = pd.DataFrame(ws.values)

# 获取 Excel 数据的 header
excel_header = list(df.iloc[0])

# 将 Excel 数据按照 CSV header 的顺序整理
data = []
for row in df.iloc[1:].values:
    d = {}
    for i, value in enumerate(row):
        d[excel_header[i]] = value
    data.append(d)

# 打开 CSV 文件并写入数据
with open(newCSVpath, 'a', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=header)
    if f.tell() == 0:
        # CSV 文件没有 header，写入 header
        writer.writeheader()
    # 写入数据
    writer.writerows(data)

