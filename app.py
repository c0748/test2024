import openpyxl
from datetime import datetime
from openpyxl.drawing.image import Image
import os

# Workbook()で新規作成
wb = openpyxl.load_workbook("files/PLOT用.xlsx", data_only=True)
ws = wb.active
values = list(ws.values)
lastrow = len(values)

# 全シートのデータを出力

print(values)
print(lastrow)

