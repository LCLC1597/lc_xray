import os
import shutil
from openpyxl import load_workbook
workbook =load_workbook('./b.xlsx')
sheet = workbook.active
list1 = [i.value for i in sheet[2]]
name = "成品"+list1[1]+"/"+"/"+list1[1]+"/"+list1[2]+"/"+list1[3]+"/"+list1[4]
os.makedirs(name)
os.chdir(name)
for rows in sheet.rows:
    if rows[0].value == "档号":
        continue
    os.mkdir(rows[0].value)
    pos=



