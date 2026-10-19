import openpyxl
from docxtpl import DocxTemplate
import datetime
import os

path  = os.path.dirname(os.path.abspath(__file__)) + "\\stu_dat.xlsx"
workbook = openpyxl.load_workbook(path)
sheet = workbook.active
list_val = list(sheet.values)
print(list_val)

doc = DocxTemplate("doc.docx")
doc.render({"message": 'Hi'})
doc.save("doc1.docx")
