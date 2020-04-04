#Chosing a specific cell and change
"""
/Font Color
/Bold text
/italic text
background color
border style
border color
text alignment

"""

import openpyxl
from openpyxl.styles import *

workbook = openpyxl.load_workbook(r"C:\Users\walee\OneDrive\Desktop\Python\Excel_Tasks\original.xlsx")

sheet = workbook['EmployeeData']

cell = sheet['B8']

font = Font(color=colors.GREEN, bold=True, italic=True)

cell.font = font

fill = PatternFill(patternType='solid', bgColor="F7FE2E")

cell.fill = fill

border = Border(left= Side(border_style='double',color="FF000000"),right=Side(border_style='double',color="FF000000")
,top=Side(border_style='double',color="FF000000"),bottom=Side(border_style='double',color="FF000000"))

cell.border = border

alignment = Alignment(horizontal='right')

cell.alignment = alignment

workbook.save(r"C:\Users\walee\OneDrive\Desktop\Python\Excel_Tasks\original.xlsx")
