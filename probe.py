import os

from openpyxl import Workbook
from openpyxl.styles import Font
import win32com.client

path = os.path.join(os.getcwd(), 'Отчеты')

excel = win32com.client.Dispatch("Excel.Application")

if not os.path.exists(path):
    os.makedirs(path)
os.chdir(path)

filename = 'ВСТМ (ССК) Смена 1 (01.10.2020-01.10.2020).xlsx'
path = os.path.join(path, filename)
work_b1 = excel.Workbooks.Open(path)
sheet = work_b1.Worksheets(1)
Selection = sheet.Range('A1:A10')
xlNone = -4142
xlDiagonalDown = 5
xlDiagonalUp = 6
xlEdgeLeft = 7
xlEdgeTop = 8

Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
Selection.Borders(xlEdgeLeft).LineStyle = xlNone
Selection.Borders(xlEdgeTop).LineStyle = xlNone