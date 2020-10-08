import os
from openpyxl import Workbook
from openpyxl.styles import Font
import win32com.client


class ExcelManager:
    def __init__(self):
        self.path = os.path.join(os.getcwd(), 'Отчеты')
        self.format_file = '.xlsx'

    def handler_excel(self, contractor):
        if not os.path.exists(self.path):
            os.makedirs(self.path)
        os.chdir(self.path)
        filename = contractor + self.format_file
        path = os.path.join(self.path, filename)
        wb = Workbook()
        wb.create_sheet('Сводка', 0)
        ws = wb.active
        # ws.auto_filter.ref = 'A2:AD2'
        wb.save(filename)
        excel = win32com.client.Dispatch("Excel.Application")
        work_b1 = excel.Workbooks.Open(path)
        sheet = work_b1.Worksheets(1)
        sheet.Cells(1, 5).Value = 'РЕЕСТР выполненных работ за период'

        sheet.Cells(1, 5).Font.Size = 12
        sheet.Cells(1, 5).Font.Bold = True

        Selection = sheet.Range('C3:G6')
        Selection.Font.Size = 10
        Selection.Font.Bold = True
        # Selection.Font = 'Times New Roman'
        Selection = sheet.Range('A1:M1000')
        Selection.font.Name = 'Times New Roman'


        sheet.Cells(3, 3).Value = 'Параметры:'
        sheet.Cells(4, 3).Value = 'период с'
        sheet.Cells(4, 4).Value = 'ДАТА С'
        sheet.Cells(4, 4).HorizontalAlignment = -4108
        sheet.Cells(5, 3).Value = 'период по'
        sheet.Cells(5, 4).Value = 'ДАТА ПО'
        sheet.Cells(5, 4).HorizontalAlignment = -4108
        sheet.Cells(6, 3).Value = 'подрядчик'

        Selection = sheet.Range('C4:C6')
        Selection.HorizontalAlignment = -4152

        sheet.Cells(4, 5).Value = 'дневная смена'
        sheet.Cells(5, 5).Value = 'ночная смена'
        sheet.Cells(6, 5).Value = 'разрывная смена'

        Selection = sheet.Range('E4:E6')
        Selection.HorizontalAlignment = -4152

        sheet.Cells(4, 6).Value = '08:00 - 19:59'
        sheet.Cells(5, 6).Value = '20:00 - 07:59'
        sheet.Cells(6, 6).Value = '08:00 - 01:00'

        sheet.Cells(4, 7).Value = '11 часовой режим'
        sheet.Cells(5, 7).Value = '11 часовой режим'
        sheet.Cells(6, 7).Value = '16 часовой режим'
        sheet.Cells(1, 1).ColumnWidth = 3
        sheet.Cells(1, 2).ColumnWidth = 8

        Selection = sheet.Range('C1:F1')
        Selection.ColumnWidth = 15

        # -4108 центр
        # -4152 право

        work_b1.Save()
        work_b1.Close()
        excel.Quit()

