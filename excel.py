import datetime
import os
import time
from pprint import pprint

from openpyxl import Workbook
from openpyxl.styles import Font
import win32com.client


class ExcelManager:
    def __init__(self):
        self.path = os.path.join(os.getcwd(), 'Отчеты')
        self.format_file = '.xlsx'

    def handler_excel(self, contractor, report_data):
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
        start_row = 8
        sheet.Cells(start_row, 2).Value = '№'
        sheet.Cells(start_row, 3).Value = 'Группировка'
        sheet.Cells(start_row, 3).ColumnWidth = 20
        sheet.Cells(start_row, 4).Value = 'Начало'
        sheet.Cells(start_row, 4).ColumnWidth = 20
        sheet.Cells(start_row, 5).Value = 'Конец'
        sheet.Cells(start_row, 5).ColumnWidth = 20
        sheet.Cells(start_row, 6).Value = 'Часы               (в Работе)'
        sheet.Cells(start_row, 7).Value = 'Часы (Дежурство)'
        sheet.Cells(start_row, 8).Value = 'Пробег'
        sheet.Cells(start_row, 9).Value = 'Часы               (в Работе) скорректир.'
        sheet.Cells(start_row, 10).Value = 'Часы (Дежурство) скорректир.'
        sheet.Cells(start_row, 11).Value = 'Пробег скоррект'

        # Selection = sheet.Range((start_row, 2), (start_row, 11))
        sheet.Cells(start_row, 2).EntireRow.HorizontalAlignment = -4108
        sheet.Cells(start_row, 2).EntireRow.WrapText = True

        Selection = sheet.Range('F1:K1')
        Selection.ColumnWidth = 12

        sheet.Cells(start_row, 12).Value = 'Нач. положение'
        sheet.Cells(start_row, 13).Value = 'Кон. положение'

        sheet.Cells(start_row, 12).ColumnWidth = 35
        sheet.Cells(start_row, 13).ColumnWidth = 35
        n = 0
        num_obj = 0
        for obj, data in report_data.items():
            num_obj += 1
            start_row += 1
            sheet.Cells(start_row, 2).Value = num_obj
            sheet.Cells(start_row, 3).Value = obj
            sheet.Cells(start_row, 6).Value = ''
            sheet.Cells(start_row, 7).Value = ''
            sheet.Cells(start_row, 8).Value = ''
            sheet.Cells(start_row, 9).Value = ''
            sheet.Cells(start_row, 10).Value = ''
            sheet.Cells(start_row, 11).Value = ''
            # sheet.Cells(start_row, 12).Value = data[0][10]
            # sheet.Cells(start_row, 13).Value = data[len(data)-1][11]

            obj_row = 1
            maxim = 99999999999999
            minim = 0
            # pprint(data)
            for number, travel in data.items():
                # pprint(travel)
                for num_con, cotr in sorted(travel.items()):
                    # pprint(cotr)

                    sheet.Cells(9, 4).Value = min(int(cotr[2]), maxim)
                    maxim = int(sheet.Cells(9, 4).Value)

                    sheet.Cells(9, 5).Value = max(int(cotr[3]), minim)
                    maxim = int(sheet.Cells(9, 5).Value)

                    start_row += 1
                    num_el = str(num_obj) + '.' + str(obj_row)
                    sheet.Cells(start_row, 2).NumberFormat = "@"
                    sheet.Cells(start_row, 2).HorizontalAlignment = -4152
                    sheet.Cells(start_row, 2).Value = num_el

                    sheet.Cells(start_row, 3).Value = f'({cotr[1]}) ' + str(num_con)

                    sheet.Cells(start_row, 4).Value = \
                        datetime.datetime.fromtimestamp(int(cotr[2])).strftime('%d.%m.%Y %H:%M:%S')
                    sheet.Cells(start_row, 5).Value = \
                        datetime.datetime.fromtimestamp(int(cotr[3])).strftime('%d.%m.%Y %H:%M:%S')

                    sheet.Cells(start_row, 6).Value = cotr[4]
                    sheet.Cells(start_row, 7).Value = cotr[5]
                    sheet.Cells(start_row, 8).Value = cotr[6]
                    sheet.Cells(start_row, 9).Value = round(float(cotr[7]), 0)
                    sheet.Cells(start_row, 10).Value = round(float(cotr[8]), 0)
                    sheet.Cells(start_row, 11).Value = round(float(cotr[9]), 0)
                    sheet.Cells(start_row, 12).Value = cotr[10]
                    sheet.Cells(start_row, 13).Value = cotr[11]
                    obj_row += 1
        #     sheet.Cells(start_row+int(data.key), 2).Value = data[1]

        sheet.Cells(9, 4).Value = \
            datetime.datetime.fromtimestamp(int(sheet.Cells(9, 4).Value)).strftime('%d.%m.%Y %H:%M:%S')
        sheet.Cells(9, 5).Value = \
            datetime.datetime.fromtimestamp(int(sheet.Cells(9, 5).Value)).strftime('%d.%m.%Y %H:%M:%S')

        work_b1.Save()
        work_b1.Close()
        excel.Quit()

