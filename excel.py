import collections
import datetime
import os
import time
from pprint import pprint

from openpyxl import Workbook
from openpyxl.styles import Font
import win32com.client


class ExcelManager:
    def __init__(self):
        self.format_file = '.xlsx'

    def handler_excel(self, contractor, report_data, smena, f_date, t_date, path, company=None):
        if company:
            contractor = f'{company} - {contractor} '
        filename = contractor + ' ' + smena + f' ({f_date}-{t_date})' + self.format_file
        path = os.path.join(path, filename)

        wb = Workbook()
        wb.create_sheet('Сводка', 0)
        ws = wb.active
        ws.auto_filter.ref = 'B8:M8'
        wb.save(filename)

        excel = win32com.client.Dispatch("Excel.Application")
        work_b1 = excel.Workbooks.Open(path)
        sheet = work_b1.Worksheets(1)
        sheet.Cells(1, 5).Value = 'РЕЕСТР выполненных работ за период'

        Selection = sheet.Range('A1:N7')
        Selection.Interior.Color = 16777215
        Selection.Interior.TintAndShade = 0
        Selection.Interior.PatternTintAndShade = 0

        sheet.Cells(1, 5).Font.Size = 12
        sheet.Cells(1, 5).Font.Bold = True

        Selection = sheet.Range('C3:G6')
        Selection.Font.Size = 10
        Selection.Font.Bold = True
        Selection = sheet.Range('A1:M500')
        Selection.font.Name = 'Times New Roman'
        Selection = sheet.Range('B8:M500')
        Selection.Font.Size = 10
        sheet.Cells(3, 3).Value = 'Параметры:'
        sheet.Cells(4, 3).Value = 'период с'
        sheet.Cells(4, 4).Value = f_date
        sheet.Cells(4, 4).HorizontalAlignment = -4108
        sheet.Cells(5, 3).Value = 'период по'
        sheet.Cells(5, 4).Value = t_date
        sheet.Cells(5, 4).HorizontalAlignment = -4108
        sheet.Cells(6, 3).Value = 'подрядчик'
        sheet.Cells(6, 4).Value = contractor.replace('(ССК)', '').replace('(ССК-РС)', '').replace('(ССК-Т)', '')
        sheet.Cells(6, 4).HorizontalAlignment = -4108
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

        Selection = sheet.Range('B8:M8')
        Selection.Interior.Color = 14803425
        Selection.Font.Bold = True

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
        sheet.Cells(start_row, 11).ColumnWidth = 15
        sheet.Cells(start_row, 2).EntireRow.HorizontalAlignment = -4108
        sheet.Cells(start_row, 2).EntireRow.WrapText = True

        Selection = sheet.Range('F1:K1')
        Selection.ColumnWidth = 12

        sheet.Cells(start_row, 12).Value = 'Нач. положение'
        sheet.Cells(start_row, 13).Value = 'Кон. положение'

        sheet.Cells(start_row, 12).ColumnWidth = 45
        sheet.Cells(start_row, 13).ColumnWidth = 45
        num_obj = 0
        time.sleep(30)
        for obj, data in report_data.items():
            sum_work = 0
            sum_duty = 0
            sum_mill = 0
            num_obj += 1
            start_row += 1
            main_row = start_row
            sheet.Cells(main_row, 2).Value = num_obj
            sheet.Cells(main_row, 3).Value = obj

            sheet.Cells(main_row, 6).Value = ''
            sheet.Cells(main_row, 7).Value = ''
            sheet.Cells(main_row, 8).Value = ''
            sheet.Cells(main_row, 9).Value = ''
            sheet.Cells(main_row, 10).Value = ''
            sheet.Cells(main_row, 11).Value = ''

            Selection = sheet.Range(f'B{main_row}:M{main_row}')
            Selection.Interior.Color = 49407
            Selection.Font.Bold = True
            obj_row = 1
            data_sort = sorted(data.items(), key=lambda x: x[0])
            data_order = collections.OrderedDict(data_sort)
            for number, travel in data_order.items():
                start_row += 1
                num_el = str(num_obj) + '.' + str(obj_row)
                sheet.Cells(start_row, 2).NumberFormat = "@"
                sheet.Cells(start_row, 2).HorizontalAlignment = -4152
                sheet.Cells(start_row, 2).Value = num_el
                sheet.Cells(start_row, 3).Value = travel[1]

                sheet.Cells(start_row, 4).Value = \
                    datetime.datetime.fromtimestamp(int(travel[2])+14400).strftime('%d.%m.%Y %H:%M:%S')

                sheet.Cells(start_row, 5).Value = \
                    datetime.datetime.fromtimestamp(int(travel[3])+14400).strftime('%d.%m.%Y %H:%M:%S')

                sheet.Cells(start_row, 6).Value = travel[4]
                sheet.Cells(start_row, 7).Value = travel[5]
                sheet.Cells(start_row, 8).Value = travel[6]
                sheet.Cells(start_row, 9).Value = round(float(travel[7]), 0)
                sheet.Cells(start_row, 10).Value = round(float(travel[8]), 0)
                sheet.Cells(start_row, 11).Value = round(float(travel[9]), 0)

                sum_work += round(float(travel[7]), 0)
                sum_duty += round(float(travel[8]), 0)
                sum_mill += round(float(travel[9]), 0)

                sheet.Cells(start_row, 12).Value = travel[10]
                sheet.Cells(start_row, 13).Value = travel[11]
                obj_row += 1

            sheet.Cells(main_row, 4).Value = sheet.Cells(main_row + 1, 4).Value
            sheet.Cells(main_row, 5).Value = sheet.Cells(start_row, 5).Value

            sheet.Cells(main_row, 12).Value = sheet.Cells(main_row + 1, 12).Value
            sheet.Cells(main_row, 13).Value = sheet.Cells(start_row, 13).Value

            sheet.Cells(main_row, 6).FormulaR1C1 = f"=SUM(R[1]C:R[{obj_row-1}]C)"
            sheet.Cells(main_row, 7).FormulaR1C1 = f"=SUM(R[1]C:R[{obj_row - 1}]C)"
            sheet.Cells(main_row, 8).FormulaR1C1 = f"=SUM(R[1]C:R[{obj_row - 1}]C)"

            Selection = sheet.Range(sheet.Cells(main_row + 1, 4), sheet.Cells(start_row, 5))
            Selection.Rows.Group()

            sheet.Cells(main_row, 9).Value = sum_work
            sheet.Cells(main_row, 10).Value = sum_duty
            sheet.Cells(main_row, 11).Value = sum_mill

        Selection = sheet.Range('B8:M' + str(start_row))
        Selection.Borders.Weight = 2

        sheet.Cells(start_row+2, 4).Value = 'Исполнитель'

        sheet.Cells(start_row + 2, 4).HorizontalAlignment = -4152
        sheet.Cells(start_row + 2, 4).RowHeight = 20

        sheet.Cells(start_row + 3, 5).Value = 'подпись'
        sheet.Cells(start_row + 3, 5).HorizontalAlignment = -4108

        sheet.Cells(start_row + 3, 6).Value = 'ФИО'

        Selection = sheet.Range(f'F{start_row+3}:G{start_row+3}')
        Selection.HorizontalAlignment = -4108
        Selection.Merge()

        sheet.Cells(start_row + 3, 9).Value = 'дата'
        sheet.Cells(start_row + 3, 9).HorizontalAlignment = -4108

        sheet.Cells(start_row + 4, 4).Value = 'Согласовано:'

        sheet.Cells(start_row + 6, 4).Value = 'Представитель Заказчика'
        sheet.Cells(start_row + 6, 4).HorizontalAlignment = -4152

        sheet.Cells(start_row + 7, 5).Value = 'подпись'
        sheet.Cells(start_row + 7, 5).HorizontalAlignment = -4108

        sheet.Cells(start_row + 7, 6).Value = 'ФИО'

        Selection = sheet.Range(f'F{start_row + 6}:G{start_row + 6}')
        Selection.HorizontalAlignment = -4108
        Selection.Merge()

        sheet.Cells(start_row + 7, 9).Value = 'дата'
        sheet.Cells(start_row + 7, 9).HorizontalAlignment = -4108

        Selection = sheet.Range(f'B{start_row+2}:L{start_row+9}')
        Selection.Interior.Color = 16777215
        Selection.Interior.TintAndShade = 0
        Selection.Interior.PatternTintAndShade = 0
        Selection.Font.Bold = True

        sheet.Cells(start_row + 2, 4).Font.Bold = True
        Selection = sheet.Range(f'E{start_row+2}:I{start_row+2}')
        xlEdgeBottom = 9
        xlContinuous = 1
        xlEdgeRight = 10
        xlEdgeLeft = 7
        xlEdgeTop = 8
        xlDash = -4115

        xlMedium = -4138
        xlDiagonalDown = 5
        xlDiagonalUp = 6
        xlInsideVertical = 11
        xlInsideHorizontal = 12

        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous

        Selection = sheet.Range(f'E{start_row + 6}:I{start_row + 6}')
        Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous

        Selection = sheet.Range(f'C{start_row + 2}:J{start_row + 8}')
        Selection.Borders(xlEdgeBottom).LineStyle = xlDash
        Selection.Borders(xlEdgeBottom).Weight = xlMedium

        Selection.Borders(xlEdgeRight).LineStyle = xlDash
        Selection.Borders(xlEdgeRight).Weight = xlMedium

        Selection.Borders(xlEdgeTop).LineStyle = xlDash
        Selection.Borders(xlEdgeTop).Weight = xlMedium

        Selection.Borders(xlEdgeLeft).LineStyle = xlDash
        Selection.Borders(xlEdgeLeft).Weight = xlMedium

        work_b1.Save()
        work_b1.Close()
        excel.Quit()
