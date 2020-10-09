from __future__ import print_function

import sys
from pprint import pprint
from wialon_app import WialonManager
from excel import ExcelManager
from time_set import time_conv


def handler_all(group_base, smena, from_time, to_time, f_date, t_date):
    report_data = None
    for group, data in group_base.items():
        print(group)
        report_data = WialonManager().exec_report(data, smena[0], from_time, to_time)
        ExcelManager().handler_excel(group, report_data, smena[1], f_date, t_date)
        print('Обработан')


def handler_single(group, data, smena, from_time, to_time, f_date, t_date):
    report_data = WialonManager().exec_report(data, smena[0], from_time, to_time)
    ExcelManager().handler_excel(group, report_data, smena[1], f_date, t_date)


if __name__ == '__main__':
    groups = WialonManager().api_get_groups()
    group_dict = {}
    while True:
        print('0. Все подрядчики')
        for num, group in enumerate(groups):
            group_dict[str(num+1)] = group
            sys.stdout.write(f'{num+1}. {group:30}     ')
            if num % 2 == 0:
                print('')
        choice_group = input('Введите номер подрядчика: ')
        if choice_group in group_dict:
            break
        else:
            print(choice_group)
            print(len(groups))
            print('Неверный выбор')
            continue

    smena_dict = \
        {
        '1': [4, 'Смена 1'],  # 1 cмена
        '2': [5, 'Смена 2'],  # 2 cмена
        '3': [1, 'Смена 1 и 2'],  # 1 и 2 смена
        '4': [6, 'Смена 3'],  # 3 cмена
    }
    while True:
        choice = input('Выберете смену:\n'
                       '1. Смена 1 (8:00 - 19:59)\n'
                       '2. Смена 2 (20:00 - 07:59)\n'
                       '3. Смена 1 и Смена 2\n'
                       '4. Смена 3 (8:00 - 01:00)\n')

        if choice in smena_dict:
            break
        else:
            print('Некорректный выбор, попробуйте еще раз')

    while True:
        period = input('Введите период в формате "ДД.ММ.ГГГГ-ДД.ММ.ГГГГ" (пример: 01.01.2020-31.01.2020)\n')
        # try:
        try:
            from_time, to_time, f_date, t_date = time_conv(period)
            break
        except:
            continue

    smena = smena_dict[choice]
    if choice_group == '0':
        handler_all(groups, smena, from_time, to_time, f_date, t_date)
    else:
        group = group_dict[str(choice_group)]
        handler_single(group, groups[group], smena, from_time, to_time, f_date, t_date)
