from __future__ import print_function

import os
import sys
from pprint import pprint
from wialon_app import WialonManager
from excel import ExcelManager
from time_set import time_conv


def handler_all(group_base, smena, from_time, to_time, f_date, t_date):
    path = os.path.join(os.getcwd(), 'Отчеты')
    if not os.path.exists(path):
        os.makedirs(path)
    os.chdir(path)
    for group, data in group_base.items():
        print(group)
        try:
            report_data = WialonManager().exec_report(data, smena[0], from_time, to_time)
            ExcelManager().handler_excel(group, report_data, smena[1], f_date, t_date, path)
            print('Обработан')
        except:
            print('Данных за период нет')


def handler_single(group, data, smena, from_time, to_time, f_date, t_date):
    path = os.path.join(os.getcwd(), 'Отчеты')
    if not os.path.exists(path):
        os.makedirs(path)
    os.chdir(path)
    report_data = WialonManager().exec_report(data, smena[0], from_time, to_time)
    ExcelManager().handler_excel(group, report_data, smena[1], f_date, t_date, path)


def handler_single_obj(obj_name, obj_data, smena, from_time, to_time, f_date, t_date):
    path = os.path.join(os.getcwd(), 'Отчеты')
    if not os.path.exists(path):
        os.makedirs(path)
    os.chdir(path)
    report_data = WialonManager().exec_report(obj_data, smena[0], from_time, to_time)
    ExcelManager().handler_excel(obj_name, report_data, smena[1], f_date, t_date, path)


if __name__ == '__main__':
    groups = WialonManager().api_get_groups()
    pprint(groups)
    group_dict = {}
    group_dict['0'] = 'all'
    while True:
        print('0. Все подрядчики')
        for num, group in enumerate(groups):
            group_dict[str(num+1)] = group
            sys.stdout.write(f'{num+1}. {group:30}     ')
            if num % 2 == 0:
                print('')
        choice_group = input('Введите номер подрядчика: ')
        if choice_group in group_dict:
            if choice_group != 0:
                print(groups[group_dict[choice_group]])
                objs = groups[group_dict[choice_group]][2]
                obj_dict = {}
                obj_dict['0'] = 'all'
                for num, obj in enumerate(objs):
                    obj_dict[num+1] = [WialonManager().api_get_obj(obj), obj]

                while True:
                    print('0. Все объекты')
                    for num, obj in obj_dict.items():
                        sys.stdout.write(f'{num}. {obj[0]:30}     ')
                        if num % 2 == 0:
                            print('')
                    choice_obj = input('Введите номер подрядчика: ')
                    if choice_obj in obj_dict:
                        break
                    else:
                        print('Неверный выбор ТС')
                        continue
            break
        else:
            print('Неверный выбор компании')
            continue

    smena_dict = \
        {
        '1': [4, 'Смена 1'],
        '2': [5, 'Смена 2'],
        '3': [1, 'Смена 1 и 2'],
        '4': [6, 'Смена 3'],
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

    elif choice_obj == '0':
        group = group_dict[str(choice_group)]
        handler_single(group, groups[group], smena, from_time, to_time, f_date, t_date)
    else:
        obj = obj_dict[str(choice_obj)]
        handler_single_obj(obj[0], obj[1], smena, from_time, to_time, f_date, t_date)
