from __future__ import print_function

import datetime
import os
import sys
from pprint import pprint
from wialon_app import WialonManager
from excel import ExcelManager
from time_set import time_conv
import time


def handler_single(group, data, smena, from_time, to_time, f_date, t_date, f_dt, t_dt, path, count_smena, company=None):
    print('Формируем отчет...')
    try:
        report_data = WialonManager().exec_report(data, smena[0], from_time, to_time)
        print('Выгружаем данные в excel...')
        ExcelManager().handler_excel(group, report_data, smena[1], f_date, t_date, f_dt, t_dt, path, count_smena, company=company)
        print('Обработан')
        print('\n')
    except:
        print('Данных за период нет')
        print('\n')
        time.sleep(1)


if __name__ == '__main__':
    path = os.path.join(os.getcwd(), 'Отчеты')
    today = str(datetime.datetime.today())
    if not os.path.exists(path):
        os.makedirs(path)
    os.chdir(path)

    while True:
        groups = WialonManager().api_get_groups()
        group_dict = {}
        count_smena = 1
        group_dict['0'] = 'all'
        while True:
            for num, group in enumerate(groups):
                group_dict[str(num+1)] = group
                sys.stdout.write(f'{num+1}. {group:30}     ')
                if num % 2 == 0:
                    print('')
            print('0. Все подрядчики')

            choice_group = input('\nВведите номер подрядчика: ')
            if choice_group in group_dict:
                if choice_group != 0:
                    # print(groups[group_dict[choice_group]])
                    objs = groups[group_dict[choice_group]][2]
                    obj_dict = {}
                    print('Загружаем машины...')
                    for num, obj in enumerate(objs):
                        obj_dict[num+1] = [WialonManager().api_get_obj(obj), obj]
                    while True:
                        for num, obj in obj_dict.items():
                            sys.stdout.write(f'{num}. {obj[0]:30}     ')
                            if int(num) % 2 == 0:
                                print('')
                        print('0. Все объекты')
                        choice_obj = int(input('\nВведите номер объекта: '))
                        if choice_obj in obj_dict or choice_obj == 0:
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
            '5': [12, 'Суточная смена'],
        }
        while True:
            choice = input('Выберете смену:\n'
                           '1. Смена 1 (8:00 - 19:59)\n'
                           '2. Смена 2 (20:00 - 07:59)\n'
                           '3. Смена 1 и Смена 2\n'
                           '4. Смена 3 (8:00 - 01:00)\n'
                           '5. Суточная смена\n'
                           )

            if choice in smena_dict:
                if choice == '3':
                    count_smena = 2
                break
            else:
                print('Некорректный выбор, попробуйте еще раз')

        while True:
            period = input('Введите период в формате "ДД.ММ.ГГГГ-ДД.ММ.ГГГГ" (пример: 01.01.2020-31.01.2020)\n')
            # try:
            try:
                from_time, to_time, f_date, t_date, f_dt, t_dt = time_conv(period)
                break
            except:
                continue

        group = group_dict[str(choice_group)]
        smena = smena_dict[choice]

        if choice_group == '0':
            for group, data in groups.items():
                print(group)
                handler_single(group, data, smena, from_time, to_time, f_date, t_date, f_dt, t_dt, path, count_smena)
        elif choice_obj == 0:
            handler_single(group, groups[group], smena, from_time, to_time, f_date, t_date, f_dt, t_dt, path, count_smena)
        else:
            obj = obj_dict[choice_obj]
            handler_single(obj[0], [obj[1], ], smena, from_time, to_time, f_date, t_date, f_dt, t_dt, path, count_smena, company=group)
