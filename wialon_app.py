import datetime
import time
from collections import defaultdict
from pprint import pprint
from wialon import flags, Wialon, WialonError
import time


class WialonManager:

    def __init__(self):
        self.token = '1e3f50514d35becfaf1b9ec8ff42f80014125DD4DBEADF212ED7DC3ED42D71466C71DF06'  # Основной
        # self.token = '290a6913b07b4afce549894ab74c1d87913BAACC55F364819505234D43FAB9FA89FC72A6'  #ССК(п) Подрядчики_api
        # self.token = 'eac53c387a819eb667e4e3fa967276ed55DE297F7272B45494FAE6F694E20D4C312F9907' #ССК(РС) Подрядчики_api
        # self.token = '526cdec32ecf25b664182796a38c3c665D97F79D99283A9CA9F4B7A1AA40E8D449D356FE' #ССК(Т) Подрядчики_api
        self.wialon = Wialon()

        try:
            login = self.wialon.token_login(token=self.token)
        except WialonError as e:
            print("Error while login")
            time.sleep(5)
            return
        self.wialon.sid = login['eid']
        self.wialon.render_set_locale({"tzOffset": 18000, "language": 'ru', "formatDate": "%Y-%m-%E %H:%M:%S"})
        self.res_id = 21922430
        self.base_group = {}


    def api_get_groups(self):
        spec = {
            'itemsType': 'avl_unit_group',
            'propName': 'sys_name',
            'propValueMask': '*',
            'sortType': 'sys_name'
        }
        interval = {"from": 0, "to": 0}
        custom_flag = flags.ITEM_RESOURCE_DATAFLAG_DRIVERS + flags.ITEM_DATAFLAG_BASE + \
                      flags.ITEM_RESOURCE_DATAFLAG_NOTIFICATIONS + 0x00001000
        units = self.wialon.core_search_items(spec=spec, force=1, flags=custom_flag, **interval)
        groups = units['items']
        for group in groups:
            if group['nm'] != 'ССК Подрядчики':
                self.base_group[group['nm']] = [group['id'], len(group['u']), group['u']]
        return self.base_group

    def api_get_obj(self, id):
        obj = self.wialon.core_search_item({'id': int(id), 'flags': 1})
        return obj['item']['nm']

    def exec_report(self, group, smena, from_time, to_time):
        # tz = 0
        tz = 7200
        result_rep = {}
        report = self.wialon.report_exec_report({
            'reportResourceId': self.res_id,
            'reportTemplateId': smena,
            'reportObjectId': group[0],
            'reportObjectSecId': 0,
            'interval': {'from': from_time, 'to': to_time, 'flags': 0}})

        rows_obj = report['reportResult']['tables'][0]['rows']

        rep_row = self.wialon.report_get_result_rows({
            "tableIndex": 0,
            "indexFrom": 0,
            "indexTo": rows_obj
        })
        # pprint(rep_row)

        for n in range(0, rows_obj):
            rep_sub_row = self.wialon.report_get_result_subrows({
                "tableIndex": 0,
                "rowIndex": n
            })
            obj_name = rep_row[n]['c'][1]
            result_rep[obj_name] = defaultdict(list)
            for row1 in rep_sub_row:
                # pprint(row1)
                if 'Outside shifts' not in row1['c']:
                    unix_key = int(row1['c'][3][:-3]) + tz
                    if smena == 12:
                        work_h = round(float(row1['c'][9]), 2)
                        time_start = int(datetime.datetime.fromtimestamp(int(row1['c'][3][:-3]) + tz).hour)
                        time_end = int(datetime.datetime.fromtimestamp(int(row1['c'][5][:-3]) + tz).hour)
                        duty_h = time_end - time_start
                    else:
                        work_h = round(float(row1['c'][9]), 2)
                        duty_h = round(float(row1['c'][10]), 2)
                    result_rep[obj_name][unix_key] = [

                        row1['c'][0],  # номер строки
                        row1['c'][1],  # имя

                        int(row1['c'][3][:-3]) + tz, #  начало
                        int(row1['c'][5][:-3]) + tz,  #  конец

                        row1['c'][6],  # часы в работе
                        row1['c'][7],  # часы в дежурстве
                        row1['c'][8],  # пробег

                        work_h,  # часы в работе (коррк)
                        duty_h,  # часы в дежурстве (коррк)
                        round(float(row1['c'][8]), 2),  # пробег (коррк)

                        row1['c'][11]['t'].replace('Road', 'Трасса').replace('km', 'км').replace('from', 'от'),
                        # нач. положение
                        row1['c'][12]['t'].replace('Road', 'Трасса').replace('km', 'км').replace('from', 'от'),
                        # кон. положение
                    ]
        return result_rep
