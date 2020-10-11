import time
from collections import defaultdict
from pprint import pprint

from wialon import flags, Wialon, WialonError
import time


class WialonManager:
    def __init__(self):
        self.token = '1e3f50514d35becfaf1b9ec8ff42f80014125DD4DBEADF212ED7DC3ED42D71466C71DF06'
        self.wialon = Wialon()

        try:
            login = self.wialon.token_login(token=self.token)
        except WialonError as e:
            print("Error while login")
            time.sleep(5)
            return
        self.wialon.sid = login['eid']
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
                self.base_group[group['nm']] = [group['id'], len(group['u'])]
        return self.base_group

    def exec_report(self, group, smena, from_time, to_time):
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
                unix_key = int(row1['c'][3][:-3]) - 7200
                result_rep[obj_name][unix_key] = [
                    row1['c'][0],  # номер строки
                    row1['c'][1],  # имя
                    int(row1['c'][3][:-3]) - 7200,  # начало
                    int(row1['c'][5][:-3]) - 7200,  # конец
                    row1['c'][6],  # часы в работе
                    row1['c'][7],  # часы в дежурстве
                    row1['c'][8],  # пробег
                    round(float(row1['c'][9]), 2),  # часы в работе (коррк)
                    round(float(row1['c'][10]), 2),  # часы в дежурстве (коррк)
                    round(float(row1['c'][8]), 2),  # пробег (коррк)

                    row1['c'][11]['t'].replace('Road', 'Трасса').replace('km', 'км').replace('from', 'от'),
                    # нач. положение
                    row1['c'][12]['t'].replace('Road', 'Трасса').replace('km', 'км').replace('from', 'от'),
                    # кон. положение
                ]

        return result_rep
