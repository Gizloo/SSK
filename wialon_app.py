import time
from collections import defaultdict
from pprint import pprint

from wialon import flags, Wialon, WialonError


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

    def exec_report(self, group):
        result_rep = defaultdict(str)
        report = self.wialon.report_exec_report({
            'reportResourceId': self.res_id,
            'reportTemplateId': 1,
            'reportObjectId': group[0],
            'reportObjectSecId': 0,
            'interval': {'from': 1601053140, 'to': 1601485140, 'flags': 0}})

        rows_obj = report['reportResult']['tables'][0]['rows']

        rep_row = self.wialon.report_get_result_rows({
            "tableIndex": 0,
            "indexFrom": 0,
            "indexTo": rows_obj
        })

        for n in range(0, rows_obj):
            print(n)
            rep_sub_row = self.wialon.report_get_result_subrows({
                                          "tableIndex": 0,
                                          "rowIndex": n
                                  })
            for row in rep_sub_row:
                pprint(row)
                print(rep_row[n]['c'][1])
                # result_rep[rep_row[n]['c'][1]] += [row['c'][1], ]
            if n > 2:
                return rep_row
        return rep_row
