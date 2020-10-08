from pprint import pprint
from wialon_app import WialonManager
from excel import ExcelManager


def handler_all(group_base):
    report_data = None
    for group, data in group_base.items():
        report_data = WialonManager().exec_report(data)
        ExcelManager().handler_excel(group, report_data)
        break


if __name__ == '__main__':
    groups = WialonManager().api_get_groups()
    handler_all(groups)
