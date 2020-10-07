from pprint import pprint

from SSK.wialon_app import WialonManager


def handler_all(group_base):
    report_data = None
    for group, data in group_base.items():
        report_data = WialonManager().exec_report(data)
        break

    # pprint(report_data)


if __name__ == '__main__':
    groups = WialonManager().api_get_groups()
    handler_all(groups)
