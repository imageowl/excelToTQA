import sys
import json
import xlrd

TQA_PATH = r'/Users/annafronhofer/PycharmProjects/pyTQA'
sys.path.insert(0, TQA_PATH)

import tqa


def json_print(j):
    print(json.dumps(j, indent=4, sort_keys=True))


def load_json_file(config_file):
    with open(config_file) as file:
        config_dict = json.load(file)

    return config_dict


def load_excel_file(excel_file, config_file):
    config_dict = load_json_file(config_file)
    wb = xlrd.open_workbook(excel_file)
    sheets_dict = config_dict['sheets']
    variable_list = []

    sched_id = get_schedule_id(config_dict, wb)

    for sheet in sheets_dict:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])

        for var in sheet['sheetVariables']:
            var_id = tqa.get_variable_id_from_string(var['name'], sched_id)[0]

            column_int = abs(65-ord(var['valueCellColumn'].upper()))  # **Need to figure out how to find this with double letters
            val = excel_sheet.cell_value(var['valueCellRow']-1, column_int)

            variable_list.append({'id': var_id, 'value': val})

            if 'metaItems' in var:
                meta_items = []
                sched_vars = tqa.get_schedule_variables(sched_id)
                for idx, s in enumerate(sched_vars['json']['variables']):
                    if s['id'] == var_id:
                        var_meta_items = s['metaItems']
                for item in var['metaItems']:
                    for i in var_meta_items:
                        if i['name'] == item['name']:
                            meta_item_id = i['id']

                    meta_column_int = abs(65 - ord(item['valueCellColumn'].upper()))  # **Need to figure out how to find this with double letters
                    meta_val = excel_sheet.cell_value(item['valueCellRow'] - 1, meta_column_int)

                    meta_items.append({'id': meta_item_id, 'value': meta_val})

                variable_list[-1]['metaItems'] = meta_items

    json_print(variable_list)


def get_schedule_id(config_dict, wb):
    machine_id = None
    schedule_id = None
    sheets_dict = config_dict['sheets']

    for sheet in sheets_dict:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])
        if 'machineCellRow' in sheet:
            machine_column_int = abs(65 - ord(sheet['machineCellColumn'].upper()))  # **Need to figure out how to find this with double letters
            machine = excel_sheet.cell_value(sheet['machineCellRow'] - 1, machine_column_int)
            machine_id = tqa.get_machine_id_from_str(machine)
        elif 'machineName' in config_dict:
            machine_id = tqa.get_machine_id_from_str(config_dict['machineName'])

        if 'scheduleCellRow' in sheet:
            sched_column_int = abs(65 - ord(sheet['scheduleCellColumn'].upper()))  # **Need to figure out how to find this with double letters
            schedule = excel_sheet.cell_value(sheet['scheduleCellRow'] - 1, sched_column_int)
            schedule_id = tqa.get_schedule_id_from_str(schedule, machine_id)
        elif 'scheduleName' in config_dict:
            schedule_id = tqa.get_schedule_id_from_str(config_dict['scheduleName'], machine_id)

    return schedule_id
