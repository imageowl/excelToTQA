import sys

TQA_PATH = r'/Users/annafronhofer/PycharmProjects/pyTQA'
sys.path.insert(0, TQA_PATH)

import tqa

import json
import xlrd


def json_print(j):
    print(json.dumps(j, indent=4, sort_keys=True))


def load_json_file(config_file):
    with open(config_file) as file:
        config_dict = json.load(file)

    return config_dict


def load_excel_file(excel_file, sheet_dict, sched_idx):
    variable_list = []

    wb = xlrd.open_workbook(excel_file)

    for sheet in sheet_dict:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])
        for var in sheet['sheetVariables']:
            var_id = tqa.get_variable_id_from_string(var['name'], sched_idx)[0]

            column_int = abs(65-ord(var['valueCellColumn'].upper()))  # **Need to figure out how to find this with double letters
            val = excel_sheet.cell_value(var['valueCellRow']-1, column_int)

            variable_list.append({'id': var_id, 'value': val})
            print('var_id: ', var_id)

            if 'metaItems' in var:
                meta_items = []
                sched_vars = tqa.get_schedule_variables(sched_idx)
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



