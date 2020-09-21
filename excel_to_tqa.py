import sys

TQA_PATH = r'/Users/annafronhofer/PycharmProjects/pyTQA'
sys.path.insert(0, TQA_PATH)

import tqa

import json
import xlrd


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
            var_id = tqa.get_variable_id_from_string(var['name'], sched_idx)

            column_int = abs(65-ord(var['valueCellColumn'].upper()))  # **Need to figure out how to find this with double letters
            val = excel_sheet.cell_value(var['valueCellRow']-1, column_int)

            variable_list.append({'id': var_id[0], 'value': val})

            if 'metaItems' in var:
                meta_items = []
                for item in var['metaItems']:
                    meta_column_int = abs(65 - ord(item['valueCellColumn'].upper()))  # **Need to figure out how to find this with double letters
                    meta_val = excel_sheet.cell_value(item['valueCellRow'] - 1, meta_column_int)

                    meta_items.append({item['name']: meta_val})

                variable_list[-1]['metaItems'] = meta_items



    print(variable_list)



