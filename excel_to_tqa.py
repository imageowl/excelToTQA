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
        print('')
        for vars in sheet['sheetVariables']:
            var_id = tqa.get_variable_id_from_string(vars['name'], sched_idx)
            variable_list.append({'id': var_id[0]})

    print(variable_list)



