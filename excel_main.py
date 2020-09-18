import sys

TQA_PATH = r'/Users/annafronhofer/PycharmProjects/pyTQA'
sys.path.insert(0, TQA_PATH)

import tqa
import excel_to_tqa


tqa.load_json_credentials('SmariCredentials.json')
print('Access Token: ', tqa.access_token)

config_dict = excel_to_tqa.load_json_file("configTest.json")

machine_idx = tqa.get_machine_id_from_str(config_dict['machineName'])
schedule_idx = tqa.get_schedule_id_from_str(config_dict['scheduleName'], machine_idx)

sheet_dict = config_dict['sheets']

excel_file_path = "/Users/annafronhofer/Desktop/testFiles/LinacCTP504Copy.xlsx"
excel_to_tqa.load_excel_file(excel_file_path, sheet_dict, schedule_idx)
