import sys

TQA_PATH = r'/Users/annafronhofer/PycharmProjects/pyTQA'
sys.path.insert(0, TQA_PATH)

import tqa
import excel_to_config

import excel_to_tqa


# excel_ex_file = "/Users/annafronhofer/Desktop/testFiles/Config Ex1.xlsx"
# excel_to_config.excel_to_config_file(excel_ex_file)


tqa.load_json_credentials('SmariCredentials.json')
print('Access Token: ', tqa.access_token)

excel_file_path = "/Users/annafronhofer/Desktop/testFiles/LinacCTP504Copy.xlsx"
config_file_path = "newConfig.json"
response = excel_to_tqa.upload_excel_file(excel_file_path, config_file_path)
print(response)
