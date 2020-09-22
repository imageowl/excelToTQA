import sys

TQA_PATH = r'/Users/annafronhofer/PycharmProjects/pyTQA'
sys.path.insert(0, TQA_PATH)

import tqa
import excel_to_tqa


tqa.load_json_credentials('SmariCredentials.json')
print('Access Token: ', tqa.access_token)

excel_file_path = "/Users/annafronhofer/Desktop/testFiles/LinacCTP504Copy.xlsx"
config_file_path = "/Users/annafronhofer/PycharmProjects/excel_to_TQA/configTest.json"
excel_to_tqa.load_excel_file(excel_file_path, config_file_path)
