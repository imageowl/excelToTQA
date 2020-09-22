import datetime
import sys
import json
import xlrd
import os.path, time

TQA_PATH = r'/Users/annafronhofer/PycharmProjects/pyTQA'
sys.path.insert(0, TQA_PATH)

import tqa


def json_print(j):
    print(json.dumps(j, indent=4, sort_keys=True))


def load_excel_file(excel_file, config_file):
    config_dict = load_json_file(config_file)  # put the info from the config file into a dictionary
    sheets_dict = config_dict['sheets']  # create a dictionary object of the excel sheets
    wb = xlrd.open_workbook(excel_file)  # load the excel workbook
    variable_list = []  # python list of variables to be used in tqa.upload_test_results

    sched_id = get_schedule_id(config_dict, wb)  # determine the id of the schedule

    for sheet in sheets_dict:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])

        for var in sheet['sheetVariables']:
            var_id = tqa.get_variable_id_from_string(var['name'], sched_id)[0]

            var_column_int = get_var_column_int(var['valueCellColumn'])
            val = excel_sheet.cell_value(var['valueCellRow']-1, var_column_int)

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

                    meta_column_int = get_var_column_int(item['valueCellColumn'])
                    meta_val = excel_sheet.cell_value(item['valueCellRow'] - 1, meta_column_int)

                    meta_items.append({'id': meta_item_id, 'value': meta_val})

                variable_list[-1]['metaItems'] = meta_items

            if 'comment' in var:
                var_comment_column_int = get_var_column_int(var['comment']['varCommentCellColumn'])
                var_comment = excel_sheet.cell_value(var['comment']['varCommentCellRow'] - 1, var_comment_column_int)
                variable_list[-1]['comment'] = var_comment

    report_date = get_report_date(config_dict, wb, excel_file)
    report_comment = get_report_comments(config_dict, wb)


    print("Schedule id: ", sched_id)
    print(variable_list)
    print("Report Date: ", report_date)
    print("Report Comment: ", report_comment)



def load_json_file(config_file):
    with open(config_file) as file:
        config_dict = json.load(file)

    return config_dict


def get_schedule_id(config_dict, wb):
    schedule_id = None

    for sheet in config_dict['sheets']:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])
        if 'machine' in sheet:  # machine name is in excel file
            machine_column_int = get_var_column_int(sheet['machine']['machineCellColumn'])
            machine = excel_sheet.cell_value(sheet['machine']['machineCellRow'] - 1, machine_column_int)
        elif 'machineName' in config_dict:  # machine name is in config file
            machine = config_dict['machineName']
        machine_id = tqa.get_machine_id_from_str(machine)

        if 'schedule' in sheet:  # schedule name is in excel file
            sched_column_int = get_var_column_int(sheet['schedule']['scheduleCellColumn'])
            schedule = excel_sheet.cell_value(sheet['schedule']['scheduleCellRow'] - 1, sched_column_int)
        elif 'scheduleName' in config_dict:  # schedule name is in config file
            schedule = config_dict['scheduleName']
        schedule_id = tqa.get_schedule_id_from_str(schedule, machine_id)

    return schedule_id


def get_var_column_int(var_col):
    col_int = None
    if isinstance(var_col, str):  # column input as letter
        var_col = var_col.upper()
        if len(var_col) == 1:  # name of column is one letter
            col_int = abs(65-ord(var_col))
        elif len(var_col) == 2:  # name of column is two letters
            first_letter = (abs(65 - ord(var_col[0])) + 1) * 26
            second_letter = abs(65 - ord(var_col[1]))
            col_int = first_letter + second_letter
    elif isinstance(var_col, int):  # column input as integer
        col_int = var_col-1  # excel file starts at 1, xlrd indexing starts at 0

    return col_int


def get_report_date(config_dict, wb, excel_file):
    report_date = None

    for sheet in config_dict['sheets']:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])
        if 'date' in sheet:  # report date is in excel file
            date_column_int = get_var_column_int(sheet['date']['dateCellColumn'])
            date = excel_sheet.cell_value(sheet['date']['dateCellRow'] - 1, date_column_int)
            report_date = xlrd.xldate_as_datetime(date, wb.datemode)

    if report_date is None:
        report_date = datetime.datetime.fromtimestamp(os.path.getmtime(excel_file))  # date last modified

    report_date = report_date.strftime('%Y-%m-%dT%H:%M')  # format date

    return report_date


def get_report_comments(config_dict, wb):
    report_comment = None

    for sheet in config_dict['sheets']:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])
        if 'reportComment' in sheet:  # report comment is in excel file
            report_comment_column_int = get_var_column_int(sheet['reportComment']['reportCommentCellColumn'])
            report_comment = excel_sheet.cell_value(sheet['reportComment']['reportCommentCellRow'] - 1,
                                                    report_comment_column_int)

    return report_comment
