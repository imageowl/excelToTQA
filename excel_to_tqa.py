import datetime
import sys
import os.path
import json
import xlrd

TQA_PATH = r'/Users/annafronhofer/PycharmProjects/pyTQA'
sys.path.insert(0, TQA_PATH)

import tqa


def json_print(j):
    print(json.dumps(j, indent=4, sort_keys=True))


def upload_excel_file(excel_file, config_file):
    config_dict = load_json_file(config_file)  # put the info from the config file into a dictionary
    sheets_dict = config_dict['sheets']  # create a dictionary object of the excel sheets
    wb = xlrd.open_workbook(excel_file)  # load the excel workbook
    variable_list = []  # python list of variables to be used in tqa.upload_test_results

    sched_id = get_schedule_id(config_dict, wb)  # determine the id of the schedule
    if sched_id is None:
        error_msg = "The schedule name and machine name must be in the config file, or their locations in the excel " \
                    "file must be in the config file."
        raise ValueError("Error: The schedule id could not be found.", error_msg)

    for sheet in sheets_dict:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])

        for var in sheet['sheetVariables']:
            var_id = tqa.get_variable_id_from_string(var['name'], sched_id)[0]

            val = get_cell_value(var['valueCellRow'], var['valueCellColumn'], excel_sheet)

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

                    meta_val = get_cell_value(item['valueCellRow'], item['valueCellColumn'], excel_sheet)

                    meta_items.append({'id': meta_item_id, 'value': meta_val})

                variable_list[-1]['metaItems'] = meta_items

            if 'comment' in var:
                var_comment = get_cell_value(var['comment']['varCommentCellRow'],
                                             var['comment']['varCommentCellColumn'], excel_sheet)
                variable_list[-1]['comment'] = var_comment

    report_date = get_report_date(config_dict, wb, excel_file)
    report_comment = get_report_comments(config_dict, wb)
    finalize = config_dict['finalize']
    mode = config_dict['mode']

    print("Schedule id: ", sched_id)
    print(variable_list)
    print("Report Date: ", report_date)
    print("Report Comment: ", report_comment)
    print("Finalize: ", finalize)
    print("Mode: ", mode)

    response = tqa.upload_test_results(schedule_id=sched_id, variable_data=variable_list, comment=report_comment,
                                       finalize=finalize, mode=mode, date=report_date, date_format='%Y-%m-%dT%H:%M')
    print("Response: ", response)


def load_json_file(config_file):
    with open(config_file) as file:
        config_dict = json.load(file)

    return config_dict


def get_cell_value(row_int, var_col, excel_sheet):
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

    value = excel_sheet.cell_value(row_int - 1, col_int)
    return value


def get_schedule_id(config_dict, wb):
    schedule_id = None

    for sheet in config_dict['sheets']:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])
        if 'machine' in sheet:  # machine name is in excel file
            machine = get_cell_value(sheet['machine']['machineCellRow'], sheet['machine']['machineCellColumn'],
                                     excel_sheet)
        elif 'machineName' in config_dict:  # machine name is in config file
            machine = config_dict['machineName']
        machine_id = tqa.get_machine_id_from_str(machine)

        if 'schedule' in sheet:  # schedule name is in excel file
            schedule = get_cell_value(sheet['schedule']['scheduleCellRow'], sheet['schedule']['scheduleCellColumn'],
                                      excel_sheet)
        elif 'scheduleName' in config_dict:  # schedule name is in config file
            schedule = config_dict['scheduleName']
        schedule_id = tqa.get_schedule_id_from_str(schedule, machine_id)

    return schedule_id


def get_report_date(config_dict, wb, excel_file):
    report_date = None

    for sheet in config_dict['sheets']:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])
        if 'date' in sheet:  # report date is in excel file
            date = get_cell_value(sheet['date']['dateCellRow'], sheet['date']['dateCellColumn'], excel_sheet)
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
            report_comment = get_cell_value(sheet['reportComment']['reportCommentCellRow'],
                                            sheet['reportComment']['reportCommentCellColumn'], excel_sheet)

    return report_comment
