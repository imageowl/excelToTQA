import datetime
from dateutil import parser
import sys
import os.path
import json
from string import ascii_uppercase
import xlrd

TQA_PATH = r'/Users/annafronhofer/PycharmProjects/pyTQA'
sys.path.insert(0, TQA_PATH)

import tqa


def upload_excel_file(excel_file, config_file):
    # get the data from the excel file and upload it to Smari

    config_dict = load_json_file(config_file)  # put the info from the config file into a dictionary
    config_sheets_dict = config_dict['sheets']  # dictionary object of the excel sheets
    wb = xlrd.open_workbook(excel_file)
    variable_list = []  # python list of variables to be used in tqa.upload_test_results

    sched_id = get_schedule_id(config_dict, wb)  # determine the id of the schedule
    if sched_id is None:
        error_msg = "The schedule name and machine name must be in the config file, or their locations in the excel " \
                    "file must be in the config file."
        raise ValueError("Error: The schedule id could not be found.", error_msg)

    for config_sheet in config_sheets_dict:
        excel_sheet = wb.sheet_by_name(config_sheet['sheetName'])

        for var in config_sheet['sheetVariables']:
            var_id = tqa.get_variable_id_from_string(var['name'], sched_id)[0]

            if "range" not in var:
                val = get_cell_value(var['valueCellRow'], var['valueCellColumn'], excel_sheet)[0]
            else:
                val = get_range_cell_values(var, excel_sheet)

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

                    if "range" not in item:
                        meta_val = get_cell_value(item['valueCellRow'], item['valueCellColumn'], excel_sheet)[0]
                    else:
                        meta_val = get_range_cell_values(item, excel_sheet)

                    meta_items.append({'id': meta_item_id, 'value': meta_val})

                variable_list[-1]['metaItems'] = meta_items

            if 'comment' in var:
                var_comment = get_cell_value(var['comment']['varCommentCellRow'],
                                             var['comment']['varCommentCellColumn'], excel_sheet)[0]
                variable_list[-1]['comment'] = var_comment

    report_date = get_report_date(config_dict, wb, excel_file)
    report_comment = get_report_comments(config_dict, wb)
    finalize = get_finalize_value(config_dict, wb)
    mode = get_mode(config_dict, wb)

    print("Schedule id: ", sched_id)
    print("Variables: ", variable_list)
    print("Report Date: ", report_date)
    print("Report Comment: ", report_comment)
    print("Finalize: ", finalize)
    print("Mode: ", mode)

    response = 0
    # response = tqa.upload_test_results(schedule_id=sched_id, variable_data=variable_list, comment=report_comment,
    #                                    finalize=finalize, mode=mode, date=report_date, date_format='%Y-%m-%dT%H:%M')
    return response


def load_json_file(config_file):
    with open(config_file) as file:
        config_dict = json.load(file)

    return config_dict


def get_cell_value(row_int, var_col, excel_sheet):
    # convert column from letter to integer and find the value in the cell
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
    return value, row_int, col_int


def get_range_cell_values(var, excel_sheet):
    vals = []
    first_val, first_row, first_col = get_cell_value(var["range"]["valueStartRow"],
                                                     var["range"]["valueStartColumn"], excel_sheet)
    last_val, last_row, last_col = get_cell_value(var["range"]["valueEndRow"],
                                                  var["range"]["valueEndColumn"], excel_sheet)
    for rowNum in range(first_row, last_row + 1):
        for colNum in range(first_col, last_col + 1):
            v = get_cell_value(rowNum, colNum + 1, excel_sheet)[0]
            vals.append(v)

    return vals


def get_schedule_id(config_dict, wb):
    # get the schedule id using the schedule name and machine id
    schedule_id = None

    for sheet in config_dict['sheets']:
        excel_sheet = wb.sheet_by_name(sheet['sheetName'])
        if 'machine' in sheet:  # machine name is in excel file
            machine = get_cell_value(sheet['machine']['machineCellRow'], sheet['machine']['machineCellColumn'],
                                     excel_sheet)[0]
        elif 'machineName' in config_dict:  # machine name is in config file
            machine = config_dict['machineName']
        machine_id = tqa.get_machine_id_from_str(machine)

        if 'schedule' in sheet:  # schedule name is in excel file
            schedule = get_cell_value(sheet['schedule']['scheduleCellRow'], sheet['schedule']['scheduleCellColumn'],
                                      excel_sheet)[0]
        elif 'scheduleName' in config_dict:  # schedule name is in config file
            schedule = config_dict['scheduleName']
        schedule_id = tqa.get_schedule_id_from_str(schedule, machine_id)

    return schedule_id


def get_report_date(config_dict, wb, excel_file):
    # to get the report date:
    #   use the date entered in the config file
    #   or use the date present in the excel file
    #   or if there is no date in the config or excel file, use the date the excel file was last modified

    report_date = None

    if "date" in config_dict:  # report date is entered in config file
        date = config_dict["date"]
        report_date = parser.parse(date)

    if report_date is None:
        for sheet in config_dict['sheets']:
            excel_sheet = wb.sheet_by_name(sheet['sheetName'])
            if 'date' in sheet:  # report date is in excel file
                date = get_cell_value(sheet['date']['dateCellRow'], sheet['date']['dateCellColumn'], excel_sheet)[0]
                report_date = xlrd.xldate_as_datetime(date, wb.datemode)

    if report_date is None:
        report_date = datetime.datetime.fromtimestamp(os.path.getmtime(excel_file))  # date last modified

    report_date = report_date.strftime('%Y-%m-%dT%H:%M')  # format date

    return report_date


def get_report_comments(config_dict, wb):
    # get the report comments from the excel file or config file, if there are any
    report_comment = None

    if "reportComment" in config_dict:  # report level comment is entered in config file
        report_comment = config_dict["reportComment"]

    if report_comment is None:
        for sheet in config_dict['sheets']:
            excel_sheet = wb.sheet_by_name(sheet['sheetName'])
            if 'reportComment' in sheet:  # report comment is in excel file
                report_comment = get_cell_value(sheet['reportComment']['reportCommentCellRow'],
                                                sheet['reportComment']['reportCommentCellColumn'], excel_sheet)[0]

    return report_comment


def get_finalize_value(config_dict, wb):
    # get the finalize value from the excel file or config file, if it is present
    finalize = None

    if "finalize" in config_dict:  # finalize value is entered in config file
        finalize = config_dict["finalize"]

    if finalize is None:
        for sheet in config_dict['sheets']:
            excel_sheet = wb.sheet_by_name(sheet['sheetName'])
            if 'finalize' in sheet:  # finalize value is in excel file
                finalize = int(get_cell_value(sheet['finalize']['finalizeCellRow'],
                                              sheet['finalize']['finalizeCellColumn'], excel_sheet)[0])

    return finalize


def get_mode(config_dict, wb):
    # get the mode from the excel file or config file, if it is present
    mode = None

    if "mode" in config_dict:  # mode is entered in config file
        mode = config_dict["mode"]

    if mode is None:
        for sheet in config_dict['sheets']:
            excel_sheet = wb.sheet_by_name(sheet['sheetName'])
            if 'mode' in sheet:  # mode is in excel file
                mode = get_cell_value(sheet['mode']['modeCellRow'], sheet['mode']['modeCellColumn'], excel_sheet)[0]

    return mode
