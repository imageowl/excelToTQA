import datetime
from dateutil import parser
import sys
import os.path
import json
import xlrd

TQA_PATH = r'/Users/annafronhofer/PycharmProjects/pyTQA'
sys.path.insert(0, TQA_PATH)

import tqa


def upload_excel_file(excel_file, config_file):
    # get the data from the excel file and upload it to Smari

    config_dict = load_json_file(config_file)  # put the info from the config file into a dictionary
    config_data_dict = config_dict['data'][0]
    excel_workbook = xlrd.open_workbook(excel_file)
    variable_list = []  # python list of variables to be used in tqa.upload_test_results

    # get the schedule id using the machine id and schedule name
    machine_name = get_header_value(config_dict, excel_workbook, 'machine')
    machine_id = tqa.get_machine_id_from_str(machine_name)
    schedule_name = get_header_value(config_dict, excel_workbook, 'schedule')
    schedule_id = tqa.get_schedule_id_from_string(schedule_name, machine_id)

    if schedule_id is None:
        error_msg = "The schedule name and machine name must be in the config file, or their locations in the excel " \
                    "file must be in the config file."
        raise ValueError("Error: The schedule id could not be found.", error_msg)

    for variable in config_data_dict['variables']:  # get all the variables and their data
        variable_id = tqa.get_variable_id_from_string(variable['name'].strip(), schedule_id)
        if len(variable_id) > 0:
            variable_id = variable_id[0]
        else:
            raise KeyError("Error: No id was found for the specified variable name: " + str(variable["name"]))
        excel_sheet = excel_workbook.sheet_by_name(variable['sheetName'].strip())

        if 'range' not in variable:  # variable only has one value
            variable_value = get_cell_value(variable['valueCellRow'], variable['valueCellColumn'], excel_sheet)[0]
        else:  # variable has multiple values
            variable_value = get_range_cell_values(variable, excel_sheet)  # python list of all variable values

        variable_list.append({'id': variable_id, 'value': variable_value})

        if 'metaItems' in variable:
            # get all the variable meta items and their values
            meta_items = get_meta_item_values(schedule_id, variable_id, variable, excel_workbook)
            variable_list[-1]['metaItems'] = meta_items

        if 'comment' in variable:
            # get the variable comment
            excel_sheet = excel_workbook.sheet_by_name(variable['comment']['sheetName'].strip())
            variable_comment = get_cell_value(variable['comment']['varCommentCellRow'],
                                              variable['comment']['varCommentCellColumn'], excel_sheet)[0]
            variable_list[-1]['comment'] = variable_comment

    # look for duplicate variables in the variable_list and merge any found
    final_variable_list = check_for_variable_duplicates(variable_list)

    # get all the inputs needed for tqa.upload_test_results
    report_comment = get_header_value(config_dict, excel_workbook, 'reportComment')
    if report_comment is None:
        report_comment = ""  # default if there are no report comments

    finalize = get_header_value(config_dict, excel_workbook, 'finalize')
    if finalize is None:
        finalize = 0  # default if finalize is not specified
    else:
        finalize = int(finalize)

    mode = get_header_value(config_dict, excel_workbook, 'mode')
    if mode is None:
        mode = 'save_append'  # default if mode is not specified

    date = get_header_value(config_dict, excel_workbook, 'date')
    if date is None:
        report_date = datetime.datetime.fromtimestamp(os.path.getmtime(excel_file))  # default if date is not specified
    if isinstance(date, float):
        report_date = xlrd.xldate_as_datetime(date, excel_workbook.datemode)
    if isinstance(date, str):
        report_date = parser.parse(date)
    report_date = report_date.strftime('%Y-%m-%dT%H:%M')  # format date

    print("Schedule id: ", schedule_id)
    print("Report Comment: ", report_comment)
    print("Finalize: ", finalize)
    print("Mode: ", mode)
    print("Report Date: ", report_date, '\n')
    json_print(final_variable_list)

    # upload the data retrieved from the excel sheet
    response = tqa.upload_test_results(schedule_id=schedule_id, variable_data=final_variable_list,
                                       comment=report_comment, finalize=finalize, mode=mode, date=report_date,
                                       date_format='%Y-%m-%dT%H:%M')
    return response


def json_print(j):
    print(json.dumps(j, indent=4))


def load_json_file(config_file):
    # load the json from the config file into a dictionary
    with open(config_file) as file:
        config_dict = json.load(file)
    return config_dict


def get_cell_value(row_int, column, excel_sheet):
    # convert column from letter to integer and find the value in the cell
    if isinstance(column, str):  # column input as letter
        # convert letter to its ascii value, then to the column index in the excel sheet
        column = column.upper()
        if len(column) == 1:  # name of column is one letter
            col_int = abs(65 - ord(column))
        elif len(column) == 2:  # name of column is two letters
            first_letter = (abs(65 - ord(column[0])) + 1) * 26
            second_letter = abs(65 - ord(column[1]))
            col_int = first_letter + second_letter
    elif isinstance(column, int):  # column input as integer
        col_int = column - 1  # excel file starts at 1, xlrd indexing starts at 0

    value = excel_sheet.cell_value(row_int - 1, col_int)
    return value, row_int, col_int


def get_range_cell_values(variable, excel_sheet):
    # find all the values associated with the specified variable
    variable_values = []
    first_val, first_row, first_col = get_cell_value(variable['range']['valueStartRow'],
                                                     variable['range']['valueStartColumn'], excel_sheet)
    last_val, last_row, last_col = get_cell_value(variable['range']['valueEndRow'],
                                                  variable['range']['valueEndColumn'], excel_sheet)
    for row_num in range(first_row, last_row + 1):
        for col_num in range(first_col, last_col + 1):
            value = get_cell_value(row_num, col_num + 1, excel_sheet)[0]
            variable_values.append(value)

    return variable_values


def get_meta_item_values(sched_id, var_id, variable, excel_workbook):
    # determine what meta items for this variable are present in the config file, and their associated values

    all_var_meta_items = []  # every possible meta item that could be present for this variable
    var_meta_items = []  # the meta items present in the config file and their associated values

    # find meta item ids by first, getting all the possible meta items for this variable, and second, picking out the
    # meta items present in the config file
    sched_variables = tqa.get_schedule_variables(sched_id)
    for sched_var in sched_variables['json']['variables']:
        if sched_var['id'] == var_id:
            # all the possible meta items for this variable
            all_var_meta_items = sched_var['metaItems']
    for var_meta_item in variable['metaItems']:
        for meta_item in all_var_meta_items:
            if meta_item['name'] == var_meta_item['name'].strip():
                # meta item present in config file
                meta_item_id = meta_item['id']

        excel_sheet = excel_workbook.sheet_by_name(var_meta_item['sheetName'].strip())
        if 'range' not in var_meta_item:  # meta item has only one value
            meta_item_value = get_cell_value(var_meta_item['valueCellRow'], var_meta_item['valueCellColumn'],
                                             excel_sheet)[0]
        else:  # meta item has multiple values
            meta_item_value = get_range_cell_values(var_meta_item, excel_sheet)

        var_meta_items.append({'id': meta_item_id, 'value': meta_item_value})

    return var_meta_items


def check_for_variable_duplicates(variables_list):
    # look for duplicate variables in variable list and combine any found

    checked_variables_list = []  # final variable list with combined duplicate variables
    temp_dict = {}  # used to create new dictionary without any duplicate variables (will convert to list)
    for var_dict in variables_list:
        if var_dict['id'] not in temp_dict:  # this variable is the first present in the list with specified id
            temp_dict[var_dict['id']] = var_dict
        else:  # this variable has already been present in the list, add it to the existing variable dictionary
            for key in var_dict:
                if key == 'value':  # add the value to the value list for specified variable
                    if not isinstance(temp_dict[var_dict['id']]['value'], list):
                        temp_dict[var_dict['id']]['value'] = [temp_dict[var_dict['id']]['value'], var_dict['value']]
                    else:
                        temp_dict[var_dict['id']]['value'].append(var_dict['value'])
                elif key == 'comment':  # append any comments to the existing variable comments
                    if 'comment' in temp_dict[var_dict['id']]:
                        temp_dict[var_dict['id']]['comment'] += ("; " + var_dict['comment'])
                    else:
                        temp_dict[var_dict['id']]['comment'] = var_dict['comment']
                elif key == 'metaItems':  # check to see if meta item already exists
                    for var_item in var_dict['metaItems']:
                        new_meta_item = True
                        if 'metaItems' not in temp_dict[var_dict['id']]:
                            temp_dict[var_dict['id']]['metaItems'] = []
                        for item in temp_dict[var_dict['id']]['metaItems']:
                            if var_item['id'] == item['id']:  # meta item already in metaItems
                                if not isinstance(item['value'], list):  # create list with all meta item values
                                    item['value'] = [item['value'], var_item['value']]
                                else:
                                    item['value'].append(var_item['value'])  # add meta item value to the value list
                                new_meta_item = False
                                break
                        if new_meta_item:  # add new meta item to meta items list
                            temp_dict[var_dict['id']]['metaItems'].append(var_item)

    for value in temp_dict.values():  # convert variables dictionary to variables list format
        checked_variables_list.append(value)

    return checked_variables_list


def get_header_value(config_dict, excel_workbook, header_name):
    value = None

    if header_name in config_dict:  # header value is entered in config file
        value = config_dict[header_name]
    elif header_name in config_dict['data'][0]:  # header value is entered in excel file
        excel_sheet = excel_workbook.sheet_by_name(config_dict['data'][0][header_name]['sheetName'].strip())
        value = get_cell_value(config_dict['data'][0][header_name]['cellRow'],
                               config_dict['data'][0][header_name]['cellColumn'], excel_sheet)[0]

    return value

