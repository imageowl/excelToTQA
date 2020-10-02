import json
import xlrd


def excel_to_config_file(excel_file):
    # take the information in the excel file and convert it to the json config file of a specific format

    config_dict = {"data": [{}]}  # dictionary that will be used to create json config file
    excel_workbook = xlrd.open_workbook(excel_file)

    sheet = excel_workbook.sheet_by_name('Config')

    find_machine(sheet, config_dict)

    find_schedule(sheet, config_dict)

    find_finalize(sheet, config_dict)

    find_mode(sheet, config_dict)

    find_date(sheet, config_dict, excel_workbook)

    find_report_comment(sheet, config_dict)

    variables_table_cell = find_value_in_sheet(sheet, 'Variables Section')
    if variables_table_cell is not None:  # find table with variables in sheet
        config_dict["data"][0]["variables"] = []
        variables_table_row, variables_table_col = variables_table_cell
        cell_is_empty = False
        row_idx = variables_table_row+2
        while cell_is_empty is False:  # add all variables to config_dict
            variable_name = sheet.cell_value(row_idx, variables_table_col).strip()
            if len(variable_name) != 0:
                variable_row = int(sheet.cell_value(row_idx, variables_table_col + 1))
                variable_col = sheet.cell_value(row_idx, variables_table_col + 2)
                variable_sheet = sheet.cell_value(row_idx, variables_table_col + 3)
                config_dict["data"][0]["variables"].append({"name": variable_name.strip(), "valueCellRow": variable_row,
                                                            "valueCellColumn": variable_col, "sheetName": variable_sheet})
                if sheet.cell_value(row_idx, variables_table_col+4).lower() == "yes":  # variable has meta items
                    # add all meta items to config_dict
                    find_meta_item(config_dict, sheet, variable_name)
    #
    #             if sheet.cell_value(row_idx, variables_table_col+5).lower() == "yes":  # variable has a variable comment
    #                 find_variable_comment(config_dict, sheet, variable_name)
            else:
                cell_is_empty = True  # no more variables present in this sheet

            row_idx += 1  # check next row for variable

    json_print(config_dict)
    # write_to_json_file(config_dict)


def find_value_in_sheet(sheet, val):
    # look through entire sheet for specified value, return the indices of the first one found
    for col_num in range(sheet.ncols):
        for row_num in range(sheet.nrows):
            if sheet.cell_value(row_num, col_num) == val:
                return row_num, col_num


def find_sheet(sheet_name, config_dict):
    for index, sheet in enumerate(config_dict["data"]):
        if sheet["sheetName"] == sheet_name:
            return index

    config_dict["data"].append({"sheetName": sheet_name})
    return -1


def find_machine(sheet, config_dict):
    machine_name = ''

    machine_name_header = find_value_in_sheet(sheet, 'Machine Name')
    if machine_name_header is not None:  # find machine name in sheet
        row, col = machine_name_header
        machine_name = sheet.cell_value(row + 1, col)
        if machine_name != '':
            config_dict["machine"] = machine_name.strip()

    if machine_name == '':
        machine_cell = find_value_in_sheet(sheet, 'machine')
        if machine_cell is not None:  # find machine name row and column in sheet
            row, col = machine_cell
            machine_row = int(sheet.cell_value(row, col + 1))
            machine_col = sheet.cell_value(row, col + 2)
            machine_sheet = sheet.cell_value(row, col + 3)
            config_dict["data"][0]["machine"] = {"cellRow": machine_row, "cellColumn": machine_col,
                                                 "sheetName": machine_sheet}


def find_schedule(sheet, config_dict):
    schedule_name = ''

    schedule_name_header = find_value_in_sheet(sheet, 'Schedule Name')
    if schedule_name_header is not None:  # find schedule name in sheet
        row, col = schedule_name_header
        schedule_name = sheet.cell_value(row + 1, col)
        if schedule_name != '':
            config_dict["schedule"] = schedule_name.strip()

    if schedule_name == '':
        schedule_cell = find_value_in_sheet(sheet, 'schedule')
        if schedule_cell is not None:  # find schedule name row and column in sheet
            row, col = schedule_cell
            schedule_row = int(sheet.cell_value(row, col + 1))
            schedule_col = sheet.cell_value(row, col + 2)
            schedule_sheet = sheet.cell_value(row, col + 3)
            config_dict["data"][0]["schedule"] = {"cellRow": schedule_row, "cellColumn": schedule_col,
                                                  "sheetName": schedule_sheet}


def find_finalize(sheet, config_dict):
    finalize_val = ''

    finalize_header = find_value_in_sheet(sheet, 'Finalize Value')
    if finalize_header is not None:  # find finalize in sheet
        row, col = finalize_header
        finalize_val = sheet.cell_value(row + 1, col)
        if finalize_val != '':
            config_dict["finalize"] = int(finalize_val)

    if finalize_val == '':
        finalize_cell = find_value_in_sheet(sheet, 'finalize')
        if finalize_cell is not None:  # find finalize value row and column in sheet
            row, col = finalize_cell
            finalize_row = int(sheet.cell_value(row, col + 1))
            finalize_col = sheet.cell_value(row, col + 2)
            finalize_sheet = sheet.cell_value(row, col + 3)
            config_dict["data"][0]["finalize"] = {"cellRow": finalize_row, "cellColumn": finalize_col,
                                                  "sheetName": finalize_sheet}


def find_mode(sheet, config_dict):
    mode_val = ''

    mode_header = find_value_in_sheet(sheet, 'Save Mode')
    if mode_header is not None:  # find mode in sheet
        row, col = mode_header
        mode_val = sheet.cell_value(row + 1, col)
        if mode_val != '':
            config_dict["mode"] = mode_val.strip()

    if mode_val == '':
        mode_cell = find_value_in_sheet(sheet, 'mode')
        if mode_cell is not None:  # find mode value row and column in sheet
            row, col = mode_cell
            mode_row = int(sheet.cell_value(row, col + 1))
            mode_col = sheet.cell_value(row, col + 2)
            mode_sheet = sheet.cell_value(row, col + 3)
            config_dict["data"][0]["mode"] = {"cellRow": mode_row, "cellColumn": mode_col, "sheetName": mode_sheet}


def find_date(sheet, config_dict, excel_workbook):
    report_date = ''

    date_header = find_value_in_sheet(sheet, 'Report Date')
    if date_header is not None:  # find date in sheet
        row, col = date_header
        date_val = sheet.cell_value(row + 1, col)
        if date_val != '':
            report_date = xlrd.xldate_as_datetime(date_val, excel_workbook.datemode)
            config_dict["date"] = str(report_date)

    if report_date == '':
        date_cell = find_value_in_sheet(sheet, 'date')
        if date_cell is not None:  # find date row and column in sheet
            row, col = date_cell
            date_row = int(sheet.cell_value(row, col + 1))
            date_col = sheet.cell_value(row, col + 2)
            date_sheet = sheet.cell_value(row, col + 3)
            config_dict["data"][0]["date"] = {"cellRow": date_row, "cellColumn": date_col, "sheetName": date_sheet}


def find_report_comment(sheet, config_dict):
    report_comment_val = ''

    report_comment_header = find_value_in_sheet(sheet, 'Report Comment')
    if report_comment_header is not None:  # find report comment in sheet
        row, col = report_comment_header
        report_comment_val = sheet.cell_value(row + 1, col)
        if report_comment_val != '':
            config_dict["reportComment"] = report_comment_val.strip()

    if report_comment_val == '':
        report_comment_cell = find_value_in_sheet(sheet, 'comment')
        if report_comment_cell is not None:  # find report level comment row and column in sheet
            row, col = report_comment_cell
            report_comment_row = int(sheet.cell_value(row, col + 1))
            report_comment_col = sheet.cell_value(row, col + 2)
            report_comment_sheet = sheet.cell_value(row, col + 3)
            config_dict["data"][0]["reportComment"] = {"cellRow": report_comment_row, "cellColumn": report_comment_col,
                                                       "sheetName": report_comment_sheet}


def find_meta_item(config_dict, sheet, variable_name):
    # add all meta items to config_dict
    config_dict["data"][0]["variables"][-1]["metaItems"] = []
    meta_items_table_row, meta_items_table_col = find_value_in_sheet(sheet, 'Meta Items Section')
    cell_is_empty = False
    row_num = meta_items_table_row
    while cell_is_empty is False:  # only look in meta items table
        found_var = sheet.cell_value(row_num, meta_items_table_col).strip()
        if len(found_var) != 0:
            if found_var == variable_name:
                meta_item_var_name = sheet.cell_value(row_num, meta_items_table_col + 1).strip()
                meta_items_row = int(sheet.cell_value(row_num, meta_items_table_col + 2))
                meta_items_col = sheet.cell_value(row_num, meta_items_table_col + 3)
                meta_items_sheet = sheet.cell_value(row_num, meta_items_table_col + 4)
                config_dict["data"][0]["variables"][-1]["metaItems"].append({"name": meta_item_var_name,
                                                                             "valueCellRow": meta_items_row,
                                                                             "valueColumn": meta_items_col,
                                                                             "sheetName": meta_items_sheet})
        else:
            cell_is_empty = True

        row_num += 1


def find_variable_comment(config_dict, sheet, variable_name):
    # add variable comment to config_dict
    comments_table_row, comments_table_col = find_value_in_sheet(sheet, 'Comments Section')
    for row_num in range(comments_table_row, sheet.nrows):  # only look in comments table
        found_var = sheet.cell_value(row_num, comments_table_col).strip()
        if found_var == variable_name:
            comment_row = int(sheet.cell_value(row_num, comments_table_col + 1))
            comment_col = sheet.cell_value(row_num, comments_table_col + 2)
            comment_sheet = sheet.cell_value(row_num, comments_table_col + 3)
            sheet_index = find_sheet(comment_sheet, config_dict)
    config_dict["data"][sheet_index]["sheetVariables"][-1]["comment"] = {"varCommentCellRow": comment_row,
                                                                           "varCommentCellColumn": comment_col}


def write_to_json_file(config_dict):
    # takes the data in the config_dict and writes it to a local file in json format
    json_object = json.dumps(config_dict, indent=4)

    with open("config_file.json", "w") as outfile:
        outfile.write(json_object)


def json_print(j):
    print(json.dumps(j, indent=4))

