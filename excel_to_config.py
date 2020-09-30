import json
import xlrd


def excel_to_config_file(excel_file):
    # take the information in the excel file and convert it to the json config file of a specific format

    config_dict = {"sheets": []}  # dictionary that will be used to create json config file
    excel_workbook = xlrd.open_workbook(excel_file)

    for sheet in excel_workbook.sheets():
        sheet_name = sheet.name
        config_dict["sheets"].append({"sheetName": sheet_name})

        find_finalize(sheet, config_dict)

        find_mode(sheet, config_dict)

        find_machine(sheet, config_dict)

        find_schedule(sheet, config_dict)

        find_date(sheet, config_dict, excel_workbook)

        find_report_comment(sheet, config_dict)

        variables_table_cell = find_value_in_sheet(sheet, 'Variables Section')
        if variables_table_cell is not None:  # find table with variables in sheet
            config_dict["sheets"][-1]["sheetVariables"] = []
            variables_table_row, variables_table_col = variables_table_cell
            cell_is_empty = False
            row_idx = variables_table_row+2
            while cell_is_empty is False:  # add all variables to config_dict
                variable_name = sheet.cell_value(row_idx, variables_table_col).strip()
                if len(variable_name) != 0:
                    variable_row = int(sheet.cell_value(row_idx, variables_table_col+1))
                    variable_col = sheet.cell_value(row_idx, variables_table_col+2)
                    config_dict["sheets"][-1]["sheetVariables"].append({"name": variable_name.strip(),
                                                                        "valueCellRow": variable_row,
                                                                        "valueCellColumn": variable_col})
                    if sheet.cell_value(row_idx, variables_table_col+3).lower() == "yes":  # variable has meta items
                        # add all meta items to config_dict
                        find_meta_item(config_dict, sheet, variable_name)

                    if sheet.cell_value(row_idx, variables_table_col+4).lower() == "yes":  # variable has a variable comment
                        find_variable_comment(config_dict, sheet, variable_name)
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
            config_dict["sheets"][-1]["finalize"] = {"finalizeCellRow": finalize_row,
                                                     "finalizeCellColumn": finalize_col}


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
            config_dict["sheets"][-1]["mode"] = {"modeCellRow": mode_row, "modeCellColumn": mode_col}


def find_machine(sheet, config_dict):
    machine_name = ''

    machine_name_header = find_value_in_sheet(sheet, 'Machine Name')
    if machine_name_header is not None:  # find machine name in sheet
        row, col = machine_name_header
        machine_name = sheet.cell_value(row + 1, col)
        if machine_name != '':
            config_dict["machineName"] = machine_name.strip()

    if machine_name == '':
        machine_cell = find_value_in_sheet(sheet, 'machine')
        if machine_cell is not None:  # find machine name row and column in sheet
            row, col = machine_cell
            machine_row = int(sheet.cell_value(row, col + 1))
            machine_col = sheet.cell_value(row, col + 2)
            config_dict["sheets"][-1]["machine"] = {"machineCellRow": machine_row, "machineCellColumn": machine_col}


def find_schedule(sheet, config_dict):
    schedule_name = ''

    schedule_name_header = find_value_in_sheet(sheet, 'Schedule Name')
    if schedule_name_header is not None:  # find schedule name in sheet
        row, col = schedule_name_header
        schedule_name = sheet.cell_value(row + 1, col)
        if schedule_name != '':
            config_dict["scheduleName"] = schedule_name.strip()

    if schedule_name == '':
        schedule_cell = find_value_in_sheet(sheet, 'schedule')
        if schedule_cell is not None:  # find schedule name row and column in sheet
            row, col = schedule_cell
            schedule_row = int(sheet.cell_value(row, col + 1))
            schedule_col = sheet.cell_value(row, col + 2)
            config_dict["sheets"][-1]["schedule"] = {"scheduleCellRow": schedule_row,
                                                     "scheduleCellColumn": schedule_col}


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
            config_dict["sheets"][-1]["date"] = {"dateCellRow": date_row, "dateCellColumn": date_col}


def find_report_comment(sheet, config_dict):
    report_comment_cell = find_value_in_sheet(sheet, 'Report Comment')
    if report_comment_cell is not None:  # find report comment in sheet
        row, col = report_comment_cell
        report_comment_val = sheet.cell_value(row + 1, col)
        config_dict["reportComment"] = report_comment_val.strip()
    else:
        report_comment_cell = find_value_in_sheet(sheet, 'comment')
        if report_comment_cell is not None:  # find report level comment row and column in sheet
            row, col = report_comment_cell
            report_comment_row = int(sheet.cell_value(row, col + 1))
            report_comment_col = sheet.cell_value(row, col + 2)
            config_dict["sheets"][-1]["reportComment"] = {"reportCommentCellRow": report_comment_row,
                                                          "reportCommentCellColumn": report_comment_col}


def find_meta_item(config_dict, sheet, variable_name):
    # add all meta items to config_dict
    config_dict["sheets"][-1]["sheetVariables"][-1]["metaItems"] = []
    meta_items_table_row, meta_items_table_col = find_value_in_sheet(sheet, 'Meta Items Section')
    for row_num in range(meta_items_table_row, sheet.nrows):  # only look in meta items table
        found_var = sheet.cell_value(row_num, meta_items_table_col).strip()
        if found_var == variable_name:
            meta_item_var_name = sheet.cell_value(row_num, meta_items_table_col + 1).strip()
            meta_items_row = int(sheet.cell_value(row_num, meta_items_table_col + 2))
            meta_items_col = sheet.cell_value(row_num, meta_items_table_col + 3)
            config_dict["sheets"][-1]["sheetVariables"][-1]["metaItems"].append({"name": meta_item_var_name,
                                                                                 "valueCellRow": meta_items_row,
                                                                                 "valueColumn": meta_items_col})


def find_variable_comment(config_dict, sheet, variable_name):
    # add variable comment to config_dict
    comments_table_row, comments_table_col = find_value_in_sheet(sheet, 'Comments Section')
    for row_num in range(comments_table_row, sheet.nrows):  # only look in comments table
        found_var = sheet.cell_value(row_num, comments_table_col).strip()
        if found_var == variable_name:
            comment_row = int(sheet.cell_value(row_num, comments_table_col + 1))
            comment_col = sheet.cell_value(row_num, comments_table_col + 2)
    config_dict["sheets"][-1]["sheetVariables"][-1]["comment"] = {"varCommentCellRow": comment_row,
                                                                  "varCommentCellColumn": comment_col}


def write_to_json_file(config_dict):
    # takes the data in the config_dict and writes it to a local file in json format
    json_object = json.dumps(config_dict, indent=4)

    with open("config_file.json", "w") as outfile:
        outfile.write(json_object)


def json_print(j):
    print(json.dumps(j, indent=4))

