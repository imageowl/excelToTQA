import json
import xlrd


def excel_to_config_file(excel_file):
    finalize_val = None
    mode_val = None
    config_dict = {"sheets": []}  # dictionary that will be used to create json config file
    wb = xlrd.open_workbook(excel_file)

    for sheet in wb.sheets():
        sheet_name = sheet.name
        config_dict["sheets"].append({"sheetName": sheet_name})

        if finalize_val is None:
            finalize_cell = find_value_in_sheet(sheet, 'Finalize')  # find finalize in sheet
            if finalize_cell is not None:
                row, col = finalize_cell
                finalize_val = int(sheet.cell_value(row+1, col))
                config_dict["finalize"] = finalize_val
            else:
                finalize_cell = find_value_in_sheet(sheet, 'finalize')
                if finalize_cell is not None:  # find finalize value row and column in sheet
                    row, col = finalize_cell
                    finalize_row = int(sheet.cell_value(row, col+1))
                    finalize_col = sheet.cell_value(row, col+2)
                    config_dict["sheets"][-1]["finalize"] = {"finalizeCellRow": finalize_row,
                                                             "finalizeCellColumn": finalize_col}

        if mode_val is None:
            mode_cell = find_value_in_sheet(sheet, 'Mode')  # find mode in sheet
            if mode_cell is not None:
                row, col = mode_cell
                mode_val = sheet.cell_value(row+1, col)
                config_dict["mode"] = mode_val.strip()
            else:
                mode_cell = find_value_in_sheet(sheet, 'mode')
                if mode_cell is not None:  # find mode value row and column in sheet
                    row, col = mode_cell
                    mode_row = int(sheet.cell_value(row, col+1))
                    mode_col = sheet.cell_value(row, col+2)
                    config_dict["sheets"][-1]["mode"] = {"modeCellRow": mode_row,
                                                         "modeCellColumn": mode_col}

        machine_name_cell = find_value_in_sheet(sheet, 'Machine Name')
        if machine_name_cell is not None:  # find machine name in sheet
            row, col = machine_name_cell
            machine_name = sheet.cell_value(row+1, col)
            config_dict["machineName"] = machine_name.strip()
        else:
            machine_cell = find_value_in_sheet(sheet, 'machine')
            if machine_cell is not None:  # find machine name row and column in sheet
                row, col = machine_cell
                machine_row = int(sheet.cell_value(row, col+1))
                machine_col = sheet.cell_value(row, col+2)
                config_dict["sheets"][-1]["machine"] = {"machineCellRow": machine_row, "machineCellColumn": machine_col}

        schedule_name_cell = find_value_in_sheet(sheet, 'Schedule Name')
        if schedule_name_cell is not None:  # find schedule name in sheet
            row, col = schedule_name_cell
            schedule_name = sheet.cell_value(row+1, col)
            config_dict["scheduleName"] = schedule_name.strip()
        else:
            schedule_cell = find_value_in_sheet(sheet, 'schedule')
            if schedule_cell is not None:  # find schedule name row and column in sheet
                row, col = schedule_cell
                schedule_row = int(sheet.cell_value(row, col+1))
                schedule_col = sheet.cell_value(row, col+2)
                config_dict["sheets"][-1]["schedule"] = {"scheduleCellRow": schedule_row, "scheduleCellColumn": schedule_col}

        date_cell = find_value_in_sheet(sheet, 'Report Date')
        if date_cell is not None:  # find date in sheet
            row, col = date_cell
            date_val = sheet.cell_value(row + 1, col)
            report_date = xlrd.xldate_as_datetime(date_val, wb.datemode)
            config_dict["date"] = str(report_date)
        else:
            date_cell = find_value_in_sheet(sheet, 'date')
            if date_cell is not None:  # find date row and column in sheet
                row, col = date_cell
                date_row = int(sheet.cell_value(row, col+1))
                date_col = sheet.cell_value(row, col+2)
                config_dict["sheets"][-1]["date"] = {"dateCellRow": date_row, "dateCellColumn": date_col}

        report_comment_cell = find_value_in_sheet(sheet, 'Report Comment')
        if report_comment_cell is not None:  # find report comment in sheet
            row, col = report_comment_cell
            report_comment_val = sheet.cell_value(row+1, col)
            config_dict["reportComment"] = report_comment_val.strip()
        else:
            report_comment_cell = find_value_in_sheet(sheet, 'report comment')
            if report_comment_cell is not None:  # find report level comment row and column in sheet
                row, col = report_comment_cell
                report_comment_row = int(sheet.cell_value(row, col+1))
                report_comment_col = sheet.cell_value(row, col+2)
                config_dict["sheets"][-1]["reportComment"] = {"reportCommentCellRow": report_comment_row,
                                                              "reportCommentCellColumn": report_comment_col}

        variables_table_cell = find_value_in_sheet(sheet, 'Variables Table')
        if variables_table_cell is not None:  # find table with variables in sheet
            config_dict["sheets"][-1]["sheetVariables"] = []
            row, col = variables_table_cell
            is_empty = False
            row_idx = row+2
            while is_empty is False:  # add all variables to config_dict
                variable_name = sheet.cell_value(row_idx, col).strip()
                if len(variable_name) != 0:
                    variable_row = int(sheet.cell_value(row_idx, col+1))
                    variable_col = sheet.cell_value(row_idx, col+2)
                    config_dict["sheets"][-1]["sheetVariables"].append({"name": variable_name.strip(),
                                                                        "valueCellRow": variable_row,
                                                                        "valueCellColumn": variable_col})
                    if sheet.cell_value(row_idx, col+3).lower() == "yes":  # variable has meta items
                        # add all meta items to config_dict
                        config_dict["sheets"][-1]["sheetVariables"][-1]["metaItems"] = []
                        meta_items_table_row, meta_items_table_col = find_value_in_sheet(sheet, 'Meta Items Table')
                        for rowNum in range(meta_items_table_row, sheet.nrows):  # only look in meta items table
                            found_var = sheet.cell_value(rowNum, meta_items_table_col).strip()
                            if found_var == variable_name:
                                meta_item_var_name = sheet.cell_value(rowNum, meta_items_table_col + 1).strip()
                                meta_items_row = int(sheet.cell_value(rowNum, meta_items_table_col + 2))
                                meta_items_col = sheet.cell_value(rowNum, meta_items_table_col + 3)
                                config_dict["sheets"][-1]["sheetVariables"][-1]["metaItems"].append({"name": meta_item_var_name,
                                                        "valueCellRow": meta_items_row, "valueColumn": meta_items_col})

                    if sheet.cell_value(row_idx, col+4).lower() == "yes":  # variable has a variable comment
                        # add variable comment to config_dict
                        comments_table_row, comments_table_col = find_value_in_sheet(sheet, 'Comments Table')
                        for rowNum in range(comments_table_row, sheet.nrows):  # only look in comments table
                            found_var = sheet.cell_value(rowNum, comments_table_col).strip()
                            if found_var == variable_name:
                                comment_row = int(sheet.cell_value(rowNum, comments_table_col+1))
                                comment_col = sheet.cell_value(rowNum, comments_table_col+2)
                        config_dict["sheets"][-1]["sheetVariables"][-1]["comment"] = {"varCommentCellRow": comment_row,
                                                                                      "varCommentCellColumn": comment_col}
                else:
                    is_empty = True  # no more variables present in this sheet

                row_idx += 1  # check next row for variable

    json_print(config_dict)
    # write_to_json_file(config_dict)


def find_value_in_sheet(sheet, val):
    # look through entire sheet for specified value
    for col_num in range(sheet.ncols):
        for row_num in range(sheet.nrows):
            if sheet.cell_value(row_num, col_num) == val:
                return row_num, col_num


def write_to_json_file(config_dict):
    json_object = json.dumps(config_dict, indent=4)

    with open("config_file.json", "w") as outfile:
        outfile.write(json_object)


def json_print(j):
    print(json.dumps(j, indent=4))

