import json

import xlrd


def json_print(j):
    print(json.dumps(j, indent=4, sort_keys=True))

def excel_to_config_file(excel_file):
    finalize = None
    mode = None
    config_dict = {"sheets": []}
    wb = xlrd.open_workbook(excel_file)

    sheet_names = wb.sheet_names()

    for name in sheet_names:
        sheet = wb.sheet_by_name(name)
        config_dict["sheets"].append({"sheetName": name})

        if finalize is None:
            finalize_cell = find_value_in_sheet(sheet, 'Finalize')
            if finalize_cell is not None:
                row, col = finalize_cell
                finalize = int(sheet.cell_value(row+1, col))
                config_dict["finalize"] = finalize

        if mode is None:
            mode_cell = find_value_in_sheet(sheet, 'Mode')
            if mode_cell is not None:
                row, col = mode_cell
                mode = sheet.cell_value(row+1, col)
                config_dict["mode"] = mode.strip()

        machine_name_cell = find_value_in_sheet(sheet, 'Machine Name')
        if machine_name_cell is not None:
            row, col = machine_name_cell
            machine_name = sheet.cell_value(row+1, col)
            config_dict["machineName"] = machine_name.strip()
        else:
            machine_cell = find_value_in_sheet(sheet, 'machine')
            if machine_cell is not None:
                row, col = machine_cell
                machine_row = int(sheet.cell_value(row, col+1))
                machine_col = sheet.cell_value(row, col+2)
                config_dict["sheets"][-1]["machine"] = {"machineCellRow": machine_row, "machineCellColumn": machine_col}

        schedule_name_cell = find_value_in_sheet(sheet, 'Schedule Name')
        if schedule_name_cell is not None:
            row, col = schedule_name_cell
            schedule_name = sheet.cell_value(row+1, col)
            config_dict["scheduleName"] = schedule_name.strip()
        else:
            schedule_cell = find_value_in_sheet(sheet, 'schedule')
            if schedule_cell is not None:
                row, col = schedule_cell
                schedule_row = int(sheet.cell_value(row, col+1))
                schedule_col = sheet.cell_value(row, col+2)
                config_dict["sheets"][-1]["schedule"] = {"scheduleCellRow": schedule_row, "scheduleCellColumn": schedule_col}

        date_cell = find_value_in_sheet(sheet, 'date')
        if date_cell is not None:
            row, col = date_cell
            date_row = int(sheet.cell_value(row, col+1))
            date_col = sheet.cell_value(row, col+2)
            config_dict["sheets"][-1]["date"] = {"dateCellRow": date_row, "dateCellColumn": date_col}

        report_comment_cell = find_value_in_sheet(sheet, 'report comment')
        if report_comment_cell is not None:
            row, col = report_comment_cell
            report_comment_row = int(sheet.cell_value(row, col+1))
            report_comment_col = sheet.cell_value(row, col+2)
            config_dict["sheets"][-1]["reportComment"] = {"reportCommentCellRow": report_comment_row,
                                                          "reportCommentCellColumn": report_comment_col}

        variables_table_cell = find_value_in_sheet(sheet, 'Variables Table')
        if variables_table_cell is not None:
            config_dict["sheets"][-1]["sheetVariables"] = []
            row, col = variables_table_cell
            is_empty = False
            row_idx = row+2
            while is_empty is False:
                variable_name = sheet.cell_value(row_idx, col).strip()
                if len(variable_name) != 0:
                    variable_row = int(sheet.cell_value(row_idx, col+1))
                    variable_col = sheet.cell_value(row_idx, col+2)
                    config_dict["sheets"][-1]["sheetVariables"].append({"name": variable_name.strip(),
                                                                        "valueCellRow": variable_row,
                                                                        "valueCellColumn": variable_col})
                    if sheet.cell_value(row_idx, col+3).lower() == "yes":
                        # meta items
                        config_dict["sheets"][-1]["sheetVariables"][-1]["metaItems"] = []
                        meta_items_table_row, meta_items_table_col = find_value_in_sheet(sheet, 'Meta Items Table')
                        for rowNum in range(meta_items_table_row, sheet.nrows):
                            found_var = sheet.cell_value(rowNum, meta_items_table_col).strip()
                            if found_var == variable_name:
                                meta_item_var_name = sheet.cell_value(rowNum, meta_items_table_col + 1)
                                meta_items_row = int(sheet.cell_value(rowNum, meta_items_table_col + 2))
                                meta_items_col = sheet.cell_value(rowNum, meta_items_table_col + 3)
                        config_dict["sheets"][-1]["sheetVariables"][-1]["metaItems"].append({"name": meta_item_var_name,
                                                        "valueCellRow": meta_items_row, "valueColumn": meta_items_col})

                    if sheet.cell_value(row_idx, col+4).lower() == "yes":
                        # variable comment
                        comments_table_row, comments_table_col = find_value_in_sheet(sheet, 'Comments Table')
                        for rowNum in range(comments_table_row, sheet.nrows):
                            found_var = sheet.cell_value(rowNum, comments_table_col).strip()
                            if found_var == variable_name:
                                comment_row = int(sheet.cell_value(rowNum, comments_table_col+1))
                                comment_col = sheet.cell_value(rowNum, comments_table_col+2)
                        config_dict["sheets"][-1]["sheetVariables"][-1]["comment"] = {"varCommentCellRow": comment_row,
                                                                                      "varCommentCellColumn": comment_col}
                else:
                    is_empty = True

                row_idx += 1

    json_print(config_dict)


def find_value_in_sheet(sheet, val):
    for col_num in range(sheet.ncols):
        for row_num in range(sheet.nrows):
            if sheet.cell_value(row_num, col_num) == val:
                return row_num, col_num
