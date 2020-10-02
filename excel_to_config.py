import json
import xlrd


def excel_to_config_file(excel_file):
    # take the information in the excel file and convert it to the json config file of a specific format

    config_dict = {"data": [{}]}  # dictionary that will be used to create json config file
    excel_workbook = xlrd.open_workbook(excel_file)

    sheet = excel_workbook.sheet_by_name('Config')

    find_header_value(sheet, config_dict, "Machine Name", "machine")
    find_header_value(sheet, config_dict, "Schedule Name", "schedule")
    find_header_value(sheet, config_dict, "Finalize Value", "finalize")
    find_header_value(sheet, config_dict, "Save Mode", "mode")
    find_header_value(sheet, config_dict, "Report Date", "date", excel_workbook)
    find_header_value(sheet, config_dict, "Report Comment", "reportComment")

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

                if sheet.cell_value(row_idx, variables_table_col+5).lower() == "yes":  # variable has a variable comment
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


def find_header_value(sheet, config_dict, header_name, header_cell_name, excel_workbook=None):
    name = ""

    name_header = find_value_in_sheet(sheet, header_name)
    if name_header is not None:  # find header name in sheet
        row, col = name_header
        name = sheet.cell_value(row + 1, col)
        if name != '':
            if isinstance(name, str):  # machine, schedule, mode or report comment
                config_dict[header_cell_name] = name.strip()
            elif isinstance(name, float):
                if name < 2:  # finalize
                    config_dict[header_cell_name] = int(name)
                else:  # report date
                    report_date = xlrd.xldate_as_datetime(name, excel_workbook.datemode)
                    config_dict[header_cell_name] = str(report_date)

    if name == '':
        cell = find_value_in_sheet(sheet, header_cell_name)
        if cell is not None:  # find machine name row and column in sheet
            row, col = cell
            value_row = int(sheet.cell_value(row, col + 1))
            value_col = sheet.cell_value(row, col + 2)
            value_sheet = sheet.cell_value(row, col + 3)
            config_dict["data"][0][header_cell_name] = {"cellRow": value_row, "cellColumn": value_col,
                                                        "sheetName": value_sheet}


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
    config_dict["data"][0]["variables"][-1]["comment"] = {"varCommentCellRow": comment_row,
                                                          "varCommentCellColumn": comment_col,
                                                          "sheetName": comment_sheet}


def write_to_json_file(config_dict):
    # takes the data in the config_dict and writes it to a local file in json format
    json_object = json.dumps(config_dict, indent=4)

    with open("config_file.json", "w") as outfile:
        outfile.write(json_object)


def json_print(j):
    print(json.dumps(j, indent=4))

