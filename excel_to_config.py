import json
import xlrd


def excel_to_config_file(excel_file):
    # take the information in the excel file and convert it to the JSON config file

    config_dict = {"data": [{}]}  # dictionary that will be used to create JSON config file
    excel_workbook = xlrd.open_workbook(excel_file)

    sheet = excel_workbook.sheet_by_name("Config")  # excel sheet with all the necessary information

    # add the values or row/column indices to the config_dict for all the data headers below
    find_header_value(sheet, config_dict, "Machine Name", "machine")
    find_header_value(sheet, config_dict, "Schedule Name", "schedule")
    find_header_value(sheet, config_dict, "Finalize Value", "finalize")
    find_header_value(sheet, config_dict, "Save Mode", "mode")
    find_header_value(sheet, config_dict, "Report Date", "date", excel_workbook)
    find_header_value(sheet, config_dict, "Report Comment", "reportComment")

    # find table with the variables in the sheet
    variables_table_cell = find_phrase_in_sheet(sheet, "Variables Section")
    if variables_table_cell is not None:
        config_dict["data"][0]["variables"] = []  # add python list for variables to config_dict
        variables_table_row, variables_table_col = variables_table_cell

        cell_is_empty = False
        row_idx = variables_table_row + 2  # start at first row of variables table
        while cell_is_empty is False:  # move down the column adding variables to the list until an empty cell is found
            variable_name = sheet.cell_value(row_idx, variables_table_col).strip()
            if len(variable_name) != 0:  # variable name found
                var_row = int(sheet.cell_value(row_idx, variables_table_col + 1))
                var_col = sheet.cell_value(row_idx, variables_table_col + 2)
                var_sheet = sheet.cell_value(row_idx, variables_table_col + 3).strip()
                config_dict["data"][0]["variables"].append({"name": variable_name.strip(), "valueCellRow": var_row,
                                                            "valueCellColumn": var_col, "sheetName": var_sheet})
                if sheet.cell_value(row_idx, variables_table_col + 4).lower() == "yes":  # variable has meta items
                    # add all meta items to config_dict
                    find_meta_item(config_dict, sheet, variable_name.strip())

                if sheet.cell_value(row_idx, variables_table_col + 5).lower() == "yes":  # variable has a comment
                    # add the variable comment to config_dict
                    find_variable_comment(config_dict, sheet, variable_name.strip())
            else:
                cell_is_empty = True  # no more variables present in this sheet

            row_idx += 1  # check next row for variable name

    write_to_json_file(config_dict)  # create the JSON config file


def find_phrase_in_sheet(sheet, phrase):
    # look through entire sheet for specified term or phrase, return the indices of the first one found
    for col_num in range(sheet.ncols):
        for row_num in range(sheet.nrows):
            if sheet.cell_value(row_num, col_num) == phrase:
                return row_num, col_num


def find_header_value(sheet, config_dict, header_name, header_cell_name, excel_workbook=None):
    header_value = ""

    value_header = find_phrase_in_sheet(sheet, header_name)
    if value_header is not None:  # find header name in sheet
        row, col = value_header
        header_value = sheet.cell_value(row + 1, col)
        if header_value != "":  # header name was found
            if isinstance(header_value, str):  # machine, schedule, mode or report comment
                config_dict[header_cell_name] = header_value.strip()
            elif isinstance(header_value, float):
                if header_value < 2:  # finalize
                    config_dict[header_cell_name] = int(header_value)
                else:  # report date
                    report_date = xlrd.xldate_as_datetime(header_value, excel_workbook.datemode)
                    config_dict[header_cell_name] = str(report_date)

    if header_value == "":
        cell = find_phrase_in_sheet(sheet, header_cell_name)
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
    meta_items_table_row, meta_items_table_col = find_phrase_in_sheet(sheet, "Meta Items Section")
    cell_is_empty = False
    row_num = meta_items_table_row + 2  # start at first row of meta items table
    while cell_is_empty is False:  # move down the column adding meta items to the list until an empty cell is found
        found_var = sheet.cell_value(row_num, meta_items_table_col).strip()
        if len(found_var) != 0:
            if found_var == variable_name:  # check if each meta item is associated with specified variable
                meta_item_var_name = sheet.cell_value(row_num, meta_items_table_col + 1).strip()
                meta_items_row = int(sheet.cell_value(row_num, meta_items_table_col + 2))
                meta_items_col = sheet.cell_value(row_num, meta_items_table_col + 3)
                meta_items_sheet = sheet.cell_value(row_num, meta_items_table_col + 4)
                config_dict["data"][0]["variables"][-1]["metaItems"].append({"name": meta_item_var_name,
                                                                             "valueCellRow": meta_items_row,
                                                                             "valueColumn": meta_items_col,
                                                                             "sheetName": meta_items_sheet})
        else:
            cell_is_empty = True  # no more meta items present in this sheet

        row_num += 1  # check next row for another meta item


def find_variable_comment(config_dict, sheet, variable_name):
    # add variable comment to config_dict
    comments_table_row, comments_table_col = find_phrase_in_sheet(sheet, "Comments Section")
    for row_num in range(comments_table_row, sheet.nrows):  # only look in comments table
        found_var = sheet.cell_value(row_num, comments_table_col).strip()
        if found_var == variable_name:  # check if each comment is associated with specified variable
            comment_row = int(sheet.cell_value(row_num, comments_table_col + 1))
            comment_col = sheet.cell_value(row_num, comments_table_col + 2)
            comment_sheet = sheet.cell_value(row_num, comments_table_col + 3)
    config_dict["data"][0]["variables"][-1]["comment"] = {"varCommentCellRow": comment_row,
                                                          "varCommentCellColumn": comment_col,
                                                          "sheetName": comment_sheet}


def write_to_json_file(config_dict):
    # takes the data in the config_dict and writes it to a local file in JSON format
    json_object = json.dumps(config_dict, indent=4)

    with open("config_output_file.json", "w") as outfile:
        outfile.write(json_object)


def json_print(j):
    print(json.dumps(j, indent=4))

