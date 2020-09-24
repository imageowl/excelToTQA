import json

import xlrd

machine_name = None
schedule_name = None
finalize = None
mode = None

def json_print(j):
    print(json.dumps(j, indent=4, sort_keys=True))

def excel_to_config_file(excel_file):
    config_dict = {"sheets": []}
    wb = xlrd.open_workbook(excel_file)

    sheet_names = wb.sheet_names()

    for name in sheet_names:
        sheet = wb.sheet_by_name(name)
        config_dict["sheets"].append({"sheetName": name})

        finalize_cell = find_value_in_sheet(sheet, 'Finalize')
        if finalize_cell is not None:
            row, col = finalize_cell
            finalize = int(sheet.cell_value(row+1, col))
            config_dict["finalize"] = finalize

        mode_cell = find_value_in_sheet(sheet, 'Mode')
        if mode_cell is not None:
            row, col = mode_cell
            mode = sheet.cell_value(row+1, col)
            config_dict["mode"] = mode

        machine_name_cell = find_value_in_sheet(sheet, 'Machine Name')
        if machine_name_cell is not None:
            row, col = machine_name_cell
            machine_name = sheet.cell_value(row+1, col)
            config_dict["machineName"] = machine_name
        else:
            machine_cell = find_value_in_sheet(sheet, 'machine')
            if machine_cell is not None:
                row, col = machine_cell
                machine_row = int(sheet.cell_value(row, col+1))
                machine_col = sheet.cell_value(row, col+2)
                config_dict["sheets"][-1]["machine"] = {"machineCellRow": machine_row, "machineCellColumn": machine_col}

        if find_value_in_sheet(sheet, 'Schedule Name') is not None:
            row, col = find_value_in_sheet(sheet, 'Schedule Name')
            schedule_name = sheet.cell_value(row+1, col)
            config_dict["scheduleName"] = schedule_name
        if find_value_in_sheet(sheet, 'schedule') is not None:
            row, col = find_value_in_sheet(sheet, 'schedule')
            schedule_row = int(sheet.cell_value(row, col+1))
            schedule_col = sheet.cell_value(row, col+2)
            config_dict["sheets"][-1]["schedule"] = {"scheduleCellRow": schedule_row, "scheduleCellColumn": schedule_col}

        if find_value_in_sheet(sheet, 'date') is not None:
            row, col = find_value_in_sheet(sheet, 'date')
            date_row = int(sheet.cell_value(row, col+1))
            date_col = sheet.cell_value(row, col+2)
            config_dict["sheets"][-1]["date"] = {"dateCellRow": date_row, "dateCellColumn": date_col}

        if find_value_in_sheet(sheet, 'report comment') is not None:
            row, col = find_value_in_sheet(sheet, 'report comment')
            report_comment_row = int(sheet.cell_value(row, col+1))
            report_comment_col = sheet.cell_value(row, col+2)
            config_dict["sheets"][-1]["reportComment"] = {"reportCommentCellRow": report_comment_row,
                                                          "reportCommentCellColumn": report_comment_col}


    json_print(config_dict)

def find_value_in_sheet(sheet, val):
    for colNum in range(sheet.ncols):
        for rowNum in range(sheet.nrows):
            if sheet.cell_value(rowNum, colNum) == val:
                return (rowNum, colNum)