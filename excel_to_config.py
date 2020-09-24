import xlrd

machine_name = None
schedule_name = None
finalize = None
mode = None

def excel_to_config_file(excel_file):
    config_dict = {"sheets": []}
    wb = xlrd.open_workbook(excel_file)

    sheet_names = wb.sheet_names()

    for name in sheet_names:
        sheet = wb.sheet_by_name(name)
        config_dict["sheets"].append({"sheetName": name})

        if find_value_in_sheet(sheet, 'Finalize') is not None:
            row, col = find_value_in_sheet(sheet, 'Finalize')
            finalize = int(sheet.cell_value(row+1, col))
            config_dict["finalize"] = finalize

        if find_value_in_sheet(sheet, 'Mode') is not None:
            row, col = find_value_in_sheet(sheet, 'Mode')
            mode = sheet.cell_value(row+1, col)
            config_dict["mode"] = mode

        if find_value_in_sheet(sheet, 'Machine Name') is not None:
            row, col = find_value_in_sheet(sheet, 'Machine Name')
            machine_name = sheet.cell_value(row+1, col)
            config_dict["machineName"] = machine_name
        elif find_value_in_sheet(sheet, 'machine') is not None:
            row, col = find_value_in_sheet(sheet, 'machine')
            machine_row = sheet.cell_value(row, col+1)
            machine_col = sheet.cell_value(row + 1, col+2)
            # machine =

        if find_value_in_sheet(sheet, 'Schedule Name') is not None:
            row, col = find_value_in_sheet(sheet, 'Schedule Name')
            schedule_name = sheet.cell_value(row+1, col)
            config_dict["scheduleName"] = schedule_name


    print(config_dict)

def find_value_in_sheet(sheet, val):
    for colNum in range(sheet.ncols):
        for rowNum in range(sheet.nrows):
            if sheet.cell_value(rowNum, colNum) == val:
                return (rowNum, colNum)