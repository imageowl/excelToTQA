# sources: https://www.geeksforgeeks.org/reading-excel-file-using-python/
# https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/

import xlrd

# filepath = "/Users/annafronhofer/Desktop/testFiles/CurrentFormat.xlsx"
filepath = "/Users/annafronhofer/Desktop/testFiles/OldFormat.xls"

# to open workbook
wb = xlrd.open_workbook(filepath)
# get sheet names
print("Sheet names: ", wb.sheet_names())

# get sheet by name
sheet1 = wb.sheet_by_name("Seasons")
print("Sheet 1 name: ", sheet1.name)
# or get sheet by index
sheet2 = wb.sheet_by_index(1)
print("Sheet 2 name: ", sheet2.name)

# for row 0 and column 0
print("Row 0, Column 0: ", sheet1.cell_value(0, 0))

# extracting number of rows
print("Number of rows: ", sheet1.nrows)
# extracting number of columns
print("Number of columns: ", sheet1.ncols, '\n')

# extracting all column names
print("Column Names: ")
for colNum in range(sheet1.ncols):
    print(sheet1.cell_value(0, colNum))

print('\n')

# extracting first column
print("First Column: ")
for rowNum in range(sheet1.nrows):
    print(sheet1.cell_value(rowNum, 0))

print('\n')

# extract a particular row value
print("First Row: ", sheet1.row_values(0))






