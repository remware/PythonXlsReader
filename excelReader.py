import xlrd
import sys, datetime

book = xlrd.open_workbook("Sample_moves.xlsx")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))

sh = book.sheet_by_index(0)
print("{0} rows {1} columns {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell B10 is {0}".format(sh.cell_value(rowx=9, colx=1)))

start_time = "7:00"
end_time = "20:00"

for rx in range(sh.nrows):
    ctys = sh.row_types(rx)
    cvals = sh.row_values(rx)
    ctype = ctys[0]
    cdate = cvals[0]
    cvalu = cvals[1]
    showval = cdate
    print("cell on {0} value {1} ".format(showval, cvalu) )