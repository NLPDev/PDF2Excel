# from tabula import convert_into
# import sys
#
# # if len(sys.argv)!=3:
# #     print("Input pdf file name, csv file name!!")
# # else:
# #     convert_into("input.pdf", "out.csv", output_format="csv")
#
# convert_into("input01.pdf", "out.csv", output_format="csv", pages='8-9')


import xlrd
import xlwt

from xlwt import Workbook

wb=xlrd.open_workbook("output.xlsx")


for aa in wb.sheet_names():
    print(wb.sheet_by_name(aa).cell_value(4,1))