import xlsxwriter
import xlrd
import re
from itertools import count

workbook = xlsxwriter.Workbook('excelsheets/result.xlsx')
worksheet = workbook.add_worksheet()

rb = xlrd.open_workbook('excelsheets/sheet1.xlsx', formatting_info=False)
sheet = rb.sheet_by_index(0)

vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
print(vals)
print(vals[1][0])
pattern = re.compile('Кафедра*', re.IGNORECASE)
number = 0
for row in vals:
    step(row, number, writeToXLSX, DOSMTE)



def step(row, number, writexlsx, DOSMTE):
    if pattern.search(row[0]):
        writexlsx(row, number)
        number += 1
    else:
        DOSMTE(row)

def writeToXLSX(row, rownumber):
    for column in row:
        worksheet.write(rownumber, count(0), column)

def DOSMTE(row):
     if row[8] == "Стаж ППС":


