import xlsxwriter
import xlrd
import re
import csv

def step(row, number, parseStep1):

    if pattern.search(row[0]):
        writeToXLSX(row, number)
        number += 1
    else:
        parseStep1(row, number)

def parseStep1(row, number):
     if row[8] == "Стаж ППС":
         writeToXLSX(row, number)
         number += 1

def writeToXLSX(row, rownumber):
    x = 0
    for column in row:
        worksheet.write(rownumber, x, column)
        x += 1

workbook = xlsxwriter.Workbook('excelsheets/result.xlsx')
worksheet = workbook.add_worksheet()

with open('excelsheets/sheet1.csv', 'r') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=';', quotechar='|')
    for row in spamreader:
        print(', '.join(row))

rb = xlrd.open_workbook('excelsheets/sheet1.xlsx', formatting_info=False)
sheet = rb.sheet_by_index(0)

vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
print(vals)
print(vals[1][0])
pattern = re.compile('Кафедра*', re.IGNORECASE)
number = 0

for row in vals:
    print(row[0])
    step(row, number, parseStep1)
workbook.close()



