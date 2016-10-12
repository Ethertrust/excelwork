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
    print(spamreader)
    vals = []
    for row in spamreader:
        vals.append(row)
print(vals)

pattern = re.compile('Кафедра*', re.IGNORECASE)
number = 0

for row in vals:
    print(row[0])
    step(row, number, parseStep1)
workbook.close()



