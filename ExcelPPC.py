import xlsxwriter
import re
import csv

def step1(row, listofnames):
    if pattern.search(row[0]):
        return True
    else:
        return parseStep2(row, listofnames)

def parseStep2(row, listofnames):
    patregalia = re.compile('доктор*', re.IGNORECASE)
    patexperience = re.compile('Стаж ППС*', re.IGNORECASE)
    namefound = False
    for pat, data in listofnames.items():
        namefound = pat.search(row[1])
        if namefound:
            print(listofnames[pat])
            if not patregalia.search(listofnames[pat][6]):
                listofnames[pat][6] = row[6]
            if not patexperience.search(listofnames[pat][8]):
                listofnames[pat][8] = row[8]
                listofnames[pat][9] = row[9]
            break
    if namefound:
        return False
    patname = re.compile(row[1] + '*', re.IGNORECASE)
    print(patname)
    listofnames[patname] = []
    listofnames[patname].append(row[0])
    listofnames[patname].append(row[1])
    listofnames[patname].append(row[2])
    listofnames[patname].append(row[3])
    listofnames[patname].append(row[4])
    listofnames[patname].append(row[5])
    listofnames[patname].append(row[6])
    listofnames[patname].append(row[7])
    listofnames[patname].append(row[8])
    listofnames[patname].append(row[9])
    return False


def writeToXLSX(row, x):
    y = 0
    for cell in row:
        worksheet.write(x, y, cell)
        y += 1

def genInt(vals, listofnames):
    x = 0
    y = 0
    for val in vals:
        if step1(val, listofnames):
            dictlist = []
            num = 1
            for pat, row in listofnames.items():
                row[0] = num
                num += 1
                writeToXLSX(row, x)
                x += 1
                dictlist.append(pat)
            for dictel in dictlist:
                del listofnames[dictel]
            del dictlist
            yield x, y
            x += 1
        y += 1

workbook = xlsxwriter.Workbook('excelsheets/result.xlsx')
worksheet = workbook.add_worksheet()
listofnames = {}

with open('excelsheets/sheet1.csv', 'r') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=';', quotechar='|')
    vals = []
    for row in spamreader:
        vals.append(row)

pattern = re.compile('Кафедра*', re.IGNORECASE)
print(pattern)
rownum = 0
for x,y in genInt(vals, listofnames):
    print(vals[y][1])
    print(x," ", y)
    writeToXLSX(vals[y], x)
    rownum = x
dictlist = []
num = 1
for pat, row in listofnames.items():
    row[0] = num
    num += 1
    rownum += 1
    writeToXLSX(row, rownum)
    dictlist.append(pat)
for dictel in dictlist:
    del listofnames[dictel]
del dictlist
workbook.close()



