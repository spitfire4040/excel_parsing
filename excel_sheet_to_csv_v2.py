import xlrd

#opens a specfic workbook, in other words an excel file
workbook = xlrd.open_workbook('FINAL_MackayTheses_Inventory.xlsx')

#opens a specific worksheet within the previously openned excel file
worksheet = workbook.sheet_by_name('Sheet1')

#accept user input
#excelFile = input("Enter the full name of excel file: ")
#worksheetName = input("Specify which sheet to read from: ")
#outputFile = input("Enter the full name of output file: ")
#columnsToBeRead = input("Enter the number of columns (horizontal) to be read: ")
#rowsToBeRead = input("Enter the number of rows (vertical) to be read: ")

#on the first row declare all values as NULL
author = 'NULL'
pubDate = 'NULL'
title = 'NULL'
thesisNumber = 'NULL'
fileName = 'NULL'
link = 'NULL'

#open text file to write to
f0 = open('MackayTheses.csv', "w")

#for some range of x rows
for x in range(0, worksheet.nrows):

    #for some range of y columns
    for y in range(0, worksheet.ncols):
        #depending on which column, update values
        if y == 0:
            author = worksheet.cell(x, y).value
        if y == 2:
            pubDate = worksheet.cell(x, y).value
        if y == 1:
            title = worksheet.cell(x, y).value
        if y == 10:
            thesisNumber = worksheet.cell(x, y).value
        if y == 6:
            fileName = worksheet.cell(x, y).value
        if y == 12:
            link = worksheet.cell(x, y).value

#If commaFlag is true, that means that commas exist in
# the data cell. They are then formatted for csv

#Write authors to file MackayTheses_Author.txt
    commaFlag = False

    for z in range(0, len(author)):
        if author[z] == ',':
            commaFlag = True
            break;
    if commaFlag == True:
        f0.write('"')

    for z in range(0, len(author)):
        if author[z] == '\n':
            break;
        f0.write(author[z])

    if commaFlag == True:
        f0.write('"')

    f0.write(',')

#write publication dates to file
    f0.write(str(pubDate))

    f0.write(',')


#Write titles to file
    commaFlag = False

    for z in range(0, len(title)):
        if title[z] == ',':
            commaFlag = True
            break;
    if commaFlag == True:
        f0.write('"')

    for z in range(0, len(title)):
        f0.write(title[z])

    if commaFlag == True:
        f0.write('"')

    f0.write(',')

#write thesisNumber to file
    f0.write(str(thesisNumber))

    f0.write(',')

#write fileName to file
    commaFlag = False

    for z in range(0, len(fileName)):
        if fileName[z] == ',':
            commaFlag = True
            break;
    if commaFlag == True:
        f0.write('"')

    for z in range(0, len(fileName)):
        f0.write(fileName[z])

    if commaFlag == True:
        f0.write('"')

    f0.write(',')

#write link to file
    commaFlag = False

    for z in range(0, len(link)):
        if link[z] == ',':
            commaFlag = True
            break;
    if commaFlag == True:
        f0.write('"')

    for z in range(0, len(link)):
        f0.write(link[z])

    if commaFlag == True:
        f0.write('"')

#make sure that last entry in file is not a new line
    if x != worksheet.nrows - 1:
        f0.write('\n')

f0.close()
