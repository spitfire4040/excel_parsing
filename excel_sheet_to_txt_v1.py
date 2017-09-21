import xlrd

#opens a specfic workbook, in other words an excel file
workbook = xlrd.open_workbook('MackayTheses_test.xlsx')

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
f0 = open('MackayTheses.txt', "w")

#for some range of x rows
#for x in range(0, worksheet.nrows):
for x in range(0, 10):

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

#Number of entry
    f0.write(str(x))
    f0.write('.')
    f0.write("\n")

#Write authors to file MackayTheses_Author.txt
    for z in range(0, len(author)):
        if author[z] == '\n':
            break;
        f0.write(author[z])

    f0.write('\n')

#write publication dates to file
    if x != 0:
        pubDate = int(pubDate)
    f0.write(str(pubDate))

    f0.write('\n')


#Write titles to file
    for z in range(0, len(title)):
        f0.write(title[z])

    f0.write('\n')

#write thesisNumber to file
    if x != 0:
        thesisNumber = int(thesisNumber)

    f0.write(str(thesisNumber))

    f0.write('\n')

#write fileName to file
    for z in range(0, len(fileName)):
        f0.write(fileName[z])

    f0.write('\n')

#write link to file
    for z in range(0, len(link)):
        f0.write(link[z])

    f0.write('\n')

#make sure that last entry in file is not a new line
    if x != worksheet.nrows - 1:
        f0.write('\n')

f0.close()
