import xlrd

#opens a specfic workbook, in other words an excel file
workbook = xlrd.open_workbook('FINAL_MackayTheses_Inventory.xlsx')

#opens a specific worksheet within the previously openned excel file
worksheet = workbook.sheet_by_name('Sheet1')

#open text file to write to
f0 = open("MackayTheses_Author.txt", "w")
f1 = open("MackayTheses_Title.txt", "w")
f2 = open("MackayTheses_pubDate.txt", "w")
f3 = open("MackayTheses_Department.txt", "w")
f4 = open("MackayTheses_Degree.txt", "w")
f5 = open("MackayTheses_ThesisDegreeName.txt", "w")
f6 = open("MackayTheses_FileName.txt", "w")

#for some range of x rows
for x in range(1, 10):
    #on the first row declare all values as NULL
    if x == 1:
        author = 'NULL'
        title = 'NULL'
        pubDate = 'NULL'
        department = 'NULL'
        degree = 'NULL'
        thesisDegreeName = 'NULL'
        fileName = 'NULL'

    #for some range of y columns
    for y in range(0,7):

        #depending on which column, update values
                if y == 0:
                    author = worksheet.cell(x, y).value
                if y == 1:
                    title = worksheet.cell(x, y).value
                if y == 2:
                    pubDate = worksheet.cell(x, y).value
                if y == 3:
                    department = worksheet.cell(x, y).value
                if y == 4:
                    degree = worksheet.cell(x, y).value
                if y == 5:
                    thesisDegreeName = worksheet.cell(x, y).value
                if y == 6:
                    fileName = worksheet.cell(x, y).value

#If commaFlag is true, that means that commas exist in
# the data cell. They are then formatted for csv
    commaFlag

#Write authors to file MackayTheses_Author.txt
    commaFlag = False

    for z in range(0, len(author)):
        if author[z] == ',':
            commaFlag = True:
            break;
    if commaFlag == True:
        f0.write('"')

    for z in range(0, len(author)):
        if author[z] == '\n':
            break;
        f0.write(author[z])

    if commaFlag == True:
        f0.write('"')


#Write titles to file
    commaFlag = False

    for z in range(0, len(title)):
        if title[z] == ',':
            commaFlag = True:
            break;
    if commaFlag == True:
        f1.write('"')

    for z in range(0, len(title)):
        f1.write(title[z])

    if commaFlag == True:
        f1.write('"')

#write publication dates to file
    commaFlag = False

    for z in range(0, len(pubDate)):
        if pubDate[z] == ',':
            commaFlag = True:
            break;
    if commaFlag == True:
        f2.write('"')

    f2.write(str(pubDate))

    if commaFlag == True:
        f2.write('"')


#write department to file
    for z in range(0, len(department)):
        f3.write(department[z])
    f3.write('\n')

#write degree to file
    for z in range(0, len(degree)):
        f4.write(degree[z])
    f4.write('\n')

#write thesisDegreeName to file
    for z in range(0, len(thesisDegreeName)):
        f5.write(thesisDegreeName[z])
    f5.write('\n')

f0.close()
f1.close()
f2.close()
f3.close()
f4.close()
f5.close()
f6.close()
