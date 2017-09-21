readFile = open("MackayTheses.csv", "r")
writeFile = open("ThesesDocument.txt", "w")

resetFlag1 = False
resetFlag2 = False
resetFlag3 = False

#Reads the first line, which is just formatting info
char = readFile.readline();
print(char);
print("HELLLLLLOOOO")

#Start at beginning of file
readFile_Amount = 1
x=0

filePos = readFile.tell()
#While we are not at the end of the file
while readFile.read(readFile_Amount):
#while x != 10:
#    x+=1
    if resetFlag1 == False:
        readFile.seek(filePos)
        resetFlag1 = True

    #print("YOOOOOOO")
    #if quotation marks are found
    filePos = readFile.tell()
    if(readFile.read(readFile_Amount) == '"'):
        if resetFlag2 == False:
            readFile.seek(filePos)
            resetFlag2 = True
        #print("DOOODUE")
        #advance to next character in file
        readFile.read(readFile_Amount)
        #Read until next quotation mark
        filePos = readFile.tell()
        while readFile.read(readFile_Amount) != '"':
            readFile.seek(filePos)
            #print("YUH")
            char = readFile.read(readFile_Amount)
            writeFile.write(char)
            filePos = readFile.tell()
