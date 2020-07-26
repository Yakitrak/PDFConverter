import sys
import os
import glob
import win32com.client

def main():
    ifConvert = input("Enter [YES] to convert files or anything else to exit: ")
    if (ifConvert.upper() == "YES"):
        callConvert()
    else:
        print("Thank you for using the PDF Converter")
        exit()
        
def callConvert():
    name = input("What is the file location to convert from: ")
    convFileType = input("Would you like to convert [PPT] or [DOC] or [BOTH]: ")
    if (convFileType.upper() == "PPT"):
        files = glob.glob(name + "\*.ppt*")
        convertPPT(files)
    elif (convFileType.upper() == "DOC"):
        files = glob.glob(name + "\*.doc*")
        convertDOC(files)
    elif (convFileType.upper() == "BOTH"):
        files = glob.glob(name + "\*.ppt*")
        convertPPT(files)
        files = glob.glob(name + "\*.doc*")
        convertDOC(files)
    else:
        print("Sorry, that is not one of the options. Please try again")
    main()

#Converts powerpoints (format 32) to PDFs
def convertPPT(files, formatType=32):
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    powerpoint.Visible = 1
    print(files)
    for filename in files:
        newname = os.path.splitext(filename)[0] + ".pdf"
        file = powerpoint.Presentations.Open(filename)
        file.SaveAs(newname, formatType)
        file.Close()
    powerpoint.Quit()
    
#Converts documents (format 17) to PDFs
def convertDOC(files, formatType=17):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = 1
    print(files)
    for filename in files:
        newname = os.path.splitext(filename)[0] + ".pdf"
        file = word.Documents.Open(filename)
        file.SaveAs(newname, formatType)
        file.Close()
    word.Quit()

main()
