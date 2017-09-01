#-------------------------------------------------------------------------------
# Name:        All in one, modular Word/Powerpoint/PDF converter, merger, spliter- BETA
# Purpose:  Title
#
# Author:      ElectroJo
#
# Created:     2017
# Copyright:   (c) ElectroJo
# Licence:     N/A
#-------------------------------------------------------------------------------

from tkinter import *
from os import walk
import win32com.client
import os
import sys
from tkinter.filedialog import askdirectory, askopenfilename

try:
    from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
except ImportError:
    print("Error, Py2PDF not installed")

import time
import re

foldername = ""
AllFiles = []
master = Tk()
window = Toplevel()
window.withdraw()
currentdir = StringVar()
currentdir.set('There is no dir selected.')
variable123 = StringVar(window)
variable123.set("Please Pick an Output FileType")
wordtonum = {".doc":0,".dot":1,".rtf":6,".txt":7,".mht":9,".htm":10,".xml":11,".docm":13,".dotx":14,".dotm":15,".docx":16,".pdf":17,".xps":18,".odt":23}
WordConvertMenu = wordtonum.copy()
PowerpointToNum = {".ppt":1,".pot":5,".rtf":6,".pps":7,"Meta Folder":15,"Each Slide as a .gif":16,"Each Slide as a .jpg":17,"Each Slide as a .png":18,"Each Slide as a .bmp":19,"Each Slide as a .tif":21,"Each Slide as a .emf":23,".pptx":24,".pptm":25,
".ppsx":28,".ppsm":29,".thmx":31,".pdf":32,".xps":33,".xml":34,".odp":35,".wmv (This Rarley Works)":37,".mp4 (This Rarley Works)":39}
PPTConvertMenu = {".ppt":1,".pot":5,".rtf":6,".pps":7,".pptx":24,".pptm":25,".ppsx":28,".ppsm":29,".xml":34,".odp":35,}
#list for adding https://msdn.microsoft.com/en-us/library/ff839952.aspx

#List of numbers for powerpoint https://msdn.microsoft.com/en-us/library/office/ff746500.aspx

#List for Excel https://msdn.microsoft.com/en-us/library/office/ff198017.aspx

##All tkinter code was learned from various wiki pages on how to use it, however all code that uses tkinter was made by me.
##Most of the wiki pages were located on http://effbot.org/

def GetInput():
    global cancle, ConvertTheDirWord, ConvertTheDirPPT
    cancle = Button(window,text="Cancel", width=10, command=WindowClear)
    MasterLable = Label(master, textvariable = currentdir).pack()
    SelectDirMaster = Button(master, text="Select a Dir", command=SelectDir).pack()
    ConvertTheDirWord = Button(master, text="Convert between Word filetypes in the current Dir", command= lambda: ConvertMenu("Word")).pack()
    ConvertTheDirPPT = Button(master, text="Convert between Powerpoint filetypes in the current Dir", command= lambda: ConvertMenu("PPT")).pack()
    MergeTen = Button(master,text="Merge PDF files",command=HowManyMerge).pack()
    Split = Button(master,text="Split a PDF file",command=PDFSplit).pack()
    QuitDisShiznit = Button(master, text="Quit", command=EndSession).pack()
    master.protocol("WM_DELETE_WINDOW", EndSession)
    window.protocol("WM_DELETE_WINDOW", EndSession)
    mainloop()

#Code modified from http://stackoverflow.com/questions/3579568/choosing-a-file-in-python-with-simple-dialog
def SelectDir():
    global foldername
    foldername = str(askdirectory()) # show an "Open" dialog box and return the path to the selected folder
    ListAllFiles(foldername)
    currentdir.set(foldername)

#Code modified from http://stackoverflow.com/questions/3207219/how-to-list-all-files-of-a-directory-in-python
def ListAllFiles(direct):
    global AllFiles
    AllFiles = []
    for (dirpath, dirnames, filenames) in walk(direct):
        AllFiles.extend(filenames)
        break

def ConvertMenu(ProgramType):
    global ConvertNow, PickAnOutput, PickDirWindow, WindowLable, ProgramChoice
    if ProgramType == "Word":
        ProgramChoice = WordConvertMenu
    elif ProgramType == "PPT":
        ProgramChoice = PPTConvertMenu
    RowNum=3
    ColumnNum=1
    WindowClear()
    master.withdraw()
    window.deiconify()
    WindowLable = Label(window, textvariable = currentdir).grid(row=1, column=2)
    for filetype in ProgramChoice:
        ProgramChoice[filetype] = Variable()
        filetype21= Checkbutton(window, text=filetype, onvalue=filetype, offvalue=".notavar", variable=ProgramChoice[filetype])
        filetype21.grid(row=RowNum,column=ColumnNum)
        if ColumnNum == 1:
            ColumnNum=2
        elif ColumnNum ==2:
            ColumnNum = 1
            RowNum+=1
        ProgramChoice[filetype].set(".notavar")
    PickDirWindow = Button(window, text="Select a Dir", command=SelectDir).grid(row=2, column=1)
    window.title("Pick Your FileTypes")
    if ProgramType == "Word":
        PickAnOutput = OptionMenu(window, variable123,".doc",".dot",".rtf",".txt",".mht",".htm",".xml",".docm",".dotx",".dotm",".docx",".pdf",".xps",".odt").grid(row=2, column=2)
    elif ProgramType == "PPT":
            PickAnOutput = OptionMenu(window, variable123,".ppt",".pot",".rtf",".pps","Meta Folder","Each Slide as a .gif","Each Slide as a .jpg","Each Slide as a .png","Each Slide as a .bmp","Each Slide as a .tif","Each Slide as a .emf",
            ".pptx",".pptm",".ppsx",".ppsm",".thmx",".pdf",".xps",".xml",".odp",".wmv (This Rarley Works)",".mp4 (This Rarley Works)").grid(row=2, column=2)
    ConvertNow = Button(window, text="Convert Now", command= lambda: ConvertButton(ProgramChoice)).grid(row=99, column=2)
    cancle.grid(row=99,column=1)

def ConvertButton(ProgramChoiceConvert):
    if ProgramChoiceConvert == WordConvertMenu:
        ProgramToNum = wordtonum
    elif ProgramChoiceConvert == PPTConvertMenu:
        ProgramToNum = PowerpointToNum
    for filetypes in ProgramChoiceConvert:
        ConvertAllInDir(ProgramToNum[variable123.get()],ProgramChoiceConvert[filetypes].get())

def ConvertAllInDir(newtype, typetocheckfor):
    for file in AllFiles:
        if file.endswith(typetocheckfor):
            ConvertFile(FileFixer(foldername,file,typetocheckfor), typetocheckfor, newtype)

#Code modified from http://stackoverflow.com/questions/6011115/doc-to-pdf-using-python
def ConvertFile(in_file, filetype, newfile):
    if in_file.endswith(filetype):
        out_file = Folderext.replace(filetype, "")
    if ProgramChoice == WordConvertMenu:
        OpenProgram = win32com.client.Dispatch('Word.Application')
        doc = OpenProgram.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=newfile)
    elif ProgramChoice == PPTConvertMenu:
        OpenProgram = win32com.client.Dispatch('Powerpoint.Application')
        doc = OpenProgram.Presentations.Open(in_file, WithWindow=False)
        doc.SaveAs(out_file, newfile)
    doc.Close()
    OpenProgram.Quit()

def FileFixer(foldername,file,Typecheck):
    global Folderext, NewPath
    Folderext = "From_"+Typecheck+"_To_"+variable123.get()
    Folderext = Folderext.replace(".","")
    NewPath = "From_"+Typecheck+"_To_"+variable123.get()
    NewPath = Folderext.replace(".","")
    NewPath = foldername+"\\"+Folderext+"\\"
    if not os.path.exists(NewPath):#Folder creation code troubleshooted with help by http://stackoverflow.com/questions/1274405/how-to-create-new-folder
        os.makedirs(NewPath)
    Folderext = foldername+"\\"+Folderext+"\\"+file
    in_file = foldername+"\\"+file
    in_file = in_file.replace("/","\\") #Fixes an error where all of the slashes will be facing the wrong direction
    in_file = in_file.replace("\\\\\\\\","\\") #Fixes the 4 slashes at the start of a file
    Folderext = Folderext.replace("/","\\")
    Folderext = Folderext.replace("\\\\\\\\","\\")
    return in_file

def FileSelector(WhichOne, other=0):
    if other == 0:
        File = str(askopenfilename())
        CurrentFileNum[WhichOne].set(File)
        return(File)
    if other == 1:
        ButtonNum[WhichOne] = "File"+str(WhichOne)
        CurrentFileNum[WhichOne] = "currentfile"+str(WhichOne)
        CurrentFileNum[WhichOne] = StringVar()
        CurrentFileNum[WhichOne].set("Pick File #"+str(WhichOne))
        ButtonNum[WhichOne] = Button(window, textvariable = CurrentFileNum[WhichOne], command= lambda WhichOne=WhichOne: Files.update({str(WhichOne):FileSelector(WhichOne, other=0)})).pack()

def PDFMergeTen(HowManyYo):
    global Files, ButtonNum, CurrentFileNum, FileSelectorNum
    Files = {}
    ButtonNum = {}
    CurrentFileNum = {}
    FileSelectorNum = {}
    WindowClear()
    master.withdraw()
    window.deiconify()
    window.title("Pick Your Files")
    HowMany = HowManyYo
    for Totals in range(1,HowMany):
        FileSelector(Totals, other=1)
    CloseWindow = Button(window,text="Merge", width=10, command= lambda: PDFFileMergeTen(Files)).pack(side=LEFT)
    cancle.pack(side=LEFT)

def HowManyMerge():
    WindowClear()
    master.withdraw()
    window.deiconify()
    window.title("How Many PDFs?")
    InputForThis = StringVar()
    EntryLable = Label(window,width=25,text="Please Enter The # of PDFs").pack()
    EntryForThis = Entry(window,textvariable=InputForThis).pack()
    MergeThem = Button(window,text="Next",command= lambda: PDFMergeTen(int(InputForThis.get())+1)).pack(side=LEFT)
    cancle.pack(side=LEFT)

def PDFFileMergeTen(FileDict,exiter="nope"):
    Count = 0
    if exiter == "exit":
        pass
    else:
        Merge = PdfFileMerger(strict=False)
        for PDFile in FileDict:
            Count += 1
            Merge.append(FileDict[str(Count)])
        Merge.write("Combined.pdf")
    Merge.close()
    WindowClear()
    for PDFiles in FileDict:
        os.remove(FileDict[PDFiles])

def PDFSplit():
    FileToSplit = []
    global File1, File2
    WindowClear()
    SplitText = StringVar()
    SplitText.set('Pick The File To Split')
    master.withdraw()
    window.deiconify()
    window.title("Pick Your File")
    File1 = Button(window, textvariable = SplitText, command= lambda: FileToSplit.append(FileSelector(1))).pack()
    CloseWindow = Button(window,text="Split", width=10, command= lambda FileToSplit=FileToSplit: PDFFileSplit(FileToSplit[0])).pack(side=LEFT)
    cancle.pack(side=LEFT)

def PDFFileSplit(File,exiter="nope"):
    if exiter == "exit":
        pass
    else:
        FileName = RemovePDF(File)
        FileToSplit = PdfFileReader(open(File, "rb"))
        SplitPoint = int(input("please enter a page to end the first half on ("+str(FileToSplit.getNumPages())+" pages total.)"))
        FileSpliter = PdfFileWriter()
        for i in range(0, SplitPoint):
            FileSpliter.addPage(FileToSplit.getPage(i))
        FileToSplitOpener = open(FileName+"_Page_1_to_"+str(SplitPoint)+".pdf", "wb")
        FileSpliter.write(FileToSplitOpener)
        FileToSplitOpener.close()

        FileToSplitP2 = PdfFileReader(open(File, "rb"))
        FileSpliterP2 = PdfFileWriter()
        for i in range(SplitPoint, FileToSplitP2.getNumPages()):
            FileSpliterP2.addPage(FileToSplitP2.getPage(i))
        FileToSplitOpenerP2 = open(FileName+"_Page_"+str(SplitPoint+1)+"_to_"+str(FileToSplitP2.getNumPages())+".pdf", "wb")
        FileSpliterP2.write(FileToSplitOpenerP2)
        FileToSplitOpenerP2.close()
    WindowClear()

def RemovePDF(file, justname=0):
        file2 = file
        if justname == 1:
            file2 = re.sub(file[0]+'[^>]+/', '', file) #Idea to use the re.sub came from http://stackoverflow.com/questions/8784396/python-delete-the-words-between-two-delimeters
        NewFileNoPDF = file2.replace("/","\\")
        NewFileNoPDF = NewFileNoPDF.replace(".pdf","")
        return NewFileNoPDF


def WindowClear():
    for widget in window.winfo_children():
        widget.forget()
        widget.grid_forget()
    variable123.set("Please Pick an Output FileType")
    window.withdraw()
    master.deiconify()

def EndSession():
    master.destroy()
    quit

count = 1
GetInput()