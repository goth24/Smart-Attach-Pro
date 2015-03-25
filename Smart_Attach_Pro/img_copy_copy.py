__author__ = 'ZA028309'

import pyscreenshot as ImageGrab
import Tkinter, tkFileDialog, Tkconstants
from Tkinter import *
import Tkinter as tk
import tkMessageBox
from tkFileDialog import askopenfilename
from PIL import ImageTk, Image
from docx import *
import time ,os,sys,datetime, string
import fnmatch,shutil
import time ,os,sys
import subprocess
from readExcelFile import get_data_from_xlsFile


class FirstFrame():
    def __init__(self,root,panel):
        self.root = root
        print "zaheer"
        self.fileName(panel)
        self.stepNumber(panel)
        self.undo(panel)
        self.bustMode(panel)
        self.singleMode(panel)
        self.get(root)


    def fileName(self,panel):
        global flText
        flText = StringVar()
        flText.set("")
        file = Label(panel,text="File Name :",fg="dark green")
        file.grid(row=0, column=0, padx=5,pady=5,sticky = W)
        fileName = Entry(panel)
        fileName.config(textvariable=flText)
        fileName.grid(row=0, column=1,pady=5,sticky = W)
        B1=Button(labelframe,text = "Load",command = self.loadPlanName)
        B1.grid(row=0, column=2, padx=5,pady=5,sticky = W)
        #self.stepNumber(panel)

    def stepNumber(self,panel):
        global myText,excelArray
        global number,fromExcelData
        myText = StringVar()
        print number
        #index = number
        #print "Index: ",index
        print excelArray
        #stepNumberValue = array[index]
        #myText.set("Step Number : %d" %(excelArray[number]))
        myText.set("Step Number : --")
        step = Label(panel,fg="dark green")
        step.config(textvariable=myText)
        step.grid(row=1, column=0, padx=5, pady=5,sticky = W)

    def undo(self,panel):
        self.B2= Button(panel,text = 'Undo', width=13, command= self.undoStep)
        #self.B2.pack(padx = 8,pady = 10,side = LEFT)
        self.B2.grid(row =2, column=0, padx=5, pady=5)

    def bustMode(self,panel):
        self.B3= Button(panel,text = 'Bust Mode', width=13, command= self.multiCapture)
        #self.B2.pack(padx = 8,pady = 10,side = LEFT)
        self.B3.grid(row =2, column=1, padx=5, pady=5)

    def singleMode(self,panel):
        self.B4= Button(panel,text = 'Single Capture', width=13, command= self.singleCapture)
        self.B4.grid(row =2, column=2, padx=5, pady=5)

    def get(self,root):
        #frame2 = Frame(root)
        self.B5= Button(text = 'Get', width=15, command= self.singleCapture)
        self.B5.pack(fill = "both",padx = 10,pady = 10)


    def loadPlanName(self):
        global flText,fileSeleted
        global excelArray
        name = askopenfilename()
        selecedFileName = name.split("/")[-1]
        print selecedFileName
        flText.set(selecedFileName)
        if (selecedFileName!= ""):
            excelArray = get_data_from_xlsFile(name)
            print "From program : ",excelArray
            fileSeleted = True



    def singleCapture(self):
        global number,myText
        global working_dir
        global bustNumber,bustStatus
        global flText,fileSeleted,excelArray

        print "Single click NUmber : ",number
        if(fileSeleted == True):
            root.withdraw()
            time.sleep(1)
            if (bustStatus == True):
                ImageGrab.grab().save(working_dir + "//screen Shot/Step %d-%d.png" %(excelArray[number],bustNumber), "PNG")
                bustStatus = False
            else:
                ImageGrab.grab().save(working_dir + "//screen Shot/Step %d.png" %excelArray[number], "PNG")

            time.sleep(1)
            root.deiconify()
            number +=1
            myText.set("Step Number : %d" %excelArray[number])
            bustNumber = 1
        else:
            tkMessageBox.showinfo("Warning", "Select a Test Plan File")

    def multiCapture(self):
        global number,myText,excelArray
        global bustNumber,bustStatus
        global working_dir
        if(fileSeleted == True):
            root.withdraw()
            time.sleep(1)
            for i in range(1):
                ImageGrab.grab().save(working_dir + "//screen Shot/Step %d-%d.png" %(excelArray[number],bustNumber), "PNG")
                time.sleep(1)
            root.deiconify()
            bustNumber +=1
            myText.set("Step Number : %d.%d" %(excelArray[number],bustNumber))
            bustStatus = True
        else:
            tkMessageBox.showinfo("Warning", "Select a Test Plan File")

    def undoStep(self):
        global bustStatus,number,bustNumber,fileSeleted,myText
        global working_dir,excelArray
        if(fileSeleted == True):
            if (bustStatus == True):
                bustNumber -=1
                if (bustNumber == 0):
                    number -=1
                    print
                    os.remove(working_dir + "//screen Shot/Step %d.png" %(excelArray[number]))
                    myText.set("Step Number : %d" %(excelArray[number]))
                else:

                    print "Bust nmber : ",bustNumber
                    os.remove(working_dir + "//screen Shot/Step %d-%d.png" %(excelArray[number],bustNumber))
                    myText.set("Step Number : %d.%d" %(excelArray[number],bustNumber))

            else:
                number -=1
                os.remove(working_dir + "//screen Shot/Step %d.png" %(excelArray[number]))
                myText.set("Step Number : %d" %(excelArray[number]))

        else:
            tkMessageBox.showinfo("Warning", "Select a Test Plan File")


        print "Done"




number = 0
myText = None
flText = None
bustNumber = 0
fromExcelData = []
excelArray = []
bustStatus = False
fileSeleted = False
working_dir = os.getcwd()
if __name__=='__main__':
    root = tk.Tk()
    root.geometry("500x180")
    #root.resizable(width=FALSE, height=FALSE)
    root.title('Smart_Attach')
    print working_dir
    PoweredData = "Smart_Attach"
    bckground = '//img/cerner_background.png'
    logo = '//img/Cerner_logo.png'
    path = working_dir + logo
    print path
    labelframe = Frame(root)
    #labelframe.grid(column = 0,row = 0)
    labelframe.pack(anchor = "s",fill = "both", expand="yes",padx = 10)
    img = ImageTk.PhotoImage(Image.open(path))
    #panel = tk.Label(labelframe,padx = 15,pady = 15, image = img)
    FirstFrame(root,labelframe)
    #root.wm_attributes("-topmost",1)
    root.mainloop()