__author__ = 'ZA028309'


#import pyscreenshot as ImageGrab
#import Tkinter, tkFileDialog, Tkconstants
from Tkinter import *
import Tkinter as tk
import ttk
#import Tkinter.ttk as ttk
import tkMessageBox
from tkFileDialog import askopenfilename
from PIL import ImageTk, Image
#from docx import *
import time ,os,sys,datetime, string
import fnmatch,shutil
import time,os,sys
import subprocess
from readExcelFileCopy import *
from imgDocCreater.imgCreatingTemp import *
import zlib
from qcTestDownload import *
from win32com.client import Dispatch
import pythoncom
#import tktable

'''
class MyDialog:
    def __init__(self, parent):
        top = self.top = tk.Toplevel(parent)
        top.geometry("300x200")
        self.myLabel = tk.Label(top, text="Client OS")
        self.myLabel.pack()

        self.myEntryBox = tk.Entry(top)
        self.myEntryBox.pack()

        self.myLabel1 = tk.Label(top, text='Enviro')
        self.myLabel1.pack()

        self.myEntryBox1 = tk.Entry(top)
        self.myEntryBox1.pack()

        self.mySubmitButton = tk.Button(top, text='Submit', command=self.send)
        self.mySubmitButton.pack()

    def send(self):
        global username
        username = self.myEntryBox.get()
        self.top.destroy()
'''
class FirstFrame():

    def __init__(self, root, childFrameone, childFrametwo,childFramethree,childFrameUtil):
        #self.root = rootLable
        self.fileName(childFrameone)
        self.utilFrame(childFrameUtil)
        self.countFrame(childFramethree)
        self.screenPrintFrame(childFramethree)

    def fileName(self, panel):
        global tsPlanName,qcSelected,tsPLanId
        tsPlanName = StringVar()
        tsPlanName.set("")
        var = IntVar()
        tsIDLabel = Label(panel, text="Test Plan ID ", fg="#474751")
        tsIDLabel.grid(row=0, column=0, padx=5, sticky=W)

        tsIdField = Entry(panel, width=15)
        tsIdField.grid(row=0, columnspan=1, padx=110, pady=5, sticky=W)

        tpNameLabel = Label(panel, text="Test Plan Name ", fg="#474751")
        tpNameLabel.grid(row=1, column=0, padx=5, sticky=W)

        tsNameField = Entry(panel,text=tsPlanName, width=55)
        tsNameField.grid(row=1, columnspan=1, padx=110, sticky=W)

        R1 = Radiobutton(panel,text="From QC", variable=var, value=1, command=lambda: self.radioValueSelect(var))
        R1.grid(row=0, columnspan=2, padx=230, sticky=W)

        R2 = Radiobutton(panel,text="From Local", variable=var, value=2, command=lambda: self.radioValueSelect(var))
        R2.grid(row=0, columnspan=2, padx=140,sticky=E)

        loadImageicon = ImageTk.PhotoImage(file=working_dir+"/icons/load_icon_2X.png")
        loadButton = Button(panel, image=loadImageicon,  text="Load",compound=TOP,  command=lambda: self.loadPlan(tsIdField), width=60, height=60)
        loadButton.image = loadImageicon
        #loadButton.grid(rowspan=1, columnspan=3, sticky=NSEW, padx=(730,20))
        loadButton.grid(row=1, column=2,rowspan=2, sticky=NE, padx=20)

        f = open(working_dir+'\\testDataListConfig.txt','r')
        listDatafile = f.readlines()
        #print listDatafile
        clientData =  listDatafile[0].split(":")[1].split(",")
        clientVariable = StringVar()
        clientLable = Label(panel,text="Client OS/Enviroment")
        clientLable.grid(row=0, column=4, padx=10, sticky=E)
        clientText = ttk.Combobox(panel, textvariable=clientVariable)
        clientText['values'] = (clientData)
        clientText.current(0)
        clientText.grid(row=0, column=5, padx=5, sticky=W)

        osVariable = StringVar()
        #osData =  listDatafile[1].split(",")
        osLable = Label(panel,text="OS")
        osLable.grid(row=0, column=6, padx=10, sticky=E)
        osText = ttk.Combobox(panel, textvariable=osVariable)
        osText['values'] = ('A', 'B', 'C')
        osText.current(0)
        osText.grid(row=0, column=7, padx=5, sticky=W)

        domainData =listDatafile[2].split(":")[1].split(",")
        domainLable = Label(panel,text="Domain")
        domainLable.grid(row=1, column=4, padx=10, sticky=E)
        domainText = ttk.Combobox(panel, textvariable=osVariable)
        domainText['values'] = domainData
        domainText.current(0)
        domainText.grid(row=1, column=5, padx=5, sticky=W)

        executionData =listDatafile[3].split(":")[1].split(",")
        executionLable = Label(panel,text="Execution Method")
        executionLable.grid(row=1, column=6, padx=5, sticky=E)
        executionText = ttk.Combobox(panel, textvariable=executionData)
        executionText['values'] = executionData
        executionText.current(0)
        executionText.grid(row=1, column=7, padx=5, sticky=W)

        solutionLable = Label(panel,text="Solution")
        solutionLable.grid(row=2, column=4, padx=5, pady=2, sticky=E)
        solutionText = Entry(panel,width=23)
        solutionText.grid(row=2, column=5, padx=5, pady=2, sticky=W)

        testDataLable = Label(panel,text="Test Data")
        testDataLable.grid(row=2, column=6, padx=5, pady=2, sticky=E)
        testDataText = Entry(panel,width=23)
        testDataText.grid(row=2, column=7, padx=5, pady=2, sticky=W)

    def radioValueSelect(self, var):
        global radioSelection
        radioSelection = str(var.get())
        print radioSelection


    def testPlans(self, panel2):
        global stpNumberValue,stpDesctiptionValue,stpExpectedValue,stpEvidenceValue,noOfScreenPrints
        table_Names = ['No', 'Steps','Status']
        treeColumn = len(table_Names)
        container = tk.Frame(panel2, height='4')
        container.pack(fill='x', expand=True)
        ##-------------------------------------------------------------------------------------------
        tree = ttk.Treeview(container,height="5", columns=('No', 'Steps','Status'),selectmode="extended")
        tree.heading('0', text='No', anchor=W)
        tree.heading('1', text='Steps', anchor=W)
        tree.heading('2', text='Status', anchor=W)
        tree.column('0', stretch=NO, minwidth=0, width=20)
        tree.column('1', stretch=NO, minwidth=0, width=300)
        tree.column('2', stretch=NO, minwidth=0, width=300)
        #tree.column('#0', stretch=NO, minwidth=0, width=0) #width 0 to not display it
        #tree.insert("","", index="end",text="1",tags=("#1"))
        tree.insert("","1","",text="1")
        #tree.ins
        tree.insert("","2","",text="Step 1")
        tree.insert("","3","",text="Pass")
        #vsb = ttk.Scrollbar(orient="vertical",command= tree.yview)
        #hsb = ttk.Scrollbar(orient="horizontal",command= tree.yview)
        #tree.configure(yscrollcommand=vsb.set)#,xscrollcommand=hsb.set)
        tree.grid(column=0, row=2, sticky='w', in_=container)
        #vsb.grid(column=treeColumn, sticky='ns', in_=container)
        #hsb.grid(column=treeColumn, row=1, sticky='ew', in_=container)
        container.grid_columnconfigure(2, weight=1)
        container.grid_rowconfigure(0, weight=1)
        container.grid(columnspan=treeColumn, sticky=W)

        #tssss = Label(panel2,width=2,height=12,text="Help")
        #tssss.grid(row=0,columnspan=4, sticky=W)


        ##---------------------------------------------------------------------------------------------
        texts = "Data"
        tpStepName = Label(panel2, width=2, text="Steps", fg="#474751")
        tpStepName.grid(row=1, column=0, padx=5, pady=3, sticky=NSEW)

        tpDescription = Label(panel2, width=42, text="Description", fg="#474751")
        tpDescription.grid(row=1, column=1, padx=5,pady=3, sticky=NSEW)

        tpExpected = Label(panel2, width=42, text="Expected Result", fg="#474751")
        tpExpected.grid(row=1, column=2, padx=5, pady=3, sticky=NSEW)

        tpEvid = Label(panel2,width=5,text="Screen", fg="#474751")
        tpEvid.grid(row=1, column=3, padx=5, pady=3, sticky=NSEW)

        #Excel row Column Data
        stepText = Text(panel2, borderwidth=2, relief="sunken", bg='gray', width=10, height=9)
        stepText.config(font=("arial", 11), undo=True, wrap='word')
        stepText.insert(END, stpNumberValue)
        stepText.grid(row=2, column=0, sticky="nsew", padx=2, pady=6)

        scrollbar1 = Scrollbar(panel2)
        scrollbar1.grid(row=2, column=1)
        text = Text(panel2, borderwidth=2, font=("arial", 11), wrap=WORD, relief="sunken", bg='gray', width=50,
                    height=11, yscrollcommand=scrollbar1.set)
        text.insert(END, stpDesctiptionValue)
        text.grid(row=2, column=1, padx=1, pady=5)
        scrollbar1.config(command=text.yview)

        scrollbar2 = Scrollbar(panel2)
        scrollbar2.grid(row=2, column=2)
        text = Text(panel2, borderwidth=2, font=("arial", 11), wrap=WORD, relief="sunken", bg='gray', width=50,
                    height=11, yscrollcommand=scrollbar2.set)
        text.insert(END, stpExpectedValue)
        text.grid(row=2, column=2, padx=1, pady=5)
        scrollbar2.config(command=text.yview)

        stepEvid = Text(panel2, borderwidth=2, font=("arial", 10), relief="sunken", bg='gray', width=18,height=9)
        stepEvid.insert(END, stpEvidenceValue)
        stepEvid.grid(row=2, column=3, sticky="nsew", padx=2, pady=6)

        if stpEvidenceValue != "":
            noOfScreenPrints+=1


        passButton = Button(panel2, width=7, text='Pass', fg="green", command=lambda: self.stepStatus("Pass"))
        passButton.config(font=("arial", 12,"bold"))
        passButton.grid(row=2, column=5, rowspan=2, sticky="n", padx=8, pady=10)

        failButton = Button(panel2, width=7, text='Fail', fg="red", command=lambda: self.stepStatus("Fail"))
        failButton.config(font=("arial", 12,"bold"))
        failButton.grid(row=2, column=5,rowspan=2, sticky="n", padx=8, pady=50)

        naButton = Button(panel2, width=7, text='N/A', fg="gray", command=lambda: self.stepStatus("N/A"))
        naButton.config(font=("arial", 12,"bold"))
        naButton.grid(row=2, column=5,rowspan=2, sticky="n", padx=8, pady=90)

        '''
        reportButton =  Button(panel2, width=7, text='Report', command=lambda: self.stepStatus("Pass"))
        reportButton.config(font=("arial", 12,"bold"))
        reportButton.grid(row=2, column=5, rowspan=2, sticky="n", padx=8, pady=150)
        '''


    def countFrame(self, childFramethree):
        global rowCount, remainingRowCount
        print rowCount
        if (rowCount ==1):
            modeRoeCount = 0
        else:
            modeRoeCount = rowCount
        remainingRowCount = StringVar()
        remainingRowCount.set("Remaining Steps : %d  " %modeRoeCount)
        stepCount = Label(childFramethree, textvariable=remainingRowCount, relief=SUNKEN, borderwidth=1, fg="#26263A", bg='gray', width=25)
        #stepCount.pack(side=LEFT, padx=5, pady=5)
        stepCount.grid(row=0, column=0, sticky="nsew", padx=10, pady=(20,20))

    def utilFrame(self,childFrameUtil):
        global crNumber,srNumber,commentstext

        crNumberLable = Label(childFrameUtil, width=7, text="CR No : ", fg="#474751")
        crNumberLable.grid(row=0, column=0,padx=5,pady=5)
        crNumberText = Entry(childFrameUtil, width=20)
        crNumberText.grid(row=0, column=1,padx=5,pady=5)
        crNumber = crNumberText

        srNumberLable = Label(childFrameUtil, width=7, text="SR No : ", fg="#474751")
        srNumberLable.grid(row=0, column=2,padx=5,pady=5)
        srNumberText =  Entry(childFrameUtil, width=20)
        srNumberText.grid(row=0, column=3,padx=5,pady=5)
        srNumber = srNumberText

        commentsLable = Label(childFrameUtil, width=10, text="Comments :", fg="#474751")
        commentsLable.grid(row=0, column=5,padx=5,pady=5)
        commentsText =  Entry(childFrameUtil, width = 120)
        commentsText.grid(row=0, column=6,padx=5,pady=5)
        commentstext = commentsText




    ########### Frame Three ##########
    def screenPrintFrame(self,childFramethree):

        iconDir = working_dir + '/icons/'

        flushImageicon = ImageTk.PhotoImage(file=iconDir+"flush_icon_2X.png")
        imgFlushButton = Button(childFramethree, text="Flush Img", image=flushImageicon,width=60, height=70,
                                compound=TOP, command=lambda: self.flushImage("Flush"))
        imgFlushButton.image = flushImageicon
        #imgFlushButton.pack(side=LEFT, padx=5, pady=5)
        imgFlushButton.grid(row=0, column=1, sticky="nsew", padx=4, pady=6)

        undoImageicon = ImageTk.PhotoImage(file=iconDir+"undo_button_2X.png")
        undoButton = Button(childFramethree, image=undoImageicon,  text="Undo", width=60, height=70,compound=TOP,
                            command=lambda: self.stepStatus("Undo"))
        undoButton.image = undoImageicon
        #undoButton.pack(side=LEFT, padx=10, pady=5)
        undoButton.grid(row=0, column=2, sticky="nsew", padx=4, pady=6)

        bustImageicon = ImageTk.PhotoImage(file=iconDir+"bust_mode_icon_2X.png")
        bustButton = Button(childFramethree, image=bustImageicon, text="Bust Mode", width=60, height=70,
                            compound=TOP, command=lambda: self.captureImage("Bust"))
        bustButton.image = bustImageicon
        #bustButton.pack(side=LEFT, padx=10, pady=5)
        bustButton.grid(row=0, column=3, sticky="nsew", padx=4, pady=6)

        singleImageicon = ImageTk.PhotoImage(file=iconDir+"single_capture_icon_2X.png")
        singleButton = Button(childFramethree, image=singleImageicon, text="Single Capture", width=60, height=70,
                              compound=TOP,command=lambda: self.captureImage("Single"))
        singleButton.image = singleImageicon
        #singleButton.pack(side=LEFT, padx=10, pady=5)
        singleButton.grid(row=0, column=4, sticky="nsew", padx=4, pady=6)

        generateImageicon = ImageTk.PhotoImage(file=iconDir+"generate_icon_2X.png")
        generateButton = Button(childFramethree, image=generateImageicon, text="Generate", width=60, height=70,
                                compound=TOP, command=self.generate)
        generateButton.image = generateImageicon
        #generateButton.pack(side=LEFT, padx=10, pady=5)
        generateButton.grid(row=0, column=5, sticky="nsew", padx=4, pady=6)

        qcRunImageicon = ImageTk.PhotoImage(file=iconDir+"qcRun_icon_2X.png")
        runQCButton = Button(childFramethree,image=qcRunImageicon, text="Run QC", width=60, height=70, compound=TOP,
                             command=self.callQC)
        runQCButton.image = qcRunImageicon
        #runQCButton.pack(side=LEFT, padx=10, pady=5)
        runQCButton.grid(row=0, column=6, sticky="nsew", padx=4, pady=6)

    def generate(self):
        global stepResult, crNumberArray, srNumberArray, commentsArray
        generateFile(selecedFileName,stepResult, crNumberArray, srNumberArray, commentsArray, noOfScreenPrints)
        time.sleep(1)
        tkMessageBox.showinfo(title="Process Completed", message="Plan Evidence Created")

    def callQC(self):
        global working_dir, selecedFileName, tsPLanId
        qcRunPlan(working_dir, selecedFileName, tsPLanId)
        print "Completed Run"
        tkMessageBox.showinfo(title="QC Process Completed", message="QC Run Completed & Evidence Attached..")



    def flushImage(self,imgDo):
        global working_dir,stpNumberValue
        mediafolder = "%s/screen Shot" %working_dir
        for r, d, f in os.walk(mediafolder):
            if(imgDo == "Flush"):
                for name in f:
                    os.remove(os.path.join(mediafolder, name))
                    print "Image Folder Cleared..!!"
            elif(imgDo == "StepClear"):
                files = [i for i in os.listdir(mediafolder) if os.path.isfile(os.path.join(mediafolder,i)) and stpNumberValue in i]
                print files
                for j in files:
                    os.remove(os.path.join(mediafolder, j))

    def captureImage(self, mode):
        global working_dir, stpNumberValue, bustFlag, selecedFileName
        print selecedFileName
        selectedDirName = selecedFileName.split(".")[0]
        screenPrintFolder = working_dir +'\\Evidence\\'+ selectedDirName+'\\screen Shot'
        if not os.path.exists(screenPrintFolder): os.makedirs(screenPrintFolder)
        root.withdraw()
        time.sleep(1)
        if (mode == "Single"):
            ImageGrab.grab().save(screenPrintFolder+'\\%s.png' %(stpNumberValue), "PNG")
            bustFlag = 0

        elif(mode == "Bust"):
            bustFlag += 1
            ImageGrab.grab().save(screenPrintFolder+'\\%s-%d.png' %(stpNumberValue, bustFlag), "PNG")
        time.sleep(1)
        root.deiconify()


    def loadPlan(self,tsIdField):
        global tsPlanName,rowCount,remainingRowCount,selecedFileName,radioSelection,tsPLanId
        global stpNumberValue,stpDesctiptionValue,stpExpectedValue,stpEvidenceValue

        print "radioSelection", radioSelection
        tsPLanId = tsIdField.get()
        print "Test ID", tsPLanId

        if(radioSelection!= ""):
            #try:
            if(radioSelection == '1'):
                print "QC call"
                selecedFileName = call_QC_Load(tsPLanId,working_dir)
                print selecedFileName

            elif(radioSelection == '2'):
                name = askopenfilename()
                print name
                selecedFileName = name.split("/")[-1]
                print selecedFileName

            tsPlanName.set(selecedFileName)
            #print tsPlanName
            rowCount = getRowCount(selecedFileName,working_dir,radioSelection)
            remainingRowCount.set(" Remaining Steps : %d  " %rowCount)

            #get the Date from the Excel Sheet (i.e step Number, Description,.....)

            stpNumberValue,stpDesctiptionValue,stpExpectedValue,stpEvidenceValue = getData(rowCount,undoFlageChange)
            print stpNumberValue
            self.testPlans(childFrametwo)

            #except:
                #print "File Not Selected"
        else:
            tkMessageBox.showinfo("Warning", "Plan from QC or Local")

    def stepStatus(self, status):

        global rowCount, remainingRowCount, nextCount, stepResult, crNumberArray, srNumberArray, commentsArray
        global tsPlanName, rowCount, remainingRowCount, bustFlag, undoFlageChange
        global stpNumberValue, stpDesctiptionValue, stpExpectedValue, stpEvidenceValue
        ######
        global crNumber,srNumber,commentstext
        global crNumber1,srNumber1,commentstext1
        ######
        crNumber1 = crNumber.get()
        srNumber1 = srNumber.get()
        commentstext1 = commentstext.get()

        utilData = [crNumber1,srNumber1,commentstext1]
        print utilData
        if(status=="Undo"):
            undoFlageChange = True
            rowCount += 2
            print "rowCounter UNdo function",rowCount
            try:
                stepResult.pop()
                crNumberArray.pop()
                srNumberArray.pop()
                commentsArray.pop()
            except:
                pass
            print stepResult, rowCount
            FirstFrame.flushImage(self, "StepClear")

        else:
            undoFlageChange = False
            print "status :",status
            stepResult.append(status)
            print "stepResult :",stepResult
            commentsArray.append(commentstext1)
            srNumberArray.append(srNumber1)
            crNumberArray.append(crNumber1)
            putData(working_dir,selecedFileName,status,utilData)
        stpNumberValue, stpDesctiptionValue, stpExpectedValue, stpEvidenceValue = getData(status, undoFlageChange)
        crNumber.delete(0,END)
        srNumber.delete(0,END)
        commentstext.delete(0,END)

        print stpNumberValue
        print stpDesctiptionValue
        print stpExpectedValue
        print stpEvidenceValue

        if(stpNumberValue != None):
            rowCount -= 1
            remainingRowCount.set(" Remaining Steps : %d  " %rowCount)
            self.testPlans(childFrametwo)
            nextCount += 1
        else:
            tkMessageBox.showinfo("Warning", "End of Steps")
        bustFlag = 0




tsPLanId = None
tsPlanName = None
selecedFileName = ''
rowCount = 1
nextCount = 1
remainingRowCount = None
stpNumberValue,stpDesctiptionValue,stpExpectedValue,stpEvidenceValue = "","","",""
crNumber,srNumber,commentstext = None,None,None
crNumber1,srNumber1,commentstext1 = "","",""
stepNumber, stepDescription, stepExpected, stepEvidence = None,None,None,None
working_dir = os.getcwd()
stepResult = []
crNumberArray =[]
srNumberArray=[]
commentsArray = []
noOfScreenPrints = 0
bustFlag = 0
undoFlageChange = False
radioSelection = ''

if __name__ == '__main__':
    root = tk.Tk()
    root.geometry("1080x790")
    root.title('Smart_Attach')
    frameLabel1 = LabelFrame(root, text="Powered by Cerner india", fg="dark gray")
    frameLabel1.pack(fill="both", expand="yes", padx=7, pady=7)

    childFrameone = LabelFrame(frameLabel1, text="Load from QC")
    childFrameone.pack(anchor="s", fill="both", padx=10, pady=5)
    print "zaheer"
    childFrametwo = LabelFrame(frameLabel1, text="Test Plan details")
    childFrametwo.place(y=400)
    childFrametwo.pack(anchor="s", fill="both", padx=10, pady=5)

    childFrameUtil = LabelFrame(frameLabel1, text="Uitility")
    childFrameUtil.place(y=400)
    childFrameUtil.pack(anchor="s", fill="both", padx=10, pady=6)

    childFramethree = LabelFrame(frameLabel1, text="Actions")
    childFramethree.place(y=400)
    childFramethree.pack(anchor="s", fill="both", padx=10, pady=(5,15))


    FirstFrame(frameLabel1, childFrameone, childFrametwo,childFramethree,childFrameUtil)

    root.mainloop()

