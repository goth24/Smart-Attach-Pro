__author__ = 'ZA028309'


#from openpyxl import Workbook
import xlrd,xlwt
import shutil
from win32com.client import Dispatch
from qcRun import *

bookpath = None
bookpathCopy = None
rowBookCounter = 1
statusList = []
def getRowCount(bookname,working_dir,radioSelection):
    global bookpath
    if (radioSelection == '1'):
        bookpath = working_dir+'\\Evidence\\'+bookname+'\\'+bookname+'.xls'
        #shutil.copy2(bookpath,working_dir+'\\Evidence\\'+bookname+'\\Copy'+bookname+'.xls')
    else:
        getbookpath = bookname.split(".")[0]
        bookpath = working_dir+'\\Evidence\\'+getbookpath+'\\'+bookname
        #shutil.copy2(bookpath,working_dir+'\\Evidence\\'+getbookpath+'\\Copy'+bookname)

        print "bookname", bookpath
    sh = fileName(bookpath)
    numberOfRows = sh.nrows
    numberOfRows-=2
    #print numberOfRows
    return numberOfRows

def fileName(wbName):

    wb = xlrd.open_workbook(wbName)
    wb.sheet_names()
    sh = wb.sheet_by_index(0)
    return sh

def getData(status,undoFlageChange):
    global bookpath,rowBookCounter
    sh = fileName(bookpath)
    fixRowValue = sh.nrows
    if(undoFlageChange == True):
        print "Before UNdo Count",rowBookCounter
        rowBookCounter -= 2
        print "Undo count :",rowBookCounter

    print "Row Value :",(fixRowValue-rowBookCounter)
    if((fixRowValue-rowBookCounter) !=0):
        stpNumber = sh.cell_value(rowBookCounter,0)
        stpDesctiption = sh.cell_value(rowBookCounter,1)
        stpExpected = sh.cell_value(rowBookCounter,2)
        stpEvidence = sh.cell_value(rowBookCounter,3)
        print stpDesctiption
        rowBookCounter +=1
    else:
        stpNumber,stpDesctiption,stpExpected,stpEvidence = (None,None,None,None)
    return stpNumber,stpDesctiption,stpExpected,stpEvidence

def putData(working_dir,selecedFileName,status,utilData):
    print "Pud Data"
    planFolder = selecedFileName.split('.')[0]
    #copyFile = working_dir+'\Evidence\\'+planFolder+'\\Copy'+selecedFileName
    copyFile = working_dir+'\Evidence\\'+planFolder+'\\'+selecedFileName
    print "copyFile",copyFile
    print working_dir,selecedFileName,status
    excel = Dispatch('Excel.Application')
    excel.Visible = False
    workbook = excel.Workbooks.Open(copyFile)
    workBook = excel.ActiveWorkbook
    sheets = workBook.Sheets('Sheet 1')
    #writeStatus(copyFile,status)
    print "rowBookCounter",rowBookCounter
    sheets.Cells(rowBookCounter, 5).Value = status
    print utilData[0],utilData[1],utilData[2]
    sheets.Cells(rowBookCounter,6).Value = utilData[0]
    sheets.Cells(rowBookCounter,7).Value = utilData[1]
    sheets.Cells(rowBookCounter,8).Value = utilData[2]
    workBook.Save()
    workbook.Close()

    #print bookpath
    #sh = fileName(bookpath)


def qcRunPlan(working_dir,selecedFileName, tsPLanId):
    global statusList
    print selecedFileName
    planFolder = selecedFileName.split('.')[0]
    #copyFile = working_dir+'\Evidence\\'+planFolder+'\\Copy'+selecedFileName
    copyFile = working_dir+'\Evidence\\'+planFolder+'\\'+planFolder+'.xls'
    print "copyFile",copyFile
    wb = xlrd.open_workbook(copyFile)
    wb.sheet_names()
    sh = wb.sheet_by_index(0)
    passRow = sh.nrows
    print "passRow",passRow
    for i in range(1,passRow):
        sts =  sh.cell_value(i,4)
        print sts
        statusList.append(sts)
    print "StatusList:",statusList
    qcRunFinal(statusList,copyFile,selecedFileName,tsPLanId)


#bookFile = "/Users/ZA028309/Desktop/Image_copy/TestFile.xls"

#d = get_data_from_xlsFile(bookFile)
#print d
