__author__ = 'ZA028309'


#from openpyxl import Workbook
import xlrd,xlwt
import shutil

bookpath = None
bookpathCopy = None
rowBookCounter = 1
def getRowCount(bookname,working_dir,radioSelection):
    global bookpath
    if (radioSelection == '1'):
        bookpath = working_dir+'\\Evidence\\'+bookname+'\\'+bookname+'.xls'
        shutil.copy2(bookpath,working_dir+'\\Evidence\\'+bookname+'\\Copy'+bookname+'.xls')
    else:
        getbookpath = bookname.split(".")[0]
        bookpath = working_dir+'\\Evidence\\'+getbookpath+'\\'+bookname
        shutil.copy2(bookpath,working_dir+'\\Evidence\\'+getbookpath+'\\Copy'+bookname)

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

def putData(working_dir,selecedFileName,status):
    print "Pud Data"
    planFolder = selecedFileName.split('.')[0]
    copyFile = working_dir+'\Evidence\\'+planFolder+'\\Copy'+selecedFileName
    print "copyFile",copyFile
    print working_dir,selecedFileName,status
    writeStatus(copyFile,status)

    #print bookpath
    #sh = fileName(bookpath)


def writeStatus(copyFile,status):

    workbook = xlwt.Workbook()
    #workbook.sheet_names()
    sheet = workbook.set_active_sheet
    #sheet = workbook.sheet_by_index(0)
    sheet.write(rowBookCounter,4,status)
    workbook.save()

    '''
    def get_data_from_xlsFile(self,wbName):
    screenPrint_numbers = []
    screenShot_steps = []

    wb = xlrd.open_workbook(wbName)
    wb.sheet_names()
    sh = wb.sheet_by_index(0)
    for data in range(1,sh.nrows):
        cellData = sh.cell_value(data,3)
        #print cellData
        if (cellData != ""):
            #print "get value"
            stpNumbers = sh.cell_value(data,0)
            stepNumber = stpNumbers.split(" ")[1]
            print stepNumber
            stepNumber = int(stepNumber)
            screenPrint_numbers.append(stepNumber)

    #print screenPrint_numbers

    return screenPrint_numbers
    '''

'''
bookFile = "/Users/ZA028309/Desktop/Image_copy/TestFile.xls"

d = get_data_from_xlsFile(bookFile)
print d
'''