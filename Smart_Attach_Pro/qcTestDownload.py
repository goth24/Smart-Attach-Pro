__author__ = 'za028309'


import win32com, win32com.client
import xlrd,xlwt,os
#import MySQLdb
import sqlite3
import ctypes,sys
import HTMLParser
import pythoncom
import base64
class MLStripper(HTMLParser.HTMLParser):
    def __init__(self):
        self.reset()
        self.fed = []


    def handle_data(self, d):
        self.fed.append(d)


    def get_fed_data(self):
        return ''.join(self.fed)


def sanitize(data):
    s = MLStripper()
    s.feed(data)
    return s.get_fed_data()


def recursiveExport(f, qc, node):

    print node.Name
    rIndex = 1
    designStepFactory = node.DesignStepFactory

    for ds in designStepFactory.NewList(''):
        StepDescription = sanitize(ds.StepDescription)
        StepName = sanitize(ds.StepName)
        StepExpectedResult = sanitize(ds.StepExpectedResult)
        Step_Evidence =  ds.Field("DS_USER_01")
        if ds.Field("DS_USER_01") is not None:
            print ds.Field("DS_USER_01")

        f.write(rIndex,0,StepName)
        f.write(rIndex,1,StepDescription)
        f.write(rIndex,2,StepExpectedResult)
        f.write(rIndex,3,Step_Evidence)
        rIndex+=1

def exportTests(qc, nodePath,working_dir):
    '''
    try:
        os.remove(fileName)
    except:
        print "No file Found"
    '''
    wb = xlwt.Workbook()
    sh = wb.add_sheet("Sheet 1")
    planName = nodePath[0].Name
    print planName
    sh.write(0,0,"Steps")
    sh.write(0,1,"Description")
    sh.write(0,2,"Expected Result")
    sh.write(0,3,"Evidence Req")
    sh.write(0,4,"Status")
    sh.write(0,5,"CR Number")
    sh.write(0,6,"SR Number")
    sh.write(0,7,"Comments")
    evid_folder = working_dir+"\Evidence\\"+planName
    print "evid_folder",evid_folder
    try:
        os.stat(evid_folder)
    except:
        os.mkdir(evid_folder)
    rIndex = 1
    for tsItem in nodePath:
        #planName = tsItem.Name
        #print "Name :",planName
        recursiveExport(sh, qc, tsItem)

    fileName = (evid_folder+'\\%s.xls'%planName)

    try:
        wb.save(fileName)
    except:
        os.remove(fileName)
        wb.save(fileName)
    return planName

def getCredentials(working_dir):
    print "Data"
    dataList = []
    f1 = open(working_dir+"\\Credentials.txt")
    qcListData = f1.readlines()
    for item in qcListData:
        dataList.append(item.split(':')[1])

    return dataList


server= r"http://qualitycenter.cerner.com/qcbin"
#username= "VS021174"
#password= "Thanthu>86"
#domainname= "IP"
#projectname= "TD_VALIDATION_TESTS"

noServer = 'Server is not available'
pwdError = 'Failed to Login'

#if __name__ == "__main__":
def call_QC_Load(tsPLanId,working_dir):

    dataList = getCredentials(working_dir)

    username = dataList[0]
    password1 = dataList[1]
    password = base64.decodestring(password1)
    domainname = dataList[2]
    projectname = dataList[3]

    print 'Logging in...'

    qc = win32com.client.Dispatch("TDApiOle80.TDConnection")#, clsctx=pythoncom.CLSCTX_LOCAL_SERVER)
    qc.InitConnection(server)
    qc.Login(username,password)
    qc.Connect(domainname, projectname)
    print "getting into QC file ",tsPLanId

    #testID = '250618'
    testID = tsPLanId


    tSetFact = qc.TestFactory
    testSetFilter = tSetFact.Filter
    d =testSetFilter.SetFilter(u"TS_TEST_ID",testID)
    h = testSetFilter.NewList()

    planName = exportTests(qc,h,working_dir)
    print "Done QC"
    return planName


def Auth(*parms):
    print "Auth"
    print parms
    try:
        qc = win32com.client.Dispatch("TDApiOle80.TDConnection")#, clsctx=pythoncom.CLSCTX_LOCAL_SERVER)
        qc.InitConnection(server)
        qc.Login(parms[0],parms[1])
        #qc.Connect(parms[2], parms[3])
        response = "Login Completed"
    except:
        sysError = sys.exc_info()[1][2][2]
        print sysError
        if sysError == noServer:
            response = "Check your Net Connection"

        elif sysError == pwdError:
            response = "check your"
        else:
            response = "Error"

    return response

