__author__ = 'za028309'

import win32com, win32com.client
import xlrd,xlwt,os

import HTMLParser
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


def recursiveExport(f, qc, node,rIndex):

    print node.Name
    rIndex = 0

    designStepFactory = node.DesignStepFactory
    print node.ID, node.Name
    for ds in designStepFactory.NewList(''):
        StepDescription = sanitize(ds.StepDescription)
        StepName = sanitize(ds.StepName)
        StepExpectedResult = sanitize(ds.StepExpectedResult)
        #StepEvidence = sanitize(ds.StepEvidence)
        StepEvid = ds.StepUser01
        print StepEvid
        f.write(rIndex,0,StepName)
        f.write(rIndex,1,StepDescription)
        f.write(rIndex,2,StepExpectedResult)
        #f.write(rIndex,3,StepEvidence)
        rIndex+=1

    #f.flush()



    # current node has more children


def exportTests(qc, nodePath):
    fileName = 'c://TextingFile.xls'
    try:
        os.remove(fileName)
    except:
        print "No file Found"
    wb = xlwt.Workbook()
    rIndex = 0
    sh = wb.add_sheet("Sheet 1")
    #mg = qc.TestSetTreeManager
    #node = mg.NodeByPath(nodePath)
    #print "Node : ",node
    #tsList = node.FindTestSets("test")
    for tsItem in nodePath:
        planName = tsItem.Name
        print "Name :",planName
        recursiveExport(sh, qc, tsItem,rIndex)
    wb.save(fileName)

server= r"http://qualitycenter.cerner.com/qcbin"
username= "VS021174"
password= "Thanthu>86"
domainname= "IP"
projectname= "TD_VALIDATION_TESTS"

if __name__ == "__main__":
    print 'Logging in...'
    qc = win32com.client.Dispatch("TDApiOle80.TDConnection")
    qc.InitConnection(server)
    qc.Login(username,password)
    qc.Connect(domainname, projectname)

    """
    Change nodePath to another "folder" in the Test Plan section of QC and run script
    """
    '''
    testID = '241526'
    tSetFact = qc.TestFactory
    testSetFilter = tSetFact.Filter
    d =testSetFilter.SetFilter(u"TS_TEST_ID",testID)
    h = testSetFilter.NewList()
    print h[0].Type
    for tsItem in h:
       print "Name :",tsItem.Name
    '''
    testID = '250618'
    tSetFact = qc.TestFactory
    testSetFilter = tSetFact.Filter
    d =testSetFilter.SetFilter(u"TS_TEST_ID",testID)
    h = testSetFilter.NewList()

    exportTests(qc,h)