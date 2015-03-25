__author__ = 'za028309'

import win32com, win32com.client
import xlrd,xlwt,os, time, sys
#import MySQLdb
import sqlite3, datetime,zipfile
from win32com.client import Dispatch
import shutil


def zip_folder(folder_path, output_path):
    """Zip the contents of an entire folder (with that folder included
    in the archive). Empty subfolders will be included in the archive
    as well.
    """
    parent_folder = os.path.dirname(folder_path)
    # Retrieve the paths of the folder contents.
    contents = os.walk(folder_path)
    try:
        zip_file = zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED)
        for root, folders, files in contents:
            # Include all subfolders, including empty ones.
            for folder_name in folders:
                absolute_path = os.path.join(root, folder_name)
                relative_path = absolute_path.replace(parent_folder + '\\',
                                                      '')
                print "Adding '%s' to archive." % absolute_path
                zip_file.write(absolute_path, relative_path)
            for file_name in files:
                absolute_path = os.path.join(root, file_name)
                relative_path = absolute_path.replace(parent_folder + '\\',
                                                      '')
                print "Adding '%s' to archive." % absolute_path
                zip_file.write(absolute_path, relative_path)
        print "'%s' created successfully." % output_path
    except IOError, message:
        print message
        sys.exit(1)
    except OSError, message:
        print message
        sys.exit(1)
    except zipfile.BadZipfile, message:
        print message
        sys.exit(1)
    finally:
        zip_file.close()


server= r"http://qualitycenter.cerner.com/qcbin"
username= "mr029157"
password= "Password1#"
domainname= "IP"
projectname= "TD_VALIDATION_TESTS"

#if __name__ == "__main__":
def qcRunFinal(ts_Status,copyFile,selecedFileName,tsPLanId):
    print tsPLanId
    #qcRunID = 2345
    #testPlanId =  tsPLanId
    print 'Logging in...'
    qc = win32com.client.Dispatch("TDApiOle80.TDConnection")
    qc.InitConnection(server)
    qc.Login(username,password)
    qc.Connect(domainname, projectname)

    #ts_Status = ['Passed','Passed','Passed','Passed','Passed','Passed',]
    working_dir = os.getcwd()
    #print working_dir
    array = []
    #### Rear from Text File
    with open(working_dir+'\\QcRunData.txt','r') as ins:
        for line in ins:
            try:
                array.append(line.split(":")[1])
                print array
            except:
                pass
    client = array[0]
    domain = array[1]
    tcExecution = array[2]
    solution = array[3]
    sysOS = array[4]
    ##############

    tsINst_Status =""
    #testID = '250618'
    testPlanId = tsPLanId
    testSetID = 'MGC-Test-I'
    #testPlanId = '266981'
    check_Status = "Fail"
    if check_Status in ts_Status:
        tsINst_Status  = check_Status
    else:
        tsINst_Status = "Passed"


    date_Time = datetime.datetime.now()
    formatDate = date_Time.strftime("Run_%m-%d_%I-%M-%S")
    #tsTreeMgr = qc.TestSetTreeManager
    #tsFolder = qc.TestSetTreeManager.NodeByPath("Root\Project Go\ORT-Orthopedics\Archive")
    tsTreeMgr = qc.TestSetTreeManager
    print "tsTreeMrg",tsTreeMgr
    tsFolder = tsTreeMgr.NodeByPath("Root\PowerChart Message Center\Archive")#"Root\Project Go\ORT-Orthopedics\Archive")
    print "tsFolder",tsFolder
    tsList = tsFolder.FindTestSets("MGC-Test-I")
    print "tsList Number",len(tsList)
    settingFileds = ["RN_USER_01","RN_USER_02","RN_USER_06"]
    for tsItem in tsList:
        print "Test Set Name:",tsItem.Name
        tsTestList = tsItem.TSTestFactory.NewList("")
        print "xxxxxxx ",len(tsTestList)
        # loop through all test cases in this list
        for tsTest in tsTestList:
            print("Test case name: " + tsTest.TestName)
            testcaseId = tsTest.TestID

            print ("Test Case:" + testcaseId)
            if testcaseId == testPlanId:
                #tsTest.Status = tsINst_Status
                #tsTestList.post()

                print "Got the plan"
                newItem = tsTest.RunFactory.AddItem(None)
                print newItem.Status
                print newItem.Name


                if check_Status in ts_Status:
                    newItem.Status = 'Failed'
                    tsINst_Status = 'Failed'
                else:
                   newItem.Status = 'Passed'
                   tsINst_Status = 'Passed'

                newItem.Name = formatDate
                print "newItem.Status", newItem.Status
                newItem.SetField(settingFileds[0],domain)
                newItem.SetField(settingFileds[1],client)
                newItem.SetField(settingFileds[2],tcExecution)
                print "Run Name :",newItem.Field("RN_RUN_NAME")
                qcRunID = newItem.Field("RN_RUN_ID")
                print "qcRunID:",qcRunID
                #print "-----",newItem.Field("RN_CYCLE_ID")
                #print "-----",newItem.Field("RN_SUBTYPE_ID")
                #print "-----",newItem.Field("RN_VC_STATUS")
                #print "-----",newItem.Field("RN_TEST_INSTANCE")

                newItem.Post()
                newItem.CopyDesignSteps()   # Copy Design Steps
                newItem.Post()
                steps = newItem.StepFactory.NewList("")
                print "Len:",len(steps)
                #step1 = steps[0]
                #step1.Status = 'Passed'
                #step1.post()
                count = 0
                runSts = ""
                for runSteps in steps:
                    print ts_Status[count]
                    if (ts_Status[count] == "Pass"):
                        runSts = "Passed"
                    elif(ts_Status[count] == "Fail"):
                        runSts = "Failed"
                    elif(ts_Status[count] == "N/A"):
                        runSts = "N/A"
                    runSteps.Status = runSts
                    runSteps.post()
                    time.sleep(1)
                    print runSteps.Status
                    count+=1

                #--------Attachment-------------#

                print copyFile
                print selecedFileName
                select_file = copyFile.split(".")[0]
                print "select_file",select_file
                folder_fileName = selecedFileName.split(".")[0]
                print "fileName",folder_fileName
                wk_dir = os.getcwd()
                rename_fileName = 'Evidence_'+folder_fileName+'_'+str(qcRunID)+'.docx'
                os.rename(select_file+'.docx',wk_dir+'\Evidence\\'+folder_fileName+'\\'+rename_fileName)
                zipFolderName = wk_dir+'\\Evidence\\'+folder_fileName+'\\Evidence_'+folder_fileName+'_'+str(qcRunID)
                zipFileName = wk_dir+'\\Evidence\\'+folder_fileName+'\\Evidence_'+folder_fileName+'_'+str(qcRunID)+'.zip'
                if not os.path.exists(zipFolderName): os.makedirs(zipFolderName)
                time.sleep(2)
                shutil.copy(wk_dir+'\Evidence\\'+folder_fileName+'\\'+rename_fileName, zipFolderName)
                zip_folder(zipFolderName,zipFileName)

                data = tsTest.Attachments
                datafile = data.AddItem(None)
                datafile.FileName = (zipFileName)
                datafile.Type = 1
                datafile.post()
                datafile.refresh()

                #break
                #break
                #tsTest.Status = tsINst_Status
                #tsTest.post()



    #tsIntanceFact = qc.TestInstanceFactory
    #testSetFilter = tsIntanceFact.Filter
    #d =testSetFilter.SetFilter(u"TS_TEST_ID",testPlanId)
    #d[0].SetField("TC_STATUS",tsINst_Status)
    #d.post()
    '''
    testSetFilter = tsFolder.Filter
    d =testSetFilter.SetFilter(u"CY_CYCLE_ID",testSetID)
    h = testSetFilter.NewList()
    print h

    for name in h:
        planName = name.Name
        print planName

    '''
    '''
    tSetFact = qc.TestFactory
    testSetFilter = tSetFact.Filter
    d =testSetFilter.SetFilter(u"TS_TEST_ID",testID)
    h = testSetFilter.NewList()

    exportTests(qc,h)
    '''
