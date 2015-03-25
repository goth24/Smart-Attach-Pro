__author__ = 'ZA028309'

from docx import *
import os,time
#from Smart_Attach import tsPlanName
import shutil,platform



def copy_images(img_dir,img_names,imdDocFolder):
    oldImgList = []
    changeList = []
    for images in img_names:
        if images == ".DS_Store":
            print "No chnages"
        else:
            print images
            changedImgName = images.replace(" ","")
            shutil.copy2(img_dir+'/'+images,imdDocFolder+'/'+changedImgName)
            changeList.append(changedImgName)
            oldImgList.append(images)

    return changeList,oldImgList

def del_temp_media(working_dir):
       mediafolder = "%s\imgDocCreater\\template\word\media" %working_dir
       #print mediafolder
       for r, d, f in os.walk(mediafolder):
              for name in f:
                     os.remove(os.path.join(mediafolder, name))
       print "Temp Folder Cleared..!!"

def del_images(working_dir,img_names):
     for names in img_names:
         os.remove(os.path.join(working_dir, names))


#if __name__ == '__main__':
def generateFile(selecedFileName,stepResult,crNumberArrya, srNumberArrya, commentsArrya,noOfScreenPrints):

    #working_dir = os.getcwd()
    #imgPath = working_dir +"/screen Shot/"
    now = time.strftime("%c")
    formated_date = time.strftime("%d-%m-%Y")
    statusLableValue = stepResult
    #statusLableValue = ['Pass', 'Pass', 'Pass', 'Pass', 'Pass', 'Pass', 'Pass', 'Pass', 'Pass', 'Pass']
    stepNumberLable = "Step No:"
    #stepLableValue = "Step 1"
    statusLable = "Status:"
    #statusLableValue = "Pass"
    commentsLable = "Comments:"
    commentsLableValue = ""
    working_dir = os.getcwd()
    selectedDir = selecedFileName.split(".")[0]
    imageDir = working_dir+'\\Evidence\\'+selectedDir+'\\screen Shot'

    osSystem = platform.system()
    print osSystem
    if (osSystem == "Windows"):
        systemName = os.getenv('COMPUTERNAME')
        print systemName
        systemName = systemName.split("-")[1]
    else:
        systemName = os.getlogin()

    print "Test PLan Name : ", selecedFileName

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

    ####
    imdDocFolder = os.path.dirname(working_dir+'\\imgDocCreater')
    #print "Root Path",imdDocFolder

    del_temp_media(imdDocFolder)

    #img_dir = working_dir +'/screen Shot'
    img_names = os.listdir(imageDir)
    #print img_names

    #copying images and changed the name of the image
    print imageDir
    print img_names
    print imdDocFolder
    changedImageList,oldImageList = copy_images(imageDir,img_names,imdDocFolder)
    '''
    indexNumb = 0
    if(img_names[0] == '.DS_Store'):
           indexNumb = 1
           print "Index is 1"
    else:
           print "indexNumb is ",indexNumb

    new_lineIndex = 0
    '''

    # Default set of relationshipships - the minimum components of a document
    relationships = relationshiplist()
     # Make a new document tree - this is the main part of a Word document
    document = newdocument()

    # This xpath location is where most interesting content lives
    body = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]

    tbl_startrow = [['Test Plan Name:',selecedFileName,'',''],['Solution(s):',solution,'',''],
            ['Created Date:',formated_date,'No. Of Steps Requiring Evidence',str(noOfScreenPrints)],['Enviroment:',client,'Operating System:',sysOS],
            ['Associate ID:',systemName,'Domain:',domain],['Test Data:','N/A','','']
            ]
    tblContent1 = table(tbl_startrow, heading=False, borders={'all': {'sz': 7, 'color': '#000000', 'val': 'single'}},
                           cwunit='pct', tblw=5300, twunit='pct')
    body.append(tblContent1)
    body.append(paragraph(""))
    body.append(paragraph("Pre-Requisites:",'u'))

    body.append(pagebreak(type='page', orient='portrait'))

    body.append(paragraph("Actual Evidence:",'b'))

    body.append(paragraph(""))

    print "Image Loop"
    #print changedImageList
    #print oldImageList
    print "print statusLableValue : ",len(statusLableValue)
    for i in range(0,len(changedImageList)):
        print "This is i: ",i

        stepLableValue = oldImageList[i].split('.')[0]
        print stepLableValue
        #######################
        stLableValue = statusLableValue[i]
        crNumberLst = crNumberArrya[i]
        srNumberLst = srNumberArrya[i]
        commentLst = commentsArrya[i]
        '''
        try:
            stLableValue = statusLableValue[i]
            crNumberLst = crNumberArrya[i]
            srNumberLst = srNumberArrya[i]
            commentLst = commentsArrya[i]
        except:
            stLableValue = "Pass"
        '''

        ########################

        tbl_rows = [ [stepNumberLable, stepLableValue, statusLable,stLableValue],
                     [commentsLable,crNumberLst, srNumberLst, commentLst]
                    ]

        tblContent = table(tbl_rows, heading=False, borders={'all': {'sz': 7, 'color': '#000000', 'val': 'single'}},
                           cwunit='pct', tblw=5300, twunit='pct')

        body.append(tblContent)

        body.append(paragraph(""))
        body.append(paragraph("Evidence:",'u'))
        body.append(paragraph(""))
        body.append(paragraph(""))

        # Add an image
        relationships, picpara = picture(relationships,changedImageList[i],
                                     'This is a test description')
        body.append(picpara)

        # Add a pagebreak
        body.append(pagebreak(type='page', orient='portrait'))

    # Create our properties, contenttypes, and other support files
    title    = 'Python docx demo'
    subject  = 'A practical example of making docx from Python'
    creator  = 'Zaheer Ahmed'
    keywords = ['python', 'Office Open XML', 'Word']

    coreprops = coreproperties(title=title, subject=subject, creator=creator,
                               keywords=keywords)
    appprops = appproperties()
    contenttype = contenttypes()
    websetting = websettings()
    wordrelationship = wordrelationships(relationships)

    selecedFileName = selecedFileName.split('.')[0]
    # Save our document
    docName = working_dir+'\\Evidence\\'+selectedDir+'\\'+selecedFileName+".docx"
    savedocx(document, coreprops, appprops, contenttype, websetting,
             wordrelationship, docName)
    del_images(working_dir,changedImageList)
