#!/usr/bin/env python

"""
This file makes a 1.docx (Word 2007) file from scratch, showing off most of the
features of python-docx.

If you need to make documents from scratch, you can use this file as a basis
for your work.

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
"""

from docx import *
import os
from Smart_Attach import selecedFileName

def getImagesList(working_dir):

    print working_dir
    imgnamesList = []
    imgDir = working_dir+'/screen Shot'
    paths = [os.path.join(imgDir,fn) for fn in next(os.walk(imgDir))[2]]
    for names in paths:
       print names
       imgName = names.split("/")[-1]
       print imgName
       imgnamesList.append(imgName)

    return imgnamesList

def del_temp_media(working_dir):
       mediafolder = "%s/template/word/media" %working_dir
       print mediafolder
       for r, d, f in os.walk(mediafolder):
              for name in f:
                     os.remove(os.path.join(mediafolder, name))
       print "Temp Folder Cleared..!!"

def copy_images(working_dir,img_dir,img_names):
       rootDirs = img_names
       pattern_find = img_names[1]
       print "img_dir ",img_dir
       imgList = []
       alist_filter = ['jpg','bmp','png','gif','PNG']
       for r,d,f in os.walk(img_dir):
           for file in f:
                  changed_file = file.replace(" ", "")
                  if file[-3:] in alist_filter:
                         sourceFolder = os.path.join(img_dir,file)
                         print "XsourceFolderX ",sourceFolder
                         print "XXworking_dirXX ",working_dir
                         shutil.copy(sourceFolder,working_dir)
                         print "XfileX ",file
                         print "Xchanged_fileX ",changed_file
                         os.renames(working_dir + '/' + file, working_dir + '/' + changed_file)
                         imgList.append(changed_file)
       #os.remove(working_dir+"/Evid_iPhone_Module.docx")
       return imgList

def del_images(working_dir,img_names,indexNumb):
    print "Delete"
    print img_names
    for i in range(indexNumb,len(img_names)):
        print img_names[i]
        img_del_path = "%s/%s" %(working_dir,img_names[i])
        os.remove(img_del_path)


#working_dir = '/Users/ZA028309/PycharmProjects/Smart_Attach'
#tsPlanName = 'PCT_FNC_WF_Orders'
'''
def docCreat(tsPlanName,working_dir):
'''
if __name__ == '__main__':

    #imgPath = working_dir +"/screen Shot/"
    testName = 'PCT_FNC_WF_Orders'
    statusLableValue = ['Pass','Pass','Pass','Pass','Pass']
    stepNumberLable = "Step No:"
    #stepLableValue = "Step 1"
    statusLable = "Status:"
    #statusLableValue = "Pass"
    commentsLable = "Comments:"
    commentsLableValue = ""
    systemName = os.getlogin()

    print "Test PLan Name : ", selecedFileName

    #####################
    working_dir1 = os.getcwd()
    #print working_dir1
    working_dir = working_dir1 + '/imgDocCreater'
    rootDirPath =  os.path.dirname(working_dir)
    print "Root Path",rootDirPath
    del_temp_media(working_dir)

    img_dir = rootDirPath +'/screen Shot'
    #print "Image DIR ....", img_dir
    img_names = os.listdir(img_dir)
    #print img_names
    #print "WK",working_dir
    change_imgNames = copy_images(working_dir,img_dir,img_names)
    indexNumb = 0
    if(change_imgNames[0] == '.DS_Store'):
           indexNumb = 1
           print "Index is 1"
    else:
           print "indexNumb is ",indexNumb

    new_lineIndex = 0


    ######################

    #imagesNameList = getImagesList(working_dir)
    #print "Done imagesNameList"

    # Default set of relationshipships - the minimum components of a document
    relationships = relationshiplist()
     # Make a new document tree - this is the main part of a Word document
    document = newdocument()

    # This xpath location is where most interesting content lives
    body = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]

    tbl_startrow = [['Test Plan Name:',selecedFileName,'',''],['Solution(s):','PathNet -- BB Transfusio','',''],
            ['Created Date:','DD-MM-YYYY','No. Of Steps Requiring Evidence',''],['Enviroment:','CITRIX','Operating System:','Windows Server 2008 Enterprise Edition'],
            ['Associate ID:',systemName,'Domain:',''],['Test Date:','','','']
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

    for i in range (indexNumb,len(change_imgNames)):
        print "change_imgNames[i]", change_imgNames[i]
        newImageName = change_imgNames[i].split('/')[-1]
        print newImageName
        print "Image Name : ",img_names
        original_img_names = img_names[i].split('/')[-1]
        stepLableValue = original_img_names.split('.')[0]
        print "Step lable Value:",statusLable


        tbl_rows = [ [stepNumberLable, stepLableValue, statusLable,statusLableValue],
                     [commentsLable,"", "", ""]
                    ]

        tblContent = table(tbl_rows, heading=False, borders={'all': {'sz': 7, 'color': '#000000', 'val': 'single'}},
                           cwunit='pct', tblw=5300, twunit='pct')

        body.append(tblContent)

        body.append(paragraph(""))
        body.append(paragraph("Evidence:",'u'))
        body.append(paragraph(""))
        body.append(paragraph(""))

        # Add an image
        print "newImageName :",newImageName
        relationships, picpara = picture(relationships,original_img_names,
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
    contenttypes = contenttypes()
    websettings = websettings()
    wordrelationships = wordrelationships(relationships)

    # Save our document
    docName = rootDirPath+"/Result Evidence/"+selecedFileName+"1.docx"
    savedocx(document, coreprops, appprops, contenttypes, websettings,
             wordrelationships, docName)
    del_images(working_dir,change_imgNames,indexNumb)
