__author__ = 'ZA028309'


import pyscreenshot as ImageGrab
import Tkinter, tkFileDialog, Tkconstants
from Tkinter import *
import Tkinter as tk
import tkMessageBox
from PIL import ImageTk, Image
from docx import *
import time ,os,sys,datetime, string
import fnmatch,shutil
import time ,os,sys
import subprocess
'''
def dd():
    x = """ this is just a word
        file to get you out
        """
    return x


working_dir = os.getcwd()
if __name__=='__main__':
    root = tk.Tk()
    root.geometry("500x350")
    root.title('Smart_Attach')
    #working_dir = os.getcwd()
    print working_dir
    PoweredData = "Smart_Attach"
    bckground = '//img/cerner_background.png'
    logo = '//img/Cerner_logo.png'
    path = working_dir + logo
    print path
    labelframe = LabelFrame(root,text= PoweredData)
    labelframe.pack(anchor = "s",fill = "both", expand="yes",padx = 10)
    img = ImageTk.PhotoImage(Image.open(path))
    #panel = tk.Label(labelframe,padx = 15,pady = 15, image = img)
    panel = tk.Label(labelframe,padx = 15,pady = 15)
    panel.pack( fill = "both",expand = "yes")

    z = StringVar()
    z.set("")

    tsxbox = Text(panel)
    tsxbox.pack()
    tsxbox.insert(END,z)
    b1 = Button(text = "d1", command = dd)
    b1.pack()
    root.mainloop()
'''

def screen():
    root.withdraw()
    ImageGrab.grab().save('/Users/ZA028309/Desktop/Image_copy/1.png')

root = tk.Tk()
b1 = Button(root, text="B1",width =20, command= screen)
b1.pack()
root.mainloop()