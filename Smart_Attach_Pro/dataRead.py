__author__ = 'MR029157'

import os

working_dir = os.getcwd()

with open(working_dir+'\\QcRunData.txt','r') as ins:
    array = []
    for line in ins:
        try:
            array.append(line.split(":")[1])
            print array
        except:
            pass

print array[0]