
import pandas as pd
import os
import datetime
import numpy as np

questdir = 'C:\Users\chouw\Desktop\Load3'

#check to make sure amount of files is correct.
# store all file names in a list.
files = []
import os
count = 0
for dirName, subdirlist, fileList in os.walk(questdir):
    for fname in fileList:
        files.append(fname)
        count += 1
#print files
print count

#Pull column headers from blank questionnaire
header = pd.read_excel('Questionnaire.xlsm', sheetname ='Asset Form', header = 2)    

#blank dataframe for storing data.
all_data = pd.DataFrame()
datenow = str(datetime.datetime.now().strftime("%Y%m%d%H%M%S"))

for fname in files:
        
    pathtoq = os.path.join(questdir,fname)
    pathtoq = pathtoq.replace('\\','/')

    def read_excel(qpath):    
        capital = pd.read_excel(qpath,sheetname = "Asset Form", header = 2)
        trimmed = capital[capital["Active Capital (C)"] == "c"]# filter for only active capital.
        return trimmed
    try:
        trimmed = read_excel(pathtoq)
    except:
        print"{} could not load".format(fname)
        continue
    #check if there are too many header columns
    while len(trimmed.columns) != 45:
        
        print trimmed.shape
        try:
            raw_input("Something wrong with {}, please check it and continue creating masterlist.".format(fname))
            trimmed = read_excel(pathtoq)
        except:
            print "{} failed".format(fname)
        # check if building Id in file matches filename
    #check if column headers were modified, if they were they will cause the header order to be alphhabetical.
    while (trimmed.columns.values != headernames).all():
        print trimmed.columns.values
        raw_input("Headers don't match in {}, please check it and continue creating masterlist.".format(fname))
        trimmed = read_excel(pathtoq)
        
    #add column for filenames.
    trimmed['filename'] = pd.Series(fname,index = trimmed.index)
    all_data = all_data.append(trimmed,ignore_index=True)

    print "{} added to masterlist.".format(fname)
    print trimmed.shape

print "write files to excel"

#add column header for filenames
headernamesfinal = np.append(headernames,'filename')
writer = pd.ExcelWriter("combine_reviewed"+datenow+".xlsx")

all_data.to_excel(writer,"Assets", columns=headernamesfinal)

writer.save()

print "write complete" 
