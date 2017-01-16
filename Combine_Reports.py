import pandas as pd
import os
import datetime
import numpy as np

#This will be the directory with all the sheets to combine.
questdir = 'C:\Users\chouw\Desktop\Load3'

#chunk of code to verify file count matches.
files = []
import os
count = 0
for dirName, subdirlist, fileList in os.walk(questdir):
    for fname in fileList:
        files.append(fname)
        count += 1
#print files
print count

#Pull header names from a blank report for use later on.
header = pd.read_excel('Empty_Questionnaire.xlsm', sheetname ='Assets', header = 2)    
headernames = header.columns.values

"""
loop for combining sheets. 
read each excel sheet with pandas. Append it to the end of the main dataframe.
"""
all_data = pd.DataFrame()
datenow = str(datetime.datetime.now().strftime("%Y%m%d%H%M%S")) # for use in output file.

#read and filter excel file function
def read_excel(qpath):    
    report = pd.read_excel(qpath,sheetname = "Asset", header = 2)
    trimmed = report[report["Active Capital"] == "c"]
    return trimmed
    
for fname in files:      
    pathtoq = os.path.join(questdir,fname)
    pathtoq = pathtoq.replace('\\','/')
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
        
    trimmed['filename'] = pd.Series(fname,index = trimmed.index) # add column at the end for filename.
    all_data = all_data.append(trimmed,ignore_index=True)
    print "{} added to masterlist.".format(fname)
    print trimmed.shape


print "write files to excel"
headernamesfinal = np.append(headernames,'filename')
writer = pd.ExcelWriter("combine_reviewed"+datenow+".xlsx")
all_data.to_excel(writer,"Assets", columns=headernamesfinal)

writer.save()

print "write complete"       
