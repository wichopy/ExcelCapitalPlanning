"""
Copyright (C) 2014-2016, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""
import numpy as np
import pandas as pd
import xlwings as xw
import os
import datetime
current_dir = 'C:\\Users\\chouw\\Desktop\\Project Sub Comp Notification\\auto_proj_updates'
@xw.sub
def Filelist():
    """
    Print list of files in current folder in range A11
    until all files printed.
    """
    wb = xw.Book.caller()
    for dirname,subdirlist,filelist in os.walk(current_dir):
        i = 0
        for fname in filelist:
            wb.sheets['PullAssetData']['A'+str(11+i)].value = fname
            i += 1
    
@xw.sub
def show_AMS():
    wb = xw.Book.caller()
    sht = wb.sheets['AMS']
    sht2 = wb.sheets['AssetExtract']
    sht3 = wb.sheets['WSR']
    sht4 = wb.sheets['CMMSDump']
#    try:
    #AMS = sht.range('A1:R1055').options(pd.DataFrame, index=False).value
    AMS = sht.range('A1').expand().options(pd.DataFrame, index=False).value
    WSR = sht3.range('A1').expand().options(pd.DataFrame, index=False).value
    AssetExtract = sht2.range('A1').expand().options(pd.DataFrame, index=False).value
    CMMSDump = sht4.range('A1').expand().options(pd.DataFrame, index=False).value
    #df = sht.Range('A1').table.options(pd.DataFrame, index=false).value
    proj_num = wb.sheets['DashBoard'].range('B1').value
    wb.sheets['DashBoard'].range('A2').value = AMS[AMS['Project Number'] == proj_num][['Project Name','Building','Project Description','Scope of Work','Forecast Substantial Completion','Actual Substantial Completion','Executive Comments']]
    wb.sheets['DashBoard'].range('A4').value = WSR[WSR['EBA Project Number'] == proj_num][['Delivered By','Project Address','Project Manager','Client Status','GC Substantial Completion','Substantial Completion']]    
    BuildingID =  wb.sheets['DashBoard'].range('C3').value
    wb.sheets['DashBoard'].range('A8').expand().clear()
    wb.sheets['DashBoard'].range('A8').value = AssetExtract[AssetExtract['BUILDING ID'] == BuildingID][['BUILDING ITEM NUMBER','MATERIAL TYPE','INSTALLATION DATE','BUILDING ITEM ZONE','STATUS','QUANTITY','ASSETPROJECTNAME']]
    wb.sheets['DashBoard'].range('I8').value = ['Comments','New Install Date','Deactivate/Create New?','New Quantity?','Different Material Type ?']
    wb.sheets['DashBoard'].range('O8').expand().clear()
    wb.sheets['DashBoard'].range('Z8').value = ['Comments','Missing unit not in RA?','Should be Deactivate?']
    wb.sheets['DashBoard'].range('O8').value = CMMSDump[CMMSDump['Building Id'] == BuildingID][['Building Item Number','Status','Description','Building Item Zone','Installation Date','Serial Number','Manufacturer','Model Number','Field Item Number','DETAIL2']]
    #wb.sheets['DashBoard'].range('A8').value = 
    #    except:
#        wb.sheets['DashBoard']['A1'].value = "Not AMS loaded"

@xw.sub
def PullUpdates():
    wb = xw.Book.caller()
    sht = wb.sheets['DashBoard']
    proj_num = wb.sheets['DashBoard'].range('B1').value
    RA = sht.range('A8').expand().options(pd.DataFrame, index=False).value
    RAUpdates = RA[(RA['Comments'].notnull()) | \
                   (RA['New Install Date'].notnull()) | \
                   (RA['Deactivate/Create New?'].notnull()) |\
                   (RA['New Quantity?'].notnull()) | \
                   (RA['Different Material Type ?'].notnull()) ]
    RAUpdates.drop('ASSETPROJECTNAME', axis=1, inplace=True)
    RAUpdates['Date Project Reviewed'] = datetime.date.today().strftime("%d/%m/%y")
    RAUpdates['Project Number'] = sht.range('B1').value
    RAUpdates['PL Adjustment'] = 0
    RAUpdates['New Install Date'].fillna('1/1/2016', inplace=True)
    RAUpdates['Comments'].fillna('Replaced as per project: {}'.format(proj_num), inplace=True)
    #RAUpdates = RA.dropna(thresh=5)
    wb.sheets['DashBoard'].range('AE8').expand().clear()
    sht.range('AE8').value = RAUpdates

@xw.sub
def PullCMMSUpdates():
    wb = xw.Book.caller()
    sht = wb.sheets['DashBoard']
    proj_num = wb.sheets['DashBoard'].range('B1').value
    CMMS = sht.range('O8').expand().options(pd.DataFrame, index=False).value
    CMMSUpdates = CMMS[(CMMS['Comments'].notnull()) | \
                   (CMMS['Missing unit not in RA?'].notnull()) | \
                   (CMMS['Should be Deactivate?'].notnull())]
    CMMSUpdates['Date Project Reviewed'] = datetime.date.today().strftime("%d/%m/%y")
    CMMSUpdates['Project Number'] = ""
    CMMSUpdates[CMMSUpdates['Should be Deactivate?'].isnull()]['Project Number'] = proj_num
    CMMSUpdates['PL Adjustment'] = 0
    #CMMSUpdates['New Install Date'].fillna('1/1/2016', inplace=True)
    CMMSUpdates['Comments'].fillna('Replaced as per project: {}'.format(proj_num), inplace=True)
    #RAUpdates = RA.dropna(thresh=5)
    sht.range('AW8').expand().clear()
    sht.range('AW8').value = CMMSUpdates

if __name__ == '__main__':
    # To run this with the debug server, set UDF_DEBUG_SERVER = True in the xlwings VBA module
    xw.serve()
