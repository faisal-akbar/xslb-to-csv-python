# import packages and set the xslb files directory path:
import os
import sys
import datetime
import win32com.client
import glob
excel = win32com.client.Dispatch("Excel.Application")
#excel.DisplayAlerts = False
#excel.Visible=False

# use glob to match the pattern 'xslb'
filePath="C:\\Users\\faisal\\Desktop\\xslbfiles\\"
excel_files = glob.glob(filePath+'*.xlsb')
#opDir=filePath+"Output\\"
#os.mkdir(opDir)
for excel1 in excel_files:
        print(excel1)
        file = excel1.split('.')[0]+'.csv'
        doc = excel.Workbooks.open(excel1)
        doc.SaveAs(Filename=file,FileFormat=24)
        doc.Close()
        excel.Quit()
