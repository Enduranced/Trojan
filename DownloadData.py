import pandas as pd
import pathlib
import time
import os 
import subprocess
### For Excel
import xlsxwriter
import xlwings as xl
import openpyxl
import win32com.client
import win32com.client as win32com
from win32com.client import DispatchEx
import File_Paths

def write_BB_query_in_excel(bbgfilepath):
    ### getting the file paths and data
    data = pd.read_csv(File_paths.input_path) ## Do not cchange ot a xlsx
    Name_Hedge_Fund = data['Name of Fund'].tolist()
    Ticker_Hedge_Fund = data['Ticker Number'].tolist()

    ## Functions to run
    workbook = xlsxwriter.Workbook(File_Paths.data_path + '\\' + 'Data' +'.xlsx')
    for i in Ticker_Hedge_Fund:
        ws = workbook.add_worksheet(i)
        make_excel(r,ws)
    workbook.close()
    path = File_Paths.data_path + '\\' + 'Data' + '.xlsx'
    run_load(len(Ticker_Hedge_Fund), path, bbgfilepath, Ticker_Hedge_Fund)
    return

def make_excel(ticker,worksheet):
    ## Adding first formula Main Data of holdings in query
    formula1 = '=BQL.Query("' + 'get(id()) for (holdings(' + '{!r}'.format(ticker) +'))"' + ',' + '"showallcols=true")'
    worksheet.write('A1', formula1)
    ## Adding the heading Extra Data
    worksheet.write('K1', "Last Price")
    worksheet.write('L1', 'Total Float')
    worksheet.write('M1', 'last done volume')
    for i in range(2,1000):
        formula2 = '=BQL(A' + str(i) + ',' + '"px_last"' + ')'
        worksheet.write('K' + str(i) ,formula2)
        formula3 = '=BDP(A' + str(i) + ',' + '"EQY_SH_OUT"' + ')'
        worksheet.write('L' + str(i), formula3)
        formula5 = '=BDP(A' + str(i) + ',' + '"EQY_FLOAT"' + ')'
        worksheet.write('M' + str(i), formula5)
    
def run_load(no_names, WB, bbgfulepath, hedgefunds): ## Later need to add this variable
    ''' open the bloomberg API for excel and give it sufficient time to load in the data before closing and saving it '''
    bb = bbgfilepath ## Need to change to local Excel API for this to work
    x1 = DispatchEx('Excel.Application')
    x1.Workbooks.Open(bb)
    x1.AddIns('Bloomberg Excel Tools').Installed = True
    wb = x1.Workbooks.Open(Filename = WB)
    x1.Visible = True
    x1.EnableEvents = False
    x1.DisplayAlerts = False

    hedgefunds2 = hedgefunds.copy()
    count = 0
    timer = 0
    while True:
        ## Check each tab if the data has loaded
        for i in hedgefunds2:
            readData = wb.Sheets(i)
            totaldata = readData.UsedRange
            if totaldata.RowsCounter > 1: ### If the additional info is added please change this value to 999
                count += 1
                hedgefunds2.remove(i)
        
        if count == no_names:
            break
        if len(hedgefunds2) > 0 and timer > 150:
            print("Please Check Each Tab to find which one did not load")
            break
        timer += 5
        time.sleep(10)
    wb.Close(True)
    x1.Quit()
    del x1
    

