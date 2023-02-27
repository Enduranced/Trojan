import pandas as pd
import File_Paths
import time
import pathlib
from openpyxl import load_workbook
import xlwings as xl
from Data_Analysis import *

def make_Excel(input_data, top, details):  ### Details capture if the user wants to have an excel of possible other factors and breakdown
    wanted_values = []
    for i in input_data:
        if i == 'OFF':
            continue
        if i == 'common_holding_freq':
            function = common_holding_freq(top)
            names = ['Stock_Ticker',]
            values = ['Frequency Of Ticker',]
            for i in function:
                names.append(i[0])
                values.append(i[1])
            inter_dic = {'Analaysis Method' : names, i : values}
            dff = pd.DataFrame(inter_dic)
            local_path = File_Paths.output_path + '\\' + "Common Holding Base on Frequency" + '.xlsx'
            writer = pd.ExcelWriter(local_path, engine  = 'xlsxwriter')
            dff.to_excel(writer, sheet_name = 'Analysis')
            writer.save()
            ### Getting the xlsxwriter objects from the dataframe writer object
            if details == str(True):
                excel_app = xl.App(visible = False)
                wb = xl.Book(local_path)
                for i in names[1:]:
                    ws = wb.sheets.add(i)
                    formula = '=BDS("' + str(i) + '",' + '"ALL_HOLDERS_PUBLIC_FILINGS"' + ',' +'"headers =T"' + ')'
                    ws.range('A1').value = formula
                xl.EnableEvents = False
                xl.DisplayAlerts = False
                wb.save()
                wb.close()
                excel_app.quit()
            else:
                continue

        if i == 'common_holding_no_share':
            function = common_holding_no_share(top)
            names = ['Stock_Ticker',]
            values = ['Total Number Of Shares Held by hedge Fund',]
            for i in function:
                names.append(i[0])
                values.append(i[1])
            inter_dic = {'Analaysis Method' : names, i : values}
            dff = pd.DataFrame(inter_dic)
            local_path = File_Paths.output_path + '\\' + "Common Holding Base on Number Of Shares" + '.xlsx'
            writer = pd.ExcelWriter(local_path, engine  = 'xlsxwriter')
            dff.to_excel(writer, sheet_name = 'Analysis')
            writer.save()
            ### Getting the xlsxwriter objects from the dataframe writer object
            if details == str(True):
                excel_app = xl.App(visible = False)
                wb = xl.Book(local_path)
                for i in names[1:]:
                    ws = wb.sheets.add(i)
                    formula = '=BDS("' + str(i) + '",' + '"ALL_HOLDERS_PUBLIC_FILINGS"' + ',' +'"headers =T"' + ')'
                    ws.range('A1').value = formula
                xl.EnableEvents = False
                xl.DisplayAlerts = False
                wb.save()
                wb.close()
                excel_app.quit()
            else:
                continue
        
        else:
            function = common_holding_vol_change(top)
            names = ['Stock_Ticker',]
            values = ['Absolute Change Of Position',]
            for i in function:
                names.append(i[0])
                values.append(i[1])
            inter_dic = {'Analaysis Method' : names, i : values}
            dff = pd.DataFrame(inter_dic)
            local_path = File_Paths.output_path + '\\' + "Common Holding absolute changes in volume" + '.xlsx'
            writer = pd.ExcelWriter(local_path, engine  = 'xlsxwriter')
            dff.to_excel(writer, sheet_name = 'Analysis')
            writer.save()
            ### Getting the xlsxwriter objects from the dataframe writer object
            if details == str(True):
                excel_app = xl.App(visible = False)
                wb = xl.Book(local_path)
                for i in names[1:]:
                    ws = wb.sheets.add(i)
                    formula = '=BDS("' + str(i) + '",' + '"ALL_HOLDERS_PUBLIC_FILINGS"' + ',' +'"headers =T"' + ')'
                    ws.range('A1').value = formula
                xl.EnableEvents = False
                xl.DisplayAlerts = False
                wb.save()
                wb.close()
                excel_app.quit()
            else:
                continue