import os
import datetime
from RPA.Excel.Files import Files as ExcelFiles
from variables import pathResultExcel
from function import CommonFunction

def parsingHTMLtdValue(StringInput):
    StartValuePos = StringInput.index(">") + 1
    EndValuePos = StringInput.index("</td>")
    NumType = StringInput.find('x:num')
    Value = StringInput[StartValuePos:EndValuePos]
    
    if NumType > 0:
        Value = float(Value)
        
    return Value
class ExcelRowItem: 
    def __init__(self, Col0, Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8, Col9, Col10, Col11, Col12, Col13, Col14, Col15, Col16, Col17, Col18, Col19, Col20, Col21, Col22, Col23): 
        self.Col0 = Col0 # Contract ID
        self.Col1 = Col1 # Contract ID/External Reference Number
        self.Col2 = Col2 # Contract Name
        self.Col3 = Col3 # Company Code
        self.Col4 = Col4 # Business Unit
        self.Col5 = Col5 # Trading Partner
        self.Col6 = Col6 # Activation Group ID
        self.Col7 = Col7 # Activation Group Status
        self.Col8 = Col8 # Accounting Start Date
        self.Col9 = Col9 # Likely Expiration Date
        self.Col10 = Col10 # Accounting Term (Month)
        self.Col11 = Col11 # Accounting Term (Days)
        self.Col12 = Col12 # Contract Currency
        self.Col13 = Col13 # Target Currency
        self.Col14 = Col14 # Posting Date
        self.Col15 = Col15 # Exchange rate to LC
        self.Col16 = Col16 # Exchange rate to GC
        self.Col17 = Col17 # Asset Class
        self.Col18 = Col18 # Finance Lease (Principal)
        self.Col19 = Col19 # Finance Lease (Interest)
        self.Col20 = Col20 # Operating Lease Costs
        self.Col21 = Col21 # Variable Lease Expense
        self.Col22 = Col22 # Low Value Lease Expense
        self.Col23 = Col23 # Short Term Lease Expense


def convertHTMLtoExcel(typeData, FileName, FilePath):
    try:
        CommonFunction.WriteLog(f"Start convert HTML to Excel - "+FileName)
        CommonFunction.WriteLog("Start at: " + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        FilePathNew = pathResultExcel+"\\"+FileName+".txt"
        FilePathOutput = pathResultExcel+"\\"+FileName+".xlsx"

        # Remove Txt File if exists
        if os.path.exists(FilePathNew):
            os.remove(FilePathNew)

        # Generate HTML to Txt Files
        with open(FilePath, 'r', encoding="utf8") as f:
            content = f.read()
            new_content = content[content.index('<table>'):content.index('</table>') + 8]
            text_file = open(FilePathNew, "w", encoding="utf8")
            text_file.write(new_content)
            text_file.close()

        countdatalines = 0
        countstartlines = 0
        indexstartlines = 0
        indexstartdata = 0
        indexfinishdata = 0
        indexdatacolumn = 0
        countstartflag = False
        output = []
        TempCol0 = ""
        TempCol1 = ""
        TempCol2 = ""
        TempCol3 = ""
        TempCol4 = ""
        TempCol5 = ""
        TempCol6 = ""
        TempCol7 = ""
        TempCol8 = ""
        TempCol9 = ""
        TempCol10 = ""
        TempCol11 = ""
        TempCol12 = ""
        TempCol13 = ""
        TempCol14 = ""
        TempCol15 = ""
        TempCol16 = ""
        TempCol17 = ""
        TempCol18 = ""
        TempCol19 = ""
        TempCol20 = ""
        TempCol21 = ""
        TempCol22 = ""
        TempCol23 = ""
        ProcessedExcelRow = 0

        # Looping Txt files and Get Excel Rows into List of Object
        CommonFunction.WriteLog("Looping Txt files and Get Excel Rows into List of Object")
        zz = 1 
        with open(FilePathNew, 'r', encoding='UTF-8') as file:
            CommonFunction.WriteLog(f"iteration: {zz}")
            while (line := file.readline().rstrip()):
                # Finding First Row of Data
                if(line == "<td colspan=1 class='style7'>Short Term Lease Expense</td>"):
                    indexstartlines = countdatalines + 3
                    countstartflag = True
                elif(line=="<td colspan=1 class='style7'>Total Remaining Undiscounted Payments (Total Lease Liability Closing Balance + ST/LT Future Interest Expense)</td>"):
                    indexstartlines = countdatalines + 3
                    countstartflag = True
                
                # Finding Start Row for Looping 1 excel Row 
                if(
                    countstartflag == True 
                    and indexstartlines <= countdatalines
                    and indexfinishdata < countdatalines
                ):
                    indexstartdata = countdatalines
                    if typeData != 'IFRS':
                        totalColumn = 23
                    else:
                        totalColumn = 21
                    indexfinishdata = countdatalines + totalColumn

                # Read Data in 1 excel Row
                if(
                    countstartflag == True 
                    and indexstartlines <= countdatalines
                    and indexstartdata <= countdatalines <= indexfinishdata
                ):
                    indexdatacolumn = countdatalines - indexstartdata
                    cleanvalue = parsingHTMLtdValue(line)
                    if(indexdatacolumn == 0): TempCol0 = cleanvalue
                    if(indexdatacolumn == 1): TempCol1 = cleanvalue
                    if(indexdatacolumn == 2): TempCol2 = cleanvalue
                    if(indexdatacolumn == 3): TempCol3 = cleanvalue
                    if(indexdatacolumn == 4): TempCol4 = cleanvalue
                    if(indexdatacolumn == 5): TempCol5 = cleanvalue
                    if(indexdatacolumn == 6): TempCol6 = cleanvalue
                    if(indexdatacolumn == 7): TempCol7 = cleanvalue
                    if(indexdatacolumn == 8): TempCol8 = cleanvalue
                    if(indexdatacolumn == 9): TempCol9 = cleanvalue
                    if(indexdatacolumn == 10): TempCol10 = cleanvalue
                    if(indexdatacolumn == 11): TempCol11 = cleanvalue
                    if(indexdatacolumn == 12): TempCol12 = cleanvalue
                    if(indexdatacolumn == 13): TempCol13 = cleanvalue
                    if(indexdatacolumn == 14): TempCol14 = cleanvalue
                    if(indexdatacolumn == 15): TempCol15 = cleanvalue
                    if(indexdatacolumn == 16): TempCol16 = cleanvalue
                    if(indexdatacolumn == 17): TempCol17 = cleanvalue
                    if(indexdatacolumn == 18): TempCol18 = cleanvalue
                    if(indexdatacolumn == 19): TempCol19 = cleanvalue
                    if(indexdatacolumn == 20): TempCol20 = cleanvalue
                    if(indexdatacolumn == 21): TempCol21 = cleanvalue
                    if typeData != 'IFRS':
                        if(indexdatacolumn == 22): TempCol22 = cleanvalue
                        if(indexdatacolumn == 23): TempCol23 = cleanvalue
                    if(countdatalines == indexfinishdata):
                        ProcessedExcelRow = ProcessedExcelRow + 1
                        # CommonFunction.WriteLog("Success Process Excel Row: " + str(ProcessedExcelRow))
                        indexstartlines = countdatalines + 3
                        output.append(ExcelRowItem(TempCol0, TempCol1, TempCol2, TempCol3, TempCol4, TempCol5, TempCol6, TempCol7, TempCol8, TempCol9, TempCol10, TempCol11, TempCol12, TempCol13, TempCol14, TempCol15, TempCol16, TempCol17, TempCol18, TempCol19, TempCol20, TempCol21, TempCol22, TempCol23))

                if(countstartflag == True):
                    # CommonFunction.WriteLog("Line{}: {}".format(countdatalines, line))
                    countstartlines += 1

                # if(countstartlines > 1000000):
                #     break

                countdatalines += 1
        
        CommonFunction.WriteLog("Exit Loop with Output no of rows: " + str(len(output)))

        # Remove Excel File if exists
        CommonFunction.WriteLog("Remove excel File if exists")
        if os.path.exists(FilePathOutput):
            os.remove(FilePathOutput)

        # Write Output to Excel
        CommonFunction.WriteLog("Write Output to Excel")
        ExcelLib = ExcelFiles()
        ExcelLib.create_workbook(path = FilePathOutput)
        ExcelLib.create_worksheet(name = "Data")
        ExcelLib.set_active_worksheet(value = "Data")

        # Write Header Data
        CommonFunction.WriteLog("Write Header Data")
        ExcelLib.set_cell_value(row = 1, column = 1, value = "Contract ID")
        ExcelLib.set_cell_value(row = 1, column = 2, value = "Contract ID/External Reference Number")
        ExcelLib.set_cell_value(row = 1, column = 3, value = "Contract Name")
        ExcelLib.set_cell_value(row = 1, column = 4, value = "Company Code")
        ExcelLib.set_cell_value(row = 1, column = 5, value = "Business Unit")
        ExcelLib.set_cell_value(row = 1, column = 6, value = "Trading Partner")
        ExcelLib.set_cell_value(row = 1, column = 7, value = "Activation Group ID")
        ExcelLib.set_cell_value(row = 1, column = 8, value = "Activation Group Status")
        ExcelLib.set_cell_value(row = 1, column = 9, value = "Accounting Start Date")
        ExcelLib.set_cell_value(row = 1, column = 10, value = "Likely Expiration Date")
        ExcelLib.set_cell_value(row = 1, column = 11, value = "Accounting Term (Month)")
        ExcelLib.set_cell_value(row = 1, column = 12, value = "Accounting Term (Days)")
        ExcelLib.set_cell_value(row = 1, column = 13, value = "Contract Currency")
        ExcelLib.set_cell_value(row = 1, column = 14, value = "Target Currency")
        if typeData != 'IFRS':
            ExcelLib.set_cell_value(row = 1, column = 15, value = "Posting Date")
            ExcelLib.set_cell_value(row = 1, column = 16, value = "Exchange rate to LC")
            ExcelLib.set_cell_value(row = 1, column = 17, value = "Exchange rate to GC")
            ExcelLib.set_cell_value(row = 1, column = 18, value = "Asset Class")
            ExcelLib.set_cell_value(row = 1, column = 19, value = "Finance Lease (Principal)")
            ExcelLib.set_cell_value(row = 1, column = 20, value = "Finance Lease (Interest)")
            ExcelLib.set_cell_value(row = 1, column = 21, value = "Operating Lease Costs")
            ExcelLib.set_cell_value(row = 1, column = 22, value = "Variable Lease Expense")
            ExcelLib.set_cell_value(row = 1, column = 23, value = "Low Value Lease Expense")
            ExcelLib.set_cell_value(row = 1, column = 24, value = "Short Term Lease Expense")
        elif typeData == 'SLANUSGAAP' or typeData == 'SLANIFRS':
            ExcelLib.set_cell_value(row = 1, column = 15, value = "Exchange rate to LC")
            ExcelLib.set_cell_value(row = 1, column = 16, value = "Exchange rate to GC")
            ExcelLib.set_cell_value(row = 1, column = 17, value = "Asset Class")
            ExcelLib.set_cell_value(row = 1, column = 18, value = "ST Future Interest Expense")
            ExcelLib.set_cell_value(row = 1, column = 19, value = "LT Future Interest Expense")
            ExcelLib.set_cell_value(row = 1, column = 20, value = "Accrued Interest Closing Balance")
            ExcelLib.set_cell_value(row = 1, column = 21, value = "ST Principal Closing Balance")
            ExcelLib.set_cell_value(row = 1, column = 22, value = "LT Principal Closing Balance")
            ExcelLib.set_cell_value(row = 1, column = 23, value = "Total Lease Liability Closing Balance (ST/LT Principal Closing Balance + Accrued Interest Closing Balance)")
            ExcelLib.set_cell_value(row = 1, column = 24, value = "Total Remaining Undiscounted Payments (Total Lease Liability Closing Balance + ST/LT Future Interest Expense)")
        else:
            ExcelLib.set_cell_value(row = 1, column = 15, value = "Posting Date")
            ExcelLib.set_cell_value(row = 1, column = 16, value = "Exchange rate to LC")
            ExcelLib.set_cell_value(row = 1, column = 17, value = "Exchange rate to GC")
            ExcelLib.set_cell_value(row = 1, column = 18, value = "Asset Class")
            ExcelLib.set_cell_value(row = 1, column = 19, value = "Finance Lease (Principal)")
            ExcelLib.set_cell_value(row = 1, column = 20, value = "Finance Lease (Interest)")
            ExcelLib.set_cell_value(row = 1, column = 21, value = "Low Value Lease Expense")
            ExcelLib.set_cell_value(row = 1, column = 22, value = "Short Term Lease Expense")
            
        ExcelLib.set_styles("A1:X1", bold = True, font_name="Calibri", size=14)

        # Write Detail Data
        
        CommonFunction.WriteLog("Write Detail Data")
        loopOutputRowNum = 2
        for loopOutput in output:
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 1, value = loopOutput.Col0)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 2, value = loopOutput.Col1)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 3, value = loopOutput.Col2)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 4, value = loopOutput.Col3)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 5, value = loopOutput.Col4)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 6, value = loopOutput.Col5)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 7, value = loopOutput.Col6)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 8, value = loopOutput.Col7)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 9, value = loopOutput.Col8)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 10, value = loopOutput.Col9)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 11, value = loopOutput.Col10)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 12, value = loopOutput.Col11)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 13, value = loopOutput.Col12)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 14, value = loopOutput.Col13)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 15, value = loopOutput.Col14)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 16, value = loopOutput.Col15)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 17, value = loopOutput.Col16)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 18, value = loopOutput.Col17)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 19, value = loopOutput.Col18)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 20, value = loopOutput.Col19)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 21, value = loopOutput.Col20)
            ExcelLib.set_cell_value(row = loopOutputRowNum, column = 22, value = loopOutput.Col21)
                
            if typeData != 'IFRS':
                ExcelLib.set_cell_value(row = loopOutputRowNum, column = 23, value = loopOutput.Col22)
                ExcelLib.set_cell_value(row = loopOutputRowNum, column = 24, value = loopOutput.Col23)
            loopOutputRowNum = loopOutputRowNum + 1

        ExcelLib.save_workbook()
        ExcelLib.close_workbook()

        CommonFunction.WriteLog("Finish at: " + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        CommonFunction.WriteLog("Finish convert HTML to Excel")
    except Exception as error:
        CommonFunction.WriteLog("Failed convert HTML to Excel")

	