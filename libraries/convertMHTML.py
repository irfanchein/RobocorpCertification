import os
import datetime
from function import CommonFunction
from RPA.Excel.Files import Files as ExcelFiles
from variables import pathResultSAP

def parsingHTMLtdValue(StringInput):
    try:
        StartValuePos = StringInput.index('title="')+7
        EndValuePos = len(StringInput) -2
        Value = StringInput[StartValuePos:EndValuePos]
            
        return Value
    except Exception as err:
        CommonFunction.WriteLog(f"error: {str(err)}")

class ExcelRowItem: 
    def __init__(self, Col0): 
        self.Col0 = Col0 
        
def convertMHTMLtoExcel(FileName, FilePath, FilePathOutput):
    try:
        CommonFunction.WriteLog(f"Start convert HTML to Excel - "+FileName)
        CommonFunction.WriteLog(f"Start at: " + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        FilePathNew = pathResultSAP+"\\"+FileName+".txt"
        
            
        # Remove Txt File if exists
        if os.path.exists(FilePathNew):
            os.remove(FilePathNew)

        # Generate MHTML to Txt Files
        with open(FilePath, 'r', encoding="utf8") as f:
            content = f.read()
            new_content = content[content.index('<table  class="list" border=1 cellspacing=0 cellpadding=1 rules=groups borderColor=black >'):content.index('</html>')]
            text_file = open(FilePathNew, "w", encoding="utf8")
            StartTable = new_content.index('<table  class="list" border=1 cellspacing=0 cellpadding=1 rules=groups borderColor=black >')
            EndTable = new_content.index('</blockquote>')
            i = 0         
            while i <= EndTable:
                try:
                    StartLine = new_content.index('<tr')
                    EndLine = new_content.index('</tr>')
                    LineContent = new_content[StartLine:EndLine]
                    ImageContent = LineContent[LineContent.index('<td ><input'):LineContent.index('</td>')]
                    if i == 0: text_file.write("<tr>\n")
                    if i > 0:
                        text_file.write("<tr>\n"+ImageContent+"\n")
                    StartTable = EndLine + 5
                    new_content = new_content[StartTable:]
                    if i == EndTable:text_file.write("</tr>\n")
                except Exception as e:
                    StartTable = EndLine + 5
                    new_content = new_content[StartTable:]
                i = i + 1
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
        ProcessedExcelRow = 0
        CommonFunction.WriteLog(f"new_content: {new_content}")

        # Looping Txt files and Get Excel Rows into List of Object
        CommonFunction.WriteLog("Looping Txt files and Get Excel Rows into List of Object")
        zz = 1 
        with open(FilePathNew, 'r', encoding='UTF-8') as file:
            while (line := file.readline().rstrip()):
                # Finding First Row of Data
                if(line == '<tr>'):
                    # CommonFunction.WriteLog("Position of Start: " + str(countdatalines))
                    indexstartlines = countdatalines + 1
                    countstartflag = True
                    
                # Finding Start Row for Looping 1 excel Row 
                if(
                    countstartflag == True 
                    and indexstartlines <= countdatalines
                    and indexfinishdata < countdatalines
                ):
                    indexstartdata = countdatalines
                    indexfinishdata = countdatalines
                    
                # Read Data in 1 excel Row
                if(
                    countstartflag == True 
                    and indexstartlines <= countdatalines
                    and indexstartdata <= countdatalines <= indexfinishdata
                ):
                    # CommonFunction.WriteLog("after if")
                    indexdatacolumn = countdatalines - indexstartdata
                    
                    cleanvalue = parsingHTMLtdValue(line)
                    if(indexdatacolumn == 0): 
                        TempCol0 = cleanvalue
                        ProcessedExcelRow = ProcessedExcelRow + 1
                        # CommonFunction.WriteLog("Success Process Excel Row: " + str(ProcessedExcelRow))
                        indexstartlines = countdatalines + 1
                        output.append(ExcelRowItem(TempCol0))
                    
                if(countstartflag == True):
                    countstartlines += 1

                countdatalines += 1
        
        CommonFunction.WriteLog("Exit Loop with Output no of rows: " + str(len(output)))

        if os.path.exists(FilePathOutput):

            # Write Output to Excel
            CommonFunction.WriteLog("Write Output to Excel")
            ExcelLib = ExcelFiles()
            ExcelLib.open_workbook(FilePathOutput)
            ExcelLib.set_active_worksheet(value=0)

            # Write Detail Data
            
            CommonFunction.WriteLog("Write Detail Data")
            loopOutputRowNum = 2
            for loopOutput in output:
                colH = ExcelLib.get_cell_value(row=loopOutputRowNum, column="H")
                if colH == None:
                    loopOutputRowNum = loopOutputRowNum + 1
                
                ExcelLib.set_cell_value(row = loopOutputRowNum, column = 1, value = loopOutput.Col0)
                loopOutputRowNum = loopOutputRowNum + 1
            
            ExcelLib.save_workbook()
            ExcelLib.close_workbook()

            CommonFunction.WriteLog("Finish at: " + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
            CommonFunction.WriteLog("Finish convert MHTML to Excel")
    except Exception as error:
        CommonFunction.WriteLog("Failed convert MHTML to Excel")
        CommonFunction.WriteLog(str(error))