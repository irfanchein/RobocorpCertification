
import datetime as datetimefunc
import glob
import shutil
import re

from convertFunction import convertHTMLtoExcel
from convertMHTML import convertMHTMLtoExcel
from datetime import datetime
from Excel import Excel
from function import selenium, CommonFunction, kill_process_username, sendNotification
from SAP import generateSAP
from time import sleep
from variables import downloadPath, excel, excelPath, dateStart, dateEnd
from variables import LeaseReconSchedulerPath, LeaseReconConfigPath, LeaseReconTemplatePath, LeaseReconSystem
from variables import maxRetry, maxWait, URL_NAKISA
from variables import pathResultExcel, pathResultSAP, PMI_UserAccount, PMI_Password, SeleniumLib, sleep3, sleep10, timeout_selenium60, timeout_selenium60, timeout_selenium60
from variables import os, sleep3, sleep5,  TablesLib, WindowsLib

def createExcelPath(type, path):
    CommonFunction.WriteLog(f"Start create excel path.")
    if str(type) == 'Active Group':
        excelPath.pathAGSLAN = path
    if str(type) == 'Unit List':
        excelPath.pathUNITSLAN = path
    if str(type) == 'USGAAP':
        excelPath.pathUSGAAP = path
    if str(type) == 'IFRS':
        excelPath.pathIFRS = path
    if str(type) == 'SLANUSGAAP':
        excelPath.pathLiabilityUSGAAP = path
    if str(type) == 'SLANIFRS':
        excelPath.pathLiabilityIFRS = path
    CommonFunction.WriteLog(f"Finish create excel path.")

def checkFileExists(prefixAct, type):
    excelRename = downloadPath+"\\"+type+".xlsx"
    excelCopy   = pathResultExcel+"\\"+type+".xlsx"
    try:
        found = False
        for ZipFiles in os.listdir(downloadPath):
            name = os.path.basename(ZipFiles)
            prefix = name[:-19]
            excelFiles = downloadPath+"\\"+name
            if prefix.lower() == prefixAct.lower():
                if len(excelFiles) > 0:
                    found = True
                    if os.path.exists(excelRename):
                        os.remove(excelRename)
                    os.rename(excelFiles, excelRename)
                    shutil.copy(excelRename, excelCopy)
                    CommonFunction.WriteLog(f"Excel {type} file downloaded {excelCopy}")
                    createExcelPath(type, path=excelCopy)
                else:    
                    CommonFunction.WriteLog(f"Finish check file doesn't exists.")
            else:
                if os.path.exists(excelCopy):
                    found = True
                    CommonFunction.WriteLog(f"Excel {type} file downloaded.")
    
                    break
                    
    except Exception as err:
        CommonFunction.WriteLog(f"Check check downloaded file error with {str(err)}.")
        CommonFunction.WriteConsole(f"Check check downloaded file error with {str(err)}.")
    return found

def removeNotif():
    CommonFunction.WriteLog(f"Start cleanup existing notif.")
    try:
        selenium.WaitElement('ng-scope', 'class', False, timeout_selenium60)
        selenium.ClickElement('ng-scope', 'class', False)
        sleep(sleep10)
        
        selenium.WaitElementHidden('ng-hide', 'class', False, timeout_selenium60)
        selenium.ClickElement('Refresh', 'span', 'title')
        sleep(sleep3)
        
        selenium.WaitElement('Remove all notifications', 'footer', True, timeout_selenium60)
        selenium.ClickElement('Remove all notifications', 'footer', True)
        sleep(sleep3)
        CommonFunction.WriteLog(f"Finish cleanup existing notif.")
    except Exception as err:
        err = str(err)
        CommonFunction.WriteLog(f"Failed cleanup existing notif. error: {err}")

def nakisa_generate_active_group_list():
    CommonFunction.WriteLog(f"Start Generate Active Group List")
    
    SeleniumLib.screenshot(locator=None, filename=None)       
        
    try:
        removeNotif()
    except:
        sleep(sleep5)
    
    try:
        sleep(sleep3)
        selenium.WaitElement('Search in Activation Groups', 'a', True, timeout_selenium60)
        sleep(sleep3)
        selenium.ClickElement('Search in Activation Groups', 'a', True)
        sleep(sleep3)
    except Exception as errWait:
        CommonFunction.WriteLog(f"[WARNING] Failed to click Search in Activation Groups with {str(errWait)}")
    
    selenium.ClickButton('Export to Excel', 'title', True)
    sleep(sleep3)

    try:
        selenium.WaitElement('Generate', 'button', True, timeout_selenium60)
        selenium.ClickElement('Generate', 'button', True)
        sleep(sleep3)
    except Exception as errBtn:
        CommonFunction.WriteLog(f"[WARNING] Failed to click Generate with {str(errBtn)}")
        sleep(sleep3)
    
    selenium.WaitElement('ng-scope', 'class', False, timeout_selenium60)
    selenium.ClickElement('ng-scope', 'class', False)
    sleep(sleep3)

    selenium.WaitElementHidden('ng-hide', 'class', False, timeout_selenium60)
    selenium.ClickElement('Refresh', 'span', 'title')
    sleep(sleep3)
    
    selenium.WaitElement('Excel Export', 'span', True, timeout_selenium60)
    selenium.ClickElement('Excel Export', 'span', True)
    sleep(sleep10)
    CommonFunction.WriteConsole(f"Click export NAKISA Active Group List.")            

def nakisa_generate_unit_List():
    CommonFunction.WriteLog(f"Start Generate Unit List")
    
    SeleniumLib.screenshot(locator=None, filename=None)  
    
    try:
        removeNotif()
    except:
        sleep(sleep5)
    
    try:
        selenium.WaitElement('Search in Units', 'a', True, timeout_selenium60)
        selenium.ClickElement('Search in Units', 'a', True)
        sleep(sleep3)
    except Exception as errWait:
        CommonFunction.WriteLog(f"[WARNING] Failed to click Search in Units with {str(errWait)}")
    
    selenium.ClickButton('Export to Excel', 'title', True)
    sleep(sleep3)

    try:
        selenium.WaitElement('Generate', 'button', True, timeout_selenium60)
        selenium.ClickElement('Generate', 'button', True)
        sleep(sleep3)
    except Exception as errBtn:
        CommonFunction.WriteLog(f"[WARNING] Failed to click Generate with {str(errBtn)}")
        sleep(sleep3)
    
    selenium.WaitElement('ng-scope', 'class', False, timeout_selenium60)
    selenium.ClickElement('ng-scope', 'class', False)
    sleep(sleep3)

    selenium.WaitElementHidden('ng-hide', 'class', False, timeout_selenium60)
    selenium.ClickElement('Refresh', 'span', 'title')
    sleep(sleep3)
    
    selenium.WaitElement('Excel Export', 'span', True, timeout_selenium60)
    selenium.ClickElement('Excel Export', 'span', True)
    sleep(sleep10)
    CommonFunction.WriteConsole(f"Click export NAKISA Unit List.")       

def OpenDisclosureReports(Name, code, monthPeriod, yearReport, generate, LeaseReconSystem, iteration, startDate, endDate):
        
    try:
        selenium.WaitElement('_navbar-menu', 'class', False, timeout_selenium60)
        selenium.ClickElement('_navbar-menu', 'class', False)
        sleep(sleep3)

        selenium.WaitElement('Disclosure Reports', 'div', True, timeout_selenium60)
        selenium.ClickElement('Disclosure Reports', 'div', True)
        sleep(sleep3)
    except Exception as er:
        WindowsLib.screenshot("desktop", "output\\Nakisa open disclosure reports error_"+str(Name)+"-"+str(code)+"-"+str(iteration)+".png")
        openNakisa()
        sleep(sleep10)
        WindowsLib.send_keys(keys="{CONTROL}l")
        WindowsLib.send_keys(keys="{ENTER}")
        sleep(sleep10)
        CommonFunction.WriteLog(f"[WARNING] Failed Open Disclosure Reports with {str(er)}")   
    
    if code != None and generate == True:
        CommonFunction.WriteLog(f"Start Open Disclosure Reports.")
        CommonFunction.WriteConsole(f"Start Open Disclosure Reports.")
        
        sleep(sleep3)
        try:
            deleteDownload('Remove already downloaded file', iteration)
        except Exception as errDel:
            CommonFunction.WriteLog(f"[INFO] Failed remove already downloaded file with: {errDel}. iteration: {iteration} - Max Retry: {maxRetry}")

        selenium.WaitElement('reports_chosen', 'id', False, timeout_selenium60)
        selenium.ClickElement('reports_chosen', 'id', False)
        sleep(sleep3)
        
        if (Name == "USGAAP"):
            ReportsChosen = SeleniumLib.does_element_contain("//div[@id='reports_chosen']/a[@class='chosen-single']/span", 'ASC 842 Cash Flow Report')
            x=0
            while ReportsChosen == False:
                selenium.WaitElement('ASC 842 Cash Flow Report', 'li', True, timeout_selenium60)
                selenium.ClickElement('ASC 842 Cash Flow Report', 'li', True)
                ReportsChosen = SeleniumLib.does_element_contain("//div[@id='reports_chosen']/a[@class='chosen-single']/span", 'ASC 842 Cash Flow Report')
                x=x+1
                if x >= maxRetry:
                    CommonFunction.WriteLog(f"[ERROR] Failed select Cash Flow USGAAP report return: {ReportsChosen}. iteration: {iteration} - Max Retry: {maxRetry}")
                    break       

        elif (Name == "IFRS"):
            ReportsChosen = SeleniumLib.does_element_contain("//div[@id='reports_chosen']/a[@class='chosen-single']/span", 'IFRS 16 Cash Flow Report')
            x=0
            while ReportsChosen == False:
                selenium.WaitElement('IFRS 16 Cash Flow Report', 'li', True, timeout_selenium60)
                selenium.ClickElement('IFRS 16 Cash Flow Report', 'li', True)
                ReportsChosen = SeleniumLib.does_element_contain("//div[@id='reports_chosen']/a[@class='chosen-single']/span", 'IFRS 16 Cash Flow Report')
                x=x+1
                if x >= maxRetry:
                    CommonFunction.WriteLog(f"[ERROR] Failed select Cash Flow IFRS report return: {ReportsChosen}. iteration: {iteration} - Max Retry: {maxRetry}")
                    break     

        elif (Name == "SLANUSGAAP" or Name == "SLANIFRS"):
            ReportsChosen = SeleniumLib.does_element_contain("//div[@id='reports_chosen']/a[@class='chosen-single']/span", 'Lease Liability Report')
            x=0
            while ReportsChosen == False:
                selenium.WaitElement('Lease Liability Report', 'li', True, timeout_selenium60)
                selenium.ClickElement('Lease Liability Report', 'li', True)
                ReportsChosen = SeleniumLib.does_element_contain("//div[@id='reports_chosen']/a[@class='chosen-single']/span", 'Lease Liability Report')
                x=x+1
                if x >= maxRetry:
                    CommonFunction.WriteLog(f"[ERROR] Failed select Liability Report return: {ReportsChosen}. iteration: {iteration} - Max Retry: {maxRetry}")
                    break

        sleep(sleep10)

        selenium.WaitElement('systems_chosen', 'id', False, timeout_selenium60)
        SystemChosen = SeleniumLib.does_element_contain("//div[@id='systems_chosen']/ul[@class='chosen-choices']/li[@class='search-choice']/span",LeaseReconSystem)
        x = 0
        while SystemChosen == False:
            SeleniumLib.input_text_when_element_is_visible("//div[@id='systems_chosen']/ul[@class='chosen-choices']/li[@class='search-field']/input", LeaseReconSystem)
            sleep(sleep3)
            selenium.ClickElement(LeaseReconSystem, 'li', True)
            WindowsLib.send_keys(keys="{ENTER}")
            SystemChosen = SeleniumLib.does_element_contain("//div[@id='systems_chosen']/ul[@class='chosen-choices']/li[@class='search-choice']/span",LeaseReconSystem)
            x=x+1
            if x >= maxRetry:
                CommonFunction.WriteLog(f"[ERROR] Failed select system return: {SystemChosen}. iteration: {iteration} - Max Retry: {maxRetry}")
                break       
        sleep(sleep10)
        
        selenium.WaitElement('companies_chosen', 'id', False, timeout_selenium60)
        CompanyCodeChosen = SeleniumLib.does_element_contain("//div[@id='companies_chosen']/ul[@class='chosen-choices']/li[@class='search-choice']/span",code)
        x = 0
        while CompanyCodeChosen == False:
            sleep(sleep3)
            selenium.ClickElement('companies_chosen', 'id', False)
            SeleniumLib.input_text_when_element_is_visible("//div[@id='companies_chosen']/ul[@class='chosen-choices']/li[@class='search-field']/input", code)
            SeleniumLib.wait_until_element_contains("//div[@id='companies_chosen']/div[@class='chosen-drop']/ul[@class='chosen-results']/li[@class='active-result highlighted']",code,timeout_selenium60)
            WindowsLib.send_keys(keys="{ENTER}")
            CompanyCodeChosen = SeleniumLib.does_element_contain("//div[@id='companies_chosen']/ul[@class='chosen-choices']/li[@class='search-choice']/span",code)
            x=x+1
            if x >= maxRetry:
                CommonFunction.WriteLog(f"[ERROR] Failed select company code return: {CompanyCodeChosen}. iteration: {iteration} - Max Retry: {maxRetry}")
                break  
        sleep(sleep10)

        selenium.WaitElement('currency_chosen', 'id', False, timeout_selenium60)
        selenium.ClickElement('currency_chosen', 'id', False)
        CurrChosen = SeleniumLib.does_element_contain("//div[@id='currency_chosen']/a[@class='chosen-single']/span", '10 -')
        x=0
        while CurrChosen == False:
            selenium.WaitElement('10 -', 'li', True, timeout_selenium60)
            sleep(sleep3)
            selenium.ClickElement('10 -', 'li', True)
            CurrChosen = SeleniumLib.does_element_contain("//div[@id='currency_chosen']/a[@class='chosen-single']/span", '10 -')
            x=x+1
            if x >= maxRetry:
                CommonFunction.WriteLog(f"[ERROR] Failed select currency return: {CurrChosen}. iteration: {iteration} - Max Retry: {maxRetry}")
                break 
        sleep(sleep3)
        
        selenium.WaitElement('standard_chosen', 'id', False, timeout_selenium60)
        selenium.ClickElement('standard_chosen', 'id', False)
        if (Name == "USGAAP" or Name == "SLANUSGAAP"):
            StandardChosen = SeleniumLib.does_element_contain("//div[@id='standard_chosen']/a[@class='chosen-single']/span", 'ASC 842')
            x=0
            while StandardChosen == False:
                selenium.WaitElement('ASC 842', 'li', True, timeout_selenium60)
                selenium.ClickElement('ASC 842', 'li', True)
                StandardChosen = SeleniumLib.does_element_contain("//div[@id='standard_chosen']/a[@class='chosen-single']/span", 'ASC 842')
                x=x+1
                if x >= maxRetry:
                    CommonFunction.WriteLog(f"[ERROR] Failed select USGAAP standard return: {StandardChosen}. iteration: {iteration} - Max Retry: {maxRetry}")
                    break 
            
        elif (Name == "IFRS" or Name == "SLANIFRS"):
            StandardChosen = SeleniumLib.does_element_contain("//div[@id='standard_chosen']/a[@class='chosen-single']/span", 'IFRS 16')
            x=0
            while StandardChosen == False:
                selenium.WaitElement('IFRS 16', 'li', True, timeout_selenium60)
                selenium.ClickElement('IFRS 16', 'li', True)
                StandardChosen = SeleniumLib.does_element_contain("//div[@id='standard_chosen']/a[@class='chosen-single']/span", 'IFRS 16')
                x=x+1
                if x >= maxRetry:
                    CommonFunction.WriteLog(f"[ERROR] Failed select IFRS standard return: {StandardChosen}. iteration: {iteration} - Max Retry: {maxRetry}")
                    break 
        sleep(sleep10)

        selenium.WaitElement('year_chosen', 'id', False, timeout_selenium60)
        sleep(sleep3)
        selenium.ClickElement('year_chosen', 'id', False)
        sleep(sleep3)
        selenium.WaitElement(yearReport, 'li', True, timeout_selenium60)
        sleep(sleep3)
        selenium.ClickElement(yearReport, 'li', True)
        sleep(sleep10)

        selenium.WaitElement('periodStart_chosen', 'id', False, timeout_selenium60)
        selenium.ClickElement('periodStart_chosen', 'id', False)
        PeriodChosenStart = SeleniumLib.get_webelements(locator='//div[@id="periodStart_chosen"]/div[@class="chosen-drop"]/ul[@class="chosen-results"]/li[@data-option-array-index]')
        SeleniumLib.press_keys(f'//div[@id="periodStart_chosen"]/div[@class="chosen-drop"]/ul[@class="chosen-results"]/li[@data-option-array-index="0"]',"ARROW_UP")
        for i in range(len(PeriodChosenStart)):
            SeleniumLib.press_keys(f'//div[@id="periodStart_chosen"]/div[@class="chosen-drop"]/ul[@class="chosen-results"]/li[@data-option-array-index="{i}"]',"ARROW_DOWN")
            DataPeriod = SeleniumLib.get_text(locator=f'//div[@id="periodStart_chosen"]/div[@class="chosen-drop"]/ul[@class="chosen-results"]/li[@data-option-array-index="{i}"]')
            month = DataPeriod[:-3]
            if month == monthPeriod:
                SeleniumLib.click_element(locator=f'//div[@id="periodStart_chosen"]/div[@class="chosen-drop"]/ul[@class="chosen-results"]/li[@data-option-array-index="{i}"]')
                break

        selenium.WaitElement('periodEnd_chosen', 'id', False, timeout_selenium60)
        selenium.ClickElement('periodEnd_chosen', 'id', False)
        PeriodChosenEnd = SeleniumLib.get_webelements(locator='//div[@id="periodEnd_chosen"]/div[@class="chosen-drop"]/ul[@class="chosen-results"]/li[@data-option-array-index]')
        SeleniumLib.press_keys(f'//div[@id="periodEnd_chosen"]/div[@class="chosen-drop"]/ul[@class="chosen-results"]/li[@data-option-array-index="0"]',"ARROW_UP")
        for i in range(len(PeriodChosenEnd)):
            SeleniumLib.press_keys(f'//div[@id="periodEnd_chosen"]/div[@class="chosen-drop"]/ul[@class="chosen-results"]/li[@data-option-array-index="{i}"]',"ARROW_DOWN")
            DataPeriod = SeleniumLib.get_text(locator=f'//div[@id="periodEnd_chosen"]/div[@class="chosen-drop"]/ul[@class="chosen-results"]/li[@data-option-array-index="{i}"]')
            month = DataPeriod[:-3]
            if month == monthPeriod:
                SeleniumLib.click_element(locator=f'//div[@id="periodEnd_chosen"]/div[@class="chosen-drop"]/ul[@class="chosen-results"]/li[@data-option-array-index="{i}"]')
                break

        selenium.WaitElement('agStatus_chosen', 'id', False, sleep3)
        StatusChosen = SeleniumLib.does_element_contain("//div[@id='agStatus_chosen']/ul[@class='chosen-choices']/li[@class='search-choice']/span",'Active')
        x = 0
        while StatusChosen == False:
            SeleniumLib.input_text_when_element_is_visible("//div[@id='agStatus_chosen']/ul[@class='chosen-choices']/li[@class='search-field']/input", 'Active')
            selenium.ClickElement('Active', 'li', True)
            WindowsLib.send_keys(keys="{ENTER}")
            StatusChosen = SeleniumLib.does_element_contain("//div[@id='agStatus_chosen']/ul[@class='chosen-choices']/li[@class='search-choice']/span",'Active')
            x=x+1
            if x >= maxRetry:
                CommonFunction.WriteLog(f"[ERROR] Failed select status return: {StatusChosen}. iteration: {iteration} - Max Retry: {maxRetry}")
                break    
                        
        selenium.WaitElement('Generate', 'Button', True, timeout_selenium60)
        selenium.ClickElement('Generate', 'Button', True)
        sleep(sleep10)
        CommonFunction.WriteConsole(f"Nakisa click generate button.")

def waitReports(Name, code, company, iteration):
    listName = re.sub('[.,]', '', company).split()
    CommonFunction.WriteLog(f"Start waiting download Disclosure Reports Files.")
    CommonFunction.WriteConsole(f"Start waiting download Disclosure Reports Files.")
    
    ReportStatusElement = "None"
    ReportRefreshElement = "None"
    x = 0
                
    try:
        sleep(sleep5)
        OpenDisclosureReports(None, None, None, None, False, None, iteration, None, None)
        sleep(sleep5) 
    except:
        sleep(sleep3)
    
    while x < maxWait:
        try:
            ReportStatusLocator = '//div[@class="report-list"]/div[@class="container-fluid"]/div[@class="launched-report"][1]/div[@class="report-info report-column"]/table[@class="report-status"]/tbody/tr[2]/td[2]'
            try:
                SeleniumLib.wait_until_element_is_visible(locator=ReportStatusLocator, timeout=timeout_selenium60)
            except:
                CommonFunction.WriteConsole(f"Nakisa file not generated. Stop waiting...")
                sleep(sleep10)
                break
            
            ReportStatusElement = SeleniumLib.find_element(locator=ReportStatusLocator)
                                    
            if ReportStatusElement.text == 'Ended':
                CommonFunction.WriteConsole(f"Check status: {ReportStatusElement.text}.")
                CommonFunction.WriteLog(f"Finish waiting button download Disclosure Reports Files: {code}")
                break
            else:
                ReportRefreshLocator = '//div[@class="report-list-caption"]/span'
                SeleniumLib.wait_until_element_is_visible(locator=ReportRefreshLocator, timeout=timeout_selenium60)
                ReportRefreshElement = SeleniumLib.find_element(locator=ReportRefreshLocator)
                
                while ReportStatusElement.text != 'Ended':
                    if ReportRefreshElement.text == 'Refreshing in 5 seconds':
                        OpenDisclosureReports(None, None, None, None, False, None, iteration, None, None)
            x=x+1
        except:
            x=x+1
            OpenDisclosureReports(None, None, None, None, False, None, x, None, None)
        
def deleteDownload(fileName, iteration):
    CommonFunction.WriteLog(f"Start delete download Disclosure Reports Files: {fileName}. iteration: {iteration} - Max Retry: {maxRetry}") 
    try:
        sleep(sleep10)   
        if iteration > 0:   
            OpenDisclosureReports(None, None, None, None, False, None, iteration, None, None)
        
        ReportDeleteLocator = '//div[@class="report-list"]/div[@class="container-fluid"]/div[@class="launched-report"][1]/div[@class="report-action-logos report-column"]/i[@class="fa fa-trash-o report-action-logo  report-delete"]'
        SeleniumLib.scroll_element_into_view(locator=ReportDeleteLocator)
        SeleniumLib.wait_until_element_is_visible(ReportDeleteLocator, timeout=timeout_selenium60)
        SeleniumLib.click_element_when_visible(ReportDeleteLocator)
        
        sleep(sleep3)
        
        selenium.WaitElement('Delete Report', 'div', True, timeout_selenium60)
        selenium.ClickElement('Delete Report', 'div', True)
        SeleniumLib.wait_until_element_is_not_visible(ReportDeleteLocator, timeout=timeout_selenium60)
        
        sleep(sleep10)
        CommonFunction.WriteLog(f"Finish delete download successfully: {fileName}.")
    except Exception as errDelete:
        errDelete = str(errDelete)
        CommonFunction.WriteLog(f"[INFO] Failed delete download Disclosure Reports Files with {errDelete}. iteration: {iteration} - Max Retry: {maxRetry}") 
    
def downloadReports(Name, code, iteration):
    CommonFunction.WriteLog(f"Start download Disclosure Reports Files.")    
    CommonFunction.WriteConsole(f"Start download Disclosure Reports Files.") 
    
    if iteration > 0:   
        OpenDisclosureReports(None, None, None, None, False, None, iteration, None, None)
        sleep(1)
            
    ReportDownloadLocator = '//div[@class="report-list"]/div[@class="container-fluid"]/div[@class="launched-report"][1]/div[@class="report-action-logos report-column"]/i[@class="fa fa-download report-action-logo report-download report-action-logo-success"]'
    SeleniumLib.scroll_element_into_view(locator=ReportDownloadLocator)
    SeleniumLib.wait_until_element_is_visible(ReportDownloadLocator, timeout=timeout_selenium60)
    if iteration == 0:
        SeleniumLib.click_element_when_visible(ReportDownloadLocator)

    Report = glob.glob(downloadPath+"\\Disclosure Reports*.xls")
    j= 0
    while len(Report) < 1:
        j = j + 1
        sleep(sleep3)
        Report = glob.glob(downloadPath+"\\Disclosure Reports.xls")
        if j > maxWait:
            CommonFunction.WriteLog(f"Reports not found, timeout reached")
            break

    if len(Report) > 0:
        if str(Name) == "SLANUSGAAP":
            fileName = "Liability "+str(code)+" USGAAP"
        elif str(Name) == "SLANIFRS":
            fileName = "Liability "+str(code)+" IFRS"
        else:
            fileName = "Cash Flow "+str(code)+" "+str(Name)
            
        if os.path.exists(downloadPath+"\\"+fileName+".xls"):
            os.remove(downloadPath+"\\"+fileName+".xls")
        os.rename(downloadPath+"\\Disclosure Reports.xls", downloadPath+"\\"+fileName+".xls")
        filePath = downloadPath+"\\"+fileName+".xls"
        CommonFunction.WriteLog(f"Finish Disclosure Reports file downloaded.")
        CommonFunction.WriteLog(f"Finish Download Disclosure Reports file downloaded.")
        
        sleep(sleep3)
        deleteDownload(fileName, iteration)

        try:
            convertHTMLtoExcel(Name, fileName, filePath)
        except Exception as er:
            error = str(er)
            CommonFunction.WriteLog(f"Failed convert HTML to Excel Disclosure Reports Files with message: {error}")
            CommonFunction.WriteLog(f"[WARNING] Failed convert HTML to Excel Disclosure Reports Files with message: {error}")
        FilePathOutput = pathResultExcel+"\\"+fileName+".xlsx"
        
        try:
            shutil.move(filePath, FilePathOutput.replace(".xlsx", ".xls"))
            CommonFunction.WriteLog(f"Name: {Name}")
            CommonFunction.WriteLog(f"path: {FilePathOutput}")    
            createExcelPath(Name, path=FilePathOutput)
        except Exception as err:
            error = str(err)
            CommonFunction.WriteLog(f"Failed set Excel Path Files with message: {error}")

def nakisa_generate_cash_flow_report(Name, code, company, System, monthPeriod, yearReport, startDate, endDate):
    try:
        CommonFunction.WriteLog(f"Start Generate Cash Flow {code}")
        CommonFunction.WriteConsole(f"Start Generate NAKISA {Name} - {code}")
        
        x = 0
        errorNakisa = False
        
        while x <= maxRetry:
            try:
                OpenDisclosureReports(Name, code, monthPeriod, yearReport, True, System, x, startDate, endDate)
                errorNakisa = False
                x=x+1
                break
            except Exception as errFilter:
                WindowsLib.screenshot("desktop", "output\\Nakisa set filter error_"+str(Name)+"-"+str(code)+"-"+str(x)+".png")
                if x == maxRetry:
                    SeleniumLib.screenshot(locator=None, filename=None)
                    CommonFunction.WriteConsole(f"[ERROR] Nakisa failed to set filter with error: {str(errFilter)}")
                    CommonFunction.WriteLog(f"[ERROR] Nakisa failed to set filter with error: {str(errFilter)}")
                    errorNakisa = True
                    break
                x=x+1
                sleep(sleep10)
                
        if errorNakisa == False:
            x=0
            while x <= maxRetry:
                try:
                    waitReports(Name, code, company, x)
                    errorNakisa = False
                    x=x+1
                    break
                except Exception as errWait:                    
                    if x == maxRetry:
                        WindowsLib.screenshot("desktop", "output\\Nakisa waiting error_"+str(Name)+"-"+str(code)+"-"+str(x)+".png")
                        SeleniumLib.screenshot(locator=None, filename=None)
                        CommonFunction.WriteConsole(f"[ERROR] Nakisa failed to wait reports with error: {str(errWait)}. Retry: {x} - Max Retry: {maxRetry}")
                        CommonFunction.WriteLog(f"[ERROR] Nakisa failed to wait reports with error: {str(errWait)}. Retry: {x} - Max Retry: {maxRetry}")
                        errorNakisa = True
                        break
                    x=x+1 
                    OpenDisclosureReports(Name, code, monthPeriod, yearReport, True, System, x, startDate, endDate)
                    sleep(sleep10)
                                      
                
        if errorNakisa == False:
            x=0
            while x <= maxRetry:
                try:
                    downloadReports(Name, code, x)
                    CommonFunction.WriteConsole(f"Finish Disclosure Reports file downloaded.")
                    CommonFunction.WriteConsole(f"Finish Generate NAKISA.")
                    errorNakisa = False
                    x=x+1
                    break
                except Exception as errDownload:
                    WindowsLib.screenshot("desktop", "output\\Nakisa download reports error_"+str(Name)+"-"+str(code)+"-"+str(x)+".png")
                    if x == maxRetry:
                        SeleniumLib.screenshot(locator=None, filename=None)
                        CommonFunction.WriteConsole(f"[ERROR] Nakisa failed to download reports with error: {str(errDownload)}. Retry: {x} - maxRetry: {maxRetry}")
                        CommonFunction.WriteLog(f"[ERROR] Nakisa failed to download reports with error: {str(errDownload)}. Retry: {x} - maxRetry: {maxRetry}")
                        errorNakisa = True
                        break
                    x=x+1
                    OpenDisclosureReports(None, None, None, None, False, None, x, None, None)
                    sleep(sleep10)
                        
    except Exception as err2:
        errors = str(err2)
        WindowsLib.screenshot("desktop", "output\\Nakisa error_"+str(Name)+"-"+str(code)+".png")
        CommonFunction.WriteLog(f"Generate Cash Flow Report error with: {errors}")
        CommonFunction.WriteConsole(f"[ERROR] Generate Cash Flow Report error with: {errors}")

def refreshWeb():
    WindowsLib.send_keys(keys="{CONTROL}l")
    WindowsLib.send_keys(keys="{ENTER}")
    sleep(sleep10)

def openNakisa():
    CommonFunction.WriteConsole(f"Start open browser.")
    try:
        SeleniumLib.close_browser()
    except:
        sleep(sleep5)
    
    selenium.OpenWeb(URL_NAKISA)
    webTitle = SeleniumLib.get_title()
    CommonFunction.WriteConsole(f"Browser Title: {webTitle}")
    
    try:
        SeleniumLib.delete_all_cookies()
    except:
        sleep(5)
    
    try:
        selenium.InputElement('j_username', 'name', False, PMI_UserAccount+"@PMINTL.NET")
        selenium.InputElement('j_password', 'name', True, PMI_Password)
        selenium.ClickButton('submit', 'name', False)
    except:
        sleep(5)
    
    refreshWeb()
    CommonFunction.WriteConsole(f"Finish open browser.")
          

def nakisaGenerate(code, company, System, USGAAP, IFRS, monthPeriod, yearReport, startDate, endDate):
    try:
        if USGAAP == 'x' and IFRS != 'x':
            nakisa_generate_cash_flow_report('USGAAP', code, company, System, monthPeriod, yearReport, startDate, endDate)
            nakisa_generate_cash_flow_report('SLANUSGAAP', code, company, System, monthPeriod, yearReport, startDate, endDate)
        elif USGAAP == 'x' and IFRS == 'x':
            nakisa_generate_cash_flow_report('USGAAP', code, company, System, monthPeriod, yearReport, startDate, endDate)
            nakisa_generate_cash_flow_report('IFRS', code, company, System, monthPeriod, yearReport, startDate, endDate)
            nakisa_generate_cash_flow_report('SLANUSGAAP', code, company, System, monthPeriod, yearReport, startDate, endDate)
            nakisa_generate_cash_flow_report('SLANIFRS', code, company, System, monthPeriod, yearReport, startDate, endDate)
        SeleniumLib.close_browser()
        sleep(sleep3)
    except Exception as errNakisa:
        error = str(errNakisa)
        try:
            SeleniumLib.close_browser()
        except:
            sleep(1)
        CommonFunction.WriteLog(f"{code} Nakisa error with {error}")
        CommonFunction.WriteConsole(f"[ERROR] {code} Nakisa error with {error}")

def Generate():
    try:
        if not os.path.exists(LeaseReconConfigPath) or not os.path.exists(LeaseReconSchedulerPath) or not os.path.exists(LeaseReconTemplatePath):
            #next development should be send email (receiver are email setup)
            sendNotification("missing")
            CommonFunction.WriteLog(f"There is missing file. Abort processing automation.")
            return False
        else:
            today = datetime.now().strftime("%d-%m-%Y")
            
            CommonFunction.WriteLog(f"today: {today}")
            CommonFunction.WriteLog(f"LeaseReconSchedulerPath: {LeaseReconSchedulerPath}")
            excel.open_workbook(LeaseReconSchedulerPath)
            
            Worksheet = excel.read_worksheet_as_table(start=2, header= False)
            
            data = TablesLib.export_table(table=Worksheet)
            
            excel.close_workbook()
            
            CommonFunction.WriteLog(f"LeaseReconConfigPath: {LeaseReconConfigPath}")
            excel.open_workbook(LeaseReconConfigPath)
            rows2 = excel.read_worksheet('Config File', start=2)
            x=0
            errCount = 0
            totalData = len(data)
            for cell in data:
                kill_process_username(process_name = "notepad.exe")
                kill_process_username(process_name = "excel.exe")
                kill_process_username(process_name = "saplogon.exe")
                
                code = cell['A'] 
                company = cell['B'] 
                ReportDate = str(cell['C']).strip() 
                ReportDate = ReportDate.replace('00', '')
                ReportDate = ReportDate.replace(':', '')
                ReportDate = ReportDate.replace(' ', '')
                ReportDate = ReportDate.replace('"', '')
                ReportDate = ReportDate.replace("'", "")
                
                CommonFunction.WriteLog(f"Report Date: {ReportDate}")
                if ReportDate == today:
                    CommonFunction.WriteConsole(f"Scheduler Date: {ReportDate}")
                    
                    if dateEnd != "None" and dateEnd != None: 
                        ReportDate = dateEnd
                        
                    CommonFunction.WriteConsole(f"Report Date: {ReportDate}")
        
                    Date = datetime.strptime(ReportDate, "%d-%m-%Y")
                    yearReport = Date.strftime('%y') 
                    year = Date.strftime('%Y') 
                    month = Date.strftime('%m')
                    monthPeriod = Date.strftime("X%m").replace('X0','X').replace('X','')
                    
                    #get end Date for Completeness check
                    startDate = dateStart
                    endDate = dateEnd
                    if dateStart.lower() == 'none':
                        vStartDate = datetimefunc.date(int(year), int(month), 1)
                        startDate = vStartDate
                    if dateEnd.lower() == 'none':
                        vEndDate = datetimefunc.date(
                            vStartDate.year + 1 if vStartDate.month == 12 else vStartDate.year
                            , 1 if vStartDate.month == 12 else vStartDate.month + 1
                            , 1
                        ) - datetimefunc.timedelta(days = 1)
                        endDate = vEndDate.strftime("%d-%m-%Y")
                        
                    if (str(code) != 'None') or (str(code) != None):
                            
                        for cell2 in rows2:
                            code2 = cell2['A'] 
                            USGAAP = cell2['C']  
                            IFRS = cell2['D'] 
                            DA = cell2['E'] 
                            DAIFRS = cell2['F'] 
                            
                            if str(code) == str(code2) :
                                i=0
                                errBrowser = False
                                while i<=maxRetry:
                                    i=i+1
                                    try:
                                        sleep(sleep10)
                                        openNakisa()
                                        webTitle = SeleniumLib.get_title()
                                        
                                        if webTitle == '' or webTitle == ' ' or webTitle == None:
                                            errBrowser = True
                                        else:
                                            errBrowser = False
                                            break
                                    except Exception as err:
                                        error = str(err)
                                        if i==maxRetry:
                                            errBrowser = True
                                            CommonFunction.WriteLog(f"Failed to open NAKISA with message: {error}")
                                            CommonFunction.WriteConsole(f"[ERROR] Failed to open NAKISA with message: {error}")
                                            break
                                
                                CommonFunction.WriteLog(f"Start Generate: {code}")
                                
                                if errBrowser==False:
                                    excelPath.pathAGSLAN            = pathResultExcel+"/Active Group.xlsx"
                                    excelPath.pathUNITSLAN          = pathResultExcel+"/Unit List.xlsx"
                                    
                                    WindowsLib.screenshot("desktop", "output\\Open browser" + "_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".png")
                                    i = 0
                                    while i <= maxRetry and not os.path.exists(excelPath.pathAGSLAN):

                                        try:                                        
                                            CommonFunction.WriteConsole(f"Start Generate NAKISA Active Group List. iteration: {i} - Max Retry: {maxRetry}")
                                            
                                            prefixAct = "LAE_" + PMI_UserAccount
                                            found = checkFileExists(prefixAct, 'Active Group')
                                            j = 0
                                                
                                            if found == False:    
                                                if i>0:
                                                    refreshWeb()
                                                    
                                                nakisa_generate_active_group_list()
                                                CommonFunction.WriteConsole(f"Downloading Active Group List...")
                                                
                                                while found == False:
                                                    
                                                    if j == maxWait:
                                                        CommonFunction.WriteLog(f"Active Group List File not found, timeout reached")
                                                        CommonFunction.WriteConsole(f"[ERROR] Active Group List File not found, timeout reached")
                                                        break
                                                    
                                                    found = checkFileExists(prefixAct, 'Active Group')
                                                    
                                                    if found == True:
                                                        CommonFunction.WriteLog(f"Finish Generate Active Group List") 
                                                        CommonFunction.WriteConsole(f"Finish Generate Active Group List")
                                                        break
                                                    
                                                    j = j + 1                                                
                                            else:
                                                CommonFunction.WriteLog(f"Finish Generate Active Group List")
                                                CommonFunction.WriteConsole(f"Finish Generate Active Group List")   
                                                break
                                            
                                            i=i+1
                                        except Exception as errAG:
                                            if i == maxRetry:
                                                CommonFunction.WriteLog(f"[ERROR] Failed to generate active group list. Error with: {str(errAG)}. Retry: {i} MaxRetry: {maxRetry}")
                                                CommonFunction.WriteConsole(f"[ERROR] Failed to generate active group list. Error with: {str(errAG)}. Retry: {i} MaxRetry: {maxRetry}")
                                                sleep(sleep10)
                                                break
                                            i=i+1
                                            refreshWeb()
                                            sleep(sleep10)
                                            
                                    i = 0
                                    while i <= maxRetry and not os.path.exists(excelPath.pathUNITSLAN):
                                        try:
                                            CommonFunction.WriteConsole(f"Start Generate NAKISA Unit List.  iteration: {i} - Max Retry: {maxRetry}")
        
                                            prefixAct = "LAE_" + PMI_UserAccount
                                            found = checkFileExists(prefixAct, 'Unit List')
                                            j = 0
                                            
                                            if found == False:
                                                if i>0:
                                                    refreshWeb()
                                                    
                                                nakisa_generate_unit_List()
                                                CommonFunction.WriteConsole(f"Downloading Unit List...")
                                                
                                                while found == False:
                                                    
                                                    if j == maxWait:
                                                        CommonFunction.WriteLog(f"Unit List File not found, timeout reached")
                                                        CommonFunction.WriteConsole(f"[ERROR] Unit List File not found, timeout reached")
                                                        break
                                                        
                                                    found = checkFileExists(prefixAct, 'Unit List')
                                                    
                                                    if found == True:
                                                        CommonFunction.WriteLog(f"Finish Generate Unit List")
                                                        CommonFunction.WriteConsole(f"Finish Generate Unit List.")
                                                        break
                                                    
                                                    j = j + 1
                                            else:
                                                CommonFunction.WriteLog(f"Finish Generate Unit List.")
                                                CommonFunction.WriteConsole(f"Finish Generate Unit List.")
                                                break
                                            
                                            i=i+1
                                        except Exception as errUnit:
                                            if i == maxRetry:
                                                CommonFunction.WriteLog(f"[ERROR] Failed to generate unit list. Error with: {str(errUnit)}. Retry: {i} MaxRetry: {maxRetry}")
                                                CommonFunction.WriteConsole(f"[ERROR] Failed to generate unit list. Error with: {str(errUnit)}. Retry: {i} MaxRetry: {maxRetry}")
                                                sleep(sleep10)
                                                break
                                            i=i+1
                                            refreshWeb()
                                            sleep(sleep10)
                                                                    
                                    nakisaGenerate(str(code), company, LeaseReconSystem, str(USGAAP).lower(), str(IFRS).lower(), monthPeriod, yearReport, startDate, endDate)
                                else:
                                    WindowsLib.screenshot("desktop", "output\\desktop no browser" + "_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".png")
                                    CommonFunction.WriteLog(f"[ERROR] No browser is open. Skip Nakisa part.")
                                    CommonFunction.WriteConsole(f"[ERROR] No browser is open. Skip Nakisa part.")
                                    
                                i=0
                                while i<= maxRetry:
                                    CommonFunction.WriteConsole(f"Start SAP Automation - {str(code)}. Iteration: {i} - Max Retry: {maxRetry}")        
                                    try:
                                        generateSAP(str(code), DA, DAIFRS, str(USGAAP).lower(), str(IFRS).lower(), year, month, startDate, endDate)
                                        CommonFunction.WriteConsole(f"Finish SAP Automation.")
                                        i=i+1
                                        break
                                    except Exception as error:
                                        if i == maxRetry:
                                            WindowsLib.screenshot(locator="desktop", filename=f'output\\errorSAP_{str(code)}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg')
                                            CommonFunction.WriteLog(f"Generate SAP error.")
                                            break 
                                        i=i+1
                                        CommonFunction.WriteLog(f"Generate SAP error with {str(error)}")
                                        CommonFunction.WriteConsole(f"[ERROR] Generate SAP error with {str(error)}")
                                        try:
                                            kill_process_username(process_name = "saplogon.exe")
                                        except Exception as errKill:
                                            CommonFunction.WriteLog(f"failed to kill SAP error with {str(errKill)}")
                                        sleep(sleep10)
                                    
                                    
                                CommonFunction.WriteLog(f"Finish Generate: {code}")
                                
                                CommonFunction.WriteLog(f"Start Copy Data Company: {code}")
                                
                                if str(USGAAP).lower() == 'x':
                                    excelPath.pathSAPFBL3NUSGAAP        =   pathResultSAP+"\\SAP "+str(code)+" FBL3N US GAAP.XLSX"
                                    excelPath.pathSAPFBL3NUSGAAPHTML    =   pathResultSAP+"\\SAP "+str(code)+" FBL3N US GAAP.HTML"
                                    convertMHTMLtoExcel(FileName=str(code)+" FBL3N US GAAP HTML", FilePath= excelPath.pathSAPFBL3NUSGAAPHTML, FilePathOutput=excelPath.pathSAPFBL3NUSGAAP)

                                if DA != None:
                                    excelPath.pathSAPDeprUSGAAP = pathResultSAP+"\\SAP "+str(code)+" Depr Simulation US GAAP.XLSX"
                                    
                                if str(IFRS).lower() == 'x':
                                    excelPath.pathSAPFBL3NIFRS      =   pathResultSAP+"\\SAP "+str(code)+" FBL3N IFRS.XLSX"
                                    excelPath.pathSAPFBL3NIFRSHTML  =   pathResultSAP+"\\SAP "+str(code)+" FBL3N IFRS.HTML"
                                    convertMHTMLtoExcel(FileName=str(code)+" FBL3N IFRS HTML", FilePath= excelPath.pathSAPFBL3NIFRSHTML, FilePathOutput=excelPath.pathSAPFBL3NIFRS)

                                if DAIFRS != None:
                                    excelPath.pathSAPDeprIFRS = pathResultSAP+"\\SAP "+str(code)+" Depr Simulation IFRS.XLSX"

                                excelPath.pathAGSLAN            = pathResultExcel+"/Active Group.xlsx"
                                excelPath.pathUNITSLAN          = pathResultExcel+"/Unit List.xlsx"
                                excelPath.pathUSGAAP            = pathResultExcel+"/Cash Flow "+str(code)+" USGAAP.xlsx"
                                excelPath.pathIFRS              = pathResultExcel+"/Cash Flow "+str(code)+" IFRS.xlsx"
                                excelPath.pathLiabilityUSGAAP   = pathResultExcel+"/Liability "+str(code)+" USGAAP.xlsx"
                                excelPath.pathLiabilityIFRS     = pathResultExcel+"/Liability "+str(code)+" IFRS.xlsx"
                                excelPath.pathSAPDeprUSGAAP     = pathResultSAP+"/SAP "+str(code)+" Depr Simulation US GAAP.XLSX"
                                excelPath.pathSAPDeprIFRS       = pathResultSAP+"/SAP "+str(code)+" Depr Simulation IFRS.XLSX"
                                excelPath.pathSAPAssetTrans     = pathResultSAP+"/SAP "+str(code)+" Asset Transaction.txt"
                                excelPath.pathSAPAssetTransFix  = pathResultSAP+"/SAP "+str(code)+" Asset Transaction.XLSX"
                                #process copy data and completeness 
                                Excel.ProcessExcelLeaseRecon(excelPath.pathAGSLAN, excelPath.pathUNITSLAN, excelPath.pathUSGAAP, excelPath.pathIFRS, excelPath.pathLiabilityUSGAAP, excelPath.pathLiabilityIFRS, excelPath.pathSAPFBL3NUSGAAP, excelPath.pathSAPFBL3NUSGAAPHTML, excelPath.pathSAPFBL3NIFRS, excelPath.pathSAPFBL3NIFRSHTML, excelPath.pathSAPDeprUSGAAP, excelPath.pathSAPDeprIFRS, excelPath.pathSAPAssetTrans, excelPath.pathSAPAssetTransFix, str(code), USGAAP, IFRS, ReportDate, str(endDate), month, year)

                                CommonFunction.WriteLog(f"Finish Copy Data Company: {code}")   
                                x=x+1  
                
                else:
                    errCount+=1  
            if errCount == totalData:
                sendNotification("schedule")
                CommonFunction.WriteLog(f"There is no schedule for today. Abort processing automation.")
                CommonFunction.WriteConsole(f"[Warning]There is no schedule for today. Abort processing automation.")
                return False
            else:
                return True
                    
    except Exception as errors:
        error = str(errors)
        CommonFunction.WriteLog(f"Generate failed with {error}")