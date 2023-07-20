import datetime
import pyperclip
import os

from SapGuiLibrary import SapGuiLibrary
from variables import DesktopLib, InstanceFileSystem, maxRetry, maxWaitSAP, pathResultSAP, sapServer, WindowsLib
from variables import sapAssetClass, sapAssetClassMethod
from variables import sapDisplayVarientIFRS, sapDisplayVarientUSGAAP, sapGLAccountIFRS, sapGLAccountUSGAAP
from variables import sapLayout, sapSortVarient, sapTransactionType, sleep10, sleep30, sapPath, totalClosePopupSAP, maxWaitLoginSAP
from Excel import Excel as excelFunc
from function import CommonFunction, kill_process_username
from time import sleep

def SAPInsertMultiSingleValuesString(ParamSAPLib, ParamList, ParamListMethod):
    try:
        #delete entire selection
        try:
            ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[16]")
        except Exception as tryDelete:
            CommonFunction.WriteLog(f"SAP Failed delete entire selection: {str(tryDelete)}")
            WindowsLib.screenshot(locator="desktop", filename=f'output\\errorSAP_SAPInsertMultiSingleValuesStringDelete_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg')                                
            
            
        iteration = 1
        accounts = ''
        
        #modify from list to row
        if ParamListMethod.lower() == "list ranges":
            ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[1]/usr/tabsTAB_STRIP/tabpINTL")
            
            for account in ParamList:
                if iteration %2 == 0:
                    accounts = accounts +"&&" + account+os.linesep
                else:
                    accounts = accounts + account
                iteration=iteration+1
        else: 
            for account in ParamList:
                if iteration >1 and iteration <= len(ParamList):
                    accounts = accounts +os.linesep + account
                else:
                    accounts = accounts + account
                iteration=iteration+1
                
        pyperclip.copy(accounts)
        ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[24]")
    except Exception as err:
        CommonFunction.WriteLog(f"SAP Failed insert multi value with error: {str(err)}")
        WindowsLib.screenshot(locator="desktop", filename=f'output\\errorSAP_SAPInsertMultiSingleValuesStringDelete_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg')                                
            
def GetSAPBeginEndDateYearMonth(ParamYear, ParamMonth):
    vReturnBeginDate = ""
    vReturnEndDate = ""
    vStartDate = datetime.date(ParamYear, ParamMonth, 1)
    vEndDate = datetime.date(
        vStartDate.year + 1 if vStartDate.month == 12 else vStartDate.year
        , 1 if vStartDate.month == 12 else vStartDate.month + 1
        , 1
    ) - datetime.timedelta(days = 1)
    vReturnBeginDate = vStartDate.strftime("%d.%m.%Y")
    vReturnEndDate = vEndDate.strftime("%d.%m.%Y")
    return vReturnBeginDate, vReturnEndDate


def PerformSAPFB3LN(ParamInstanceFileSystem, ParamSAPLib
, ParamSAPGLAccountList, ParamSAPCompanyCode, ParamSAPBeginDate, ParamSAPEndDate, ParamSAPLayout, ParamSAPAssetClassMethod
, ParamSAPDownloadPath, ParamSAPDownloadFile, ParamFileCheckRetry, ParamFileCheckSleep, ParamFlag
):
    try:
        CommonFunction.WriteLog(f"PerformSAPFB3LN {ParamSAPCompanyCode}")
        CommonFunction.WriteConsole(f"Start SAP FBL3N - {ParamSAPCompanyCode}")
        
        ParamSAPLib.run_transaction("FBL3N")
        
        err = 0
        errFmt = 0
        
        ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH")
        
        SAPInsertMultiSingleValuesString(
            ParamSAPLib, ParamSAPGLAccountList, 'None'
        )
        
        ParamSAPLib.send_vkey(vkey_id = 8) # Press F8
        sleep(maxWaitSAP)
        
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtSD_BUKRS-LOW"
            , text = ParamSAPCompanyCode
        )
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtSO_BUDAT-LOW"
            , text = ParamSAPBeginDate
        )
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtSO_BUDAT-HIGH"
            , text = ParamSAPEndDate
        )
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtPA_VARI"
            , text = ParamSAPLayout
        )

        ParamSAPLib.select_radio_button(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/radX_AISEL"
        )

        ParamSAPLib.send_vkey(vkey_id = 8) # Press F8
        sleep(maxWaitSAP)
        #try to get error message
        try:
            val = ParamSAPLib.get_value("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]")
            idx = val.index('No')
            if idx >=0 :
                CommonFunction.WriteLog(f"Generate SAP FB3LN {str(ParamFlag)} error with {str(val)}")
                err = 1
        except Exception as error:
            error = str(error)
            CommonFunction.WriteLog(f"Try get SAP message error with {error}")
            
        CommonFunction.WriteLog(f"try SAP err: {err}")
        if err == 0:
            ParamSAPLib.send_vkey(vkey_id = 16) # Press Shift + F4
            ParamSAPLib.select_radio_button(
                element_id = "/app/con[0]/ses[0]/wnd[1]/usr/radRB_OTHERS"
            )
            
            try:
                ParamSAPLib.select_from_list_by_label(
                    element_id = "/app/con[0]/ses[0]/wnd[1]/usr/cmbG_LISTBOX"
                    , value = "Excel - Office Open XML Format (XLSX)"
                )
                errFmt = errFmt -1
                
            except Exception as e:
                err = str(e)
                
                errFmt = errFmt +1
                try:
                    ParamSAPLib.select_from_list_by_label(
                        element_id = "/app/con[0]/ses[0]/wnd[1]/usr/cmbG_LISTBOX"
                        , value = "Excel (in Office 2007 XLSX Format)"
                    )
                    errFmt = errFmt -1
                    
                except Exception as e2:
                    err2 = str(e2)
                    
                    errFmt = errFmt +1
            
            if errFmt > 0:
                CommonFunction.WriteLog(f"PerformSAPFB3LN Select from All Available Formats failed to find Excel Format.")
            else:
                ParamSAPLib.send_vkey(vkey_id = 0) # Press Enter
                ParamSAPLib.input_text(
                    element_id = "/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_PATH"
                    , text = ParamSAPDownloadPath
                )
                ParamSAPLib.input_text(
                    element_id = "/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_FILENAME"
                    , text = ParamSAPDownloadFile
                )
                ParamSAPLib.send_vkey(vkey_id = 0) # Press Enter

                SAPDownloadFilePath = ParamSAPDownloadPath + "\\" + ParamSAPDownloadFile

                # Retry until file exist in Folder
                vTempRetry = 1
                while(ParamInstanceFileSystem.does_file_not_exist(path = SAPDownloadFilePath) and vTempRetry < ParamFileCheckRetry):
                    CommonFunction.WriteLog(f"Failed when Check SAP File {ParamSAPDownloadFile} - Attempt # {str(vTempRetry)}")
                    sleep(ParamFileCheckSleep)
                    vTempRetry = vTempRetry + 1

                if(ParamInstanceFileSystem.does_file_not_exist(path = SAPDownloadFilePath)):
                    CommonFunction.WriteLog(f"Finish Automation since File SAP {ParamSAPDownloadFile} cannot be generated")
                    # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Connection issue:  File SAP " + ParamSAPDownloadFile + " cannot be generated.")
                else:
                    # Kill Excel Process that is opening
                    kill_process_username(process_name = "notepad.exe")
                    kill_process_username(process_name = "excel.exe")
                
                sleep(sleep10)
                ParamSAPLib.send_vkey(vkey_id = 9) # Press F9
                sleep(5)
                
                #Try to remove popup
                try:
                    WindowsLib.screenshot(locator="desktop", filename=f'output\\SAPFBL3N_Click_F9_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg')  
                    CommonFunction.WriteLog("Start try close popup in SAP FBL3N")
                    WindowsLib.send_keys(keys="C")
                    sleep(5)
                    CommonFunction.WriteLog("Finish try close popup in SAP FBL3N")
                except:
                    CommonFunction.WriteLog("SAP PopUp doesn't appear in SAP FBL3N")
                    
                sleep(sleep10)
                                       
                ParamSAPLib.select_radio_button(element_id ="/app/con[0]/ses[0]/wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]")
                ParamSAPLib.send_vkey(vkey_id = 0) # Press Enter
                
                ParamHTML = ParamSAPDownloadFile.replace(".XLSX", ".HTML").strip()
                ParamSAPLib.send_vkey(vkey_id = 0) # Press Enter
                ParamSAPLib.input_text(
                    element_id = "/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_PATH"
                    , text = ParamSAPDownloadPath
                )
                ParamSAPLib.input_text(
                    element_id = "/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_FILENAME"
                    , text = ParamHTML
                )
                ParamSAPLib.send_vkey(vkey_id = 0) # Press Enter

                SAPDownloadFilePath = ParamSAPDownloadPath + "\\" + ParamHTML

                # Retry until file exist in Folder
                vTempRetry = 1
                while(ParamInstanceFileSystem.does_file_not_exist(path = SAPDownloadFilePath) and vTempRetry < ParamFileCheckRetry):
                    CommonFunction.WriteLog(f"Failed when Check SAP File {ParamSAPDownloadFile} - Attempt # {str(vTempRetry)}")
                    sleep(ParamFileCheckSleep)
                    vTempRetry = vTempRetry + 1

                if(ParamInstanceFileSystem.does_file_not_exist(path = SAPDownloadFilePath)):
                    CommonFunction.WriteLog(f"Finish Automation since File SAP FBL3N HTML cannot be generated")
                    # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Connection issue:  File SAP SAP FBL3N HTML cannot be generated.")
                else:
                    # Kill Excel Process that is opening
                    kill_process_username(process_name = "notepad.exe")
                    kill_process_username(process_name = "excel.exe")
                    # Press Shift F3 twice to close the FB3LN Page and return to main menu
                    ParamSAPLib.send_vkey(vkey_id = 15) # Press Shift F3
                    ParamSAPLib.send_vkey(vkey_id = 15) # Press Shift F3
                    CommonFunction.WriteConsole(f"Finish SAP FBL3N.")
                sleep(sleep10)
    except Exception as error:
        CommonFunction.WriteLog(f"Generate SAP FB3LN error with: {error}")
         
        try:
            WindowsLib.screenshot(locator="desktop", filename=f'output\\SAPFBL3N_Error_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg') 
            fiscal = ParamSAPLib.get_value("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]")
            idx = fiscal.index('Fiscal')
            if idx >= 0 :
                CommonFunction.WriteLog(f"Generate SAP FB3LN "+str(ParamFlag) +" error Fiscal year change not yet made.")
                # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Generate SAP FB3LN "+str(ParamFlag) +" error Fiscal year change not yet made.")
                ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]")
            else:
                val = fiscal
                CommonFunction.WriteLog(f"val: {str(val)}")
                if val =='' or val == ' ' or val ==None:
                    CommonFunction.WriteLog(f"Generate SAP FB3LN "+str(ParamFlag) +" error Connection issue")
                else:
                    CommonFunction.WriteLog(f"Generate SAP FB3LN {str(ParamFlag)} error with {str(val)}")
                # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Connection issue:  Generate SAP FB3LN "+str(ParamFlag) +" error")
            ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]")
            CommonFunction.WriteLog(f"Fiscal year change not yet made.")
        except Exception as error1:
            CommonFunction.WriteLog(f"Generate SAP FB3LN error with: {str(error1)}")
            try:
                noData = ParamSAPLib.get_value("/app/con[0]/ses[0]/wnd[1]/usr/txtMESSTXT1")
                if noData == 'No data was selected':
                    CommonFunction.WriteLog(f"Generate SAP FB3LN "+str(ParamFlag) +" error data was not found.")
                    # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Does not exist: Generate SAP FB3LN "+str(ParamFlag) +" error data was not found.")
                    ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]")
            except Exception as error2:
                CommonFunction.WriteLog(f"Generate SAP FB3LN error with: {str(error2)}")
                 
                try:
                    val = ParamSAPLib.get_value("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]")
                    CommonFunction.WriteLog(f"val: {str(val)}")
                    if val =='' or val == ' ' or val ==None:
                        # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Other Reason: Generate SAP FB3LN "+str(ParamFlag) +" error failed to get element.")
                        CommonFunction.WriteLog(f"Generate SAP FB3LN {str(ParamFlag)} error failed to get element.")
                    else:
                        # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Other Reason: Generate SAP FB3LN "+str(ParamFlag) +" error with "+str(val))
                        CommonFunction.WriteLog(f"Generate SAP FB3LN {str(ParamFlag)} error with {str(val)}")
                except Exception as error3:
                    # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Generate SAP FB3LN "+str(ParamFlag) +" error failed to get element.")
                    CommonFunction.WriteLog(f"Generate SAP FB3LN error with: {str(error3)}")
                     
                    sleep(sleep10)
        ParamSAPLib.send_vkey(vkey_id = 15) # Press Shift F3
            
def PerformSAPDepreciationSimulation(ParamInstanceFileSystem, ParamSAPLib
, ParamSAPCompanyCode, ParamSAPAssetClassMethod, ParamAssetClass, ParamSAPReportYear, ParamSAPDepreciationArea, ParamSAPLayout
, ParamSAPDownloadPath, ParamSAPDownloadFile, ParamFileCheckRetry, ParamFileCheckSleep, ParamFlag
):
    try:
        CommonFunction.WriteLog(f"PerformSAPDepreciationSimulation {ParamSAPCompanyCode}")
        CommonFunction.WriteConsole(f"Start SAP Depreciation Simulation - {ParamSAPCompanyCode}")
        
        ParamSAPLib.run_transaction("S_ALR_87012936")
        
        err = 0
        errFmt = 0
        
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtBUKRS-LOW"
            , text = ParamSAPCompanyCode
        )
        if ParamSAPAssetClassMethod.lower() == 'ranges':
            ParamSAPLib.input_text(
                element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtSO_ANLKL-LOW"
                , text = ParamAssetClass[0]
            )
            ParamSAPLib.input_text(
                element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtSO_ANLKL-HIGH"
                , text = ParamAssetClass[1]
            )
            
        else:
            ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[0]/usr/btn%_SO_ANLKL_%_APP_%-VALU_PUSH")
                
            SAPInsertMultiSingleValuesString(
                ParamSAPLib, ParamAssetClass, ParamSAPAssetClassMethod.lower()
            )

            ParamSAPLib.send_vkey(vkey_id = 8) # Press F8
            sleep(maxWaitSAP)    

        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtBERDATUM"
            , text = "31.12." + str(ParamSAPReportYear)
        )
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtBEREICH1"
            , text = ParamSAPDepreciationArea
        )
        ParamSAPLib.send_vkey(vkey_id = 19) # Press SHIFT + F7
        ParamSAPLib.select_radio_button(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/radP_MONTH"
        )
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtP_VARI"
            , text = ParamSAPLayout
        )
        ParamSAPLib.send_vkey(vkey_id = 8) # Press F8
        sleep(maxWaitSAP)
        
        #try to get error message
        try:
            val = ParamSAPLib.get_value("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]")
            idx = val.index('No')
            if idx >=0 :
                # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Other Reason: Generate Depreciation Simulation "+str(ParamFlag) +" error with "+str(val))
                CommonFunction.WriteLog(f"Generate Depreciation Simulation {str(ParamFlag)} error with {str(val)}")
                err = 1
        except Exception as error:
            error = str(error)
            CommonFunction.WriteLog(f"Try get SAP message error with {error}")
        
        CommonFunction.WriteLog(f"try SAP err: {err}")
        if err == 0:
            ParamSAPLib.send_vkey(vkey_id = 16) # Press Shift + F4
            ParamSAPLib.select_radio_button(
                element_id = "/app/con[0]/ses[0]/wnd[1]/usr/radRB_OTHERS"
            )
            
            try:
                ParamSAPLib.select_from_list_by_label(
                    element_id = "/app/con[0]/ses[0]/wnd[1]/usr/cmbG_LISTBOX"
                    , value = "Excel - Office Open XML Format (XLSX)"
                )
                errFmt = errFmt -1
            except Exception as e:
                err = str(e)
                
                errFmt = errFmt +1
                try:
                    ParamSAPLib.select_from_list_by_label(
                        element_id = "/app/con[0]/ses[0]/wnd[1]/usr/cmbG_LISTBOX"
                        , value = "Excel (in Office 2007 XLSX Format)"
                    )
                    errFmt = errFmt -1
                except Exception as e2:
                    err2 = str(e2)
                    errFmt = errFmt +1
                    
            if errFmt > 0:
                CommonFunction.WriteLog(f"PerformSAPDepreciationSimulation Select from All Available Formats failed to find Excel format.")
            else:
                ParamSAPLib.send_vkey(vkey_id = 0) # Press Enter
                ParamSAPLib.input_text(
                    element_id = "/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_PATH"
                    , text = ParamSAPDownloadPath
                )
                ParamSAPLib.input_text(
                    element_id = "/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_FILENAME"
                    , text = ParamSAPDownloadFile
                )
                ParamSAPLib.send_vkey(vkey_id = 0) # Press Enter

                SAPDownloadFilePath = ParamSAPDownloadPath + "\\" + ParamSAPDownloadFile

                # Retry until file exist in Folder
                vTempRetry = 1
                while(ParamInstanceFileSystem.does_file_not_exist(path = SAPDownloadFilePath) and vTempRetry < ParamFileCheckRetry):
                    CommonFunction.WriteLog(f"Failed when Check SAP File {ParamSAPDownloadFile} - Attempt # {str(vTempRetry)}")
                    sleep(ParamFileCheckSleep)
                    vTempRetry = vTempRetry + 1

                if(ParamInstanceFileSystem.does_file_not_exist(path = SAPDownloadFilePath)):
                    CommonFunction.WriteLog(f"Finish Automation since File SAP {ParamSAPDownloadFile} cannot be generated")
                    # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Connection issue:  File SAP " + ParamSAPDownloadFile + " cannot be generated.")
                else:
                    # Kill Excel Process that is opening
                    kill_process_username(process_name = "notepad.exe")
                    kill_process_username(process_name = "excel.exe")
                    # Press Shift F3 twice to close the FB3LN Page and return to main menu
                    ParamSAPLib.send_vkey(vkey_id = 15) # Press Shift F3
                    ParamSAPLib.send_vkey(vkey_id = 15) # Press Shift F3
                    CommonFunction.WriteConsole(f"Finish SAP Depreciation Simulation - {ParamSAPCompanyCode}")
                sleep(sleep10)
    except Exception as error:
        try:
            WindowsLib.screenshot(locator="desktop", filename=f'output\\SAPDepreciationSimulation_Error_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg') 
            fiscal = ParamSAPLib.get_value("/app/con[0]/ses[0]/wnd[1]/usr/txtMESSTXT1")
            idx = fiscal.index('Fiscal')
            if idx >= 0 :
                CommonFunction.WriteLog(f"Generate SAP Depreciation Simulation {str(ParamFlag)} error Fiscal year change not yet made.")
                # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Generate SAP Depreciation Simulation "+str(ParamFlag) +" error Fiscal year change not yet made.")
                ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]")
            else:
                # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Connection issue:  Generate SAP Depreciation Simulation "+str(ParamFlag) +" error")
                CommonFunction.WriteLog(f"Generate SAP Depreciation Simulation {str(ParamFlag)} error")
            CommonFunction.WriteLog(f"Fiscal year change not yet made.")
        except Exception as error:
            CommonFunction.WriteLog(f"Generate SAP Depreciation Simulation error with: {str(error)}")
             
            try:
                noData = ParamSAPLib.get_value("/app/con[0]/ses[0]/wnd[1]/usr/txtMESSTXT1")
                if noData == 'No data was selected':
                    CommonFunction.WriteLog(f"Generate SAP Depreciation Simulation "+str(ParamFlag) +" error data was not found.")
                    # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Does not exist: Generate SAP Depreciation Simulation "+str(ParamFlag) +" error data was not found.")
                    ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[1]/tbar[0]/btn[0]")
            except Exception as error2:
                CommonFunction.WriteLog(f"Generate SAP Depreciation Simulation error with: {str(error2)}")
                 
                try:
                    val = ParamSAPLib.get_value("/app/con[0]/ses[0]/wnd[0]/sbar/pane[0]")
                    CommonFunction.WriteLog(f"val: {str(val)}")
                    if val =='' or val == ' ' or val ==None:
                        CommonFunction.WriteLog(f"Generate SAP Depreciation Simulation {str(ParamFlag)} error failed to get element.")
                        # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Other Reason: Generate SAP Depreciation Simulation "+str(ParamFlag) +" error failed to get element.")
                    else:
                        CommonFunction.WriteLog(f"Generate SAP Depreciation Simulation  {str(ParamFlag)} error with {str(val)}")
                        # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Other Reason: Generate SAP Depreciation Simulation  "+str(ParamFlag) +" error with "+str(val))
                except Exception as error3:
                    # excelFunc.UpdateStatus(ParamSAPCompanyCode, 'Partially Completed', "Other Reason : Generate SAP Depreciation Simulation "+str(ParamFlag) +" error failed to get element.")
                    CommonFunction.WriteLog(f"Generate SAP Depreciation Simulation error with: {str(error3)}")
                     
                    sleep(sleep10)
        ParamSAPLib.send_vkey(vkey_id = 15) # Press Shift F3
            
def PerformSAPAssetTransaction(ParamInstanceApplication, ParamInstanceFileSystem, ParamSAPLib
, ParamSAPTransactionTypeList, ParamSAPCompanyCode, ParamSAPAssetClassMethod, ParamAssetClass
, ParamSAPBeginDate, ParamSAPEndDate, ParamSAPDepreciationArea, ParamSAPSortVariant, ParamSAPLayout
, ParamSAPDownloadPath, ParamSAPDownloadFile, ParamFileCheckRetry, ParamFileCheckSleep
):
    try: 
        CommonFunction.WriteLog(f"PerformSAPAssetTransaction - {ParamSAPCompanyCode}")
        CommonFunction.WriteConsole(f"Start SAP Asset Transaction - {ParamSAPCompanyCode}")
        
        ParamSAPLib.run_transaction("S_ALR_87012048")
        # ParamSAPLib.send_vkey(vkey_id = 19) # Press SHIFT + F7
        
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtBUKRS-LOW"
            , text = ParamSAPCompanyCode
        )
        
        if ParamSAPAssetClassMethod.lower() == 'ranges':
            ParamSAPLib.input_text(
                element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtSO_ANLKL-LOW"
                , text = ParamAssetClass[0]
            )
            ParamSAPLib.input_text(
                element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtSO_ANLKL-HIGH"
                , text = ParamAssetClass[1]
            )
        else:
            ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[0]/usr/btn%_SO_ANLKL_%_APP_%-VALU_PUSH")
            
            SAPInsertMultiSingleValuesString(
                ParamSAPLib, ParamAssetClass, ParamSAPAssetClassMethod.lower()
            )
                
            ParamSAPLib.send_vkey(vkey_id = 8) # Press F8
            sleep(maxWaitSAP)
            
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtBERDATUM"
            , text = ParamSAPEndDate
        )
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtBEREICH1"
            , text = ParamSAPDepreciationArea
        )
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtSRTVR"
            , text = ParamSAPSortVariant
        )
        ParamSAPLib.select_radio_button(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/radXEINZEL"
        )
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtP_VARI"
            , text = ParamSAPLayout
        )

        ParamSAPLib.send_vkey(vkey_id = 19) #shift + f7
        ParamInstanceApplication.send_keys(keys="{PAGEDOWN}") #page down
        
        ParamSAPLib.click_element(element_id = "/app/con[0]/ses[0]/wnd[0]/usr/btn%_SO_BWASL_%_APP_%-VALU_PUSH")
        
        SAPInsertMultiSingleValuesString(
            ParamSAPLib, ParamSAPTransactionTypeList, 'None'
        )
        
        ParamSAPLib.send_vkey(vkey_id = 8) # Press F8
        sleep(maxWaitSAP)
        
        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtSO_BUDAT-LOW"
            , text = ParamSAPBeginDate
        )

        ParamSAPLib.input_text(
            element_id = "/app/con[0]/ses[0]/wnd[0]/usr/ctxtSO_BUDAT-HIGH"
            , text = ParamSAPEndDate
        )
        
        ParamSAPLib.send_vkey(vkey_id = 8) # Press F8
        sleep(maxWaitSAP)
        try:
            ParamSAPLib.send_vkey(vkey_id = 9) # Press F9
            sleep(5)
            
            #Try to remove popup
            try:
                WindowsLib.screenshot(locator="desktop", filename=f'output\\SAPAssetTransaction_Click_F9_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg')  
                CommonFunction.WriteLog("Start try close popup in SAP Asset Transaction")
                WindowsLib.send_keys(keys="C")
                sleep(5)
                CommonFunction.WriteLog("Finish try close popup in SAP Asset Transaction")
            except:
                CommonFunction.WriteLog("SAP PopUp doesn't appear in SAP Asset Transaction")
            
            sleep(sleep10)
            
            try:
                CommonFunction.WriteLog(f"Start SAP Asset Transaction choose radio.")
                ParamSAPLib.select_radio_button(element_id ="/app/con[0]/ses[0]/wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]")
                ParamSAPLib.send_vkey(vkey_id = 0) # Press Enter
                CommonFunction.WriteLog(f"Finish SAP Asset Transaction choose radio.")
            except Exception as errTryRadio:
                errTryRadio = str(errTryRadio)
                WindowsLib.screenshot(locator="desktop", filename=f'output\\SAPAssetTransaction_FailedRadioButton_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg') 
                CommonFunction.WriteLog(f"SAP Asset Transaction failed to choose radio button. {errTryRadio}")
                                
            try:
                CommonFunction.WriteLog(f"Start SAP Asset Transaction choose path.")    
            
                ParamSAPLib.input_text(
                    element_id = "/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_PATH"
                    , text = ParamSAPDownloadPath
                )
                ParamSAPLib.input_text(
                    element_id = "/app/con[0]/ses[0]/wnd[1]/usr/ctxtDY_FILENAME"
                    , text = ParamSAPDownloadFile
                )
                ParamSAPLib.send_vkey(vkey_id = 0) # Press Enter
                
                CommonFunction.WriteLog(f"Finish SAP Asset Transaction choose path.")
            except Exception as errTryPath:
                errTryPath = str(errTryPath)
                WindowsLib.screenshot(locator="desktop", filename=f'output\\SAPAssetTransaction_FailedChoosePath_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg') 
                CommonFunction.WriteLog(f"SAP Asset Transaction failed to choose path. {errTryPath}")
            
            SAPDownloadFilePath = ParamSAPDownloadPath + "\\" + ParamSAPDownloadFile

            # Retry until file exist in Folder
            vTempRetry = 1
            while(ParamInstanceFileSystem.does_file_not_exist(path = SAPDownloadFilePath) and vTempRetry < ParamFileCheckRetry):
                CommonFunction.WriteLog(f"Failed when Check SAP File {ParamSAPDownloadFile} - Attempt # {str(vTempRetry)}")
                sleep(ParamFileCheckSleep)
                vTempRetry = vTempRetry + 1

            if(ParamInstanceFileSystem.does_file_not_exist(path = SAPDownloadFilePath)):
                WindowsLib.screenshot(locator="desktop", filename=f'output\\SAPAssetTransaction_FileDoesntExists_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg') 
                CommonFunction.WriteLog(f"Finish Automation since File SAP {ParamSAPDownloadFile} cannot be generated")
            else:
                # Kill Excel Process that is opening
                kill_process_username(process_name = "notepad.exe")
                kill_process_username(process_name = "excel.exe")
                # Press Shift F3 twice to close the FB3LN Page and return to main menu
                ParamSAPLib.send_vkey(vkey_id = 15) # Press Shift F3
                ParamSAPLib.send_vkey(vkey_id = 15) # Press Shift F3
                CommonFunction.WriteConsole(f"Finish SAP Asset Transaction - {ParamSAPCompanyCode}")
            sleep(sleep10)
        except Exception as error:
                WindowsLib.screenshot(locator="desktop", filename=f'output\\SAPAssetTransaction_NoData_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg') 
                CommonFunction.WriteLog(f"Generate SAP Assets Transaction error with: {str(error)}")
                sleep(sleep10)
    except Exception as error:
        WindowsLib.screenshot(locator="desktop", filename=f'output\\SAPAssetTransaction_Error_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.jpg')  
        CommonFunction.WriteLog(f"Generate SAP Assets Transaction error with: {str(error)}")
        sleep(sleep10)
        ParamSAPLib.send_vkey(vkey_id = 15) # Press Shift F3
            

def generateSAP(ParamExcelCompanyCode, ParamExcelDepreciationArea, ParamExcelDepreciationAreaIFRS, ParamExcelUSGAAPFlag, ParamExcelIFRSFlag, ParamExcelYear, ParamExcelMonth, ParamStartDate, ParamEndDate):   
    CommonFunction.WriteLog(f"USGAAP: {ParamExcelUSGAAPFlag} - IFRS: {ParamExcelIFRSFlag} - DepArea: {ParamExcelDepreciationArea} - DepAreaIFRS: {ParamExcelDepreciationAreaIFRS}")
    
    SAPConnectionServer = sapServer
    CommonFunction.WriteLog(f"SAPConnectionServer: {str(SAPConnectionServer)}")
        
    ExcelInputBeginDate, ExcelInputEndDate = GetSAPBeginEndDateYearMonth( int(ParamExcelYear), int(ParamExcelMonth) )
    try:
        if ParamStartDate.lower() != 'none' and ParamEndDate.lower() != 'none':
            ExcelInputBeginDate = str(ParamStartDate).replace("-", ".")
            ExcelInputEndDate = str(ParamEndDate).replace("-", ".")
    except:
        CommonFunction.WriteLog(f"SAP Start Date: {ParamStartDate} - End Date: {ExcelInputEndDate}")
    
    SAPDownloadPath = pathResultSAP
    SAPDownloadFile01 = "SAP "+ ParamExcelCompanyCode+" FBL3N US GAAP.XLSX"
    SAPDownloadFile02 = "SAP "+ ParamExcelCompanyCode+" FBL3N IFRS.XLSX"
    SAPDownloadFile03 = "SAP "+ ParamExcelCompanyCode+" Depr Simulation US GAAP.XLSX"
    SAPDownloadFile04 = "SAP "+ ParamExcelCompanyCode+" Depr Simulation IFRS.XLSX"
    SAPDownloadFile05 = "SAP "+ ParamExcelCompanyCode+" Asset Transaction.txt"
    FileCheckRetry = 10
    FileCheckSleep = 5

    CommonFunction.WriteLog(f"Start SAP Automation")

    # Step 1 - Open SAP Application
     
    vIntRetry = 0   
    while vIntRetry <= maxRetry:
        try:
            kill_process_username(process_name = "saplogon.exe")
            CommonFunction.WriteLog(f"Opening SAP with open path.")
            DesktopLib.open_application(name_or_path = r""+sapPath)
            sleep(sleep30) 
            list = WindowsLib.list_windows() 
            
            for app in list:
                if app['name'] == 'saplogon.exe':
                    vIntRetry=3
                    CommonFunction.WriteLog(f"Open SAP succesfully.")
                    CommonFunction.WriteConsole(f"Open SAP succesfully.")
                    break
                
            vIntRetry = vIntRetry + 1
        except Exception as errTryOpen:
            if vIntRetry == maxRetry:
                CommonFunction.WriteLog(f"[ERROR] Failed to open SAP with {str(errTryOpen)}")
                break
            sleep(3)
        vIntRetry = vIntRetry + 1
                     
    # Step 1.6 - Kill Excel Process that is opening
    CommonFunction.WriteLog(f"Kill Notepad")
    kill_process_username(process_name = "notepad.exe")
    kill_process_username(process_name = "excel.exe")


    # Step 2 - Open SAP Gui Library - Attach to Existing Application and Open Connection to ACQ
    SAPLib = SapGuiLibrary()
    SAPConnectionServer = SAPConnectionServer.replace('^', '').strip()
    iteration =0
    
    while iteration <= maxRetry:
        try:
            CommonFunction.WriteConsole(f"Connecting SAP...")
            SAPLib.connect_to_session()
            CommonFunction.WriteConsole(f"Session connected...")
            break
        except Exception as errTryConnect:
            if iteration == maxRetry:
                CommonFunction.WriteConsole(f"SAP connect to session failed: {errTryConnect}")
                break
            iteration=iteration+1
        
    
    iteration =0
    while iteration <= maxRetry:
        try:
            CommonFunction.WriteLog(f"Try Opening SAP. iteration: {iteration} - Max Retry: {maxRetry}")
            SAPLib.open_connection(SAPConnectionServer)
            CommonFunction.WriteConsole(f"SAP Connected.")
            break
        except Exception as errCon:
            if iteration == maxRetry:
                CommonFunction.WriteLog(f"Error Retry Opening SAP. iteration: {iteration} - Max Retry: {maxRetry}")
                CommonFunction.WriteConsole(f"Error Retry Opening SAP. iteration: {iteration} - Max Retry: {maxRetry}")
                # excelFunc.UpdateStatus(ParamExcelCompanyCode, 'Partially Completed', "Other Reason: Generate SAP error with "+str(errCon))
                break
            iteration=iteration+1
    
    try :
        CommonFunction.WriteLog("Start Check Pop Up Login Two Server")
        sleep(5)
        SAPLib.select_radio_button("/app/con[1]/ses[0]/wnd[1]/usr/radMULTI_LOGON_OPT1")
        SAPLib.click_element("/app/con[1]/ses[0]/wnd[1]/tbar[0]/btn[0]")
        WindowsLib.send_keys(keys="{ENTER}")
        
        CommonFunction.WriteLog("Finish Check Pop Up Login Two Server")
    except :
        CommonFunction.WriteLog(f"Connect Server : {SAPConnectionServer}")
        
    CommonFunction.WriteConsole(f"Loading SAP Transaction...")
    sleep(maxWaitLoginSAP) #wait loading login
    
    iteration=0
    while iteration <= totalClosePopupSAP:
        try:
            CommonFunction.WriteLog("Start try close popup")
            
            WindowsLib.send_keys(keys="{ENTER}")
            sleep(5)
            SAPLib.send_vkey(vkey_id = 0) # Press Enter
            sleep(5)
            CommonFunction.WriteLog("Finish try close popup")
            
            iteration=iteration+1
        except:
            CommonFunction.WriteLog("SAP PopUp doesn't appear")
            break
    
    CommonFunction.WriteConsole(f"SAP Run Transaction.")
        
    # Step PDD V - 3.15 - US GAAP balance
    if not os.path.exists(SAPDownloadPath+"\\"+SAPDownloadFile01):
        if(ParamExcelUSGAAPFlag == 'x'):
            PerformSAPFB3LN(
                InstanceFileSystem, SAPLib
                , sapGLAccountUSGAAP
                , ParamExcelCompanyCode, ExcelInputBeginDate, ExcelInputEndDate
                , sapLayout, sapAssetClassMethod
                , SAPDownloadPath, SAPDownloadFile01, FileCheckRetry, FileCheckSleep, "US GAAP"
            )

    # Step PDD V - 3.17 - IFRS balances
    if not os.path.exists(SAPDownloadPath+"\\"+SAPDownloadFile02):
        if(ParamExcelIFRSFlag == 'x'):
            PerformSAPFB3LN(
                InstanceFileSystem, SAPLib
                , sapGLAccountIFRS
                , ParamExcelCompanyCode, ExcelInputBeginDate, ExcelInputEndDate
                , sapLayout, sapAssetClassMethod
                , SAPDownloadPath, SAPDownloadFile02, FileCheckRetry, FileCheckSleep, "IFRS"
            )

    #Step PDD V - 3.18 - Depreciation Simulation US GAAP - S_ALR_87012936
    if not os.path.exists(SAPDownloadPath+"\\"+SAPDownloadFile03):
        if(ParamExcelUSGAAPFlag == 'x'):
            PerformSAPDepreciationSimulation(
                InstanceFileSystem, SAPLib
                , ParamExcelCompanyCode
                , sapAssetClassMethod, sapAssetClass
                , ParamExcelYear, ParamExcelDepreciationArea
                , sapDisplayVarientUSGAAP
                , SAPDownloadPath, SAPDownloadFile03, FileCheckRetry, FileCheckSleep, "US GAAP"
            )

    # Step PDD V - 3.20 - Depreciation Simulation IFRS - S_ALR_87012936
    if not os.path.exists(SAPDownloadPath+"\\"+SAPDownloadFile04):
        if(ParamExcelIFRSFlag == 'x'):
            PerformSAPDepreciationSimulation(
                InstanceFileSystem, SAPLib
                , ParamExcelCompanyCode
                , sapAssetClassMethod, sapAssetClass
                , ParamExcelYear, ParamExcelDepreciationAreaIFRS
                , sapDisplayVarientIFRS
                , SAPDownloadPath, SAPDownloadFile04, FileCheckRetry, FileCheckSleep, "IFRS"
            )

    # Step PDD V - 3.21 - Asset Transaction - S_ALR_87012048
    if not os.path.exists(SAPDownloadPath+"\\"+SAPDownloadFile05):
        PerformSAPAssetTransaction(
            WindowsLib, InstanceFileSystem, SAPLib
            , sapTransactionType
            , ParamExcelCompanyCode
            , sapAssetClassMethod, sapAssetClass
            , ExcelInputBeginDate, ExcelInputEndDate, ParamExcelDepreciationArea
            , sapSortVarient, sapLayout
            , SAPDownloadPath, SAPDownloadFile05, FileCheckRetry, FileCheckSleep
        )
        
    sleep(5)
    CommonFunction.WriteLog(f"Finish SAP Automation")
    kill_process_username(process_name = "saplogon.exe")
        