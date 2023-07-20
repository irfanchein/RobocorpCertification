import datetime, os
import subprocess

from RPA.Desktop.OperatingSystem import OperatingSystem
from robot.api import logger
from sendmail import send
from variables import area, EdgeOptions, PMI_UserAccount, maxRetry, MailMissingFilesSubject, MailMissingFilesMessage, MailMissingScheduleSubject, MailMissingScheduleMessage
from variables import ExcelLib, ENV_VAULT, SeleniumLib, LeaseReconSchedulerPath
from variables import LeaseReconEmails, LeaseReconListCompany, pathResultExcel, pathResultSummary
from variables import MailSender, MailSenderName, MailPassword, MailHost, mailSuccess, mailSuccessPartial, MailUserName, MailSupportAddr


class CommonFunction:
    def GetConfig(inSecret, inVariable):
        Secret = ENV_VAULT.get_secret(inSecret)
        return Secret[inVariable]

    def WriteLog(inMessage):
        logger.info("PY Log: " + inMessage + " - on: " + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        
    def WriteConsole(inMessage):
        logger.console(inMessage)

def kill_process_username(process_name):
    try:
        InstanceOS = OperatingSystem()
        result = subprocess.run([r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe', 'Get-WmiObject Win32_Process -Filter "name=\'' + process_name + '\'" | Select Handle, @{Name="UserName";Expression={$_.GetOwner().Domain+"\"+$_.GetOwner().User}} | Sort-Object Handle'], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, shell=True)
        output = result.stdout.decode('utf-8').rstrip()
        output2 = output.split('\n')
        vOutput = []
        vInt = 1
        for vloop in output2:
            if(vInt <= 3):
                vInt = vInt + 1
            else:
                vTemp = vloop.replace("\r", "").replace(" ", "").replace("PMI", "")
                vTempIndex = 0
                for i, c in enumerate(vTemp):
                    if(c.isdigit() == False):
                        vTempIndex = i
                        break
                vOutput.append(vTemp[0:vTempIndex] + "###&&&$$$" + vTemp[vTempIndex:])
                vInt = vInt + 1

        for vloop in vOutput:
            vPID = vloop.split("###&&&$$$")[0]
            vUsername = vloop.split("###&&&$$$")[1]
            if(vUsername == PMI_UserAccount):
                try:
                    InstanceOS.kill_process_by_pid(int(vPID))
                except:
                    CommonFunction.WriteLog(f"Kill process failed for PID : {vPID} and username: {vUsername} ")   
    
    except Exception as errKillProcess:
        CommonFunction.WriteLog(f"Kill process failed for username: {vUsername} with {str(errKillProcess)}")       

class selenium:
    def OpenWeb(link):
        SeleniumLib.open_available_browser(
            url = link,
            maximized=True,
            browser_selection="edge",
            options=EdgeOptions
        )
        SeleniumLib.set_browser_implicit_wait(value=1)
        SeleniumLib.set_selenium_implicit_wait(value=1)

    def ClickButton(name, type, visible):
        if(visible==True):
            SeleniumLib.click_button_when_visible(locator='//button[@'+type+'="'+name+'"]')
        else:
            SeleniumLib.click_button(locator='//button[@'+type+'="'+name+'"]')

    def ClickElement(name, type, contain):
        if (contain == True):
            SeleniumLib.click_element_when_visible(locator='xpath://'+type+'[contains(text(), "'+name+'")]')
        elif(contain != True and contain != False):
            SeleniumLib.click_element_when_visible(locator='xpath://'+type+'[@'+contain+'="'+name+'"]')
        else:
            SeleniumLib.click_element_when_visible(locator=type+':'+name)

    def WaitElement(name, type, contain, time):
        if (contain == True):
            SeleniumLib.wait_until_element_is_enabled(locator='xpath://'+type+'[contains(text(), "'+name+'")]', timeout=time)
        elif(contain != True and contain != False):
            SeleniumLib.wait_until_element_is_enabled(locator='xpath://'+type+'[@'+contain+'="'+name+'"]', timeout=time)
        else:
            SeleniumLib.wait_until_element_is_enabled(locator=type+':'+name, timeout=time)

    def WaitElementHidden(name, type, contain, time):
        if (contain == True):
            SeleniumLib.wait_until_element_is_not_visible(locator="xpath://"+type+"[contains(text(), '"+name+"')]", timeout=time)
        elif(contain != True and contain != False):
            SeleniumLib.wait_until_element_is_not_visible(locator="xpath://"+type+"[@"+contain+"='"+name+"']", timeout=time)
        else:
            SeleniumLib.wait_until_element_is_not_visible(locator=type+':'+name, timeout=time)

    def InputElement(name, type, contain, texts):
        if (contain == False):        
            SeleniumLib.input_text(locator=type+':'+name, text=texts)
        else:
            SeleniumLib.input_password(locator=type+':'+name, password=texts)

def createRunTime(id, startTime, endTime):
        try:
            path = pathResultExcel+"\\runtime.xlsx"
            if os.path.exists(path):
                os.remove(path)

            ExcelLib.create_workbook(path)
            
            ExcelLib.set_cell_value(row=1, column="A", value="ID")
            ExcelLib.set_cell_value(row=1, column="B", value="Start Time")
            ExcelLib.set_cell_value(row=1, column="C", value="End Time")
        
            ExcelLib.set_cell_value(row=2, column="A", value=str(id))
            if startTime.lower() != 'none' and startTime.lower() != '' and startTime != None:
                ExcelLib.set_cell_value(row=2, column="B", value=str(startTime))
            if endTime.lower() != 'none' and endTime.lower() != '' and endTime != None:
                ExcelLib.set_cell_value(row=2, column="C", value=str(endTime))

            ExcelLib.save_workbook(path)
        except Exception as errRuntime:
            error = str(errRuntime)
            CommonFunction.Writelog(f"Failed to create run time error with {error}")

def sendNotification(type):
    iteration = 0
    while iteration <= maxRetry: 
        try:
            CommonFunction.WriteLog(f"Start Send Mail. Iteration: {iteration} - Max Retry: {maxRetry}")
            CommonFunction.WriteConsole(f"Start Send Mail. Iteration: {iteration} - Max Retry: {maxRetry}")
            if type.lower() == "missing":
                MailTo = MailSupportAddr
                MailCC = ""
                MailSubject = MailMissingFilesSubject + " in " + area
                text = MailMissingFilesMessage
                fileAttached = None
            else:
                MailTo = MailSupportAddr
                MailCC = ""
                MailSubject = MailMissingScheduleSubject + " in " + area
                text = MailMissingScheduleMessage
                fileAttached = LeaseReconSchedulerPath
            
            MailHTMLBody = """
            <html>
            <head></head>
            <body>
                """ + text + """
            </body>
            </html>
            """
            EmailResponse = send(MailSender, MailSenderName, MailTo, MailCC, MailUserName, MailPassword, MailHost, MailSubject, None, MailHTMLBody, fileAttached)
            if EmailResponse == 0:
                CommonFunction.WriteLog("Send Mail Success.")
                CommonFunction.WriteLog("Finish automation.")
                CommonFunction.WriteConsole("Finish automation.")
                break
            else:
                CommonFunction.WriteLog("Send Mail Failed.")           
            iteration = iteration+1
            
        except Exception as errMail:
            if iteration == maxRetry:
                CommonFunction.WriteLog(f"Send Mail Failed. {errMail}")
                CommonFunction.WriteConsole(f"[Error] Send Mail Failed. {errMail}")
                break
            iteration = iteration+1
   
def LeaseReconSendMail():
    iteration = 0
    while iteration <= maxRetry: 
        try:
            
            CommonFunction.WriteLog(f"Start Send Mail. Iteration: {iteration} - Max Retry: {maxRetry}")
            CommonFunction.WriteConsole(f"Start Send Mail. Iteration: {iteration} - Max Retry: {maxRetry}")
            
            #delete runtime
            try:
                path = pathResultExcel+"\\runtime.xlsx"
                os.remove(path)
            except Exception as errRmv:
                err = str(errRmv)
            
            resultPath = None
            dt = datetime.datetime.now().strftime("%d-%m-%Y")
            endTime = datetime.datetime.now().strftime("%H_%M")
            end= str(endTime)
            prefixAct = "Scheduler with status "+str(dt)
            fileName = ""
            found = 0
            for ZipFiles in os.listdir(pathResultSummary):
                name = os.path.basename(ZipFiles)
                prefix = name[:32]
                postfix = name[33:]
                ext = postfix[5:]
                if prefixAct == prefix and ext == ".xlsx":
                    found = 1
                    resultPath = pathResultSummary+"\\"+name
                    fileName = name
                    
            if found == 0:
                resultPath = pathResultSummary+"\\"+fileName+".xlsx"
                fileName = prefixAct+" "+end
                
            CommonFunction.WriteLog(f"scheduler path: {resultPath}")
            if resultPath != '':
                ExcelLib.open_workbook(path=resultPath)
                rows = ExcelLib.read_worksheet('Sheet1', start=2)
                for cell in rows:
                    code = cell['A'] #company code - 1002
                    status = cell['D'] #status completed, partially completed, incomplete

                    if status != 'Completed' and status != None and code != None:
                        LeaseReconListCompany.append(str(code))
            
                content = '<br>'.join(LeaseReconListCompany)

                MailTo = LeaseReconEmails
                MailCC = ""
                MailSubject = str(fileName)
                if LeaseReconListCompany != []:
                    text = mailSuccessPartial.replace('listCompany', content)
                    text = text.replace('filename', str(fileName))
                else:                        
                    text = mailSuccess.replace('filename', str(fileName))

                MailHTMLBody = """
                <html>
                <head></head>
                <body>
                    """ + text + """
                </body>
                </html>
                """
                EmailResponse = send(MailSender, MailSenderName, MailTo, MailCC, MailUserName, MailPassword, MailHost, MailSubject, None, MailHTMLBody,resultPath)
                if EmailResponse == 0:
                    CommonFunction.WriteLog("Send Mail Success.")
                    CommonFunction.WriteLog("Finish automation.")
                    CommonFunction.WriteConsole("Finish automation.")
                    break
                else:
                    CommonFunction.WriteLog("Send Mail Failed.")           
            iteration = iteration+1
            
        except Exception as errMail:
            if iteration == maxRetry:
                CommonFunction.WriteLog(f"Send Mail Failed. {errMail}")
                CommonFunction.WriteConsole(f"[Error] Send Mail Failed. {errMail}")
                break
            iteration = iteration+1