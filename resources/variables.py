import datetime
import os

from datetime import date
from robot.api import logger
from RPA.Robocorp.Vault import Vault
from RPA.Browser.Selenium import Selenium
from RPA.Desktop import Desktop
from RPA.Desktop.OperatingSystem import OperatingSystem
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.Tables import Tables
from RPA.Windows import Windows
from selenium.webdriver.chrome.options import Options as Chrome_Options
from selenium.webdriver.edge.options import Options as Edge_Options

ENV_VAULT = Vault()
TODAY = datetime.datetime.now()

area = str(os.getenv("area")).strip()
ChromeOptions = Chrome_Options()
curdir = os.getcwd()
DesktopLib = Desktop()
EdgeOptions = Edge_Options()
excel = Files()
ExcelLib = Files()
ExcelTargetLib = Files()
InstanceOS = OperatingSystem()
InstanceFileSystem = FileSystem()
SeleniumLib = Selenium()
TablesTargetLib = Tables()
TablesLib = Tables()
today = date.today()
WindowsLib = Windows()
SeleniumLib.page
def GetConfig(inSecret, inVariable):
   Secret = ENV_VAULT.get_secret(inSecret)
   return Secret[inVariable]

def WriteLog(inMessage):
        logger.info("PY Log: " + inMessage + " - on: " + datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
#config
MailUserName = str(GetConfig('GeneralCredentials', 'EmailUsername')).strip()
MailPassword = str(GetConfig('GeneralCredentials','EmailPassword')).strip()
MailHost = str(GetConfig('GeneralCredentials','EmailHost')).strip()
PMI_UserAccount = str(os.getlogin()).strip()
PMI_Password = GetConfig('RobAccountCredentials',PMI_UserAccount).strip()
dateStart= str(os.getenv('startDate')).strip()
dateEnd= str(os.getenv('endDate')).strip()

LeaseReconPath = str(os.getenv("LeaseReconPath")).strip().replace("robUser", PMI_UserAccount)
LeaseReconInputPath = LeaseReconPath+str(os.getenv("LeaseReconInputPath")).strip()
LeaseReconResultPath = LeaseReconPath+str(os.getenv("LeaseReconResultPath")).strip()

#var
assetNbrLength = int(os.getenv("assetNbrLength"))
BrowserProfile = "user-data-dir=C:\\Users\\"+PMI_UserAccount+"\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default"
ChromeDriverPath = str(os.getenv("chrome_driver_path")).strip()
downloadPath = 'C:\\Users\\'+PMI_UserAccount+'\\Downloads'
EdgeDriverPath = str(os.getenv("edge_driver_path")).strip()
FilePassword = str(GetConfig('RPA_GL_R2R_Lease_Reconciliation','ExcelLockPassword')).strip()
Font = str(os.getenv("Font")).strip()

LeaseReconConfigName = str(os.getenv("LeaseReconConfigName")).strip()
LeaseReconConfigPath = LeaseReconInputPath+"/"+LeaseReconConfigName
LeaseReconEmails = str(os.getenv("LeaseReconEmails")).strip()
LeaseReconListCompany = []
LeaseReconSchedulerName = str(os.getenv("LeaseReconSchedulerName")).strip()
LeaseReconSchedulerPath = LeaseReconInputPath+"/"+LeaseReconSchedulerName
LeaseReconTemplateName= str(os.getenv("LeaseReconTemplateName")).strip()
LeaseReconTemplatePath = LeaseReconInputPath+"/"+LeaseReconTemplateName
LeaseReconSystem = str(os.getenv("LeaseReconSystem")).strip()

MailMissingFilesSubject = str(os.getenv('mailMissingFilesSubject')).strip()
MailMissingFilesMessage = str(os.getenv('mailMissingFilesMessage')).replace('^', '').strip()
MailMissingScheduleSubject = str(os.getenv('mailMissingScheduleSubject')).strip()
MailMissingScheduleMessage = str(os.getenv('mailMissingScheduleMessage')).replace('^', '').strip()

MailSender = str(os.getenv('mailSender')).strip()
MailSenderName = str(os.getenv('mailSenderName')).strip()
mailSuccess = str(os.getenv('mailSuccess')).replace('^', '').strip()
mailSuccessPartial = str(os.getenv('mailSuccessPartial')).replace('^', '').strip()
MailSupportAddr = str(os.getenv('mailSupportAddress')).strip()
maxRetry = int(os.getenv("MAX_TRY"))
maxWait = int(os.getenv("MAX_WAIT"))
maxWaitSAP = int(os.getenv("MAX_WAIT_SAP"))
maxWaitLoginSAP = int(os.getenv("MAX_WAIT_LOGIN_SAP"))
NAKISA_URL = str(os.getenv("nakisa_url_start")).strip()
NAKISA_URL = NAKISA_URL.replace("https://", "")
NAKISA_URL = NAKISA_URL.replace(":", "")
NAKISA_URL = NAKISA_URL.replace("//", "")
output = LeaseReconResultPath+"\\file" 
outputExcel = LeaseReconResultPath+"\\excel"
outputFolder = str(datetime.datetime.now().strftime("%Y-%m")).strip()
if dateEnd != None and dateEnd != "None":
   ed = datetime.datetime.strptime(dateEnd, "%d-%m-%Y")
   outputFolder = str(ed.strftime("%Y-%m"))
outputJson = output+"\\json"
pathResultExcel = output+"\\"+area+"\\NAKISA"
pathResultSAP= output+"\\"+area+"\\SAP"
pathResultSummary= output+"\\"+area+"\\Summary"
RemoteDebug = str(os.getenv("remote_debug")).strip()
RobotStatus = "300"
sapAssetClass = eval(os.getenv("sapAssetClass").replace('^', '').strip())
sapAssetClassMethod =str(os.getenv("sapAssetClassMethod")).strip()
sapDisplayVarientUSGAAP=str(os.getenv("sapDisplayVarientUSGAAP")).strip()
sapDisplayVarientIFRS=str(os.getenv("sapDisplayVarientIFRS")).strip()
sapGLAccountUSGAAP = eval(os.getenv("sapGLAccountUSGAAP").replace('^', '').strip())
sapGLAccountIFRS = eval(os.getenv("sapGLAccountIFRS").replace('^', '').strip())
sapLayout = str(os.getenv("sapLayout")).strip()
sapPath = str(os.getenv("SAPApplicationPath")).strip()
sapServer = str(os.getenv("SAPConnectionServer")).strip()
sapSortVarient = str(os.getenv("sapSortVarient")).strip()
sapTitle = str(os.getenv("SAPApplicationTitle")).strip()
sapTransactionType = eval(os.getenv("sapTransactionType").replace('^', '').strip())
SessionStart = datetime.datetime.utcnow()
totalClosePopupSAP = int(os.getenv("totalClosePopupSAP"))

#business sharepoint used for scheduler and results
sharepointBusinessURL = str(os.getenv("sharepointBusinessURL")).strip()
sharepointBusinessPath = str(os.getenv("sharepointBusinessPath")).strip()
sharepointBusinessFolder = str(os.getenv("sharepointBusinessFolder")).strip()

#config sharepoint used for config and template
sharepointConfigURL = str(os.getenv("sharepointConfigURL")).strip()
sharepointConfigPath = str(os.getenv("sharepointConfigPath")).strip()
sharepointConfigFolder = str(os.getenv("sharepointConfigFolder")).strip()

sharepointFolderInput = str(os.getenv("sharepointFolderInput")).strip()
sharepointFolderOutput = str(os.getenv("sharepointFolderOutput")).strip()

sharepointFolderMonth = outputFolder

sharepointFolderDetail = str(os.getenv("sharepointFolderDetail")).strip()
sharepointFolderSummary = str(os.getenv("sharepointFolderSummary")).strip()

sharepointBusinessFullPath = sharepointBusinessPath+sharepointBusinessFolder+"/"
sharepointBusinessFullPath = sharepointBusinessFullPath.lower().replace("none/", "")

sharepointConfigFullPath = sharepointConfigPath+sharepointConfigFolder+"/"
sharepointConfigFullPath = sharepointConfigFullPath.lower().replace("none/", "")

SPOclientId = str(GetConfig('GeneralCredentials', "SharepointClientID")).strip()
SPOtenantId = str(GetConfig('GeneralCredentials', "SharepointTenantID")).strip()
SPOsecretId = str(GetConfig('GeneralCredentials', "SharepointSecretID")).strip()

sharepointURL  = 'https://accounts.accesscontrol.windows.net/'+ SPOtenantId +'/tokens/OAuth/2'


sleep1 = 1
sleep3 = 3
sleep5 = 5
sleep10 = 10
sleep15 = sleep10 + 5
sleep30 = sleep10 * 3
timeout_selenium10 = 10
timeout_selenium30 = timeout_selenium10 * 3
timeout_selenium60 = timeout_selenium10 * 6
timeout_selenium90 = timeout_selenium10 * 9

URL_NAKISA = "https://"+PMI_UserAccount+"@PMINTL.NET:"+PMI_Password+"@"+NAKISA_URL

#Example browser arguments
# options.addArguments("--no-sandbox"); // Bypass OS security model, MUST BE THE VERY FIRST OPTION
# options.addArguments("--headless");
# options.setExperimentalOption("useAutomationExtension", false);
# options.addArguments("start-maximized"); // open Browser in maximized mode
# options.addArguments("disable-infobars"); // disabling infobars
# options.addArguments("--disable-extensions"); // disabling extensions
# options.addArguments("--disable-gpu"); // applicable to windows os only
# options.addArguments("--disable-dev-shm-usage"); // overcome limited resource problems

browserArgNoSandbox = "--no-sandbox"
browserArgStartMaximized = "--start-maximized"
browserArgDisableSHM = "--disable-dev-shm-usage"

ChromeOptions.add_argument(browserArgStartMaximized)
ChromeOptions.add_argument(RemoteDebug)

EdgeOptions.add_argument(browserArgNoSandbox)
EdgeOptions.add_argument(BrowserProfile)
EdgeOptions.add_argument(browserArgDisableSHM)
EdgeOptions.add_argument(browserArgStartMaximized)
EdgeOptions.add_argument(RemoteDebug)

class excelPath:
   pathAGSLAN='None'
   pathUNITSLAN='None'
   pathUSGAAP='None'
   pathIFRS='None'
   pathLiabilityUSGAAP='None'
   pathLiabilityIFRS='None'
   pathSAPFBL3NUSGAAP='None'
   pathSAPFBL3NUSGAAPHTML='None'
   pathSAPFBL3NIFRS='None'
   pathSAPFBL3NIFRSHTML='None'
   pathSAPDeprUSGAAP='None'
   pathSAPDeprIFRS='None'
   pathSAPAssetTrans='None'
   pathSAPAssetTransFix='None'
   pathResultFile = 'None'
