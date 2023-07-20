import datetime, glob, os, time
import helper_spo as spo

from function import CommonFunction, createRunTime
from time import sleep
from variables import downloadPath, ExcelLib, InstanceFileSystem, maxRetry
from variables import LeaseReconInputPath, LeaseReconSchedulerPath, LeaseReconResultPath
from variables import LeaseReconConfigName, LeaseReconConfigPath, LeaseReconResultPath, LeaseReconSchedulerName, LeaseReconTemplateName, LeaseReconTemplatePath
from variables import pathResultExcel, pathResultExcel, pathResultSAP, pathResultSummary, SPOclientId, SPOsecretId, SPOtenantId
from variables import sharepointURL, sharepointBusinessURL, sharepointBusinessFullPath, sharepointConfigURL, sharepointConfigFullPath, sharepointFolderDetail, sharepointFolderInput, sharepointFolderMonth, sharepointFolderOutput, sharepointFolderSummary


access_token = spo.getBearerToken(SPOclientId, SPOtenantId, SPOsecretId, sharepointURL)

def DownloadInputFile():
    iteration = 0
    while iteration <= maxRetry:
        try:
            CommonFunction.WriteConsole(f"Start")
            CommonFunction.WriteConsole(f"Start Download Input File (Config, Template & Scheduler). Iteration: {iteration} - Max Retry: {maxRetry}")
            CommonFunction.WriteLog(f"sharepointURL: {sharepointURL}")
        
            CommonFunction.WriteLog(f"sharepointBusinessURL: {sharepointBusinessURL}")
            CommonFunction.WriteLog(f"sharepointBusinessFullPath: {sharepointBusinessFullPath}")
            
            CommonFunction.WriteLog(f"sharepointConfigURL: {sharepointConfigURL}")
            CommonFunction.WriteLog(f"sharepointConfigFullPath: {sharepointConfigFullPath}")
            
            CommonFunction.WriteLog(f"sharepointFolderInput: {sharepointFolderInput}")
            CommonFunction.WriteLog(f"sharepointFolderOutput: {sharepointFolderOutput}")
            CommonFunction.WriteLog(f"sharepointFolderMonth: {sharepointFolderMonth}")
            CommonFunction.WriteLog(f"sharepointFolderDetail: {sharepointFolderDetail}")
            CommonFunction.WriteLog(f"sharepointFolderSummary: {sharepointFolderSummary}")
            
            #check sharepointFolder 
            if not os.path.exists(LeaseReconInputPath):
                try:
                    os.makedirs(LeaseReconInputPath)
                except:
                    sleep(3)

            if not os.path.exists(pathResultExcel):
                try:
                    os.makedirs(pathResultExcel)
                except:
                    sleep(3)

            if not os.path.exists(pathResultSAP):
                try:
                    os.makedirs(pathResultSAP)
                except:
                    sleep(3)
                
            if not os.path.exists(pathResultSummary):
                try:
                    os.makedirs(pathResultSummary)
                except:
                    sleep(3)
                
            # Clean Up Files 
            if os.path.exists(downloadPath):
                try:
                    InstanceFileSystem.empty_directory(path = downloadPath)
                except:
                    sleep(3)
                
            if os.path.exists(LeaseReconInputPath):
                try:
                    InstanceFileSystem.empty_directory(path = LeaseReconInputPath)
                except:
                    sleep(3)
            
            if os.path.exists(pathResultExcel):
                try:
                    InstanceFileSystem.empty_directory(path = pathResultExcel)
                except:
                    sleep(3)
            
            if os.path.exists(pathResultSAP):
                try:
                    InstanceFileSystem.empty_directory(path = pathResultSAP)
                except:
                    sleep(3)
            
            if os.path.exists(pathResultSummary):
                try:
                    InstanceFileSystem.empty_directory(path = pathResultSummary)
                except:
                    sleep(3)
            
            for filename in glob.glob(LeaseReconResultPath+"/sap-screenshot*"):
                try:
                    os.remove(filename) 
                except:
                    sleep(3)
            
            #Download File & Write to Local
            if sharepointFolderInput.lower() == 'none':
                SPOTemplatePath    = sharepointConfigFullPath + str(LeaseReconTemplateName)
            else:
                SPOTemplatePath    = sharepointConfigFullPath + sharepointFolderInput + "/" + str(LeaseReconTemplateName)
                        
            vfile_url1 = sharepointConfigURL + "_api/web/GetFileByServerRelativeUrl('"+ SPOTemplatePath +"')/$value"
            one_drive_template_file = spo.getSharepointFileExcel(access_token, vfile_url1)
            msg =str(one_drive_template_file)
            msg = msg.find("does not exist")
            if msg == -1:
            # Perform retry as it may error
                for vLoopRetry in range(5):
                    try:
                        write_mapping_path = LeaseReconTemplatePath
                        write_mapping_mode = 'wb'
                        with open(write_mapping_path, write_mapping_mode) as outfile:
                            outfile.write(one_drive_template_file)
                            break
                    except:
                        CommonFunction.WriteConsole("Fail Writing Lease Recon Template Excel File on Attempt No: " + str(vLoopRetry + 1))
                        time.sleep(5)
                CommonFunction.WriteLog("Finish download template from sharepoint")
            
            
            if sharepointFolderInput.lower() == 'none':
                SPOConfigPath    = sharepointConfigFullPath + str(LeaseReconConfigName)
            else:
                SPOConfigPath    = sharepointConfigFullPath + sharepointFolderInput + "/" + str(LeaseReconConfigName)
            
            id = datetime.datetime.now().strftime("%d-%m-%Y")
            startTime = datetime.datetime.now().strftime("%H_%M")
            CommonFunction.WriteLog(f"Start Time Generate "+startTime)
            createRunTime(id, startTime, '')
            #sharepoint each server
            # Download File & Write to Local
            vfile_url1 = sharepointConfigURL + "_api/web/GetFileByServerRelativeUrl('"+ SPOConfigPath +"')/$value"
            one_drive_template_file1 = spo.getSharepointFileExcel(access_token, vfile_url1)
            msg1 =str(one_drive_template_file1)
            msg1 = msg1.find("does not exist")
            if msg1 == -1:
            # Perform retry as it may error
                for vLoopRetry in range(5):
                    try:
                        write_mapping_path = LeaseReconConfigPath
                        write_mapping_mode = 'wb'
                        with open(write_mapping_path, write_mapping_mode) as outfile:
                            outfile.write(one_drive_template_file1)
                            break
                    except:
                        CommonFunction.WriteConsole("Fail Writing Lease Recon Config File on Attempt No: " + str(vLoopRetry + 1))
                        time.sleep(5)
                CommonFunction.WriteLog("Finish download config from sharepoint")

            if sharepointFolderInput.lower() == 'none':
                SPOSchedulerPath    = sharepointBusinessFullPath + str(LeaseReconSchedulerName)
            else:
                SPOSchedulerPath    = sharepointBusinessFullPath + sharepointFolderInput + "/"  + LeaseReconSchedulerName

            # Download File & Write to Local
            vfile_url2 = sharepointBusinessURL + "_api/web/GetFileByServerRelativeUrl('"+ SPOSchedulerPath +"')/$value"
            one_drive_template_file2 = spo.getSharepointFileExcel(access_token, vfile_url2)
            msg2 =str(one_drive_template_file2)
            msg2 = msg2.find("does not exist")
            if msg2 == -1:
                # Perform retry as it may error
                for vLoopRetry2 in range(5):
                    try:
                        write_mapping_path2 = LeaseReconSchedulerPath
                        write_mapping_mode2 = 'wb'
                        with open(write_mapping_path2, write_mapping_mode2) as outfile2:
                            outfile2.write(one_drive_template_file2)
                            break
                    except:
                        CommonFunction.WriteConsole("Fail Writing Lease Recon Scheduler File on Attempt No: " + str(vLoopRetry2 + 1))
                        time.sleep(5)
                
                CommonFunction.WriteLog("Finish download scheduler from sharepoint")
            
            CommonFunction.WriteLog(f"SPOConfigPath: {SPOConfigPath}")
            CommonFunction.WriteLog(f"SPOTemplatePath: {SPOTemplatePath}")
            CommonFunction.WriteLog(f"SPOSchedulerPath: {SPOSchedulerPath}")
            
            if msg != -1 or msg1 != -1 or msg2 != -1:
                endTime = datetime.datetime.now().strftime("%H_%M")
                end= str(endTime)
                createRunTime(id, startTime, end)
                if msg != -1:
                    CommonFunction.WriteConsole("Failed download template from sharepoint")
                    CommonFunction.WriteLog("Failed download template from sharepoint")
                if msg1 != -1:
                    CommonFunction.WriteConsole("Failed download config from sharepoint")
                    CommonFunction.WriteLog("Failed download config from sharepoint")
                if msg2 != -1:
                    CommonFunction.WriteConsole("Failed download scheduler from sharepoint")
                    CommonFunction.WriteLog("Failed download scheduler from sharepoint")
                CommonFunction.WriteLog("Finish automation.")
                CommonFunction.WriteLog("There is no config or scheduler. Abort processing automation.")
            else:
                CommonFunction.WriteConsole(f"Finish Download Input File (Config, Template & Scheduler).")
                break
            
            iteration=iteration+1
        except Exception as errSPO:
            if iteration == maxRetry:
                CommonFunction.WriteLog(f"Sharepoint download error with: {errSPO}")
                CommonFunction.WriteConsole(f"[ERROR] There is no config or scheduler. Abort processing automation.")
                break
            iteration=iteration+1
            
def UploadFile(fileName, filePath, typeFile):
    iteration = 0
    while iteration <= maxRetry: 
        try:
            file_url = sharepointBusinessURL + "_api/web/folders"
            
            fileName = fileName.replace('/', '')
            
            if sharepointFolderOutput.lower() == 'none':
                new_folder_url = sharepointBusinessFullPath
            else:
                new_folder_url = sharepointBusinessFullPath+sharepointFolderOutput
                            
            new_folder_url1 = new_folder_url+"/"+sharepointFolderMonth
            new_folder_url2 = new_folder_url1+"/"+sharepointFolderDetail 
            new_folder_url3 = new_folder_url1+"/"+sharepointFolderSummary 
            
            detailFolder = new_folder_url1
            summaryFolder = new_folder_url1
            
            try:
                request_result = spo.createFolderSharepoint(access_token, file_url, new_folder_url)
            except:
                sleep(3)
            
            try:
                request_result1 = spo.createFolderSharepoint(access_token, file_url, new_folder_url1)
            except:
                sleep(3)
                
            if str(typeFile).lower() == 'summary':
                CommonFunction.WriteLog(f"Start upload summary file in sharepoint. Iteration: {iteration} - Max Retry: {maxRetry}")
                CommonFunction.WriteConsole(f"Start upload summary file in sharepoint. Iteration: {iteration} - Max Retry: {maxRetry}")
                CommonFunction.WriteLog(f"summaryFolder: {summaryFolder}")
                
                if sharepointFolderSummary.lower() != 'none':
                    try:
                        request_result3 = spo.createFolderSharepoint(access_token, file_url, new_folder_url3) 
                        summaryFolder = new_folder_url3
                    except:
                        sleep(3)
                
                write_template_path = filePath
                target_folder_url = summaryFolder
                file_url = sharepointBusinessURL + "_api/web/GetFolderByServerRelativeUrl('" + target_folder_url + "')/Files/add(url='"+fileName+"',overwrite=true)"

                with open(write_template_path, 'rb') as file_input:
                    request_result = spo.uploadFileSharepoint(access_token, file_url, file_input)
                    
                CommonFunction.WriteLog(f"Finish upload {fileName} in sharepoint.")
                iteration=iteration+1
                break
                
            elif str(typeFile).lower() == 'detail' and sharepointFolderDetail.lower() != 'none':
                CommonFunction.WriteLog(f"Start upload detail file in sharepoint. Iteration: {iteration} - Max Retry: {maxRetry}")
                CommonFunction.WriteLog(f"detailFolder: {detailFolder}")
                
                try:
                    request_result2 = spo.createFolderSharepoint(access_token, file_url, new_folder_url2) 
                    detailFolder = new_folder_url2
                except:
                    sleep(3)

                write_template_path = filePath
                target_folder_url = detailFolder
                file_url = sharepointBusinessURL + "_api/web/GetFolderByServerRelativeUrl('" + target_folder_url + "')/Files/add(url='"+fileName+"',overwrite=true)"

                with open(write_template_path, 'rb') as file_input:
                    request_result = spo.uploadFileSharepoint(access_token, file_url, file_input)
                
                CommonFunction.WriteLog(f"Finish upload {fileName} in sharepoint.")
                iteration=iteration+1
                break
            
            iteration=iteration+1
            break
            
        except Exception as errSPO:
            if iteration == maxRetry:
                CommonFunction.WriteLog(f"Sharepoint upload error with: {errSPO}")
                if os.path.exists(str(filePath)):
                    CommonFunction.WriteConsole(f"[ERROR] Sharepoint upload error with: {errSPO}")
                break
            iteration=iteration+1
            
def uploadStatus():
    iteration = 0
    while iteration <= maxRetry: 
        try:
            CommonFunction.WriteConsole(f"Start upload scheduler with status. Iteration: {iteration} - Max Retry: {maxRetry}")
            path = pathResultExcel+"\\runtime.xlsx"
            ExcelLib.open_workbook(path)
            row = ExcelLib.read_worksheet('Sheet', start=2)
            for cell in row:
                start = cell['B']
                end = cell['C']
            ExcelLib.close_workbook()    

            ExcelLib.open_workbook(path=LeaseReconSchedulerPath)
            dt = datetime.datetime.now().strftime("%d-%m-%Y")
            
            endTime = datetime.datetime.now().strftime("%H_%M")
            end= str(endTime)
            
            # Once all process steps completed for all company codes Robot goes to sharepointFolder named YYYYMM
            # (where YYYY-year, MM-month) and updates the scheduler file name with End Time.,
            # The file name format is “Scheduler with status dd-mm-yyyy hh-mm to hh-mm”, where the
            # end time is the time at the end of the name.
            fileName3 = "Scheduler with status "+str(dt)+" "+end+".xlsx"
            fileName2 = "Scheduler with status "+str(dt)+" "+start+" to "+end+".xlsx"
            
            schedulerPath = pathResultSummary+"\\"+fileName2
            schedulerMailPath = pathResultSummary+"\\"+fileName3
            if os.path.exists(schedulerPath):
                os.remove(schedulerPath)
            ExcelLib.save_workbook(path=schedulerPath)
            ExcelLib.save_workbook(path=schedulerMailPath)
            CommonFunction.WriteLog("schedulerPath: "+fileName2)
            CommonFunction.WriteLog("schedulerMailPath: "+fileName3)
            
            if sharepointFolderOutput.lower() == 'none':
                new_folder_url = sharepointBusinessFullPath
            else:
                new_folder_url = sharepointBusinessFullPath+sharepointFolderOutput
                            
            new_folder_url1 = new_folder_url+"/"+sharepointFolderMonth
            new_folder_url3 = new_folder_url1+"/"+sharepointFolderSummary 
            
            summaryFolder = new_folder_url1
            
            if sharepointFolderSummary.lower() != 'none':
                summaryFolder = new_folder_url3
                
            write_template_path = schedulerPath
            target_folder_url = summaryFolder
            file_url = sharepointBusinessURL + "_api/web/GetFolderByServerRelativeUrl('" + target_folder_url + "')/Files/add(url='"+fileName2+"',overwrite=true)"

            with open(write_template_path, 'rb') as file_input:
                request_result = spo.uploadFileSharepoint(access_token, file_url, file_input)
            
            CommonFunction.WriteConsole(f"Finish upload scheduler with status.")
                    
            iteration=iteration+1
            break
        except Exception as errSched:
            if iteration == maxRetry:
                CommonFunction.WriteConsole(f"[ERROR] Sharepoint upload scheduler error with: {str(errSched)}")
                CommonFunction.WriteLog(f"[ERROR] Sharepoint upload scheduler error with: {str(errSched)}")
                break
            iteration=iteration+1
