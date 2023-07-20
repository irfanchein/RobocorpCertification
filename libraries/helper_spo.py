import requests
import json


from function import CommonFunction

def getBearerToken(client_id, tenant_id, client_secret, url):
#     sbody = 'grant_type=client_credentials&client_id=bee18c35-545e-4ad9-afdd-7808d81d3be4@8b86a65e-3c3a-4406-8ac3-19a6b5cc52bc&client_secret=os~2J-g4p7sJuc~OHI~AUmvN39RdSO-Q5i&resource=00000003-0000-0ff1-ce00-000000000000/pmicloud.sharepoint.com@8b86a65e-3c3a-4406-8ac3-19a6b5cc52bc'
    sbody = 'grant_type=client_credentials&client_id=' + client_id +'@' + tenant_id + '&client_secret='+ client_secret  +'&resource=00000003-0000-0ff1-ce00-000000000000/pmicloud.sharepoint.com@' + tenant_id
    
    try:
        headers ={}
        headers['Content-Type'] = 'application/x-www-form-urlencoded'
        # CommonFunction.WriteLog(sbody)
        req = requests.post(url, headers=headers, data=sbody)
        # CommonFunction.WriteLog(req.status_code)
        req_json = json.loads(req.text)
        CommonFunction.WriteLog('Calling getBearerToken...')
        return req_json['access_token']
    except Exception as e:
        CommonFunction.WriteLog(e)
        return 'error:' + e

def getSharepointList(access_token,list_name,spo_url):
    try:
        url =spo_url + "_api/web/lists/getbytitle('" + list_name + "')/items"
#         url = spo_url + '_api/web?$select=Title'
        headers ={}
        headers['Authorization'] = 'Bearer ' + access_token
        headers['Accept'] = 'application/json'
        req = requests.get(url, headers=headers)
        # CommonFunction.WriteLog(f"{req}")
        # CommonFunction.WriteLog(f"{req.text}")
        req_json = json.loads(req.text)
        # CommonFunction.WriteLog(f"{req_json}")
        CommonFunction.WriteLog(f"response status: {req.status_code}")
        # CommonFunction.WriteLog(f"{req.headers}")
        CommonFunction.WriteLog('Calling getSharepointList...')
        return req_json
    except Exception as e:
        return 'error:' + e 

def getSharepointFile(token,file_url):
    
    try:
#         url =spo_url + "_api/web/lists/getbytitle('" + list_name + "')/items"
#         url = spo_url + '_api/web?$select=Title'
        headers ={}
        headers['Authorization'] = 'Bearer ' + token
#         headers['Accept'] = 'application/json'
        req = requests.get(file_url, headers=headers)
#         CommonFunction.WriteLog(req)
#         req_json = json.loads(req.text)
        # CommonFunction.WriteLog(req.status_code)
        # CommonFunction.WriteLog(req.headers)
        CommonFunction.WriteLog('Finish Getting Flat File From Sharepoint...')
        return req.text
    except Exception as e:
        return 'error:' + e 

def getSharepointFileExcel(token,file_url):
    try:
        headers ={}
        headers['Authorization'] = 'Bearer ' + token
        req = requests.get(file_url, headers=headers)
        CommonFunction.WriteLog(f'request status: {req}')
        CommonFunction.WriteLog('Finish Getting Excel File From Sharepoint...')
        return req.content
    except Exception as e:
        return 'error:' + e 


def createFolderSharepoint(token, file_url, new_folder_url):
    try:
        sbody = "{ '__metadata':{ 'type': 'SP.Folder' }, 'ServerRelativeUrl':'" + new_folder_url + "' }"
        headers ={}
        headers['Authorization'] = 'Bearer ' + token
        headers['Accept'] = 'application/json;odata=verbose'
        headers['Content-Type'] = 'application/json;odata=verbose'

        req = requests.post(file_url, headers = headers, data=sbody)
        CommonFunction.WriteLog('Finish Create Folder in Sharepoint...')
        return req
    except Exception as e:
        return 'error:' + e 

def uploadFileSharepoint(token, file_url, file_content):
    try:
        sbody = file_content
        headers ={}
        headers['Authorization'] = 'Bearer ' + token

        req = requests.post(file_url, headers = headers, data = sbody)
        CommonFunction.WriteLog('Finish Upload File in Sharepoint...')
        return req
    except Exception as e:
        return 'error:' + e