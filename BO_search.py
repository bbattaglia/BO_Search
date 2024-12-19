import requests
import xlsxwriter

# Global Variables
username = 'test'
password = 'test' 
URL = 'http://botst01.apps.TEST.net/' 

#### Functions 
def login(URL,headers):
    print("Getting token...")
    data_get = '''
        <attrs xmlns="http://www.sap.com/rws/bip">
            <attr name="password" type="string">Srvicbo</attr>
            <attr name="auth" type="string" possibilities="secEnterprise,secLDAP,secWinAD,secSAPR3">secEnterprise</attr>
            <attr name="userName" type="string">SRVICBO</attr>
        </attrs>
    '''
    r = requests.post(URL + 'biprws/logon/long', data=data_get, headers=headers)
    if r.ok:
        x_sap_logontoken = r.headers['X-SAP-LogonToken']
        print("Token: " + x_sap_logontoken)
        return x_sap_logontoken
    else:
        print("HTTP %i - %s, Message %s" % (r.status_code, r.reason, r.text))

def set_headers():
    headers = {
                'userName': 'SRVICBO',
                'password':'Srvicbo',
                'auth':'secEnterprise',
                'xmlns':'http://www.sap.com/rws/bip',
                'Content-Type':'application/xml',
                'Accept':'application/xml'
                }
    return headers

def set_headers_search(x_sap_logontoken):
    headers = {'x-sap-logontoken': x_sap_logontoken,
               'Content-Type':'application/json',
               'Accept':'application/json'
                }
    return headers

def searchUniverse(URL,x_sap_logontoken,workbook):
    headers = set_headers_search(x_sap_logontoken)
    parentId = []
    query = {
        "query": "SELECT * FROM CI_APPOBJECTS WHERE SI_KIND='Universe'"
        }
    r = requests.post(URL + 'biprws/v1/cmsquery?pagesize=8000', headers=headers, json=query)
    if r.ok:
        row = 0
        worksheetUniverse = workbook.add_worksheet("Universe")
        worksheetReport = workbook.add_worksheet("Universe-Report")
        worksheetFolder = workbook.add_worksheet("Universe Folder")
        print('OK Universe')
        response=r.json()
        entries = response['entries']
        for i in range(len(entries)):
            col = 0
            id = entries[i].get('SI_ID', '')
            name = entries[i].get('SI_NAME', '')
            description = entries[i].get('SI_DESCRIPTION','Missing')
            folder = entries[i].get('SI_PARENT_FOLDER_CUID','Missing')
            report = entries[i]['SI_WEBI']
            totalReport = report['SI_TOTAL']
            worksheetUniverse.write(i, col, id)
            worksheetUniverse.write(i, col+1, name)
            worksheetUniverse.write(i, col+2, description)
            worksheetUniverse.write(i, col+3, folder)
            worksheetUniverse.write(i, col+4, totalReport)
            parentId.append(folder)
            for j in range(totalReport):
                worksheetReport.write(row, 0, id)
                worksheetReport.write(row, 1, report[str(j+1)])
                row += 1
        parentId = list(dict.fromkeys(parentId))
       
        for i in range (len(parentId)):
            col = 0
            query = {
                "query": "SELECT SI_ID,SI_NAME,SI_PATH FROM CI_APPOBJECTS WHERE SI_CUID =\'{}\'".format(parentId[i])
            }
            r = requests.post(URL + 'biprws/v1/cmsquery?pagesize=8000', headers=headers, json=query)
            if r.ok:
                universePath = ""
                response = r.json()
                entries = response['entries']
                folderPath = entries[0]['SI_PATH']
                numFolder = folderPath['SI_NUM_FOLDERS']
                for j in range(numFolder):
                    universePath += folderPath['SI_FOLDER_NAME'+str(j+1)]+'/'
                    print(universePath)
                worksheetFolder.write(i, col,parentId[i])
                worksheetFolder.write(i, col+1, universePath)
            else:
                print("HTTP %i - %s, Message %s" % (r.status_code, r.reason, r.text))
    else:
         print("HTTP %i - %s, Message %s" % (r.status_code, r.reason, r.text))



def getUsers(URL,x_sap_logontoken,workbook):
    headers = set_headers_search(x_sap_logontoken)
    query = {
        "query": " Select SI_NAME, SI_USERFULLNAME, SI_ID, SI_NAMEDUSER,SI_LASTLOGONTIME, SI_EMAIL_ADDRESS From CI_SYSTEMOBJECTS Where SI_KIND='User' Order by SI_NAME "
        }
    r = requests.post(URL + 'biprws/v1/cmsquery?pagesize=8000', headers=headers, json=query)
    if r.ok:
        print('OK')
        worksheet = workbook.add_worksheet("User")
        response=r.json()
        entries = response['entries']
        for i in range(len(entries)):
            col = 0
            id= entries[i].get('SI_ID', '')
            username = entries[i].get('SI_NAME', '')
            email = entries[i].get('SI_EMAIL_ADDRESS','Missing')
            name = entries[i].get('SI_USERFULLNAME','Missing')
            lastAccess = entries[i].get('SI_LASTLOGONTIME','1-gen-1900 00.00')
            worksheet.write(i, col, id)
            worksheet.write(i, col+1, username)
            worksheet.write(i, col+2, email)
            worksheet.write(i, col+3, name)
            worksheet.write(i, col+4, lastAccess)        
    else:
         print("HTTP %i - %s, Message %s" % (r.status_code, r.reason, r.text))
   

def searchReport(URL,x_sap_logontoken,workbook):
    headers = set_headers_search(x_sap_logontoken)
    parentId = []
    query = {
        "query": "SELECT SI_DESCRIPTION,SI_ID,SI_NAME,SI_OWNER,SI_CREATION_TIME,SI_PARENT_FOLDER_CUID,SI_UPDATE_TS FROM CI_INFOOBJECTS WHERE SI_KIND IN ('Webi') AND SI_ANCESTOR = 23"
        }
    r = requests.post(URL + 'biprws/v1/cmsquery?pagesize=20000', headers=headers, json=query)
    if r.ok:
        print('OK')
        worksheet = workbook.add_worksheet("Report")
        response=r.json()
        entries = response['entries']
        for i in range(len(entries)):
            col = 0
            reportDescription = entries[i]['SI_DESCRIPTION']
            reportID = entries[i]['SI_ID']
            reportName = entries[i]['SI_NAME']
            reportOwner = entries[i]['SI_OWNER']
            reportCreation = entries[i]['SI_CREATION_TIME']
            reportFolder = entries[i].get('SI_PARENT_FOLDER_CUID','Missing')
            reportLastUpdate = entries[i].get('SI_UPDATE_TS','Missing') 
            worksheet.write(i, col, reportID)
            worksheet.write(i, col+1, reportName)
            worksheet.write(i, col+2, reportOwner)
            worksheet.write(i, col+3, reportCreation)
            worksheet.write(i, col+4, reportLastUpdate)
            worksheet.write(i, col+5, reportDescription)
            worksheet.write(i, col+6, reportFolder)
            parentId.append(reportFolder)
        parentId = list(dict.fromkeys(parentId))
        worksheet = workbook.add_worksheet("Report Folder")
        for i in range (len(parentId)):
            col = 0
            query = {
                "query": " SELECT SI_ID,SI_NAME,SI_PATH FROM CI_INFOOBJECTS WHERE SI_CUID =\'{}\'".format(parentId[i])
            }
            r = requests.post(URL + 'biprws/v1/cmsquery?pagesize=20000', headers=headers, json=query)
            if r.ok:
                reportPath = ""
                response = r.json()
                entries = response['entries']
                folderPath = entries[0]['SI_PATH']
                numFolder = folderPath['SI_NUM_FOLDERS']
                for j in range(numFolder):
                    reportPath += folderPath['SI_FOLDER_NAME'+str(j+1)]+'/'
                    print(reportPath)
                worksheet.write(i, col,parentId[i])
                worksheet.write(i, col+1, reportPath)
            else:
                print("HTTP %i - %s, Message %s" % (r.status_code, r.reason, r.text))
        else:
            print("HTTP %i - %s, Message %s" % (r.status_code, r.reason, r.text))

    

def main():
    headers=set_headers()
    x_sap_logontoken =login(URL,headers)
    workbook = xlsxwriter.Workbook('BO_Extraction.xlsx')
    searchReport(URL,x_sap_logontoken,workbook)
    #searchUniverse(URL,x_sap_logontoken,workbook)
    #getUsers(URL,x_sap_logontoken,workbook)
    workbook.close()
# Main execution    
main()
