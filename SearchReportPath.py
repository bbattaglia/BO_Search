import requests

# Global Variables
username = 'test'
password = 'test' 
URL = 'http://boprd-app01.apps.TEST.net/'
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
        print('OK Universe')
        response=r.json()
        print(response)
    else:
         print("HTTP %i - %s, Message %s" % (r.status_code, r.reason, r.text))


def searchFolder(URL,x_sap_logontoken,folder_id, path = []):
    headers = set_headers_search(x_sap_logontoken)
    query = {
        "query": f"SELECT SI_NAME, SI_ID, SI_PARENTID FROM CI_INFOOBJECTS WHERE SI_ID = '{folder_id}'"
    }
    r = requests.post(URL + 'biprws/v1/cmsquery?pagesize=40000', headers=headers, json=query)
    if r.ok:
        response=r.json()
        entries = response['entries']
        if entries:  # Se ci sono voci, chiamiamo ricorsivamente searchFolder per ciascuna voce
            for entry in entries:
                path.append(entry['SI_NAME'])
                searchFolder(URL, x_sap_logontoken, entry['SI_PARENTID'],path)
        else:
            print("PATH: ", "/".join(path[::-1]))
                
    else:
         print("HTTP %i - %s, Message %s" % (r.status_code, r.reason, r.text)) 



def searchReport(URL,x_sap_logontoken,report_name):
    headers = set_headers_search(x_sap_logontoken)
    query = {
        "query": f"SELECT SI_DESCRIPTION, SI_ID, SI_NAME, SI_OWNER, SI_CREATION_TIME, SI_PARENTID, SI_UPDATE_TS FROM CI_INFOOBJECTS WHERE SI_KIND IN ('Webi') AND SI_NAME LIKE '%{report_name}%'"
    }
    r = requests.post(URL + 'biprws/v1/cmsquery?pagesize=40000', headers=headers, json=query)
    if r.ok:
        response=r.json()
        entries = response['entries']
    else:
         print("HTTP %i - %s, Message %s" % (r.status_code, r.reason, r.text))

    for entry in entries:
        searchFolder(URL,x_sap_logontoken,entry['SI_PARENTID'])


def main():
    headers = set_headers()
    x_sap_logontoken = login(URL, headers)
    
    while True:
        report_name = input("Inserisci il nome del report (o 'exit' per terminare): ")
        
        if report_name.lower() == 'exit':
            print("Programma terminato.")
            break
        
        searchReport(URL, x_sap_logontoken, report_name)

main()
