import requests
import json
import time

#define all the user variables
client_id = 'ENTER CLIENT_ID HERE'
client_secret = 'ENTER CLIENT_SECRET HERE'
username = 'ENTER PARTNER CENTER USERNAME HERE'
password = 'ENTER PARTNER CENTER PASSWORD HERE'

# function to request an access token
def requestAccessToken(client_id, client_secret, username, password):

    data = {'resource': 'https://api.partnercenter.microsoft.com', 'client_id': client_id, 'client_secret': client_secret, 'grant_type': 'password', 'scope': 'openid', 'username': username,'password': password}
    URL = 'https://login.microsoftonline.com/30db41dc-10f8-4a58-99aa-40d995d8ab8d/oauth2/token'
    headers = {'return-client-request-id': 'true', 'Content-Type': 'application/x-www-form-urlencoded'}

    response = requests.post(URL, data=data, headers=headers)
    #print(response.text)

    #extract just the access token from the returned JSON blob
    access_token = response.json()["access_token"]
    refresh_token = response.json()["refresh_token"]
    #other items provided in response: token_type, scope, expires_in, ext_expires_in, expires_on, not_before, resource, id_token, 

    return access_token

access_token = requestAccessToken(client_id, client_secret, username, password)

print(f"The access_token value is:\n{access_token}")

# Function to create query
def createQuery(access_token):
    
    reqUrl = "https://api.partnercenter.microsoft.com/insights/v1/mpn/ScheduledQueries"

    headersList = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    #payload = "{\r\n  \"Name\": \"All Training 070122a\",\r\n  \"Description\": \"List of all training items\",\r\n  \"Query\": \"SELECT PGAMpnId,TrainingActivityId,TrainingTitle,TrainingType,IndividualFirstName,IndividualLastName,Email,CorpEmail,TrainingCompletionDate,ExpirationDate,ActivationStatus,Month,IcMCP,MCPID,MPNId,PartnerName,PartnerCityLocation,PartnerCountryLocation from TrainingCompletions TIMESPAN LIFETIME\"\r\n}"
    #payload = "{\r\n  \"Name\": \"Office 365 Usage\",\r\n  \"Description\": \"List of all Office 365 usage\",\r\n  \"Query\": \"SELECT PGAMpnId,CustomerTenantId,CustomerTpid,WorkloadName,Month,PaidAvailableUnits,MonthlyActiveUsers,CustomerName,CustomerMarket,CustomerSegment,MPNId,PartnerName,PartnerLocation,PartnerAttributionType,IsDuplicateRowForPGA from OfficeUsage TIMESPAN LIFETIME\"\r\n}"
    payload = "{\r\n  \"Name\": \"All Training\",\r\n  \"Description\": \"List of all training items\",\r\n  \"Query\": \"SELECT TrainingActivityId from TrainingCompletions TIMESPAN LAST_6_MONTHS\"\r\n}"

    response = requests.request("POST", reqUrl, data=payload,  headers=headersList)

    #print(response.text)

    responsejson = json.loads(response.text)

    #print(responsejson)

    #extracting the queryId from the returned JSON object using this tutorial: https://stackoverflow.com/questions/23306653/python-accessing-nested-json-data
    queryId = responsejson['value'][0]['queryId']

    return queryId

queryId = createQuery(access_token)

print(f"\nThe queryId is {queryId}")

#Function to create the report

def createQuery(access_token, queryId):

    reqUrl = "https://api.partnercenter.microsoft.com/insights/v1/mpn/ScheduledReport"

    headersList = {
    "Accept": "application/json",
    "User-Agent": "Thunder Client (https://www.thunderclient.io)",
    "X-AadUserTicket": access_token,
    "Content-Type": "application/json" 
    }

    payload = "{\n  \"ReportName\": \"All Training Report\",\n  \"Description\": \"Report listing all the training activities\",\n  \"QueryId\": \"" + queryId + "\",\n  \"ExecuteNow\": true\n}"

    response = requests.request("POST", reqUrl, data=payload,  headers=headersList)

    #print(response.text)
    reportId = json.loads(response.text)['value'][0]['reportId']
    print("\n" + json.loads(response.text)['message'])
    print(f"\nThe reportId is {reportId}")

    return reportId

reportId = createQuery(access_token, queryId)

#This function checks the status of the report, and the URL (if ready)
def checkExecutionStatus(access_token, reportId):
    reqUrl = "https://api.partnercenter.microsoft.com/insights/v1/mpn/ScheduledReport/execution/" + reportId
    headersList = {
    "X-AadUserTicket": access_token,
    "Accept": "application/json" 
    }

    payload = ""

    response = requests.request("GET", reqUrl, data=payload,  headers=headersList)

    #print(response.text)
    executionStatus = json.loads(response.text)['value'][0]['executionStatus']
    reportLink = json.loads(response.text)['value'][0]['reportAccessSecureLink']

    print(f"\nExecution status: {executionStatus}")
    
    return executionStatus, reportLink

#Function to download a file from a URL - used in the downloadReport Function
def saveLink(URL, outfile):
    the_book = requests.get(URL, stream=True)
    with open(outfile, 'wb') as f:
      for chunk in the_book.iter_content(1024 * 1024 * 2):  # 2 MB chunks
        f.write(chunk)

#Function to continually check whether the report is ready to be downloaded
def downloadReport():
    executionStatus = "Pending"

    while executionStatus != "Completed":
        executionStatus, reportLink = checkExecutionStatus(access_token, reportId)

        if executionStatus == "Completed":
            saveLink(reportLink,"TrainingReport.csv")
            print(reportLink)
            print("Report successfully downloaded")
            break
        elif executionStatus == "Pending":
            print("Sleeping for 60 seconds before retrying...")
            time.sleep(60)

downloadReport()
