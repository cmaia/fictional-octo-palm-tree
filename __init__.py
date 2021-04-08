import logging
from msal.application import ConfidentialClientApplication
from datetime import datetime
import requests
from dateutil import relativedelta
import json

import azure.functions as func


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')
    
    #MS DOCS Graph
    scopes = ['https://graph.microsoft.com/.default']

    #Preparing the service principal to authenticate in MSGraph
    app = ConfidentialClientApplication(
        "<client_id>",
        authority="https://login.microsoftonline.com/<tenant_id>",
        client_credential="<XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX>"
    )
    #Get acess token provided by Azure. Dont forget the API permissions
    result = app.acquire_token_for_client(scopes=scopes)

	currentMonthDT = datetime.today().date()
	oneMonthAgo = currentMonthDT - relativedelta.relativedelta(months=1) # Be carefull with the date format when call MS Graph API, this needs to be in ISO 8601
	#endpointTest = 'https://graph.microsoft.com/beta/users?$filter=signInActivity/lastSignInDateTime le <2021-04-01>&$select=displayName,userPrincipalName'
	endpoint = 'https://graph.microsoft.com/beta/users?$filter=signInActivity/lastSignInDateTime le {0}&$select=displayName,userPrincipalName'.format(str(oneMonthAgo))
	
  #The headers needed to call the API in MSGRaph.
  #In production, I advise that you store the credentials with secure environment variables.
  http_headers = {
	'Authorization': 'Bearer ' + result['access_token'],
	'Accept': 'application/json',
	'Content-Type': 'application/json'
	}
  logging.info('Calling MSGraph - requesting the json')
  data = requests.get(endpoint, headers=http_headers, stream=False).json()
  logging.info('MSGRAPH response stored.')

	if not data['value']:
		return func.HttpResponse('empty', status_code=404)
	else:
    http_headers_response = {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-cache',
    }
		return func.HttpResponse(json.dumps(data), status_code=200, headers=http_headers_response)
  #End of code
