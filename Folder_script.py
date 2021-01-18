#Meta actual: modificar linea 31 "flow = InstalledAppFlow.from_client_secrets_file('client_id.json', SCOPES)" para que funcione con web en vez de installed.

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
#########################
#pickle for credentials and oauth
def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_id.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    #build spreadsheet service
    service3 = build('sheets', 'v4', credentials=creds)
    
    #create file for a student list   
    spreadsheet = {
    'properties': {
        'title': 'lista'
    }
}
    spreadsheet = service3.spreadsheets().create(body=spreadsheet,
                                    fields='spreadsheetId').execute()
    spid = spreadsheet.get('spreadsheetId')
#    sprange = "A1:C1"
    
    #pickle the spreadsheet id for later
    with open('spid.pickle', 'wb') as spreadsheet:
            pickle.dump(spid, spreadsheet)
            
    batch_update_values_request_body = {
    # The new values to apply to the spreadsheet.
    "value_input_option": "RAW",
    'data': [
    {
  "range": "A1:E1",
  "majorDimension": "COLUMNS",
  "values": [
    ["Curso"], ["Estudiante"], ["Correo Electronico"], ["ID de Registro"], ["ID de Carpeta"]
  ]
}]
}

    
    request = service3.spreadsheets().values().batchUpdate(spreadsheetId=spid, body=batch_update_values_request_body).execute()
    


# leer el archivo y crear las listas en el programa


    
#Execute    
if __name__ == '__main__':
    main()
