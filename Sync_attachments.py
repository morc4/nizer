from __future__ import print_function
import pickle
import os.path
import base64
import binascii
from googleapiclient.discovery import build
from apiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
from datetime import datetime
import time

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
#########################
def store_doc(service, folderId, file_name, file_path):
	file_metadata = {
	'name': file_name,
	'mimeType': 'application/vnd.google-apps.spreadsheet',
	'parents': [folderId]
        }
	media = MediaIoBaseUpload(file_path, mimetype='application/vnd.google-apps.spreadsheet')
	service.files().create(body=file_metadata, media_body=media, fields = 'id').execute()



#########################
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
            
    #build the services for each google thing
    service2 = build('drive', 'v3', credentials=creds)
    service3 = build('sheets', 'v4', credentials=creds)
    service = build('gmail', 'v1', credentials=creds)
    
    # create main folder for the folders of courses/classes
    folder_metadata = {
    'name': 'Estudiantes',
    'mimeType': 'application/vnd.google-apps.folder'
}
    folder = service2.files().create(body=folder_metadata).execute()
    folder_id = folder.get('id')
    #cursos list should be feed with a for loop taking cursos from lista de estudiantes
    
    # get the id of the spreadsheet that has the student's list
    
    with open('spid.pickle', 'rb') as spreadsheet:
            list_id = pickle.load(spreadsheet)
    
    lista_cursos_cruda = service3.spreadsheets().values().batchGet(spreadsheetId=list_id, ranges=["A2:B999"], majorDimension="ROWS").execute()

    #get the base list out of the sub lists and dictionaries

    rangos_cursos = lista_cursos_cruda["valueRanges"]
    lista_b = rangos_cursos[0]
    lista_b = lista_b["values"]
    curso_dict = {}

    #create and feed dictionary with students
    for pair in lista_b:
        if pair[0] not in curso_dict:
            curso_dict[pair[0]] = [pair[1]]
        else:
            curso_dict[pair[0]].append(pair[1])
    
    print(curso_dict)
   
    #create folders for each classroom
    id_cursos_list = {}
    id_folder_list = []
    
    for curso in curso_dict.keys():
        metadata_folder = {
        'name': curso,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [folder_id]
}
        folder_curso = service2.files().create(body=metadata_folder).execute()
        id_curso = folder_curso.get('id')
        id_cursos_list[curso] = id_curso
        


    #estudiantes list should be feed with a loop from the list after it has been updated by the teacher with the data

    
    id_estudiantes_list = {}
    id_log_list = []
    
    for cursox, lista in curso_dict.items():
        for estudiante in lista:
            metadata_estudiante = {
            'name': estudiante,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [id_cursos_list[cursox]]}
            folder_estudiante = service2.files().create(body=metadata_estudiante).execute()
            
            id_estudiante = folder_estudiante.get('id')
            id_folder_list.append([id_estudiante])
            
            #add students and id of student folder to id_estudiantes_list
            id_estudiantes_list[estudiante] = id_estudiante

            #create student log:
            base_info_student = {
            'properties': {
                'title': estudiante
                }
            }
            log_estudiante = service3.spreadsheets().create(body=base_info_student, fields='spreadsheetId').execute()
            
            logid = log_estudiante.get('spreadsheetId')
            logrange = "A1:E1"
            
            logbody = {
            "value_input_option": "RAW",
    'data': [
    {
  "range": "A1:E1",
  "majorDimension": "COLUMNS",
  "values": [
    ["Fecha"], ["Texto"], ["Link Correo"], ["Link Archivo"], ["ID"]
  ]
}]
}
            make_log = service3.spreadsheets().values().batchUpdate(spreadsheetId=logid, body=logbody).execute()    
      
            id_log_list.append([logid])
      
            #move log to student folder
            move_log = service2.files().update(fileId=logid, addParents=id_estudiante, enforceSingleParent=True).execute()                        
    
        #update student list with the log id of individual student
            
    list_body = {
            "value_input_option": "RAW",
    'data': [
    {
  "range": "D2:D999",
  "majorDimension": "ROWS",
  "values": id_log_list
}]
}

    update_list_with_log = service3.spreadsheets().values().batchUpdate(spreadsheetId=list_id, body=list_body).execute()
    
    #add student folder id to the main students list
    
    folder_id_body = {
            "value_input_option": "RAW",
    'data': [
    {
  "range": "E2:E999",
  "majorDimension": "ROWS",
  "values": id_folder_list
}]
}

    update_list_with_log = service3.spreadsheets().values().batchUpdate(spreadsheetId=list_id, body=folder_id_body).execute()

######################################################################### Mail Part of stuff
    # get the id of the spreadsheet that has the student's list
    
    
            
    lista_cursos_cruda = service3.spreadsheets().values().batchGet(spreadsheetId=list_id, ranges=["B2:E999"], majorDimension="ROWS").execute()
   
    #get the base list out of the sub lists and dictionaries
    
    rangos_cursos = lista_cursos_cruda["valueRanges"]
    lista_b = rangos_cursos[0]
    lista_b = lista_b["values"]
    curso_dict = {}

    #create and feed dictionary with mail (pair[1]) as the key and a list as the value containing name (pair[0]), log id (pair[2]) and folder id (pair[3])
    for pair in lista_b:
        if pair[1] not in curso_dict:
            curso_dict[pair[1]] = [pair[0], pair[2], pair[3]]
        else:
            curso_dict[pair[1]].append([pair[0], pair[2], pair[3]])
    
    #list the mails from the current year belonging to the students in the student's list

    search_query = ""
    
    def ListMessagesMatchingQuery(service=service, user_id="me", query=search_query):
        try:
            response = service.users().messages().list(userId=user_id,
                                               q=query).execute()
            messages = []
            if 'messages' in response:
                messages.extend(response['messages'])

            while 'nextPageToken' in response:
                page_token = response['nextPageToken']
                response = service.users().messages().list(userId=user_id, q=query, pageToken=page_token).execute()
                messages.extend(response['messages'])

            return messages
        except errors.HttpError:
            print('An error occurred: %s' % error)

    #read student log
    
     #los correos vienen revueltos, toca buscar la forma con un for loop de que me los de separados por estudiante. -> por cada estudiante actualizar meta data, correr funcion de leer correos, actualizar log de estudiante, sincronizar archivos adjuntos.
    
    currentYear = datetime.now().year
    
    message_bodies = []
    
    for student_mail in curso_dict.keys():
        search_query = "after: 01/01/{}, from: {}, has:attachment".format(currentYear, student_mail)
        
        mails = ListMessagesMatchingQuery(query=search_query)
        
        #Extract information from messages
        
        for message_id in mails:
            message = service.users().messages().get(userId="me", id=message_id['id']).execute()
            
            #get and format date for the student log
            
            message_date = int(message.get("internalDate"))
            
            actual_date = time.localtime(message_date/1000)
            
            fecha = time.strftime('%d/%m/%Y %H:%M:%S',  actual_date)
            
            # get snippet for the student log
            message_snippet = message.get("snippet")
            
            # message link
            
            message_link = "https://mail.google.com/mail/u/0/#all/{0}".format(message_id['id'])
                        
            #get attachments
                        
            for part in message['payload']['parts']:
                if part['filename']:
                    if 'data' in part['body']:
                        data = part['body']['data']
                    else:
                        att_id = part['body']['attachmentId']
                        att = service.users().messages().attachments().get(userId='me', messageId=message_id['id'], id=att_id).execute()
                        data = att['data']
                    file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                    path = part['filename']
                    
                    #create/download file
                    with open(path, 'wb') as f:
                        f.write(file_data)

                    attachment_metadata = {
                    'name': '{}'.format(path), 
                    'parents': curso_dict[student_mail][2]
                    }
                    
                    #upload file
                    media = MediaFileUpload('{}'.format(path))
                    attachment_file = service2.files().create(body=attachment_metadata, media_body=media, fields='id').execute()
                    attachment_id_in_drive = attachment_file.get('id')
                    
                    #move file to student folder
                    move_attachment = service2.files().update(fileId=attachment_id_in_drive, addParents=curso_dict[student_mail][2], enforceSingleParent=True).execute() 
                    
                    #delete local copy of file in profepro server
                    os.remove(path)
                    
                    #consolidate info to upload to student log, we need date of sent, text, message ID and a link to the message/mail. date = fecha
                    
                    #update sheet
                    
                    student_log_in_dict = curso_dict[student_mail][1]
                     
                    update_log_body = {
                    
                    "value_input_option": "RAW",
                     
                    'data': [
                    {
                    "range": "A2:E2",
                    "majorDimension": "COLUMNS",
                    "values": [[fecha], [message_snippet], [message_link], ["https://drive.google.com/file/d/{}".format(attachment_id_in_drive)], [attachment_id_in_drive]]
}]
}
 
                    insert_row_body = {
                    "requests": 
                    [{"insertRange": {
                    "range": {"sheetId": 0, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 5},
                    "shiftDimension": "ROWS"}

      }]}
                    
                    update_rows = service3.spreadsheets().batchUpdate(spreadsheetId=student_log_in_dict, body=insert_row_body).execute()
                    
                    update_log = service3.spreadsheets().values().batchUpdate(spreadsheetId=student_log_in_dict, body=update_log_body).execute()
                    

############################

if __name__ == '__main__':
    main()
            
