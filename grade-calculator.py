#!/usr/bin/env python
# coding: utf-8

# In[16]:


from __future__ import print_function

import math
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1ec7TLRpt8YBE2hrmtuGXNm6409XMvlsTqXEt2L55Pmw'
SAMPLE_RANGE_NAME = 'engenharia_de_software!A4:G27'


def main():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    # Ler informacoes do Google Sheets
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    sheet_values = result['values']

    # ALL THE CODE ABOVE WAS INSPIRED BY TUTORIALS AND APIs DOCUMENTATION THAT I STUDIED TO COMPLETE THIS CHALLENGE
    # https://developers.google.com/sheets/api/guides/concepts
    # https://www.hashtagtreinamentos.com/integracao-do-python-com-google-sheets-python?gad_source=1&gclid=Cj0KCQiAw6yuBhDrARIsACf94RWkUFLIqcxWsUDffYDRjafKu0S-hKamiXL7WPKh85Qd1i4XQEEFLvUaAv0DEALw_wcB
    # https://developers.google.com/sheets/api/quickstart/python?hl=pt-br
    # https://www.youtube.com/watch?v=l7pL_Y3fw-o

    absence_limit = 15  # 25% of the total number of classes wich is 60

    add_values = []

    for line in sheet_values:
        absences = line[2]
        if int(absences) > absence_limit:       # checking if the student failed due to absence
            add_values.append(["Reprovado por Falta", "0"])
        else:
            avg = (int(line[3]) + int(line[4]) + int(line[5])) / 3
            if int(avg) < 50:                   # checking if the student failed by grade
                add_values.append(["Reprovado por Nota", "0"])
            elif int(avg) < 70:                 # checking if the student has to do the final test and calculating
                approval_grade = 100 - avg      # how much he needs to pass
                approval_grade = math.ceil(approval_grade)
                add_values.append(["Exame Final", str(approval_grade)])
            elif int(avg) >= 70:                # checking if the student was approved
                add_values.append(["Aprovado", "0"])

    sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,          # updating the values in the sheet through
                          range="G4", valueInputOption="USER_ENTERED",  # the API passing the necessary parameters
                          body={'values': add_values}).execute()


if __name__ == '__main__':
    main()
