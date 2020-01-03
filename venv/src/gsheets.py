from __future__ import print_function
import os
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient import errors
from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.pickle.
# SCOPES = ['https://www.googleapis.com/auth/script.projects']
# SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
spreadSheetId = "1xRwSYtnmB4Y3_2f8ZPHxLuzMy7WuuZIW8jOY0nsIzN8" # RFS_0AXX Backend
#spreadSheetId = "1yFZTL6BqbUXGBtcSoeXGCsO9mlcT3gxBBGt_v2NWprQ" # Test
#spreadSheetId = "1vJWylr8CFs-sWflGHp6Opmd8FpME_xvbrbOm0PzucwE" # RFS_9Txx Backend

class GSheet:
    def __init__(self):
        self.sheet = None
        self.data = {}

    def getData(self):
        """Calls the Apps Script API.
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
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('sheets', 'v4', credentials=creds)
        self.sheet = service.spreadsheets()
        sheet_props = self.sheet.get(spreadsheetId=spreadSheetId, fields="sheets.properties").execute()
        sheet_names = [ sheet_prop["properties"]["title"] for sheet_prop in sheet_props["sheets"]]

        for sname in sheet_names:
            if not sname.startswith("RFS") and sname != "Email-Verifikation":
                continue
            rows = self.sheet.values().get(spreadsheetId=spreadSheetId, range=sname).execute().get('values', [])
            self.data[sname] = rows
        return self.data

    def addColumn(self, sheetName, colName):
        col = 0;
        for row in self.data[sheetName]:
            col = max(col, len(row))
        self.addValue(sheetName, 0, col, colName);

    def addValue(self, sheetName, row, col, val):
        # row, col are 0 based
        values = [[val]]
        body = {"values": values}
        range = sheetName + "!" + chr(ord('A') + col) + str(row+1)  # 0,0-> A1, 1,2->C2 2,1->B3
        result = self.sheet.values().update(spreadsheetId=spreadSheetId, range=range, valueInputOption = "RAW", body = body).execute()
