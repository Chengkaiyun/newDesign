from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import globals as Var
import sheetDC
import sheetUF
import newDesign as index

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1BD1A01VL2Bx_bMBgoc3eCwb9cT4BWI2fzuQdQop8yUo'
deviceWhich_RANGE_NAME = 'device b2c!A1:Z'
designer_RANGE_NAME = 'designer!A1:Z'
color_RANGE_NAME = 'SSA!A2:Z'


def main():
    global SPREADSHEET_ID, deviceWhich_RANGE_NAME, designer_RANGE_NAME, color_RANGE_NAME, sheet

    creds = None

    # token.pickle stores the user's access and refresh tokens, and is created automatically
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    # =============================== upload file用 =============================== #

    Var.needColor = []
    Var.allColor = []
    Var.allBrand = []
    Var.ChooseSheet = ""
    Var.deviceB2C = []
    Var.deviceB2B = []
    Var.deviceKroma = []
    Var.deviceWhich = []
    Var.allBrandB2C = []
    Var.allBrandB2B = []
    Var.allBrandKroma = []
    Var.SEcountry = ['tw', 'jp', 'other']
    sheetUF.mainUF()
    index.okLabel("UFno")

    # =============================== delete color 用 =============================== #
    Var.needColor = []
    Var.allColor = []
    Var.preOrder = []
    Var.preOrderDevice = []
    Var.preOrderDate = []
    Var.Allcountry = ['tw/b2b', 'io/eu', 'fr', 'de', 'jp', 'es']
    sheetDC.mainDC()
