import os.path
import pickle
import re
import json

from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.exceptions import RefreshError
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from conf import SPREADSHEET_CREDENTIALS_FILE_PATH, SPREADSHEET_TOKEN_FILE_PATH
from errors import handle_error


def validate_sheet_range(s):
    valid_range = re.compile(
        r'(?P<sheet>.+)!(?P<col1>[A-Z]+)(?P<row1>\d+)(:(?P<col2>[A-Z]+)(?P<row2>\d+))?')
    return valid_range.match(s)


def read_sheet_range(api, spreadsheet_id, sheet_range):
    try:
        response = (api.values()
                    .get(spreadsheetId=spreadsheet_id, range=sheet_range)
                    .execute()
                    .get('values', [['0']]))

        # responses trim blank cells, act as if they are 0-filled
        match = validate_sheet_range(sheet_range)
        length = int(match.group('row2')) - int(match.group('row1')) + 1
        for lst in response:
            if len(lst) < 1:
                lst.append('0')
        flat = [val.strip().lower() for row in response for val in row]
        while len(flat) < length:
            flat.append('0')

        return flat

    except HttpError as error:
        handle_error('sheets_api', val=error._get_reason())


def write_to_cell(api, spreadsheet_id, cell, val):
    try:
        api.values().update(
            spreadsheetId=spreadsheet_id,
            range=cell,
            valueInputOption='RAW',
            body={'values': [[val]]}
        ).execute()
    except HttpError as error:
        handle_error('sheets_api', val=error._get_reason())


# https://developers.google.com/sheets/api/quickstart/python
def create_service():
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets']

        if not os.path.exists(SPREADSHEET_CREDENTIALS_FILE_PATH):
            handle_error('no_credentials')

        with open(SPREADSHEET_CREDENTIALS_FILE_PATH, 'r') as f:
            credentials_info = json.load(f)

        if 'installed' in credentials_info or 'web' in credentials_info:
            creds = None
            if os.path.exists(SPREADSHEET_TOKEN_FILE_PATH):
                with open(SPREADSHEET_TOKEN_FILE_PATH, 'rb') as token:
                    creds = pickle.load(token)
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    try:
                        creds.refresh(Request())
                    except RefreshError:
                        os.remove(SPREADSHEET_TOKEN_FILE_PATH)
                        flow = InstalledAppFlow.from_client_secrets_file(
                            SPREADSHEET_CREDENTIALS_FILE_PATH, scopes)
                        creds = flow.run_local_server(port=0)
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        SPREADSHEET_CREDENTIALS_FILE_PATH, scopes)
                    creds = flow.run_local_server(port=0)
                with open(SPREADSHEET_TOKEN_FILE_PATH, 'wb') as token:
                    pickle.dump(creds, token)
        elif credentials_info.get('type') == 'service_account':
            creds = Credentials.from_service_account_file(
                SPREADSHEET_CREDENTIALS_FILE_PATH, scopes=scopes)
        else:
            handle_error('invalid_credentials')

        service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
        return service.spreadsheets()

    except HttpError as error:
        handle_error('sheets_api', val=error._get_reason())
    except Exception as error:
        handle_error('sheets_api', val=str(error))
