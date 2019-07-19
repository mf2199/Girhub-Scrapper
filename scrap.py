from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import github_utils
import sheet

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def authenticate():
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
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)
    return service


service = authenticate()

sheet_id = sheet.create_new_sheet(
    service, 'QLogic Python Internal Review'
)

sheet.create_title_row(
    service,
    sheet_id,
    [
        ['Priority', 80, 'CENTER', ['Paused', 'Low', 'Medium', 'High', 'Critical', 'Done']],
        ['Issue', 50, 'CENTER', None],
        ['Status', None, 'CENTER', ['PENDING', 'OPEN', 'SUBMITTED', 'ON HOLD', 'MERGED', 'REOPENED', 'CLOSED', 'IRRELEVANT']],
        ['Created', None, 'CENTER', None],
        ['Description', 450, None, None],
        ['Repository', None, 'CENTER', None],
        ['API', None, 'CENTER', None],
        ['Assignee', None, 'CENTER', ['Ilya', 'Hemang', 'Maxim', 'Paras', 'Sumit', 'Sangram', 'N/A']],
        ['Task', None, 'CENTER', None],
        ['Opened', None, 'CENTER', None],
        ['Internal PR', None, 'CENTER', None],
        ['Public PR', None, 'CENTER', None],
        ['Comment', 550, None, None],
    ]
)

rows, count = github_utils.build_whole_table()

sheet.save_to_sheet(service, sheet_id, rows, count)

# print(sheet.read_sheet(service, sheet_id, 'Лист1'))

# sheet.update_list(
#     service,
#     '1uLnTYWCePtuvqDaFbSqIqqZfHzwbBnmaR_Pyqm5D6HM',
#     'Лист1'
# )
