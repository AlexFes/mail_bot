from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from utils.client import EmailClient
from utils.mail import Email

# usage: export email_addr=youraddress@yandex.ru email_passwd=yourpassword sheet_id=googledocid
email_addr = os.environ['email_addr']
email_passwd = os.environ['email_passwd']
sheet_id = os.environ['sheet_id']

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def authenticate():
    creds = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds


def get_inbox():
    pass


def get_sheet_lines():
    pass


def write_line():
    pass


def start(sheet):
    result = sheet.values().get(spreadsheetId=sheet_id,
                                range='B3:B10').execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print(values)


if __name__ == '__main__':
    credentials = authenticate()
    service = build('sheets', 'v4', credentials=credentials)
    spreadsheet = service.spreadsheets()

    start(spreadsheet)


# if __name__ == "__main__":
#     with EmailClient(email_addr, email_passwd) as client:
#         count = client.get_mails_count()
#
#         for i in range(1, count +1):
#             mail = client.get_mail_by_index(i)
#             content = mail.__repr__()
#             print(content)
