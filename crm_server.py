from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from utils.client import EmailClient
from utils.mail import Email

# Usage: export email_addr= email_passwd= sheet_id=
email_addr = os.environ['email_addr']
email_passwd = os.environ['email_passwd']
sheet_id = os.environ['sheet_id']
header = ["Project", "Sent", "Email", "Name",
           "Source", "Direction", "Cost, $", "Payment", "Currency",
           "Deadline", "Manager", "Status", "Notes",
           "Reason for lost/Shifted start", "CLIENT | Data of last letter", "OUR | Data of last letter"]

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


def read_line(sheet):
    pass


def write_line(sheet, row, line):
    header_body = {
        "values": [line]
    }

    result = sheet.values().update(spreadsheetId=sheet_id,
                        range="A{}:P{}".format(row, row),
                        valueInputOption="RAW",
                        body=header_body).execute()

    print('{} cells updated at row={}'.format(result.get('updatedCells'), row))
    

def write_our_date(sheet, row, date):
    pass


def write_client_date(sheet, row, date):
    pass


def get_inbox():
    with EmailClient(email_addr, email_passwd) as client:
        count = client.get_mails_count()
        inbox_list = []

        for i in range(1, count + 1):
            mail = client.get_mail_by_index(i)
            content = mail.__repr__()
            inbox_list.append(content)

        return inbox_list


def start(sheet):
    # Setup header row
    write_line(sheet, 1, header)

if __name__ == '__main__':
    # Authenticate
    credentials = authenticate()
    service = build('sheets', 'v4', credentials=credentials)
    spreadsheet = service.spreadsheets()

    # Start polling yandex mail
    start(spreadsheet)
