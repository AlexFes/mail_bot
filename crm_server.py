from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from datetime import datetime

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


def read_line(sheet, row):
    result = sheet.values().get(
        spreadsheetId=sheet_id, range="A{}:P{}".format(row, row)).execute()
    return result.get('values', [])


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
            content = mail.get_line_data()
            inbox_list.append(content)

        return inbox_list


def get_table(inbox):
    table = {}
    result = []

    for mail in reversed(inbox):
        sender, subject, date = mail
        date_obj = datetime.strptime(date, '%a, %d %b %Y %H:%M:%S %z')

        if sender[1] not in table:
            table[sender[1]] = (sender[0], subject, str(date_obj.date()))
        else:
            continue

    for key, value in table.items():
        result.append(['', value[2], key, value[0], '', '', '', '', '', '', '', '', value[1], '', '', ''])

    return result


def update_sheet(sheet, table):
    for ind, line in enumerate(table):
        write_line(sheet, ind + 3, line)


def start(sheet):
    new_table = get_table(get_inbox())
    update_sheet(sheet, new_table)


if __name__ == '__main__':
    # Authenticate
    credentials = authenticate()
    service = build('sheets', 'v4', credentials=credentials)
    spreadsheet = service.spreadsheets()

    # Setup header row
    write_line(spreadsheet, 1, header)

    # Start polling yandex mail
    start(spreadsheet)
