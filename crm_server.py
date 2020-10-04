from __future__ import print_function
import pickle
import time
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


def write_line(sheet, row, line):
    header_body = {
        "values": [line]
    }

    result = sheet.values().update(spreadsheetId=sheet_id,
                        range="A{}:P{}".format(row, row),
                        valueInputOption="RAW",
                        body=header_body).execute()

    print('{} cells updated at row={}'.format(result.get('updatedCells'), row))
    

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
    inbox.sort(key=lambda tup: tup[2])

    for mail in inbox:
        sender, subject, date, to = mail

        # Client mail
        if sender[1] != email_addr:
            # New line
            if sender[1] not in table:
                table[sender[1]] = {
                    "sender": sender[0],
                    "subject": subject,
                    "date": str(date.date()),
                    "client_date": str(date.date()),
                    "our_date": ""
                }
            # Update line
            else:
                table[sender[1]]["client_date"] = str(date.date())

        # Our mail
        else:
            # New line
            if to[1] not in table:
                table[to[1]] = {
                    "sender": to[0],
                    "subject": subject,
                    "date": str(date.date()),
                    "client_date": "",
                    "our_date": str(date.date())
                }
            # Update line
            else:
                table[to[1]]["our_date"] = str(date.date())

    for key, value in table.items():
        result.append([value["date"], key, value["sender"],
                       value["subject"], value["client_date"], value["our_date"]])

    return result


def update_sheet(sheet, table):
    data = []

    for ind, line in enumerate(table):
        data.append({
            "range": "B{}:D{}".format(ind + 3, ind + 3),
            "values": [line[:3]]
        })
        data.append({
            "range": "M{}".format(ind + 3),
            "values": [[line[3]]]
        })
        data.append({
            "range": "O{}:P{}".format(ind + 3, ind + 3),
            "values": [line[4:]]
        })

    body = {
        "valueInputOption": "RAW",
        "data": data
    }

    result = sheet.values().batchUpdate(spreadsheetId=sheet_id, body=body).execute()
    print('{0} cells updated.'.format(result.get('totalUpdatedCells')))


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
    while True:
        start(spreadsheet)
        time.sleep(10)