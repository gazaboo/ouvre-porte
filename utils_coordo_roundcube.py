import imaplib
import pprint
from imap_tools.imap_utf7 import decode
from imap_tools import MailBox, Q, OR
import datetime as dt 
import re
from collections import namedtuple
from gspread.urls import SPREADSHEETS_API_V4_BASE_URL


def get_email_folders(email_info):
    # connect to host using SSL and login to server
    imap = imaplib.IMAP4_SSL(email_info.host)
    imap.login(email_info.login, email_info.password)

    # Get the folders
    raw_folders = [decode(folder) for folder in imap.list()[1]]
    folders = [folder.split(' "/" ')[1].replace('"','') for folder in raw_folders]

    return folders


def get_email_from_last(num_days: 1, folders, email_info):
    Mail = namedtuple('Mail', ['from_', 'date', 'object', 'body'])
    mails_roundcube = []

    with MailBox(email_info.host).login(email_info.login, email_info.password) as mailbox:
        for folder in folders:
            mailbox.folder.set(folder)
            query_result = [Mail(msg.from_, msg.date,msg.subject, msg.text) 
                                for msg in mailbox.fetch(Q(date_gte=dt.date.today() - dt.timedelta(days=num_days)))]
            mails_roundcube = mails_roundcube + query_result
    
    return mails_roundcube  

def select_last_mail_of_each_sender(mails_roundcube):
    Mail = namedtuple('Mail', ['from_', 'date', 'object', 'body'])
    
    # On recupère le sender de chaque mail 
    senders = [sender.from_ for sender in mails_roundcube]

    # Keep only the last one
    mail_Dict = {}
    for sender in senders:
        last_mails_of_sender = Mail(sender, dt.datetime.today() - dt.timedelta(days=180), '', '')
        for mail in mails_roundcube:
            if (mail.date.date() >= last_mails_of_sender.date.date()) and mail.from_==sender:
                last_mails_of_sender = mail
        mail_Dict[sender] = last_mails_of_sender

    return mail_Dict


def update_all_note_on_sheet(coordo_sheet, mail_Dict):
    
    worksheet_list = coordo_sheet.worksheets()
    for sheet in worksheet_list:
        try:
            col_mails          = sheet.find("Mails accueillants").col
            col_date_echange   = sheet.find("Dernier mail reçu").col
            liste_accueillants_raw = sheet.col_values(col_mails)
            
            # Get all cells "colonne dernier échange"
            cell_list = sheet.range(1, col_date_echange, 300, col_date_echange)

            # On met la date du jour en face et la note
            for sender, mail_recu in mail_Dict.items():
                for i, mail_accueillants in enumerate(liste_accueillants_raw):
                    found_row = re.search(sender.lower(), mail_accueillants.lower())
                    if found_row is not None and len(mail_recu.body)>0: 
                        cell_list[i].value = mail_recu.date.strftime('%d %B, %Y')
                        note = mail_recu.body.replace('\r', ' ')
                        insert_note(sheet, i, col_date_echange-1, note)
                        print("Sheet : ", sheet, " - Mail reçu --> ", mail_recu.from_)
            sheet.update_cells(cell_list)
        except:
            pass


def insert_note(worksheet, row, col, note):
    """
    Insert note into the google worksheet for a certain cell.
    
    Compatible with gspread.
    """
    spreadsheet_id = worksheet.spreadsheet.id
    worksheet_id = worksheet.id

    url = f"{SPREADSHEETS_API_V4_BASE_URL}/{spreadsheet_id}:batchUpdate"
    payload = {
        "requests": [
            {
                "updateCells": {
                    "range": {
                        "sheetId": worksheet_id,
                        "startRowIndex": row,
                        "endRowIndex": row + 1,
                        "startColumnIndex": col,
                        "endColumnIndex": col + 1
                    },
                    "rows": [
                        {
                            "values": [
                                {
                                    "note": note
                                }
                            ]
                        }
                    ],
                    "fields": "note"
                }
            }
        ]
    }
    worksheet.spreadsheet.client.request("post", url, json=payload)


def generate_logs(log_sheet):
    ######################################
    # Log for debugging purposes
    now = dt.datetime.today()
    now = now.strftime("%H:%M, le %d-%m-%Y")
    log_sheet.append_row([now, "MAJ - sheet coordo"], value_input_option='RAW')
