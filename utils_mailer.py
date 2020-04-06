from datetime import datetime
import smtplib
from email.message import EmailMessage
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
from collections import namedtuple

def send_mail_to_mediateurs(client, email_info):
    mediation_sheet = client.open("helper_python_op").worksheet('liste_mediation')
    log_sheet       = client.open("helper_python_op").worksheet('logs')
    coordo_sheet    = client.open("Coordo/Mediation")
    Accueil = namedtuple('Accueil', ['Mediateur', 'Accueilli', 'Mail_mediateur'])
    mediations = [Accueil(*accueil) for accueil in mediation_sheet.get_all_values()]

    for mediation in mediations:
        sheet      = coordo_sheet.worksheet(mediation.Accueilli)
        table_html = get_planning(sheet)
        html       = create_email_mediation(mediation, table_html)
        send_email(mediation, html, email_info)
        generate_logs(log_sheet, mediation)
        break

    print("Mails envoyes...")
        # time.sleep(60)


def create_email_mediation(info_accueil, table_html):
    ######################################
    # Write the HTML email
    with open('style.css', 'r') as f:
        style = f.read()

    html = f'''
    <html>
    <head>
        <style>{style}</style>
    </head>
    <body>
        <p> Bonjour {info_accueil.Mediateur}, </p>
        <p> 
            Ceci est un message automatique de rappel concernant la mediation de <b> {info_accueil.Accueilli} </b> <br>
            Inutile de repondre a cet email.
        </p>
        <p> Voici le planning actuel : <p>
        {table_html}
        <p> En cas de besoin le tableau de suivi est situe ici : https://docs.google.com/spreadsheets/d/1U5etwfJKZMBZxaDkRMg1qhkzfRSyZ3RmrYpyf-GRrY8/edit#gid=0<p>

    </body>
    </html>
    '''
    return html

def get_planning(sheet):
    ######################################
    # Process the data
    begin_row = 0
    for i, val in enumerate(sheet.get_all_values()): 
        if  any(list(map(lambda x: 'ACCUEIL' in x, val))):
            begin_row = i+1
            break

    mail_content_all_cols = sheet.get_all_values()[begin_row:]
    mail_content = [x[:3] for x in mail_content_all_cols]
    
    html = "<table class=mystyle>"
    html += "<tr>"
    html += "\n".join(map(lambda x: "<th>" + x + "</th>", mail_content[0])) 
    html += "<tr>"
    for row in mail_content[1:]:
       html += "<tr>"
       # Make <tr>-pairs, then join them.
       html += "\n".join(map(lambda x: "<td>" + x + "</td>", row)) 
       html += "</tr>"
    html += "</table>"

    return html


def send_email(info_accueil, html, email_info):
    import os
    # email_host = os.getenv('EMAIL_HOST')

    print(f'[Ouvre-Porte] Mediation {info_accueil.Accueilli} - Mail automatique')
    ######################################
    # Send the email
    msg = EmailMessage()
    msg['Subject'] = f'[Ouvre-Porte] Mediation {info_accueil.Accueilli} - Mail automatique'
    msg['From'] = email_info.login
    msg['To'] = info_accueil.Mail_mediateur

    msg.add_alternative(html, subtype='html')

    with smtplib.SMTP_SSL(email_info.host, 465) as smtp:
        # PASSWORD = os.getenv('LOP_PASS')
        smtp.login(email_info.login, email_info.password)
        smtp.send_message(msg)


def generate_logs(log_sheet, infos):
    ######################################
    # Log for debugging purposes
    now = datetime.today()
    now = now.strftime("%d-%m-%Y")
    log_sheet.append_row([now.__str__(), infos.Mediateur], value_input_option='RAW')


def connect_to_drive(creds_raw):
    import json 
    scope = ["https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive"]

    creds_json = json.loads(creds_raw)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
    client = gspread.authorize(creds)
    return client