import utils_mailer as mailer
import utils_coordo_roundcube as ucr
import os
from collections import namedtuple
import time

# Email OP info
Email_info = namedtuple('Email_info', ['host', 'login', 'password'])
email_OP = Email_info(os.getenv('EMAIL_HOST'), 
                      os.getenv('SENDING_EMAIL_OP'), 
                      os.getenv('LOP_PASS'))

# while(True):
# Info google Drive
credentials = os.getenv('GOOGLE_APP_CREDS')

# Get access to google drive
client = mailer.connect_to_drive(credentials)

###########################################################
# Envoi d'un mail automatique de rappel aux mediateurs OP
###########################################################

# Send the planning to the "mediateurs"
#mailer.send_mail_to_mediateurs(client, email_OP)

###########################################################
# Mettre en lien le sheet coordo avec la boite mail de l'OP
###########################################################

# Get all folders in Email - Roundcube Ouvre Porte 
folders = ucr.get_email_folders(email_OP)

# Get email from the last "num_days" 
num_days = 1
mails_roundcube = ucr.get_email_from_last(num_days, folders, email_OP)

# Get only the last sent mail from each sender (if multiple email 
# sent the same day, only the last one will be visible on google sheet )
mail_Dict = ucr.select_last_mail_of_each_sender(mails_roundcube)

# Get the sheet to use 
coordo_sheet = client.open("Coordo/Mediation")

# Mise a jour de toutes les notes sur le sheet
ucr.update_all_note_on_sheet(coordo_sheet, mail_Dict)

# On génère les logs pour garder une trace et débuguer
log_sheet = client.open("helper_python_op").worksheet('logs')
ucr.generate_logs(log_sheet)

    # Sleep
    # time.sleep(600)