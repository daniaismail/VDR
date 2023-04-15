import imaplib
import email
import smtplib
import sys
from datetime import datetime, timedelta
import os
import fnmatch
from time import strftime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import decode_header
import logging

today = datetime.today()
t = today.strftime("%d%m%y")
logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%d%m%Y %I:%M:%S%p', filename=f"C:/Users/mvmwe/PycharmProjects/VDR/log/log {t}.txt", level=logging.DEBUG)
yesterday = datetime.today() - timedelta(1)
cutoff = today - timedelta(days=10)
dt = cutoff.strftime('%d-%b-%Y')
file_path = r'C:\Users\mvmwe\Dropbox\MVMCC\VDR'
folder_year_path = os.path.join(file_path, str(yesterday.year))
folder_month_path = os.path.join(folder_year_path, strftime("%b").upper())
error_file = []

if not os.path.exists(file_path):
    os.mkdir(file_path)

if not os.path.exists(folder_year_path):
    os.mkdir(folder_year_path)

if not os.path.exists(folder_month_path):
    os.mkdir(folder_month_path)

message = "VDR Directory: " + folder_month_path
# EMAIL INFO
email_vdr = "vdr@meridiansurveys.com.my"
pwd_vdr = "T%zf5ccq;ZMc"
server_mssb = "meridian-svr.meridiansurveys.com.my"
email_wgg = "mvmcc@wild-geese-group.com"
pwd_wgg = "s9nPviD\\"
server_wgg = "mail.wild-geese-group.com"

# READ FROM TXT FILE AND APPEND INTO LIST
vesselemail_list = []
clientname_list = []
vesselname_list = []
vesselemail_wgg_list = []
clientname_wgg_list = []
vesselname_wgg_list = []

logging.debug('Logging started..')
logging.debug('Open email-address-vdr.txt')
try:
    with open(r'C:\Users\mvmwe\PycharmProjects\VDR\email-address-vdr.txt') as f:
        try:
            for line in f:
                vesselemail, vesselname, clientname = line.split(',')
                vesselemail_list.append(vesselemail)
                vesselname_list.append(vesselname)
                clientname_list.append(clientname.replace("\n", ""))
        except:
            logging.error('Error info in text file. Check for typo')
except:
    logging.exception('Error info in text file or check txt file name')
    sys.exit()

logging.debug('Open email-address-wgg.txt')
try:
    with open(r'C:\Users\mvmwe\PycharmProjects\VDR\email-address-wgg.txt') as f:
        try:
            for line in f:
                vesselemailwgg, vesselnamewgg, clientnamewgg = line.split(',')
                vesselemail_wgg_list.append(vesselemailwgg)
                vesselname_wgg_list.append(vesselnamewgg)
                clientname_wgg_list.append(clientnamewgg.replace("\n", ""))
        except:
            logging.error('Error info in text file. Check for typo')
except:
    logging.exception('Error info in text file or check txt file name')
    sys.exit()

def dwl_vdr(email_add, password, server, vesselEmail, vesselName, clientName):
    imap = imaplib.IMAP4_SSL(server, 993)
    imap.login(email_add, password)
    imap.select('INBOX')

    index = 0

    while index < len(vesselEmail) and index < len(vesselName):
        try:
            client_folder = os.path.join(folder_month_path, clientName[index])

            if not os.path.exists(client_folder):
                os.mkdir(client_folder)

            folder_name = os.path.join(client_folder, vesselName[index])
            if not os.path.exists(folder_name):
                os.mkdir(folder_name)
            logging.debug(f'Download from email {vesselEmail[index]}')
            typ, data = imap.search(None, '(SINCE %s)' % (dt,), '(FROM %s)' % (vesselEmail[index],))

            for num in data[0].split():
                typ, data = imap.fetch(num, '(RFC822)')
                raw_email = data[0][1]
                raw_email_string = raw_email.decode('ISO-8859–1')
                email_message = email.message_from_string(raw_email_string)
                subject_name = email_message['subject']

                # att_path = "No attachment found from email " + subject_name
                logging.debug(f'Email subject: {subject_name}')
                for part in email_message.walk():
                    try:
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue

                        fileName = part.get_filename()
                        if fnmatch.fnmatch(fileName, "*.xls*"):
                            try:
                                att_path = os.path.join(folder_name, fileName)
                                print(att_path)
                                if not os.path.isfile(att_path):
                                    fp = open(att_path, "wb")
                                    fp.write(part.get_payload(decode=True))
                                    fp.close()
                            except TypeError as e:
                                error_file.append(fileName)
                                continue
                            logging.debug(f'File downloaded: {fileName}')

                        elif fnmatch.fnmatch(fileName, "*.doc*"):
                            try:
                                att_path = os.path.join(folder_name, fileName)
                                print(att_path)
                                if not os.path.isfile(att_path):
                                    fp = open(att_path, "wb")
                                    fp.write(part.get_payload(decode=True))
                                    fp.close()
                            except TypeError as e:
                                error_file.append(fileName)
                                continue
                            logging.debug(f'File downloaded: {fileName}')

                        elif fnmatch.fnmatch(fileName, "*.pdf"):
                            try:
                                att_path = os.path.join(folder_name, fileName)
                                print(att_path)
                                if not os.path.isfile(att_path):
                                    fp = open(att_path, "wb")
                                    fp.write(part.get_payload(decode=True))
                                    fp.close()
                            except TypeError as e:
                                error_file.append(fileName)
                                continue
                            logging.debug(f'File downloaded: {fileName}')

                        else:
                            logging.debug(f'No file downloaded from {subject_name}')

                    except (OSError, TypeError) as f:
                        print(f)
                        continue

                # print(att_path)

                for response in data:
                    if isinstance(response, tuple):
                        data = email.message_from_bytes(response[1])
                        subject = decode_header(data["Subject"])[0][0]
                        if isinstance(subject, bytes):
                            # if it's a bytes type, decode to str
                            subject = subject.decode()
                        logging.debug(f'Deleting email {subject}')
                imap.store(num, "+FLAGS", "\\Deleted")

            index += 1

        except (OSError, TypeError) as k:
            print(k)
            continue

    imap.expunge()
    imap.close()
    imap.logout()
    logging.debug('Job finished..')

def send_email(finishmessage):
    msg = MIMEMultipart()
    message = str(finishmessage)

    #message parameters
    password = "dc)in]}Xzk&%"
    msg['From'] = "mvmcc@meridiansurveys.com.my"
    msg['To'] = "mvmcc@meridiansurveys.com.my"
    msg['Subject'] = "VDR Download"

    msg.attach(MIMEText(message, 'plain'))

    server = smtplib.SMTP('meridian-svr.meridiansurveys.com.my: 587')
    server.starttls()

    server.login(msg['From'], password)

    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()

try:
    logging.debug('Downloading from vdr@meridiansurveys.com.my..')
    dwl_vdr(email_vdr, pwd_vdr, server_mssb, vesselemail_list, vesselname_list, clientname_list)
    logging.debug('Downloading from mvmcc@wild-geese-group.com..')
    dwl_vdr(email_wgg, pwd_wgg, server_wgg, vesselemail_wgg_list, vesselname_wgg_list, clientname_wgg_list)
    send_email(message)

except Exception as e:
    logging.error("Error occurred", exc_info=True)
