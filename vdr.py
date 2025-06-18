import imaplib
import email
import smtplib
import sys
import calendar
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
yesterday = today - timedelta(days=16)
year = yesterday.year
y_month = yesterday.month
month = calendar.month_abbr[y_month].upper()
dt = yesterday.strftime('%d-%b-%Y')
logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%d%m%Y %I:%M:%S%p', filename=f"C:/Users/user/PycharmProjects/VDR/log/log {t}.txt", level=logging.DEBUG)
file_path = r'C:/Users/user/Dropbox/MVMCC\VDR'
folder_full_path = os.path.join(file_path, str(year), str(month))
error_file = []

if not os.path.exists(file_path):
    os.makedirs(file_path)
else:
    print(f"Folder '{file_path}' already exists")

if not os.path.exists(folder_full_path):
    os.makedirs(folder_full_path)
else:
    print(f"Folder '{folder_full_path}' already exists")

message = "VDR Directory: " + folder_full_path
# EMAIL INFO
email_vdr = "vdr@meridiansurveys.com.my"
pwd_vdr = "T%zf5ccq;ZMc"
email_mvmcc = "mvmcc@meridiansurveys.com.my"
pwd_mvmcc = "dc)in]}Xzk&%"
server_mssb = "meridian-svr.meridiansurveys.com.my"

# READ FROM TXT FILE AND APPEND INTO LIST
vesselemail_list = []
clientname_list = []
vesselname_list = []

logging.debug('Logging started..')
logging.debug('Open email-address.txt')
try:
    with open(r'C:\Users\user\PycharmProjects\VDR\email-address.txt') as f:
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

# DOWNLOAD VDRS FROM EMAIL FUNCTION
def dwl_vdr(email_add, password, server, vesselEmail, vesselName):
    imap = imaplib.IMAP4_SSL(server, 993)
    imap.login(email_add, password)
    imap.select('INBOX')

    index = 0

    while index < len(vesselEmail) and index < len(vesselName):
        try:
            client_folder = os.path.join(folder_full_path, clientname_list[index])

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

                '''
                raw_email_string = raw_email.decode('ISO-8859–1')
                email_message = email.message_from_string(raw_email_string)
                subject_name = email_message['subject']
                
                '''
                email_message = email.message_from_bytes(raw_email)
                subject, encoding = decode_header(email_message["Subject"])[0]

                if isinstance(subject, bytes):
                    subject = subject.decode(encoding or "UTF-8", errors="replace")

                print(f'Subject: {subject}')

                # att_path = "No attachment found from email " + subject_name
                logging.debug(f'Email subject: {subject}')
                for part in email_message.walk():
                    try:
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue

                        fileName = part.get_filename()
                        if fileName:
                            decode_name, encoding = decode_header(fileName)[0]
                            if isinstance(decode_name, bytes):
                                decode_name = decode_name.decode(encoding or "UTF-8", errors="replace")

                        if fnmatch.fnmatch(decode_name, "*.xls*"):
                            try:
                                att_path = os.path.join(folder_name, decode_name)
                                print(att_path)
                                if not os.path.isfile(att_path):
                                    fp = open(att_path, "wb")
                                    fp.write(part.get_payload(decode=True))
                                    fp.close()
                            except TypeError as e:
                                error_file.append(decode_name)
                                continue
                            logging.debug(f'File downloaded: {decode_name}')

                        elif fnmatch.fnmatch(decode_name, "*.doc*"):
                            try:
                                att_path = os.path.join(folder_name, decode_name)
                                print(att_path)
                                if not os.path.isfile(att_path):
                                    fp = open(att_path, "wb")
                                    fp.write(part.get_payload(decode=True))
                                    fp.close()
                            except TypeError as e:
                                error_file.append(decode_name)
                                continue
                            logging.debug(f'File downloaded: {decode_name}')

                        elif fnmatch.fnmatch(decode_name, "*.pdf"):
                            try:
                                att_path = os.path.join(folder_name, decode_name)
                                print(att_path)
                                if not os.path.isfile(att_path):
                                    fp = open(att_path, "wb")
                                    fp.write(part.get_payload(decode=True))
                                    fp.close()
                            except TypeError as e:
                                error_file.append(decode_name)
                                continue
                            logging.debug(f'File downloaded: {decode_name}')

                        else:
                            logging.debug(f'No file downloaded from {subject}')

                    except (OSError, TypeError) as f:
                        print(f)
                        continue

                # print(att_path)

                '''for response in data:
                    if isinstance(response, tuple):
                        data = email.message_from_bytes(response[1])
                        subject = decode_header(data["Subject"])[0][0]
                        if isinstance(subject, bytes):
                            # if it's a bytes type, decode to str
                            subject = subject.decode()
                        logging.debug(f'Deleting email {subject}')
                imap.store(num, "+FLAGS", "\\Deleted")
                '''

            index += 1

        except (OSError, TypeError) as k:
            print(k)
            continue

    #imap.expunge()
    imap.close()
    imap.logout()
    logging.debug('Downloading finished..')


'''
# SEND EMAIL WHEN FINISHED FUNCTION
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
'''

# DELETE EMPTY FOLDER FUNCTION
def delete_empty_folders(folder_full_path):
    for folder_name, subfolders, files in os.walk(folder_full_path, topdown=False):
        for subfolder in subfolders:
            folder_path = os.path.join(folder_name, subfolder)
            if not os.listdir(folder_path):
                os.rmdir(folder_path)
                logging.debug(f"Deleted empty folder: {folder_path}")
                print(f"Deleted empty folder: {folder_path}")


try:
    logging.debug('Downloading..')
    dwl_vdr(email_vdr, pwd_vdr, server_mssb, vesselemail_list, vesselname_list)
    dwl_vdr(email_mvmcc, pwd_mvmcc, server_mssb, vesselemail_list, vesselname_list)
    delete_empty_folders(folder_full_path)
    #send_email(message)

except Exception as e:
    logging.error("Error occurred", exc_info=True)
