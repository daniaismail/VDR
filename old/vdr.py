import imaplib
import smtplib
import email
from datetime import datetime, timedelta
import os
import fnmatch
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import decode_header

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

file_path = r'C:\Users\MVMWEB\Dropbox\MVMCC\VDR'
#date = (datetime.today()).strftime("%d-%b-%Y")
today = datetime.today()
yesterday = datetime.today()-timedelta(1)
cutoff = today - timedelta(days=7)
dt = cutoff.strftime('%d-%b-%Y')
error_file = []
folder_month = yesterday.month

folder_month_path = os.path.join(file_path, str(folder_month))
if not os.path.exists(folder_month_path):
    os.mkdir(folder_month_path)

message = "VDR Directory: " + folder_month_path

fr_vdr = ["vanessa9@ipsignature3.net", "master.mk16@stationsatcommail.com", "mk6@stationsatcommail.com",
          "jmseribesut@jmfleet.com.my", "setiateguh@stationsatcom.commbox.com",
        "express.alpha@jvcmega.com.my", "ntp29@ipsignature3.net", "ntp23@ipsignature3.net",
          "centusone@centus.commbox.com", "centusten@centus.commbox.com", "centusthree@centus.commbox.com", "centustwo@centus.commbox.com", "exn@eopl.gtmailplus.com", "greatship_maya@greatshipglobal.com", "jmpurnama@jmfleet.com.my", "master.grace@keyfieldoffshore.com", "taqwaadam14@gmail.com", "pacific.harrier@swireships.com", "phviking@gtmailplus.com", "skpatriot@skom.com.my", "yinsonhermes@satcomglobalmail.com", "yinsonperwira@stationsatcommail.com"]

f_name_vdr = ["jxnippon\\533170182", "jxnippon\\533016800", "jxnippon\\533002530",
              "petrofac\\533000479", "petrofac\\533062000",
              "pflng\\533170609", "pflng\\533130433", "pflng\\533130422",
              "pttep\\533180103", "pttep\\533132209", "pttep\\533180174", "pttep\\533180104", "pttep\\533130820", "pttep\\563381000", "pttep\\533000492", "pttep\\533132128", "pttep\\533000856", "pttep\\525005353", "pttep\\533131177", "pttep\\533180152", "pttep\\533130329", "pttep\\533640000"]

fr_mvmcc = ["skprodigy@skom.com.my", "ntp27@ipsignature3.net"]
f_name_mvmcc = ["pflng\\533180070", "pflng\\533130498"]

fr_wgg = ["ntp37@ipsignature3.net", "ksp.pioneer@outlook.com", "ntp28@ipsignature3.net"]
f_name_wgg = ["pflng\\533131114", "pflng\\533150095", "pflng\\533130499"]

email_archive3 = "mvmcc.archive3@gmail.com"
pwd_archive3 = "S#@ZX3ng"

email_vdr = "vdr@meridiansurveys.com.my"
pwd_vdr = "T%zf5ccq;ZMc"

email_vdr1 = "vdr1@meridiansurveys.com.my"
pwd_vdr1 = "x2c(UR*{gfT#"

email_vdr2 = "vdr2@meridiansurveys.com.my"
pwd_vdr2 = "C{v@*}Zmq5?#"

email_mvmcc = "mvmcc@meridiansurveys.com.my"
pwd_mvmcc = "dc)in]}Xzk&%"

email_wgg = "mvmcc@wild-geese-group.com"
pwd_wgg = "s9nPviD\\"

server_gmail = "imap.gmail.com"
server_wgg = "mail.wild-geese-group.com"
server_mssb = "meridian-svr.meridiansurveys.com.my"

clients = ["petrofac", "pflng", "pttep", "jxnippon"]

if not os.path.exists(file_path):
    os.mkdir(file_path)

for client in clients:
    folder = os.path.join(folder_month_path, client)
    if not os.path.exists(folder):
        os.mkdir(folder)

def dwl_vdr(email_add, password, server, fr, f_name):
    imap = imaplib.IMAP4_SSL(server, 993)
    imap.login(email_add, password)
    imap.select('INBOX')
    #typ, data = imap.search(None, '(SENTSINCE {0})'.format(date))

    index = 0

    while index < len(fr) and index < len(f_name):
        try:
            folder_name = os.path.join(folder_month_path, f_name[index])
            if not os.path.exists(folder_name):
                os.mkdir(folder_name)

            typ, data = imap.search(None, '(SINCE %s)' % (dt,), '(FROM %s)' % (fr[index],))

            for num in data[0].split():
                typ, data = imap.fetch(num, '(RFC822)')
                raw_email = data[0][1]
                raw_email_string = raw_email.decode('ISO-8859–1')
                email_message = email.message_from_string(raw_email_string)
                subject_name = email_message['subject']

                att_path = "No attachment found from email " + subject_name
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

                    except (OSError, TypeError) as e:
                        error_file.append(fileName)
                        continue

                print(att_path)

                for response in data:
                    if isinstance(response, tuple):
                        data = email.message_from_bytes(response[1])
                        subject = decode_header(data["Subject"])[0][0]
                        if isinstance(subject, bytes):
                            # if it's a bytes type, decode to str
                            subject = subject.decode()
                        print("Deleting", subject)
                imap.store(num, "+FLAGS", "\\Deleted")

            index += 1

        except OSError as e:
            error_file.append(fileName)
            continue

        except TypeError as e:
            error_file.append(fileName)
            continue

    print(error_file)
    imap.expunge()
    imap.close()
    imap.logout()

dwl_vdr(email_wgg, pwd_wgg, server_wgg, fr_wgg, f_name_wgg)
dwl_vdr(email_mvmcc, pwd_mvmcc, server_mssb, fr_mvmcc, f_name_mvmcc)
#dwl_vdr(email_vdr, pwd_vdr, server_mssb, fr_vdr, f_name_vdr)
dwl_vdr(email_vdr2, pwd_vdr2, server_mssb, fr_vdr, f_name_vdr)
#dwl_vdr(email_archive3, pwd_archive3, server_gmail, fr_vdr, f_name_vdr)
#send_email(message)
