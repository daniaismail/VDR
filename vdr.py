import imaplib
import email
from datetime import datetime, timedelta
import os
import fnmatch
from xlrd import XLRDError

date = (datetime.today()).strftime("%d-%b-%Y")
today = datetime.today()
cutoff = today - timedelta(days=5)
dt = cutoff.strftime('%d-%b-%Y')
error_file = []

'''
fr = ["centustwo@ipsignature3.net"]
f_name = ["centus two"]

email_add = "dania@meridiansurveys.com.my"
password = "e*eM-@FfK$w*"

email_add = "mvmcc@wild-geese-group.com"
password = "s9nPviD\\"

mail.wild-geese-group.com
meridian-svr.meridiansurveys.com.my

fr = ["ntpxxxVII@amosconnect.com"]
f_name = ["ntp 37"]

email_add = "vdr@meridiansurveys.com.my"
password = "T%zf5ccq;ZMc"

fr = ["jmmurni@jmfleet.com.my", "omniemery1@iconoffshore.com.my", "vanessa9@ipsignature3.net", "jmseribesut@jmfleet.com.my", "setiateguh@stationsatcom.commbox.com", "setiazaman@stationsatcom.commbox.com", "nauticatg.puteri28@gmail.com", "express.alpha@jvcmega.com.my", "9mvp9@amosconnect.com", "centusone@ipsignature3.net", "centusthree@ipsignature3.net", "centustwo@ipsignature3.net", "exh@eopl.gtmailplus.com", "iconamara@iconoffshore.com.my", "jmabadi@jmfleet.com.my", "jmpermai@jmfleet.com.my", "taqwaadam@singtelmailbiz.com", "skpatriot@skom.com.my", "yinsonperwira@stationsatcommail.com", "sjane.captain@sapuraenergy.com", "skatomik@skom.com.my", "skline79@skom.com.my", "perdanamarathon@intraoil.com.my", "setiajihad@stationsatcommail.com", "jmehsan@jmfleet.com.my", "ese@eopl.gtmailplus.com", "skpride@skom.com.my", "ebr@eopl.gtmailplus.com", "ete@eopl.gtmailplus.com", "Express85@gtmailplus.com", "jmpurnama@jmfleet.com.my", "exn@eopl.gtmailplus.com", "taqwaadam14@gmail.com", "fcbmasindah@outlook.com", "dayang_almira@ipsignature3.net"]
f_name = ["jm murni", "omni emery 1", "vanessa 9", "jm seri besut", "setia teguh", "setia zaman", "ntp 28",  "express alpha", "ntp 29", "centus one", "centus three", "centus two", "executive honour", "icon amara", "jm abadi", "jm permai", "mv taqwa adam", "sk patriot", "yinson perwira", "sapura jane", "sk atomik", "sk line 79", "p marathon", "setia jihad", "jm ehsan", "executive stride", "sk pride", "executive brilliance", "executive tide", "express 85", "jm purnama", "executive excellence", "mv taqwa adam", "mas indah", "dayang almira"]

email_add = "vdr1@meridiansurveys.com.my"
password = "x2c(UR*{gfT#"

fr = ["skprodigy@skom.com.my", "9mvp7@amosconnect.com", "9mvp8@amosconnect.com"]
f_name = ["sk prodigy", "ntp 27", "ntp 28"]

email_add = "mvmcc@meridiansurveys.com.my"
password = "dc)in]}Xzk&%"

email_add = "mvmcc@meridiansurveys.com.my"
password = "dc)in]}Xzk&%"

'''
fr = ["jmmurni@jmfleet.com.my", "omniemery1@iconoffshore.com.my", "vanessa9@ipsignature3.net", "jmseribesut@jmfleet.com.my", "setiateguh@stationsatcom.commbox.com", "setiazaman@stationsatcom.commbox.com", "nauticatg.puteri28@gmail.com", "express.alpha@jvcmega.com.my", "9mvp9@amosconnect.com", "centusone@ipsignature3.net", "centusthree@ipsignature3.net", "centustwo@ipsignature3.net", "exh@eopl.gtmailplus.com", "iconamara@iconoffshore.com.my", "jmabadi@jmfleet.com.my", "jmpermai@jmfleet.com.my", "taqwaadam@singtelmailbiz.com", "skpatriot@skom.com.my", "yinsonperwira@stationsatcommail.com", "sjane.captain@sapuraenergy.com", "skatomik@skom.com.my", "skline79@skom.com.my", "perdanamarathon@intraoil.com.my", "setiajihad@stationsatcommail.com", "jmehsan@jmfleet.com.my", "ese@eopl.gtmailplus.com", "skpride@skom.com.my", "ebr@eopl.gtmailplus.com", "ete@eopl.gtmailplus.com", "Express85@gtmailplus.com", "jmpurnama@jmfleet.com.my", "exn@eopl.gtmailplus.com", "taqwaadam14@gmail.com", "fcbmasindah@outlook.com", "dayang_almira@ipsignature3.net"]
f_name = ["jm murni", "omni emery 1", "vanessa 9", "jm seri besut", "setia teguh", "setia zaman", "ntp 28",  "express alpha", "ntp 29", "centus one", "centus three", "centus two", "executive honour", "icon amara", "jm abadi", "jm permai", "mv taqwa adam", "sk patriot", "yinson perwira", "sapura jane", "sk atomik", "sk line 79", "p marathon", "setia jihad", "jm ehsan", "executive stride", "sk pride", "executive brilliance", "executive tide", "express 85", "jm purnama", "executive excellence", "mv taqwa adam", "mas indah", "dayang almira"]

email_add = "vdr1@meridiansurveys.com.my"
password = "x2c(UR*{gfT#"


file_path = r'C:\Users\MVMWEB\pythonProject\Email\Downloaded'

if not os.path.exists(file_path):
    os.mkdir(file_path)

imap = imaplib.IMAP4_SSL("meridian-svr.meridiansurveys.com.my", 993)
imap.login(email_add, password)
imap.select('Inbox')
#typ, data = imap.search(None, '(SENTSINCE {0})'.format(date))

index = 0

while index < len(fr) and index < len(f_name):
    try:
        folder_name = os.path.join(file_path, f_name[index])
        if not os.path.exists(folder_name):
            os.mkdir(folder_name)

        typ, data = imap.search(None, '(SINCE %s)' % (dt,), '(FROM %s)' % (fr[index],))

        for num in data[0].split():
            typ, data = imap.fetch(num, '(RFC822)')
            raw_email = data[0][1]
            raw_email_string = raw_email.decode('windows-1252')
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
                            pass

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
                            pass

                    elif fnmatch.fnmatch(fileName, "*.dat*"):
                        try:
                            att_path = os.path.join(folder_name, fileName + "")
                            if not os.path.isfile(att_path):
                                fp = open(att_path, "wb")
                                fp.write(part.get_payload(decode=True))
                                fp.close()
                            else:
                                #num += 1
                                fp = open(att_path, "wb")
                                fp.write(part.get_payload(decode=True))
                                fp.close()
                        except TypeError as e:
                            error_file.append(fileName)
                            pass

                except TypeError as e:
                    error_file.append(fileName)
                    pass

            print(att_path)
        index += 1

    except TypeError as e:
        print(e)
        pass

print(error_file)
imap.close()
imap.logout()
