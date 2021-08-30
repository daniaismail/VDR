import imaplib
import email
from datetime import datetime, timedelta
import os
import fnmatch

file_path = r'C:\Users\MVMWEB\pythonProject\Email\Downloaded'
date = (datetime.today()).strftime("%d-%b-%Y")
today = datetime.today()
cutoff = today - timedelta(days=2)
dt = cutoff.strftime('%d-%b-%Y')
error_file = []

'''
email_add = "dania@meridiansurveys.com.my"
password = "e*eM-@FfK$w*"

'''
email_vdr1 = "vdr1@meridiansurveys.com.my"
pwd_vdr1 = "x2c(UR*{gfT#"

email_vdr = "vdr@meridiansurveys.com.my"
pwd_vdr = "T%zf5ccq;ZMc"

email_mvmcc = "mvmcc@meridiansurveys.com.my"
pwd_mvmcc = "dc)in]}Xzk&%"

email_wgg = "mvmcc@wild-geese-group.com"
pwd_wgg = "s9nPviD\\"

server_wgg = "mail.wild-geese-group.com"
server_mssb = "meridian-svr.meridiansurveys.com.my"

fr_vdr1 = ["jmmurni@jmfleet.com.my", "omniemery1@iconoffshore.com.my", "vanessa9@ipsignature3.net", "jmseribesut@jmfleet.com.my", "setiateguh@stationsatcom.commbox.com", "setiazaman@stationsatcom.commbox.com", "express.alpha@jvcmega.com.my", "9mvp9@amosconnect.com", "centusone@ipsignature3.net", "centusthree@ipsignature3.net", "centustwo@ipsignature3.net", "exh@eopl.gtmailplus.com", "iconamara@iconoffshore.com.my", "jmabadi@jmfleet.com.my", "jmpermai@jmfleet.com.my", "skpatriot@skom.com.my", "yinsonperwira@stationsatcommail.com", "skatomik@skom.com.my", "skline79@skom.com.my", "jmehsan@jmfleet.com.my", "ebr@eopl.gtmailplus.com", "Express85@gtmailplus.com", "jmpurnama@jmfleet.com.my", "exn@eopl.gtmailplus.com", "taqwaadam14@gmail.com", "fcbmasindah@outlook.com", "dayang_almira@ipsignature3.net", "ebn@eopl.gtmailplus.com", "vdr.mkbumimas@gmail.com", "pacific.harrier@spoships.com", "master@tourmaline.ss.commbox.com", "lsk.bridge@emas.com"]
f_name_vdr1 = ["hess\\jm murni", "hess\\omni emery 1", "jxnippon\\vanessa 9", "petrofac\\jm seri besut", "petrofac\\setia teguh", "petrofac\\setia zaman", "pflng\\express alpha", "pflng\\ntp 29", "pttep\\centus one", "pttep\\centus three", "pttep\\centus two", "pttep\\executive honour", "hess\\icon amara", "pttep\\jm abadi", "pttep\\jm permai", "pttep\\sk patriot", "pttep\\yinson perwira", "petrofac\\sk atomik", "pttep\\sk line 79", "petrofac\\jm ehsan", "pttep\\executive brilliance", "petrofac\\express 85", "pttep\\jm purnama", "pttep\\executive excellence", "pttep\\mv taqwa adam", "jxnippon\\mas indah", "jxnippon\\dayang almira", "pttep\\executive benevolence", "pttep\\mas baiduri", "pttep\\pacific harrier", "pttep\\tourmaline", "pttep\\jp88 stork"]

fr_vdr = ["iconamara@iconoffshore.com.my", "dayang_almira@ipsignature3.net", "lsk.bridge@emas.com", "vdr.mkbumimas@gmail.com", "pacific.harrier@spoships.com", "master@tourmaline.ss.commbox.com"]
f_name_vdr = ["icon amara", "dayang almira", "lewek stork", "mas baiduri", "pacific harrier", "tourmaline"]

fr_mvmcc = ["skprodigy@skom.com.my", "9mvp7@amosconnect.com", "9mvp8@amosconnect.com"]
f_name_mvmcc = ["pflng\\sk prodigy", "pflng\\ntp 27", "pflng\\ntp 28"]

fr_wgg = ["ntpxxxVII@amosconnect.com"]
f_name_wgg = ["pflng\\ntp 37"]

if not os.path.exists(file_path):
    os.mkdir(file_path)

def dwl_vdr(email_add, password, server, fr, f_name):
    imap = imaplib.IMAP4_SSL(server, 993)
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

dwl_vdr(email_wgg, pwd_wgg, server_wgg, fr_wgg, f_name_wgg)
dwl_vdr(email_mvmcc, pwd_mvmcc, server_mssb, fr_mvmcc, f_name_mvmcc)
dwl_vdr(email_vdr, pwd_vdr, server_mssb, fr_vdr1, f_name_vdr1)
#dwl_vdr(email_vdr, pwd_vdr, server_mssb, fr_vdr, f_name_vdr)