from email import header
import re
import os
from exchangelib import Q
import imaplib
import email, email.parser
import smtplib

# def att_dow_rep():
server =imaplib.IMAP4_SSL('outlook.office365.com',993)
server.login('ashu098765432@outlook.com','AshuGupta18')
server.select()
typ, data = server.search(None, '(SUBJECT "")')
mail_ids = data[0]
id_list = mail_ids.split()
path = r"C:\Users\Ashu\Desktop\Sequalstring\Emailtask"

for num in data[0].split():
    typ, data = server.fetch(num, '(RFC822)' )
    raw_email = data[0][1]
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)
    for part in email_message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()
        if bool(fileName):
            if 'image' in fileName or 'IMG' in fileName:
                pass
            elif fileName.endswith('.xlsx') or fileName.endswith('.csv'):
                filePath = os.path.join(path, fileName)
                if not os.path.isfile(filePath) :
                    fp = open(filePath, 'wb')
                    # fp.write(part.get_payload(decode=True))
                    fp.close()
                aps = smtplib.SMTP('smtp.office365.com', 587)
                aps.starttls()
                em_mssg = "Your Mail has been received and all the necessary docs have been downloaded"
                massg = email.message_from_string(em_mssg)
                massage = massg
                aps.login('ashu098765432@outlook.com','AshuGupta18')
                aps.sendmail('ashu098765432@outlook.com','anand.singh.act@outlook.com', massage.as_string())
                print("Mail Send")
server.logout
print("Attachment downloaded from mail")
# att_dow_rep()