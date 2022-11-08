from exchangelib import DELEGATE, Account, Credentials, EWSTimeZone
import re
import os
credentials = Credentials(username='ashu098765432@outlook.com',password='AshuGupta18')
account = Account(primary_smtp_address='ashu098765432@outlook.com', credentials=credentials, autodiscover=True, access_type=DELEGATE)
# print(account,'>>>>>>>>>>>>>>>>>>>>>>>>>')
for item in account.inbox.filter(is_read=False,)[::-1]:
    mail_sender = item.sender
    mail_sender = str(mail_sender)
    mail_sender = re.findall('email_address\=\'.*?\'',mail_sender)[0]
    print(mail_sender)
    mail_sender = mail_sender.replace('email_address=\'','').replace('\'','').replace("'","_").strip()
    mail_subject = item.subject
    print(mail_subject)
    mail_id = item.message_id
    print(mail_id)
    mail_recipients = str(item.to_recipients)
    print(mail_recipients)
    mail_cc = str(item.cc_recipients)
    print(mail_cc)