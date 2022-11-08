import imaplib

from jmespath import search
import imap
import importlib
import email

imap = imaplib.IMAP4_SSL('outlook.office365.com')
imap.login('ashu098765432@outlook.com', 'AshuGupta18') 
imap.select('Inbox')
_, search_data = imap.search(None, 'ALL')
for num in search_data[0].split():
    _, data = imap.fetch(num, '(RFC822)')
    _, b = data[0]
    email_message = email.message_from_bytes(b)
    # print(email_message)
    for header in ['subject','to','from','date']:
        print("{}: {}".format(header,email_message[header]))
    for part in email_message.walk():
        if part.get_content_type() == "text/plain" or part.get_content_type == "text/html":
            body = part.get_payload(decode=True)
            email_data = body.decode()
            print(email_data)




#     # def email_forwarding(email_data):
#     #     toaddr = 'rupeshsc1999@outlook.com'
#     #     # bcc = em
#     #     fromaddr = "ashu098765432@outlook.com"
#     #     mss = "Escalation Mail.Please Look into it ASAP"
#     #     alll = mss + email_data
#     #     message_text = email.message_from_string(alll)
#     #     message = message_text
#     #     # message.replace_header("From", "ashu098765432@outlook.com")
#     #     # toaddrs = [*toaddr]
#     #     # print(toaddrs, '------------------Finalll')
#     #     smtp = smtplib.SMTP('smtp.office365.com', 587)
#     #     smtp.starttls()
#     #     smtp.login('ashu098765432@outlook.com','AshuGupta18')
#     #     smtp.sendmail(fromaddr, toaddr, message.as_string(message))
#     #     print("Forward Mails Sent successfully")
#     #     smtp.quit()

#     # def email_forwardin(email_dat):
#     #     toaddr = 'rupeshsc1999@outlook.com'
#     #     # bcc = em
#     #     fromaddr = "ashu098765432@outlook.com"
#     #     mssges = "This is RPA related"
#     #     email_dat = mssges + email_data
#     #     message_text = email.message_from_string(email_dat)
#     #     message = message_text
#     #     # message.replace_header("From", "ashu098765432@outlook.com")
#     #     # toaddrs = [*toaddr]
#     #     # print(toaddrs, '------------------Finalll')
#     #     smtp = smtplib.SMTP('smtp.office365.com', 587)
#     #     smtp.starttls()
#     #     smtp.login('ashu098765432@outlook.com','AshuGupta18')
#     #     smtp.sendmail(fromaddr, toaddr, message.as_string(mssges))
#     #     print("RPA Mails Sent successfully")
#     #     smtp.quit()

#     # def email_forwardi(email_da):
#     #     toaddr = 'anand.singh.act@outlook.com'
#     #     # bcc = em
#     #     fromaddr = "ashu098765432@outlook.com"
#     #     mssges = "This is AI and ML related"
#     #     email_da = mssges+email_data
#     #     message_text = email.message_from_string(email_da)
#     #     message = message_text
#     #     # message.replace_header("From", "ashu098765432@outlook.com")
#     #     smtp = smtplib.SMTP('smtp.office365.com', 587)
#     #     smtp.starttls()
#     #     smtp.login('ashu098765432@outlook.com','AshuGupta18')
#     #     smtp.sendmail(fromaddr, toaddr, message.as_string(mssges))
#     #     print("AI and ML Mails Sent successfully")
#     #     smtp.quit()


#  # message.replace_header("From", "ashu098765432@outlook.com")
#             # toaddrs = [*toaddr]
#             # print(toaddrs, '------------------Finalll')



#     # imapHost = 'outlook.office365.com'
#     # imap = imaplib.IMAP4_SSL(imapHost)
#     # imap.login('ashu098765432@outlook.com', 'AshuGupta18') 
#     # imap.select('Inbox')
#     # tmp, data = imap.search(None, 'ALL')
#     # for num in data[0].split():
#     #     # print(tmp)
#     #     tmp, data = imap.fetch(num, '(RFC822)')
#     #     body = data[0][1]
#     #     email_message = email.message_from_bytes(body)
#     #     email_data = body.decode("utf-8")
