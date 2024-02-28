from builtins import isinstance, print, str, tuple
from robot.libraries.BuiltIn import BuiltIn
import time
import imaplib
import email

def fetchCode(Upcomming_EMAIL_Address,Data):
    username = "Your Email Id"
    password = "Your Password"
    imap_server = "outlook.office365.com"
    value = ""
    time.sleep(60)
    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(username, password)
    imap.select("INBOX")
    result, data = imap.search(None, "From", Upcomming_EMAIL_Address)
    if result == "OK":
        email_number = str(data[0]).split(' ')[-1].replace("'", "")
        res, msg = imap.fetch(str(email_number), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                for part in msg.walk():
                    content_type = part.get_content_type()
                if content_type == "text/html":
                    body = part.get_payload(decode=True).decode()
                    if Data == "code":
                        variable = body
                        Print(body)
                        variable = body.split('>')[12].split('</strong')[0]
                    elif Data == "link":
                        variable = body.split('href="')[1].split('">')[0]
                    else:
                        variable = body.split("href=")[1].split(">")[0]
                        variable = variable.strip('"')
                Print(variable)
                return (variable)
    else:
        print("Email Not Found")
