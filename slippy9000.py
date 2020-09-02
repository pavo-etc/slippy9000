import imaplib
import email
from email.header import decode_header
import webbrowser
import os

def monthAsNum(month_name):
    switch = {  "Jan" : "01",
                "Feb" : "02",
                "Mar" : "03",
                "Apr" : "04",
                "May" : "05",
                "Jun" : "06",
                "Jul" : "07",
                "Aug" : "08",
                "Sep" : "09",
                "Oct" : "10",
                "Nov" : "11",
                "Dec" : "12"    }
    return switch.get(month_name, "month_error")

def main():
    with open("secrets.txt","r") as f:
        secrets = f.read().split("\n")
    username = secrets[0]
    password = secrets[1]

    imap = imaplib.IMAP4_SSL("imap.gmail.com")

    imap.login(username, password)
    print(f"LOGGED IN: {username}")

    status, messages = imap.select("INBOX")

    # total number of emails
    messages = int(messages[0])

    print("------------SEARCHING Inbox------------")
    xero_emails = 0
    pdfs_saved = 0
    for i in range(1, messages+1):
        
        res, msg = imap.fetch(str(i), "(RFC822)") # fetch email message by ID
        
        for response in msg:
            if not isinstance(response, tuple):
                continue
            
            message = email.message_from_bytes(response[1])
            
            subject = decode_header(message["Subject"])[0][0]
            sender = message["From"]
            
            print(f"\nSubject:\t{subject}")
            print(f"Sender: \t{sender}")
            
            if sender != "noreply@post.xero.com": # ignore other senders
                break
            else:
                xero_emails += 1
            
            if subject[0:39] != "Payslip for Zach Manson for Week ending":
                print("\tIGNORING: Not a payslip!")
                break
            date_in_subject = subject[-4:] + monthAsNum(subject[-8:-5]) + subject[-11:-9]
            filepath = os.path.join(os.getcwd(), f"PaySlip {date_in_subject}.pdf")
            if os.path.isfile(filepath):
                print("\tIGNORING: Already downloaded attachment!")
                break
            
            for part in message.walk():
                content_disposition = str(part["Content-Disposition"])
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    print(f"Attachment:\t{filename}")
                    
                    print("\tDOWNLOADING attachment...")
                    open(filepath, "wb").write(part.get_payload(decode=True))
                    pdfs_saved += 1
                    print("\tDOWNLOADED attachment!")
    
    print("\n------------SEARCHED Inbox-------------")
    print(f"{xero_emails} emails from Xero in inbox.")
    print(f"{pdfs_saved} payslips downloaded from inbox.")
    imap.close()
    imap.logout()

if __name__ == "__main__":
    print("RUNNING slippy9000...")
    main()
