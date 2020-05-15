#! /usr/bin/python3
import argparse
from sys import argv
from smtplib import SMTP
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate


'''
    - Email attachments to office addresses!
    - Requires access to an inbox to send the email from.
    - https://twitter.com/__mez0__
    - https://github.com/mez-0
'''

def get_args():
    parser = argparse.ArgumentParser(description="Fish.")
    parser.add_argument("-u", "--sender-address", metavar="", required=True, help="Username to authenticate with (Office)")
    parser.add_argument("-p", "--sender-password", metavar="", help="Password to authenticate with (Office)")
    parser.add_argument("-t", "--target-address", metavar="", required=True, help="Victim Email Address")
    parser.add_argument("-a", "--attachment", metavar="", required=True, help="Attachment to send")
    parser.add_argument("-e", "--html-email", metavar="", required=True, help="HTML Body for Email")
    parser.add_argument("--smtp-server", metavar="", default="smtp.office365.com", type=str, help="SMTP Server Address")
    parser.add_argument("--smtp-port", metavar="", default=587, type=int, help="SMTP Server Port")
    args = parser.parse_args()
    return args    

def get_html(filename):
    # https://beefree.io/templates/
    try:
        with open(filename, 'r') as f:
            print(f'[+] Reading from {filename}')
            return f.read()
    except Exception as e:
        print(f'[!] Unable to open {filename}: ' + e)
        quit()

def build_email(sender_email, target_email, email_html, attachment_file):
    html = get_html(email_html)
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = target_email
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = 'test'

    msg.attach(MIMEText(html, 'html'))

    files = [attachment_file]

    try:
        for f in files:
            with open(f, "rb") as fil:
                part = MIMEApplication(fil.read(),Name=basename(f))
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
            msg.attach(part)
    except Exception as e:
        print(f'[!] Unable to open {attachment_file}: ' + str(e))
        quit()

    print(f'[+] Sending {len(html)} to {target_email}!')
    return msg

def connect(sender_email, sender_password, target_email):
    args = get_args()
    smtp_server = args.smtp_server
    smtp_port = args.smtp_port
    if sender_email != None and sender_password != None:
        try:
            conn = SMTP(smtp_server, smtp_port)
            print(f'[+] Connected to {smtp_server}:{smtp_port}!')

            conn.starttls()
            conn.set_debuglevel(False)
            conn.login(sender_email, sender_password)
            print(f'[+] Authenticated as {sender_email}!')
            return conn
        except Exception as e:
            print('Received an error: ' + str(e))
            quit()
    else:
        print('[!] Please Specify Credentials!')
        quit()

def send_message(sender_email, sender_password, target_email, email_html, attachment_file):
    html = get_html(email_html)
    
    msg = build_email(sender_email, target_email, email_html, attachment_file)

    conn = connect(sender_email, sender_password, target_email)

    try:
        conn.sendmail(sender_email, target_email, msg.as_string())
        print('[+] Email Succesfully sent!')
    except Exception as e:
        print('[!] Email Failed: ' + str(e))
    finally:
        conn.close()

def main():
    args = get_args()

    send_message(args.sender_address, args.sender_password, args.target_address, args.html_email, args.attachment)

if __name__ == '__main__':
    main()