#! /usr/bin/python3
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

smtp_server = "smtp.office365.com"
smtp_port = 587

def get_html(filename):
    # https://beefree.io/templates/
    try:
        with open(filename, 'r') as f:
            print(f'[+] Reading from {filename}')
            return f.read()
    except Exception as e:
        print(f'[!] Unable to open {filename}: ' + e)
        quit()

def send_message(sender_email, sender_password, target_email, email_html, attachment_file):
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

    try:
        conn = SMTP(smtp_server, smtp_port)
    except Exception as e:
        print('Received an error: ' + e)
        quit()

    print(f'[+] Connected to {smtp_server}:{smtp_port}!')
    conn.starttls()
    conn.set_debuglevel(False)

    try:
        conn.login(sender_email, sender_password)
        conn.sendmail(sender_email, target_email, msg.as_string())
    except Exception as e:
        print('[!] Failed to send:' + e)
        quit()
    finally:
        print('[+] Email sent!')
        conn.quit()
        return 0

def main():
    try:
        sender_email = argv[1]
        sender_password = argv[2]
        target_email = argv[3]
        email_html = argv[4]
        attachment_file = argv[5]
    except Exception as e:
        print('Usage: python3 fishy.py <sender_email> <sender_password> <target_email> <email_html> <attachment_file>')
        quit()

    send_message(sender_email, sender_password, target_email, email_html, attachment_file)

if __name__ == '__main__':
    main()