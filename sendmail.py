#! /usr/bin/python3
import sys
import subprocess
import pkg_resources

# Install missing requirements
required = {'O365', 'datetime'}
installed = {pkg.key for pkg in pkg_resources.working_set}
missing = required - installed
if missing:
    python = sys.executable
    subprocess.check_call([python, '-m', 'pip', 'install', *missing], stdout=subprocess.DEVNULL)

from O365 import Account, FileSystemTokenBackend, Connection, MSGraphProtocol
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.message import EmailMessage
import argparse
import os.path, time
from datetime import datetime, timedelta


infer_fields=False

parser = argparse.ArgumentParser(description='Send an e-mail from command line.')

def str2bool(v):
    if isinstance(v, bool):
       return v
    if v.lower() in ('yes', 'true', 't', 'y', '1'):
        return True
    elif v.lower() in ('no', 'false', 'f', 'n', '0'):
        return False
    else:
        raise argparse.ArgumentTypeError('Boolean value expected.')

parser.add_argument('-u', '--username', action='store', help="Username/ClientID for authentication and from address.", required=True)
parser.add_argument('-p', '--password', action='store', help="Authentication Password/Client Secret.", required=True)
parser.add_argument('-t', '--to', action='store', help="Email of recipient or a comma separated list (Use -t or -e).", required=False)
parser.add_argument('-e', '--emailfile', action='store', help="Text file containing list of emails (Use -t or -e).", required=False)
parser.add_argument('-s', '--subject', action='store', help="Email subject.", required=True)
parser.add_argument('-m', '--message', action='store', help="Email message (Use -b or -m).", required=False)
parser.add_argument('-b', '--bodylink', action='store', help="Text file containing email message (Use -b or -m).", required=False)
parser.add_argument('-a', '--attachment', action='store', help="Email attachment", required=False)
parser.add_argument('--smtp', type=str2bool, nargs='?', const=True, default=False, help="Use SMTP to send e-mail? Otherwise will use API connection.")
parser.add_argument('-d', '--delay', action='store', metavar='N', nargs='+', type=int, default=6, help="How many seconds to delay between sending each e-mail (Defaults is 6 seconds).", required=False)

args = parser.parse_args()
email_user = args.username
email_password = args.password
delay = args.delay

# Assign email body content based on command line variable.
if args.message is None and args.bodylink is None:
    print("Either an inline message body or text file must be provided to procede. Use the -m or -b option.")
    raise argparse.ArgumentTypeError('Either an inline message body or text file must be provided to procede. Use the -m or -b option.')
elif args.message is not None and args.bodylink is not None:
    print("Either an inline message body or text file must be provided to procede. Use the -m or -b option.")
    raise argparse.ArgumentTypeError('Either an inline message body or text file must be provided to procede. Use the -m or -b option.')
elif args.bodylink is None:
    body = args.message
elif args.message is None:
    with open('body.txt','r') as file:
        body = file.read()

# Load the proper list of e-mail addresses.
if args.to is None and args.emailfile is None:
    print("Destination emails must be provided by command line or text file.")
    raise argparse.ArgumentTypeError("Destination emails must be provided by command line or text file.")
elif args.to is not None and args.emailfile is not None:
    print("Emails can only be provided by one method.")
    raise argparse.ArgumentTypeError("Emails can only be provided by one method.")
elif args.to is not None:
    email_send = args.to.split(",")
elif args.emailfile is not None:
    filepath = args.emailfile
    with open(filepath) as fp:
        line = fp.readline()
        email_file = open(filepath, "r")
        email_send = email_file.read().split()
        email_file.close()

if (args.smtp is True):
        msg = MIMEMultipart()
        msg['From'] = email_user
        msg['To'] = email_send
        msg['Subject'] = args.subject 
        msg.attach(MIMEText(body,'plain'))
        if args.attachment is not None:
            attachment  =open(args.attachment,'rb')
            part = MIMEText('application','octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',"attachment; filename= "+args.attachment)
            msg.attach(part)
            text = msg.as_string()
            server = smtplib.SMTP('smtp.office365.com',587)
            server.starttls()
            server.login(email_user,email_password)
            server.sendmail(email_user,email_send,text)
            server.quit()
else:
    client_id = args.username
    client_secret = args.password
    credentials = (client_id, client_secret)
    scopes = ['https://graph.microsoft.com/Mail.ReadWrite', 'https://graph.microsoft.com/Mail.Send', 'basic', 'offline_access']
    account = Account(credentials, scopes=scopes)
    connection = Connection(credentials, scopes=scopes)
    # If a token has never been acquired, create one.
    if os.path.isfile('o365_token.txt'):
        print("Token file exists")
    else:
        print("Creating token")
        account.authenticate()
    print("Refreshing token")
    connection.refresh_token()
    for i in email_send:
        print("Sending Email to: {}".format(i))
        m = account.new_message()
        m.to.add(i)
        m.subject = "{}".format(args.subject)
        m.body = body
        m.attachments.add(args.attachment)
        m.send()
        time.sleep(delay) # Sleep between emails
