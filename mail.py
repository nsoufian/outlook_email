from exchangelib import Credentials, Account, Message, Mailbox
from time import sleep
import argparse

def getAccount(login, password):
    credentials = Credentials(login, password)
    return Account(login, credentials=credentials, autodiscover=True)

def sendEmail(account, subject, body, recipeint):
    m = Message(
        account=account,
        subject=subject,
        body=body,
        to_recipients=[
            Mailbox(email_address=recipeint),
        ],
    )
    m.send()
    print('Successfully try to send email to {} '.format(recipeint))

def sendEmails(account, subject, body, recipeints, pause_delay=1):
    for recipeint in recipeints:
        sendEmail(account, subject, body, recipeint)
        sleep(pause_delay)
        


def main(login, password,subject, body, recipeints, delay):
    account = getAccount(login, password)
    sendEmails(account, subject, body, recipeints, delay)

if __name__ == '__main__':
    parser = argparse.ArgumentParser('Send email via outlook...')
    parser.add_argument('--subject', '-s', required=True,type=str, help="Subject of the email to send.")
    parser.add_argument('--body', '-b', required=True, type=argparse.FileType('r'), help="Path of file contains the body of the email to send.")
    parser.add_argument('--login', '-l', required=True, type=str, help="Outlook account login.")
    parser.add_argument('--password', '-p', required=True, type=str, help="Outlook account password.")
    parser.add_argument('--emails', '-e', required=True, type=argparse.FileType('r'), help="Path of file contains emails separated by new line..")
    parser.add_argument('--delay', '-d', type=int, default=1, help="Number of second to wait between each email.")
    args = parser.parse_args()

    subject = args.subject
    body = args.body.read()
    login = args.login
    password= args.password
    emails =[email[:-1] for email in args.emails.readlines()]
    delay = args.delay
    main(login, password, subject, body, emails, delay)