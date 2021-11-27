# Outlook python sender
python script to send email via outlook exchange.
## Install

to start the installation open the terminal.

Clone the repository
```sh
git clone https://github.com/nsoufian/outlook_email.git
```

Move to directory
```sh
cd outlook_email
```

Install dependencies

```sh
python3 -m pip install --user -r requirements.txt
```

## Usage 

```
python3 mail.py --help
usage: Send email via outlook... [-h] --subject SUBJECT --body BODY --login LOGIN --password PASSWORD --emails EMAILS [--delay DELAY]

optional arguments:
  -h, --help            show this help message and exit
  --subject SUBJECT, -s SUBJECT
                        Subject of the email to send.
  --body BODY, -b BODY  Path of file contains the body of the email to send.
  --login LOGIN, -l LOGIN
                        Outlook account login.
  --password PASSWORD, -p PASSWORD
                        Outlook account password.
  --emails EMAILS, -e EMAILS
                        Path of file contains emails separated by new line..
  --delay DELAY, -d DELAY
                        Number of second to wait between each email.

```