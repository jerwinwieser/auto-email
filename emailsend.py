import smtplib

from string import Template

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd

def read_template(filename):    
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

ADDRESS, PASS = pd.read_csv('__credentials__/details.txt').values[0]

contacts = pd.read_csv('contacts.csv')
names = contacts['name'].values
emails = contacts['email'].values

message_template = read_template('message.txt')

def main():
    # set up the SMTP server
    s = smtplib.SMTP(host='smtp.gmail.com', port=587)
    s.starttls()
    s.login(ADDRESS, PASS)

    # For each contact, send the email:
    for name, email in zip(names, emails):
        print('Sending email to : ' + name.title())

        # create a message
        msg = MIMEMultipart()

        # add in the actual person name to the message template
        message = message_template.substitute(PERSON_NAME=name.title())

        # setup the parameters of the message
        msg['From']=ADDRESS
        msg['To']=email
        msg['Subject']="This is TEST"
      
        # add in the message body
        msg.attach(MIMEText(message, 'plain'))
      
        # send the message via the server set up earlier.
        s.send_message(msg)
        del msg
      
    # Terminate the SMTP session and close the connection
    s.quit()

if __name__ == '__main__':
    main()