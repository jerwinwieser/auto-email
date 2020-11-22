import smtplib

from string import Template

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import os

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def attach_document(filename):
    # open pdf file in binary mode
    with open(filename, "rb") as attachment:
        # add file as application/octet-stream
        # email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )
    return part

ADDRESS, PASS = pd.read_csv('__credentials__/details.txt').values[0]

contacts = pd.read_csv('contacts.csv', names = ['name', 'email'])

print(contacts)

names = contacts['name'].values
emails = contacts['email'].values

message_template = read_template('message.txt')
 
def main():
    # set up the SMTP server
    s = smtplib.SMTP(host='smtp.gmail.com', port=587)
    s.starttls()
    s.login(ADDRESS, PASS)
 
    # attach these files
    dirnames = os.listdir('__apply__')
    i = 0

    # For each contact, send the email:
    for name, email, dirname in zip(names, emails, dirnames):
        print('Composing email for : ' + name.title())
      
        # create a message
        msg = MIMEMultipart()

        # add in the actual person name to the message template
        message = message_template.substitute(PERSON_NAME=name.title())
       
        dirname = '__apply__/' + dirname
        filenames = os.listdir(dirname)
        
        filenames = [dirname + '/' + filename for filename in filenames]
        print(filenames)

        for filename in filenames:
            print('Attaching file : ' + filename)
            
            # add attachment to email
            part = attach_document(filename)

            # add attachment to message and convert message to string
            msg.attach(part)
        
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
    print('Email has been send')
