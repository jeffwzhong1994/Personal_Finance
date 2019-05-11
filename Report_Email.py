# -*- coding: utf-8 -*-
import smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import formatdate
from email import encoders
import schedule 
import time

MY_ADDRESS = 'Myaddress'
PASSWORD = 'mypassword'
PATH = '/Users/Email_pipeline/whatever/'

def get_contacts(filename):
    """
    Return two lists names, emails containing names and email addresses
    read from a file specified by filename.
    """
    names = []
    emails = []
    with open(filename, mode='r', encoding='utf-8') as contacts_file:
        for a_contact in contacts_file:
            names.append(a_contact.split()[0])
            emails.append(a_contact.split()[1])

    return names, emails

def read_template(filename):
    """
    Returns a Template object comprising the contents of the 
    file specified by filename.
    """
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def main():
    names, emails = get_contacts('%smycontacts.txt' % (PATH)) # read contacts
    message_template = read_template('%smessage.txt' % (PATH))

    # set up the SMTP server
    s = smtplib.SMTP(host='smtp.google.com', port = 465)
    s.starttls()
    s.login(MY_ADDRESS, PASSWORD)

    # For each contact, send the email:
    for name, email in zip(names, emails):
        msg = MIMEMultipart()       # create a message

        # add in the actual person name to the message template
        message = message_template.substitute(PERSON_NAME=name.title())

        # Prints out the message body for our sake
        print(message)

        # setup the parameters of the message
        msg['From']=MY_ADDRESS
        msg['To']=email
        msg['Subject']="钟文略自研的财务自动管理系统v2.1"
        
        # add in the message body
        msg.attach(MIMEText(message, 'plain'))
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open("%sMay.xlsx" %(PATH), "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename = "May_10.xlsx"')
        msg.attach(part)
        # send the message via the server set up earlier.
        s.send_message(msg)
        del msg

    # Terminate the SMTP session and close the connection
    s.quit()

if __name__ == '__main__':
    main()