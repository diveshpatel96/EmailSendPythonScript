from random import randint
import pandas as pd
from email.message import EmailMessage
import smtplib
import logging
import time
import sys
import pdfkit
import os

path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

logging.basicConfig(filename='mail.log', level=logging.DEBUG)

totalSend = 1
if(len(sys.argv) > 1):
    totalSend = int(sys.argv[1])

emaildf = pd.read_csv('gmail.csv')
contactsData = pd.read_csv('contacts.csv')
subjects = pd.read_csv("subjects.csv")
bodies = ['body.txt', 'body2.txt', 'body3.txt', 'body4.txt', 'body5.txt']



def send_mail(name, email, emailId, password, bodyFile, subjectPhrase):
    newMessage = EmailMessage()

    # Invoice Number and Subject
    invoiceNo = randint(10000000, 99999999)
    subject = subjectPhrase + str(invoiceNo) + " of item"

    newMessage['Subject'] = subject
    newMessage['From'] = emailId
    newMessage['To'] = email
    transaction_id = randint(10000000, 99999999)

    # Mail Body Content
    body = open(bodyFile, 'r').read()
    body = body.replace('$name', name)
    body = body.replace('$email', email)
    body = body.replace('$invoice_no', str(transaction_id))

    # Mail PDF File
    html = open('html_code.html', 'r').read()
    html = html.replace('$name', name)
    html = html.replace('$email', email)
    html = html.replace('$invoice_no', str(transaction_id))
    # saving the changes to html_code.html
    with open('html_code.html', 'w') as f:
        f.write(html)

    file = "Bill" + str(invoiceNo) + ".pdf"
    pdfkit.from_file('html_code.html', file, configuration=config)

    html = open('html_code.html', 'r').read()
    html = html.replace(str(transaction_id), '$invoice_no')
    with open('html_code.html', 'w') as f:
        f.write(html)

    newMessage.set_content(body)

    try:
        with open(file, 'rb') as f:
            file_data = f.read()
            file_name = f.name
        newMessage.add_attachment(
            file_data, maintype='application', subtype='octet-stream', filename=file_name)

        with smtplib.SMTP_SSL('smtp.mail.ru',465) as smtp:
            smtp.login(emailId, password)
            smtp.send_message(newMessage)
            smtp.quit()

        os.remove(file)

        print(f"send to {email} by {emailId} successfully : {totalSend}")
        logging.info(
            f"send to {email} by {emailId} successfully : {totalSend}")

    except smtplib.SMTPResponseException as e:
        error_code = e.smtp_code
        error_message = e.smtp_error
        print(f"send to {email} by {emailId} failed")
        logging.info(f"send to {email}  by {emailId} failed")
        print(f"error code: {error_code}")
        print(f"error message: {error_message}")
        logging.info(f"error code: {error_code}")
        logging.info(f"error message: {error_message}")

        remove_email(emailId, password)


def start_mail_system():
    global totalSend
    j = 0  # for sender emails
    k = 0  # for bodies
    l = 0  # for subjects
    

    for i in range(len(contactsData)):
        emaildf = pd.read_csv('gmail.csv')
        if(j >= len(emaildf)):
            j = 0
        time.sleep(0.2)
        send_mail(contactsData.iloc[i]['name'], contactsData.iloc[i]['email'], emaildf.iloc[j]['email'],
                  emaildf.iloc[j]['password'], bodies[k], subjects.iloc[l]['subject'])
        totalSend += 1
        j = j + 1
        k = k + 1
        l = l + 1
       
        if j == len(emaildf):
            j = 0
        if k == len(bodies):
            k = 0
        if l == len(subjects):
            l = 0
        
    quit()


def remove_email(emailId, password):
    df = pd.read_csv('gmail.csv')
    index = df[df['email'] == emailId].index
    df.drop(index, inplace=True)
    df.to_csv('gmail.csv', index=False)
    print(f"{emailId} removed from gmail.csv")
    logging.info(f"{emailId} removed from gmail.csv")


try:
    for i in range(6):
        start_mail_system()
except KeyboardInterrupt as e:
    print(f"\n\ncode stopped by user")
