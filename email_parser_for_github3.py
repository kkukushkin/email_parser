#!/usr/bin/env python
# coding: utf-8

import imaplib
import base64
import email
import getpass
from email.header import Header, decode_header, make_header
from email import policy
import quopri
import pandas as pd

   
def parse_my_email(email_host, email_user, email_pass, local_address, num_mails):
    mail = imaplib.IMAP4_SSL(email_host)
    mail.login(email_user, email_pass)
    mail.select('INBOX')       #folder can be changed. Names of folders can be found by command: print(mail.list())
    if what_to_look_for == 'ALL':                                             
        (typ, msgnums) = mail.search(None, 'ALL')
    else:
        (typ, msgnums) = mail.search(None, what_to_look_for, choose_email, 'SINCE', date_from, 'BEFORE', date_till) 
    parse_all(msgnums, mail, local_address, num_mails)
    mail.close()
    mail.logout()
    print(f'Please, find your file here: {local_address}mail_parsing_result.csv')
     
def decode_mime_words(header):                       #function helps to avoid headers decoding error
    return u''.join(
        word.decode(encoding or 'utf8') if isinstance(word, bytes) else word 
        for word, encoding in email.header.decode_header(header))

def parse_all(msgnums, mail, local_address, num_mails):
    all_together = []                                           #parsing mails
    for messages in msgnums:
        for message in messages.split()[0:int(num_mails)]:            
            typ, raw_data1 = mail.fetch(message, '(RFC822)')
            message = email.message_from_string(raw_data1[0][1].decode('UTF-8'))
            date = message['Date']
            header = message['subject']
            workon = decode_mime_words(header)
            for p in message.walk():
                if p.get_content_type() == "text/plain":
                    body = p.get_payload(decode=True).decode(p.get_content_charset())
                    all_together.append([date, workon, body])
                    columns = ('date', 'header', 'body')
                    df = pd.DataFrame(data = all_together, columns = columns)
                    df.index = range(1,len(df)+1)
                    df.to_csv(f'{local_address}mail_parsing_result.csv', sep=';', index = True, encoding='utf-8-sig') 

if __name__ == '__main__':                             #getting information from user
    email_host = input('Enter your email host: e.g. imap.yandex.ru ')
    email_user = input('Enter your email address: ')
    email_pass = getpass.getpass(prompt = 'Enter your email password: ')
    folder = input('Enter email folder (INBOX or Sent): ') 
    what_to_look_for = input('Please, enter FROM or TO, or ALL: ')
    if what_to_look_for == 'ALL':
        num_mails = input('Enter quantity of mails you want to parse: ')
        local_address = input('Enter your local address, where csv file should be saved:')
        parse_my_email(email_host, email_user, email_pass, local_address, num_mails)
    else:
        choose_email = input('Please, enter email of a person you need emails from or to: ')
        date_from = input('Please, enter start date (in format e.g. 1-Dec-2019): ')
        date_till = input('Please, enter finish date (in format e.g. 1-Dec-2019): ')
        num_mails = input('Enter quantity of mails you want to parse: ')
        local_address = input('Enter your local address, where csv file should be saved:')
        parse_my_email(email_host, email_user, email_pass, local_address, num_mails)
