# -*- coding: utf-8 -*-
"""
Created on Sun Feb 16 16:57:34 2020

@author: slice
"""
import imaplib
import smtplib
import time
import email
import config
import email
# -------------------------------------------------
#
# Utility to read email from Gmail Using Python
#
# ------------------------------------------------

def read_email_from_outllook():
    try:
        mail = imaplib.IMAP4_SSL(config.imap_server,config.imap_port)
        mail.login(config.UserName,config.Password)
        mail.select('inbox')

        type, data = mail.search(None, 'UNSEEN')
        mail_ids = data[0]

        id_list = mail_ids.split()   
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])


        for i in range(latest_email_id,first_email_id, -1):
            typ, data = mail.fetch(i, '(RFC822)' )

            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1])
                    email_subject = msg['subject']
                    email_from = msg['from']
                    print ('From : ' + email_from + '\n')
                    print ('Subject : ' + email_subject + '\n')

    except:
        print('Exception occured')
read_email_from_outllook()