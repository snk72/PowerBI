#!/usr/bin/python

import smtplib

sender = 'sindre.klepp@leabank.no'
receivers = ['sindre.klepp@leabank.no']

message = """From: From Person <from@fromdomain.com>
To: To Person <to@todomain.com>
Subject: SMTP e-mail test

This is a test e-mail message.
"""

try:
   smtpObj = smtplib.SMTP('leabank-no.mail.protection.outlook.com')  #! Open the command prompt or terminal on your computer.Enter the command: nslookup -type=mx yourdomain.com (replace “yourdomain.com” with your email domain).
   smtpObj.sendmail(sender, receivers, message)         
   print ("Successfully sent email")
except :
   print ("Error: unable to send email")