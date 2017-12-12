import plotly.plotly as py
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sb 
import smtplib

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

d = pd.read_excel("emaildata.xlsx")
d['Name'] = d['first name'] +' ' + d['last name']

for emails,names,bid,comp_name,livedate,phonenumber in zip(d['email'],d['Name'],d['businessid'],d['businessname'],d['livedate'],d['phone number']):
    fromaddr = "youraddress"
    toaddr = emails
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = 'Automated Email Test'
    info = {"1": names, "2": comp_name, "3": str(phonenumber), "4": bid,"5": str(livedate)}
    body = """<head>
<style>
div {
    padding-bottom: 30px
}
</style>
</head>


<h1>Dear %s,</h1>


<body>If you are reading this that means an html email was automatically sent via a python script, pulling data from a csv file. It is sent using a windows batch file. Below is information from that csv regarding you:</body>


<p>Company Name: <strong>%s</strong></p>
<p>Phone Number: <strong>%s</strong></p>
<p>Business ID: <strong>%s</strong></p>
<div>LiveDate: <strong>%s</strong></div>
 
<div>Below is a link to a website:</div>
<p><strong><a href="https://www.google.com/">HERE</a></strong></p>







</html>""" %(names,comp_name,str(phonenumber),str(bid),str(livedate))
    msg.attach(MIMEText(body,'html'))
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.starttls()
    server.login(fromaddr,'yourpassword')
    server.sendmail(fromaddr, toaddr, msg.as_string())
    server.quit()
