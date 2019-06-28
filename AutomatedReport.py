#! python3
import smtplib
import pyodbc
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#insert your db server information 
server = ''
conn = pyodbc.connect(server)
cursor = conn.cursor()
#query from database to get information
query = """SELECT * """ 

#using pandas to read sql query then write to excel file
pandas = pd.read_sql(query,conn)
conn.close()
writer = pd.ExcelWriter('YourExcelFileName.xlsx',engine='xlsxwriter')
pandas.to_excel(writer,'Sheet1',index=False)

#changing the layout of the excel file to make it more readable
worksheet = writer.sheets['Sheet1']
#set excel sheet zoom
worksheet.set_zoom(90)
#set column sizes so the excel columns wouldnt be squished up 
worksheet.set_column('A:B',15)
worksheet.set_column('C:D',50)
worksheet.set_column('G:H',30)
writer.save()

#Log in to email account.
fromaddr = "from@gmail.com"
toaddr = "to@gmail.com"
msg = MIMEMultipart()
msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = "Email Subject"
part = MIMEBase('application', "octet-stream")
 # This is the same file name from above
part.set_payload(open("YourExcelFileName.xlsx", "rb").read())   
#Need to encode so that query can be read 
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="YourExcelFileName.xlsx"')
body = "email body here"
msg.attach(MIMEText(body, 'plain'))
msg.attach(part)
#configure client based on your service
smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
smtpObj.ehlo()
smtpObj.starttls()
smtpObj.ehlo()
#log in to your email account 
smtpObj.login('from@gmail.com', 'pass')
smtpObj.sendmail(fromaddr, toaddr, msg.as_string())
server = smtplib.SMTP('mail')
server.quit()

