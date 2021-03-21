import pandas as pd
import smtplib 
import string

from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 

from PIL import Image, ImageDraw, ImageFont
from pandas import ExcelWriter
from pandas import ExcelFile
p_wE_d ="" #Enter password 
fromaddr = "" # Enter Email id

# reading the excel file
df = pd.read_excel('file.xlsx', sheet_name='Sheet1')

#loop for printing certificate
for i in df.index:
	image = Image.open('cert.jpg')
	draw = ImageDraw.Draw(image)
	font = ImageFont.truetype('Montserrat-ExtraBoldItalic_8abc4d055eace304d7ae98636dde5f4d.ttf', size=65)

	color = 'rgb(61, 90, 241)'
	name = df['Name'][i] 
	name.upper()
	# position = df['position'][i]
	print(i+1,name)
	# print(i)
	draw.text((940, 440), name, fill=color, font=font)

	imageName = "Certificate "+name+".pdf"
	image.save(imageName)
	
	# intialization of sending the email
	toaddr = df['Email ID'][i] # fetches email id from excel file
	msg = MIMEMultipart() 
	msg['From'] = fromaddr 

	# storing the receivers email address 
	msg['To'] = toaddr 

	# storing the subject 
	msg['Subject'] = "Certificate for Participation in WebCraft Competition 2021"

	# string to store the body of the mail 
	body = '''Thank you for participating in WebCraft Competiton 2021
			  
	Thanks and Regards
	Team CSI 
			'''

	# attach the body with the msg instance 
	msg.attach(MIMEText(body, 'plain')) 

	# open the file to be sent 
	filename = name +".pdf"
	attachment = open(imageName, "rb")
	p = MIMEBase('application', 'octet-stream') 

	# To change the payload into encoded form 
	p.set_payload((attachment).read()) 

	# encode into base64 
	encoders.encode_base64(p) 

	p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 

	# attach the instance 'p' to instance 'msg' 
	msg.attach(p) 

	# creates SMTP session 
	s = smtplib.SMTP('smtp.gmail.com', 587) 

	# start TLS for security 
	s.starttls() 

	# Authentication 
	s.login(fromaddr, p_wE_d) 

	# Converts the Multipart msg into a string 
	text = msg.as_string() 

	# sending the mail 
	s.sendmail(fromaddr, toaddr, text) 

	# terminating the session 
	


s.quit() 