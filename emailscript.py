import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from validate_email_address import validate_email
from dotenv import load_dotenv

# change these as per use
load_dotenv()
your_email = os.getenv("email") # | "jaypatel****@*****.***"
your_password = os.getenv("password") # | "***************"
 
# establishing connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)
 
# reading the spreadsheet
email_list = pd.read_excel('/Users/jyptel/Desktop/Projects/Python Script/Email script/Seedfunded Startup from 2012-2024/output_chunk_17.xlsx')
# email_list = pd.read_csv('/Users/jyptel/Desktop/Projects/Python Script/Jay-companies-1.csv')
 
# getting the names and the emails
names = email_list['First Name']
emails = email_list['Email']
Company_names = email_list['Company name']

# iterate through the records
for i in range(len(emails)):
    
    # for every record get the name and the email addresses
    
    name = names[i]
    email = emails[i]
    company_name = Company_names[i]
    resume_files = "Jay_patel_resume.pdf"

    if not validate_email(email):
        print(f"Skipping email to {email} as it's not a email valid email address")
        continue
    
    subject = "Full Stack Developer | AWS | ML experienced"
    
    body = f"Hi {name}, \n\nMy name is Jay, I am seeking a learning opportunity with a {company_name} Start-Up where I can learn and grow in a new unique environment. Startups often bring fresh, innovative ideas to the table. As a developer, I thrive on creativity and the opportunity to work on cutting-edge projects. Helping startups grow allows me to be part of this exciting journey of innovation. \n\nI have 3+ years of experience as a software developer. I have worked on the most complex project of an e-commerce application, where I have implemented and learned several backend technologies to improve the performance of the website using techniques like code refactoring, database optimization and caching Strategies using Python and Javascript languages. \n\nI was responsible for adding advanced security features such as authentication and authorization using JWT tokenization. I have ensured that the functions returned are being tested before it goes to the deployment stage. \n\nI have Integrated Restful APIs seamlessly into the backend using technologies such as Python and Javascript. Also been part of handling React and Angular for optimal user experiences across diverse devices and screen sizes. Also integrated several 3rd Party APIs such as ChatGPT Api, Weather API and Payment Gateway API seamlessly into the React and Angular Framework. \n\nStrong working experience with Agile Methodologies and Git Best Practices with good communications. I have Utilized AWS Knowledge to Deploy, manage and scale applications. \n\nPerforming unit testing on functions and methods to ensure they behave as expected. used message broker for building distributed systems. \n\nExperience with working independently and rapidly programming, debugging and dealing with production Django code every day. \n\nBest regards,\nJay Patel\njaypatel97043@gmail.com\n(201) 360-1940"

    # create a MIMEText object to represent the email content
    message = MIMEMultipart()
    message['From'] = your_email
    message['To'] = email
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    with open(resume_files, "rb") as f:
        resume_attachment = MIMEApplication(f.read(), _subtype="pdf")  # Change the _subtype as needed
        resume_attachment.add_header('content-disposition', 'attachment', filename=f.name)
        message.attach(resume_attachment)

    server.sendmail(your_email, email, message.as_string())
server.close()