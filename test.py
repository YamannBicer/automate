import smtplib
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.header import Header
import dotenv
import os
import datetime

dotenv.load_dotenv(override=True)

def send_email():
    # email_sender = os.getenv("EMAIL_SENDER")
    # email_password = os.getenv("EMAIL_PASSWORD")
    # to_emails = os.getenv("TO_EMAILS")
    # cc_emails = os.getenv("CC_EMAILS")

    email_sender = "yamanomated@gmail.com"
    email_password = os.getenv("GMAIL_PASSWORD")
    to_emails = ["yaman.bicer@puretechnology.com.tr", "fatih.gultekin@pureenergy.com.tr",
                 "feyza.kesilmis@pureenergy.com.tr"]
    cc_emails = ["yamanbicer@gmail.com"]




    # semtp_server="smtp.office365.com"
    smtp_server="smtp.gmail.com"
    smtp_port=587

    subject = f"Türkiye DSG P&L Daily Report {datetime.date.today() - datetime.timedelta(1)}"
    body = """
    Dear Puritans,
    
    Attached is the Türkiye DSG P&L Daily Report for 12/03/2025. It provides an overview of the previous day's management performance, including the daily P&L amount.
    
    Best
    """
    report = os.path.abspath(f"reports-pdf/TürkiyeDSGDailyReport_{datetime.date.today() - datetime.timedelta(1)}.pdf")

    message = MIMEMultipart()
    message["Subject"] = subject
    message["From"] = email_sender
    message["To"] = to_emails
    message["Cc"] = cc_emails


    # Attach the email body
    message.attach(MIMEText(body, "plain"))

    # Open, encode, and attach the file
    with open(report, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode the file in Base64
    encoders.encode_base64(part)

    # Properly encode the filename to handle non-ASCII characters
    filename = os.path.basename(report)
    encoded_filename = Header(filename, 'utf-8').encode()
    part.add_header("Content-Disposition", f'attachment; filename="{encoded_filename}"')
    message.attach(part)

    # recipients = to_emails + cc_emails
    recipients = to_emails

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.ehlo()           # Identify ourselves to the server
        server.starttls()       # Upgrade to a secure TLS connection
        server.ehlo()           # Re-identify as a secure connection
        server.login(email_sender, email_password)
        server.sendmail(email_sender, recipients, message.as_string())

    print("Email sent successfully!")

send_email()