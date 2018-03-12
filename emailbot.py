import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

EMAIL_USERNAME = "adapbot@gmail.com"
EMAIL_PASSWORD = "AdapBot143"
SMTP_SERVER = "smtp.gmail.com:587"


def send_mail(attached_file_name, recipients):
    """
    This is an helper function for sending the data to your email
    :param attached_file_name: file to send (full path)
    :param recipients: comma separated list of email recipients
    :return: None
    """
    email_message = MIMEMultipart()
    email_message['Subject'] = 'Experiment Data'
    email_message['From'] = EMAIL_USERNAME
    email_message['To'] = ", ".join(recipients)
    email_message['Date'] = formatdate(localtime=True)
    email_message.attach(MIMEText("Your experiment has finished!"))
    attach = MIMEBase('application', 'octet-stream')
    attach.set_payload(open(attached_file_name, 'rb').read())
    encoders.encode_base64(attach)
    attach.add_header('Content-Disposition', 'attachment; filename="{}"'.format(attached_file_name))
    email_message.attach(attach)


    send_email_connection = smtplib.SMTP(SMTP_SERVER)
    send_email_connection.starttls()
    send_email_connection.login(EMAIL_USERNAME, EMAIL_PASSWORD)

    send_email_connection.sendmail(
        EMAIL_USERNAME,
        recipients,
        email_message.as_string()
    )

    send_email_connection.quit()
