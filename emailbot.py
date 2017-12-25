import smtplib
import email
import email.mime.application

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
    email_message = email.mime.Multipart.MIMEMultipart()
    email_message['Subject'] = 'Experiment Data'
    email_message['From'] = EMAIL_USERNAME
    email_message['To'] = ", ".join(recipients)
    
    body = email.mime.Text.MIMEText("""Your experiment has 
    finished!Here is the data.""")
    email_message.attach(body)

    file_reader = open(attached_file_name, 'rb')
    attachment = email.mime.application.MIMEApplication(
        file_reader.read(),
        _subtype="xlsx"
    )

    file_reader.close()

    attachment.add_header(
        'Content-Disposition',
        'attachment',
        filename=attached_file_name
    )

    email_message.attach(attachment)

    send_email_connection = smtplib.SMTP(SMTP_SERVER)
    send_email_connection.starttls()
    send_email_connection.login(EMAIL_USERNAME, EMAIL_PASSWORD)

    send_email_connection.sendmail(
        EMAIL_USERNAME,
        recipients,
        email_message.as_string()
    )

    send_email_connection.quit()
