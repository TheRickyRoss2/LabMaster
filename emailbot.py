import smtplib
import mimetypes
import email
import email.mime.application

def sendMail(filename, recipients):
    
    msg = email.mime.Multipart.MIMEMultipart()
    msg['Subject'] = 'Experiment Data'
    msg['From'] = 'adapbot@gmail.com'
    msg['To'] = ", ".join(recipients)
    
    body = email.mime.Text.MIMEText("""Your experiment has finished! Here is the data.
    """)
    msg.attach(body)
    
    fp = open(filename, 'rb')
    att = email.mime.application.MIMEApplication(fp.read(), _subtype="xlsx")
    fp.close()
    att.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(att)
    
    s = smtplib.SMTP('smtp.gmail.com:587')
    s.starttls()
    s.login('adapbot@gmail.com', 'AdapBot143')
    s.sendmail('adapbot@gmail.com', recipients, msg.as_string())
    s.quit()