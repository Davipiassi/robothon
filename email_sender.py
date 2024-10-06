import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

class EmailSender:
    def __init__(self, email, password):
        self.email = email
        self.password = password
        self.server = None
        
    def connect(self):
        try:
            self.server = smtplib.SMTP('smtp.gmail.com', 587)
            self.server.starttls()
            self.server.login(self.email, self.password)
        except Exception:
            raise
    
    def disconnect(self):
        if self.server:
            self.server.quit()
            self.server = None
            
    def setup_message_receivers(self, receivers):
        return ', '.join(receivers)
    
    def setup_message_html_content(self, message, content):
        message.attach(MIMEText(content, 'html'))     
        
    def setup_message_files(self, message, files):
        for file in files:
            with open(file, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                filename = os.path.basename(file)
                part.add_header('Content-Disposition', f'attachment; filename={filename}')
                message.attach(part)
            
    def send_email(self, receivers, subject, content, files):
        if not self.server:
            raise Exception('No SMTP server connected')
        
        message = MIMEMultipart()
        message['From'] = self.email
        message['To'] = self.setup_message_receivers(receivers)
        message['Subject'] = subject
        
        self.setup_message_html_content(message, content)
        self.setup_message_files(message, files)
        
        self.server.sendmail(self.email, receivers, message.as_string())