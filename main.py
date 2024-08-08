import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from jinja2 import FileSystemLoader, Environment
import pandas as pd


class EmailSender:
    def __init__(self, smtp_server, smtp_port, smtp_username, smtp_password):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.smtp_username = smtp_username
        self.smtp_password = smtp_password
        self.env = Environment(loader=FileSystemLoader('templates'))

    def render_mail(self, template_name, **kwargs):
        """Render email content from a Jinja2 template."""
        template = self.env.get_template(template_name)
        return template.render(**kwargs)

    def create_message(self, receiver, body, subject='测试邮件', sender=None):
        """Create a MIME email message."""
        if sender is None:
            sender = self.smtp_username

        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))
        return msg

    def send_email(self, receiver, template_name, **template_vars):
        """Send an email with the rendered template."""
        body = self.render_mail(template_name, **template_vars)
        msg = self.create_message(receiver, body)

        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()  # If using STARTTLS
                server.login(self.smtp_username, self.smtp_password)
                server.send_message(msg)
            print('邮件发送成功')
        except Exception as e:
            print(f'邮件发送失败: {e}')


# Usage example
if __name__ == '__main__':

    sheet_data = pd.read_excel('excel/records.xlsx').to_dict(orient='records')

    smtp_server = 'smtp.163.com'
    smtp_port = 25  # Use 587 for STARTTLS
    smtp_username = input('input your mail addr\n')
    smtp_password = input('input your password\n')
    receiver_addr = input("input your receiver's mail addr\n")


    data = {'Receiver': 'winterbear',
            'Region_A': 'RA',
            'Region_B': 'RB',
            'data': sheet_data} # Change it based on your template

    email_sender = EmailSender(smtp_server, smtp_port, smtp_username, smtp_password)
    email_sender.send_email(receiver_addr, 'Email.html',**data)
