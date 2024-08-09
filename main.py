import configparser
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from jinja2 import FileSystemLoader, Environment
import pandas as pd


class EmailSender:
    def __init__(self, smtp_server, smtp_port, smtp_username, smtp_password):
        """
        初始化EmailSender类, 设置SMTP服务器配置和Jinja2模板环境

        :param smtp_server: SMTP服务器地址
        :param smtp_port: SMTP服务器端口
        :param smtp_username: SMTP服务器用户名
        :param smtp_password: SMTP服务器密码
        """
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.smtp_username = smtp_username
        self.smtp_password = smtp_password
        self.env = Environment(loader=FileSystemLoader('templates'))

    def render_mail(self, template_name, **kwargs):

        """
        渲染邮件内容使用Jinja2模板

        :param template_name: 模板文件名
        :param kwargs: 模板变量
        :return: 渲染后的邮件内容
        """
        template = self.env.get_template(template_name)
        return template.render(**kwargs)

    def create_message(self, receiver, body, subject='测试邮件', sender=None):

        """
        创建MIME邮件消息

        :param receiver: 接收者邮件地址
        :param body: 邮件内容
        :param subject: 邮件主题
        :param sender: 发送者邮件地址, 默认为SMTP用户名
        :return: MIME邮件消息
        """
        if sender is None:
            sender = self.smtp_username

        # 创建一个MIMEMultipart对象
        msg = MIMEMultipart()
        msg['From'] = sender
        msg['To'] = receiver
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))  # 添加HTML邮件内容
        return msg

    def send_email(self, receiver, template_name, **template_vars):

        """
        发送带有渲染模板的邮件

        :param receiver: 接收者邮件地址
        :param template_name: 模板文件名
        :param template_vars: 模板变量
        """
        # 渲染邮件内容
        body = self.render_mail(template_name, **template_vars)
        # 创建邮件消息
        msg = self.create_message(receiver, body)

        try:
            # 连接到SMTP服务器并发送邮件
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()  # 如果使用STARTTLS
                server.login(self.smtp_username, self.smtp_password)
                server.send_message(msg)
            print('邮件发送成功')
        except Exception as e:
            print(f'邮件发送失败: {e}')

def get_smtp_config(config_file='config.ini'):
    """
    获取SMTP配置，如果配置文件中没有相关信息，则要求用户输入并更新配置文件

    :param config_file: 配置文件名
    :return: SMTP服务器地址、端口、用户名和密码
    """
    config = configparser.ConfigParser()
    config.read(config_file)

    if 'SMTP' not in config:
        config['SMTP'] = {}

    smtp_server = config['SMTP'].get('server', '')
    smtp_port = config['SMTP'].getint('port', 0)
    smtp_username = config['SMTP'].get('username', '')
    smtp_password = config['SMTP'].get('password', '')

    if not smtp_server:
        smtp_server = input('请输入SMTP服务器地址: ')
        config['SMTP']['server'] = smtp_server

    if not smtp_port:
        smtp_port = int(input('请输入SMTP服务器端口: '))
        config['SMTP']['port'] = str(smtp_port)

    if not smtp_username:
        smtp_username = input('请输入SMTP用户名: ')
        config['SMTP']['username'] = smtp_username

    if not smtp_password:
        smtp_password = input('请输入SMTP密码: ')
        config['SMTP']['password'] = smtp_password

    with open(config_file, 'w') as configfile:
        config.write(configfile)

    return smtp_server, smtp_port, smtp_username, smtp_password
# 使用示例

if __name__ == '__main__':
    # 从Excel文件读取数据
    sheet_data = pd.read_excel('excel/records.xlsx').to_dict(orient='records')

    # 获取SMTP服务器配置
    smtp_server, smtp_port, smtp_username, smtp_password = get_smtp_config()
    # 获取接收者地址
    receiver_addr = input("请输入接收者的邮箱地址: ")

    # 准备模板数据
    data = {
        'Receiver': 'winterbear',
        'Region_A': 'RA',
        'Region_B': 'RB',

        'data': sheet_data  # 根据你的模板进行调整
    }

    # 创建EmailSender实例并发送邮件
    email_sender = EmailSender(smtp_server, smtp_port, smtp_username, smtp_password)
    email_sender.send_email(receiver_addr, 'Email.html', **data)
