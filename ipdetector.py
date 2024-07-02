import smtplib
import configparser
from email.mime.text import MIMEText
import requests
import time
import win32serviceutil
import win32service
import win32event
import servicemanager


def get_public_ip():
    response = requests.get('https://api.ipify.org?format=json')
    if response.status_code == 200:
        return response.json()['ip']
    else:
        raise Exception('Failed to get public IP')


class MyService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'MyService'
    _svc_display_name_ = 'My Service Display Name'

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def main(self):
        # 服务的主要逻辑
        config = configparser.ConfigParser()
        config.read('config.ini')
        mail_host = config['Default']['host']
        mail_code = config['Default']['password']
        sender = config['Default']['account']
        receivers = config['Default']['receivers']
        public_ip = get_public_ip()
        print(f'Your public IP is: {public_ip}')
        while True:
            if get_public_ip() == public_ip:
                time.sleep(5)
            else:
                public_ip = get_public_ip()
                try:
                    message = MIMEText(
                        f'IP发生变更，新的公网IP地址为：{public_ip}'
                    )
                    message['From'] = sender
                    message['Subject'] = 'IP变更提醒'
                    mail = smtplib.SMTP(mail_host, 25)
                    mail.login(sender, mail_code)
                    mail.sendmail(sender, receivers, message.as_string())
                    print('发送成功')
                except smtplib.SMTPException as e:
                    print(f'发送失败，错误原因：{e}')


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(MyService)
