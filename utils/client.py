import logging
import poplib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

from utils.mail import Email


logger = logging.getLogger(__name__)



class EmailClient(object):
    def __init__(self, email_account, passwd):
        self.email_account = email_account
        self.password = passwd
        self.server = self.connect(self)

    @staticmethod
    def connect(self):
        server = poplib.POP3_SSL("pop.yandex.ru")
        logger.info(server.getwelcome().decode('utf8'))
        server.user(self.email_account)
        server.pass_(self.password)
        return server

    def send_mail(self, to, subject, text):
        smtp_server = smtplib.SMTP("smtp.yandex.ru", 587)
        smtp_server.starttls()
        smtp_server.login(self.email_account, self.password)

        msg = MIMEText(text, "plain", "utf-8")
        msg['From'] = self.email_account
        msg['To'] = to
        msg['Subject'] = Header(subject)
        print(msg.as_string())
        smtp_server.sendmail(msg['From'], [msg['To']], msg.as_string())
        #
        # msg = MIMEMultipart()
        # msg['From'] = self.email_account
        # msg['To'] = to
        # msg['Subject'] = subject
        # msg.attach(MIMEText(text, 'plain'))

        # body = "\r\n".join(("From: %s" % self.email_account, "To: %s" % to, "Subject: %s" % Header(subject, 'utf-8'), "", MIMEText(text, 'plain', 'utf-8').as_string()))
        # body = "\r\n".join(("From: %s" % self.email_account, "To: %s" % to, "Subject: %s" % subject, "", text))
        # smtp_server.sendmail(self.email_account, [to], body)
        # smtp_server.sendmail(msg['From'], msg['To'], msg.as_string())
        smtp_server.quit()

    def get_mails_list(self):
        _, mails, _ = self.server.list()
        return mails

    def get_mails_count(self):
        mails = self.get_mails_list()
        return len(mails)

    def get_mail_by_index(self, index):
        resp_status, mail_lines, mail_octets = self.server.retr(index)
        return Email(mail_lines)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is None:
            logger.info('exited normally\n')
            self.server.quit()
        else:
            logger.error('raise an exception! ' + str(exc_type))
            self.server.close()
            return False # Propagate



if __name__ == '__main__':
    useraccount = "XXXXX"
    password = "XXXXXX"

    client = EmailClient(useraccount, password)
    num = client.get_mails_count()
    print(num)
    for i in range(1, num):
        print(client.get_mail_by_index(i))
