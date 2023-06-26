import configparser
import csv
import os
import re
import smtplib
import sys
from datetime import datetime as dt
from dataclasses import dataclass

from dotenv import load_dotenv
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class NotEmail(Exception):
    def __init__(self, *args):
        if args:
            self.message = args[0]
        else:
            self.message = None

    def __str__(self):
        if self.message:
            return 'Error, {0} '.format(self.message)
        else:
            return 'Error has been raised'


@dataclass()
class LogMessageMail:
    def message_log(self, unique, mes, type_mes='INFO'):
        now = dt.now()
        time_log = now.strftime("%d.%m.%Y-%H_%M_%S")
        time_for_file = now.strftime("%d.%m.%Y")
        mes_data = [
            [unique, mes, type_mes, time_log]
        ]
        log_dir = os.path.join(os.getcwd(), 'log_send')
        if not os.path.exists(log_dir):
            os.mkdir(log_dir)
        with open(r'log_send\send_mails_' + time_for_file + '.csv', 'a', encoding='cp1251',
                  newline="") as file:
            # !!!! чтобы в windows разделитель в виде ; bcgjkmpet
            writer = csv.writer(file, delimiter=";")
            writer.writerows(
                mes_data
            )


@dataclass()
class SendMail(LogMessageMail):
    """
    Отправляем мэйл
    """
    theme_global: str = None
    html_global: str = None
    pattern_for_email_re: str = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
    id_image_for_include: str = None
    files_image_for_include: list = None

    def __post_init__(self):
        config = configparser.ConfigParser()
        # file_settings = os.path.join(os.getcwd(), 'settings_server_mail.ini')
        # config.read(file_settings, encoding='utf-8-sig')
        load_dotenv()
        self.email = str(os.getenv('LOGIN_MAIL')),
        self.password = str(os.getenv('PASSWORD')),
        self.ip_server = str(os.getenv('IP_SERVER')),
        self.port = str(os.getenv('PORT')),
        self.from_str = str(os.getenv('FROM_STR'))
        # self.theme = str(config['Mail']['theme_mail'])
        # try:
        #     self.email = str(config['MailServer']['sender_email']),
        #     self.password = str(config['MailServer']['password']),
        #     self.ip_server = str(config['MailServer']['ip_server']),
        #     self.port = str(config['MailServer']['port']),
        #     self.from_str = str(config['MailServer']['from_str'])
        #     # self.theme = str(config['Mail']['theme_mail'])
        # except(Exception,):
        #     raise FileNotFoundError(
        #         'Файл конфигурации отправки settings_mail.ini не найден или содержит ошибки.'
        #     )
        if self.theme_global is None:
            self.theme_global = ''

    def check_email(self, list_mails):
        """Проверяем входящий список e-mail на соответствие формату"""
        # Проверяем e-mail
        def regular_chack(email):
            """Проверяем email"""
            match_email = re.fullmatch(self.pattern_for_email_re, email)
            if not match_email:
                self.message_log(email, 'Значение исключенноб т.к. не соответствует формату e-mail.', 'ERROR')
            else:
                return email
        return list(filter(regular_chack, list_mails))

    def push_mail_group(
            self,
            emails_for_send: list,
            html=None,
            text=None,
            cc_emails: list = None,
            bcc_emails: list = None,
            theme_param=None,
            dict_included_files=None

    ):
        """
        Отправка письма как в текстовом варианте так и в HTML
        :param list_image_for_include: Графические файлы для включения в шаблон HTML
        :param dict_included_files: Словарь с файлом(и) для отправки(путей) где ключ это наименование файла в письме
        :param theme_param: Можно указать и в параметре, свойства классв
        :param emails_for_send: Cписок email графа "Кому"
        :param cc_emails: Список email графа "Копия"(необязательный)
        :param bcc_emails: Список скрытых email графа скрытая "Скрытая копия"(необязательный)
        :param theme: Тема письма. Если не указана то берется из файла settings_mail.ini
        :param html: Текст в формате HTML
        :param text: Альтернативный текст
        :return:
        """
        global server
        # Проверяем основвной список emails
        emails_for_send = self.check_email(emails_for_send)
        print('Отфильтрованный список: ', emails_for_send)
        # Выбираем тему.
        if theme_param is not None:
            theme = theme_param
        elif self.theme_global is not None:
            theme = self.theme_global
        else:
            theme = ''
        # "alternative" - Делающий возможным выбор клиенту между текстом и HTML(но в результате тестирования)
        message = MIMEMultipart('mixed')
        message["Subject"] = theme
        # От кого(Можно указать текст)
        message["From"] = self.from_str
        message["To"] = ', '.join(emails_for_send)
        if cc_emails is not None:
            cc_emails = self.check_email(cc_emails)
            message["cc"] = ', '.join(cc_emails)
        else:
            cc_emails = []
        if bcc_emails is not None:
            bcc_emails = self.check_email(bcc_emails)
            message["bcc"] = ', '.join(bcc_emails)
        else:
            bcc_emails = []
        # Проверяем осатлся ли хоть один e-mail из 3 групп
        # после проверки, если нет то возвращаем ошибку
        if len(emails_for_send) == 0 and len(cc_emails) == 0 and len(bcc_emails) == 0:
            raise NotEmail(
                'В результате проверки не найден адресса e-mail для отправки.'
            )
        # Важен порядок обработки
        # if text is None:
        #     text = 'The content is only available in HTML'
        # part1 = MIMEText(text, "plain")
        if html is not None:
            part2 = MIMEText(html, "html")
        else:
            part2 = MIMEText(self.html_global, "html")
        # message.attach(part1)
        message.attach(part2)

        if self.files_image_for_include is not None:
            for num, img in enumerate(self.files_image_for_include):
                # print(img)
                image = MIMEImage(open(img, 'rb').read())
                id_image_name = f'{self.id_image_for_include}{num}'
                # print('Индентификатор имени графического файла для HTML', id_image_name)
                image.add_header('Content-ID', f'<{id_image_name}>')
                message.attach(image)
        if dict_included_files is not None:
            # Если есть файлы для включенияв письмо то добавляем в список:
            for key, file in dict_included_files.items():
                attachment = MIMEBase('application', "octet-stream")
                header = 'Content-Disposition', 'attachment; filename="%s"' % key
                try:
                    with open(file, "rb") as fh:
                        data = fh.read()
                    print('Файл для включения', file)
                    attachment.set_payload(data)
                    encoders.encode_base64(attachment)
                    attachment.add_header(*header)
                    message.attach(attachment)
                except IOError:
                    msg = "Error opening attachment file %s" % file
                    print(msg)
                    sys.exit(1)
        print('Сервер: ', f'{self.ip_server[0]}:{self.port[0]}')
        print('Отправляем от имени: ', self.email[0])
        emails = emails_for_send + bcc_emails + cc_emails
        server = smtplib.SMTP(f'{self.ip_server[0]}:{self.port[0]}')
        server.login(self.email[0], self.password[0])
        server.sendmail(self.email[0], emails, message.as_string())
        server.quit()


if __name__ == '__main__':
    # для примера
    send = SendMail(
        theme_global='Тестовое сообщение',
        id_image_for_include='include',
        files_image_for_include=[os.path.join(os.getcwd(), 'img', 'logo.png')]
    )
    try:
        send.push_mail_group(
            html=''''
            <html>
            <head>
                <meta charset="utf-8">
            </head>
                <body>
                    <div style="text-align: center;">
                        <h1>TEST</h1>
                        <div style="text-align: center;"><img src="cid:include0"></div>
                    </div>
                </body>
            </html>
            '''
            ,

            emails_for_send=['agent-smit777@yandex.ru', 'fling-shrn@mail.ru'],
            # emails_for_send=['blablabla'],

            text='Тест',
            bcc_emails=[],
            theme_param='Тестовое письма c разметкой HTML.',
            dict_included_files={
                'file_1.pdf': os.path.join(os.getcwd(), 'files_for_include', 'test_file.pdf'),
                'file_2.pdf': os.path.join(os.getcwd(), 'files_for_include', 'test_file_2.pdf')
            }
        )
    except(NotEmail,) as er:
        print(er)
