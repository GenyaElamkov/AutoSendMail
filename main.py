"""
Скрипт в автоматическом режиме 1 раз в месяц отправляет PDF файл на почту.
"""

import mimetypes
import os
import smtplib
import ssl
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, make_msgid

from dotenv import load_dotenv

load_dotenv()

# Email от кого отправлять.
from_mail = os.getenv("FROM_MAIL")
# Пароль отправителя.
password = os.getenv("PASSWORD")
# Отображение имени отправителя рядом с почтой.
sender_name = os.getenv("SENDER_NAME")

# Email кому отправлять.
to_mail = os.getenv("TO_MAIL")
# Отображение имени почты кому приходит письмо.
recipient_name = os.getenv("RECIPIENT_NAME")

# Хост для исходящий сообщений.
host_smtp = "smtp.mail.ru"
# Порт для исходящих сообщений.
port = 465

# Тема письма.
subject = "Отчет по монополию СПЭС ???"


def send_email(dir_name: str) -> str:
    context = ssl.create_default_context()

    try:
        with smtplib.SMTP_SSL(host_smtp, port, context=context) as s:
            s.login(from_mail, password)
            msg = MIMEMultipart()
            msg['From'] = formataddr((sender_name, from_mail))
            msg['To'] = formataddr((recipient_name, to_mail))
            msg['Subject'] = subject
            msg['Message-ID'] = make_msgid()

            print("Collecting...")
            files = os.listdir(dir_name)
            for file in files:
                ftype, encoding = mimetypes.guess_type(file)
                file_type, subtype = ftype.split("/")

                with open(f"Pattern/{file}", "rb") as f:
                    att = MIMEApplication(f.read(), subtype)

                att.add_header('Content-Disposition', 'attachment',
                               filename=file)
                msg.attach(att)
            print("Sending...")
            s.sendmail(from_mail, to_mail, msg.as_string())

        return 'Сообщение отправлено успешно!'
    except Exception as _ex:
        return f"{_ex}\nПожалуйста, проверьте свой логин или пароль!"


def main() -> None:
    dir_name = "Pattern"
    if not os.path.isdir(dir_name):
        os.mkdir(dir_name)

    print(send_email(dir_name))


if __name__ == "__main__":
    main()
