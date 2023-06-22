"""
Скрипт в автоматическом режиме 1.txt раз в месяц отправляет PDF файл на почту.
"""
import base64
import mimetypes
import os
import smtplib
import ssl
from datetime import datetime
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
subject = "Ответ СПЭС 99 (САО)"


def send_email(dir_name: str) -> str:
    """
    Отправляет файлы на почту.
    :param dir_name: имя папки где лежат файлы.
    :return: Сообшеение об успещно отправлении или Error.
    """
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

                with open(f"{dir_name}/{file}", "rb") as f:
                    att = MIMEApplication(f.read(), subtype)

                att.add_header('Content-Disposition', 'attachment',
                               filename=file)
                msg.attach(att)
            print("Sending...")
            s.sendmail(from_mail, to_mail, msg.as_string())

        return 'Сообщение отправлено успешно!'
    except Exception as _ex:
        return f"{_ex}\nПожалуйста, проверьте свой логин или пароль!"


def get_image_file_as_base64_data(file_path: str) -> str:
    """
    Декодирование img.
    """
    with open(file_path, 'rb') as image_file:
        return base64.b64encode(image_file.read()).decode()


# res = get_image_file_as_base64_data("Pattern/image.png")


def main() -> None:
    dir_pattern = "Pattern"
    send_name = "Send"
    if not os.path.isdir(send_name):
        os.mkdir(send_name)

    out_path = f"{send_name}/cao.pdf"

    # Конвертируем в pdf.
    # Отправляем на почту.
    # print(send_email(send_name))


if __name__ == "__main__":
    main()
