"""
Скрипт в автоматическом режиме 1 раз в месяц отправляет PDF файл на почту.
"""
import mimetypes
import os
import smtplib
import ssl
from dataclasses import dataclass
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, make_msgid

from dotenv import load_dotenv  # type: ignore

load_dotenv()


@dataclass
class UserMail:
    # Email от кого отправлять.
    from_mail: str
    # Пароль отправителя.
    password: str
    # Отображение имени отправителя рядом с почтой.
    sender_name: str
    # Email кому отправлять.
    to_mail: str
    # Отображение имени почты кому приходит письмо.
    recipient_name: str
    # Хост для исходящий сообщений.
    host_smtp: str
    # Порт для исходящих сообщений.
    port: int
    # Тема письма.
    subject: str


def get_mime_application(path: str) -> MIMEApplication:
    ftype, encoding = mimetypes.guess_type(path)
    if ftype is None:
        subtype = "octet-stream"
    else:
        file_type, subtype = ftype.split("/")

    with open(f"{path}", "rb") as f:
        att = MIMEApplication(f.read(), subtype)

    att.add_header('Content-Disposition', 'attachment', filename=os.path.basename(path))
    return att


def send_email(path: str, user_mail: UserMail) -> str:
    """
    Отправляет файлы на почту.
    """
    context = ssl.create_default_context()

    try:
        with smtplib.SMTP_SSL(user_mail.host_smtp, user_mail.port, context=context) as s:
            s.login(user_mail.from_mail, user_mail.password)
            msg = MIMEMultipart()
            msg['From'] = formataddr((user_mail.sender_name, user_mail.from_mail))
            msg['To'] = formataddr((user_mail.recipient_name, user_mail.to_mail))
            msg['Subject'] = user_mail.subject
            msg['Message-ID'] = make_msgid()

            att = get_mime_application(path)
            msg.attach(att)

            s.sendmail(user_mail.from_mail, user_mail.to_mail, msg.as_string())

        return 'Сообщение отправлено успешно!'
    except smtplib.SMTPAuthenticationError as e:
        print(f"Ошибка авторизации: {e}")
        return "Ошибка: Неверный логин или пароль. Используйте App Password."
    except smtplib.SMTPRecipientsRefused as e:
        print(f"Получатель отклонён: {e}")
        return "Ошибка: Email получателя отклонён сервером."
    except smtplib.SMTPException as e:
        print(f"SMTP ошибка: {e}")
        return f"SMTP ошибка: {e}"
    except Exception as e:
        print(f"Неизвестная ошибка: {e}")
        return f"Ошибка: {e}"


def main() -> None:
    user = UserMail(
        from_mail=os.getenv("FROM_MAIL"),
        password=os.getenv("PASSWORD"),
        sender_name=os.getenv("SENDER_NAME"),
        to_mail=os.getenv("TO_MAIL"),
        recipient_name=os.getenv("RECIPIENT_NAME"),
        host_smtp="smtp.mail.ru",
        port=465,
        subject="ответ по СПЭС 3246 (САО)",
    )

    # Путь где лежит файл
    pdf_path: str = os.getenv("PDF_PATH")
    if not os.path.isfile(pdf_path):
        print(f"Файл не найден: {pdf_path}")
        return
    print(send_email(pdf_path, user))


if __name__ == '__main__':
    main()
