"""
Скрипт в автоматическом режиме 1 раз в месяц отправляет PDF файл на почту.
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
from pathlib import Path

import comtypes.client
import jinja2
import pdfkit
from docxtpl import DocxTemplate
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


def set_templates(date_now: str, path_template: str,
                  path_save: str) -> None:
    path_true = Path(path_template).exists()
    if not path_true:
        raise FileNotFoundError('[!] Отсутствует шаблон.')

    doc = DocxTemplate(path_template)

    context = {"from_data": date_now}

    doc.render(context)
    doc.save(path_save)


def get_image_file_as_base64_data(file_path:str) -> str:
    with open(file_path, 'rb') as image_file:
        return base64.b64encode(image_file.read()).decode()


def set_templates_html(date_now: str, output_path: str):
    context = {"from_data": date_now,
               "img_string": get_image_file_as_base64_data("Pattern/image.png")
               }
    template_loader = jinja2.FileSystemLoader("Pattern")
    template_env = jinja2.Environment(loader=template_loader)

    template = template_env.get_template("report.html")
    output_text = template.render(context)

    config = pdfkit.configuration(wkhtmltopdf="wkhtmltox/bin/wkhtmltopdf.exe")
    pdfkit.from_string(output_text, output_path, configuration=config,
                       css="Pattern/my_style.css",
                       options={"enable-local-file-access": True})


def convert(input_path: str, out_path: str) -> None:
    wdFormatPDF = 17

    input_path = os.path.abspath(input_path)
    out_path = os.path.abspath(out_path)
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(input_path)
    doc.SaveAs(out_path, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


def main() -> None:
    dir_pattern = "Pattern"
    send_name = "Send"
    if not os.path.isdir(send_name):
        os.mkdir(send_name)

    date_now = f"{datetime.now().strftime('%d.%m.%Y')}г."
    path_template = f"{dir_pattern}/report.docx"
    # input_path = f"{dir_pattern}/tmp.docx"
    out_path = f"{send_name}/cao.pdf"
    # set_templates(date_now, path_template, input_path)
    set_templates_html(date_now, out_path)

    # Конвертируем в pdf.
    # convert(input_path, out_path)
    # Удаляем tmp файл.
    # os.remove(input_path)
    # Отправляем на почту.
    print(send_email(send_name))


if __name__ == "__main__":
    main()
