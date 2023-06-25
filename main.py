"""
Скрипт в автоматическом режиме 1.txt раз в месяц отправляет PDF файл на почту.
"""
import mimetypes
import os
import smtplib
import ssl
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, make_msgid

from dotenv import load_dotenv
from fpdf import FPDF

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


def converts_html_in_pdf(dir_pattern: str, out_path: str, date_now: str):
    """
    Конвертирует html текс  и добавляет img в pdf файл.
    :param out_path: Путь куда положить pdf файл.
    :param dir_pattern: Директория от куда брать img.
    """
    # Размер страницы.
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_margins(30, 10, 20)
    # Шрифт.
    pdf.add_font("DejaVu", "", "font/timesnewromanpsmt.ttf")
    pdf.set_font("DejaVu", size=14)

    pdf.image(f"{dir_pattern}/image.png", w=20, h=30, x=115, y=200)
    # Шаблон html файла.
    pdf.write_html(f"""
    <table>
        <tr>
            <td width="60%"></td>
            <td width="40%">
Заместителю начальника
ФГКУ &laquo;УВО ВНГ России
по городу Москве&raquo;
полковнику полиции
Р.А. Морозову
            </td>
        </tr>
    </table>
    <p> </p>
    <center>РАПОРТ</center>
    <table>
    <tr>
    <td align="justify">
    Во исполнение сопроводительного письма электронного сообщения ФГКУ &laquo;УВО ВНГ России по городу Москве&raquo; № 99 от 18.01.2023г., в целях реализации положения Указа Президента РФ от 21.12.17г. № 618 &laquo;Об основных направлениях государственной политики по развитию конкуренции&raquo;, приказа Росгвардии от 07.02.20г. № 24 &laquo;Об утверждении Положения об организации в ВНГ РФ системы внутреннего обеспечения соответствия требованиям антимонопольного законодательства РФ&raquo;, приказа командующего Центральным Оршанско-Хинганским Краснознаменным округом ВНГ РФ от 16.03.20г. № 137 &laquo;Об организации обеспечения соответствия требований антимонопольного законодательства РФ в Центральном округе ВНГ РФ&raquo;, Федерального закона от 05.04.13г. № 44-ФЗ &laquo;О контрактной системе в сфере закупок товаров, работ, услуг для обеспечения государственных и муниципальных нужд&raquo; и организации функционирования антимонопольной системы в ФГКУ &laquo;УВО ВНГ России по городу Москве&raquo; и его структурных подразделениях в МОВО по САО ФГКУ &laquo;УВО ВНГ России по городу Москве&raquo; не имеется нарушений антимонопольного законодательства при осуществлении закупок товаров, работ, услуг в соответствии с законодательством РФ о контрактной системе в сфере закупок для обеспечения государственных и муниципальных нужд.
    </td>
    </tr>
    </table>
    <table>
        <tr>
            <td>
Заместитель начальника МОВО по САО
ФГКУ &laquo;УВО ВНГ России по городу Москве&raquo;
            </td>
        </tr>
        <tr>
            <td>подполковник полиции</td>
            <td align="right">А.В. Худяков</td>
        </tr>
        <tr><td>{date_now}</td></tr>
    </table>
    """)

    pdf.output(out_path)


def main() -> None:
    dir_pattern = "Pattern"
    # Яндекс cloud разрешает только отправлять с директории tmp.
    send_name = "tmp"
    if not os.path.isdir(send_name):
        os.mkdir(send_name)

    # test.
    out_path = f"{send_name}/cao.pdf"
    # yandex cloud.
    # out_path = f"/{send_name}/cao.pdf"
    date_now = f"{datetime.now().strftime('%d.%m.%Y')}г."
    # Конвертируем в pdf.
    converts_html_in_pdf(dir_pattern, out_path, date_now)
    # Отправляем на почту.
    # print(send_email(send_name))
    # os.remove(out_path)

if __name__ == "__main__":
    main()
