import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pathlib

from save_to_excel import get_and_save_to_file
from functions import make_word
 
def main():
    fromaddr = "sashanovizov@mail.ru"
    toaddr = "sashanovizov@mail.ru"
    
    with open('./pass.txt', 'r', encoding='utf-8') as file:
        password = file.readline().strip()

    path = pathlib.Path('C:\\Users\\Sasha\\Учеба\\Программирование и прочее\\python\\ТЗ\\result.xlsx')

    rows_count = get_and_save_to_file(path)

    body = f'Программа обработала {rows_count} {make_word(rows_count)}'

    send_mail(
        fromaddr=fromaddr, 
        toaddr=toaddr,
        mypass=password,
        subj='Результат работы скрипта',
        body=body,
        file=path
    )

def send_mail(fromaddr, toaddr, mypass, subj, body, file: pathlib.Path):
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = subj
    
    msg.attach(MIMEText(body, 'plain'))

    with open(file, 'rb') as f:
        part = MIMEApplication(
                    f.read(),
                    Name=file.name
                )
        part['Content-Disposition'] = 'attachment; filename="%s"' % file.name
        msg.attach(part)
    
    server = smtplib.SMTP_SSL('smtp.mail.ru', 465)
    server.login(fromaddr, mypass)
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()
    
if __name__ == '__main__':
    main()