from api_solution import request_data 
from pathlib import Path
from save_to_excel import save_to_file
from send_mail import send_mail
from functions import calculate_date, make_word

from site_solution import *

def main():
    "Вызовите функцию api() для решения через API или site() для решения через браузер. Необходимо указать пароль для доступа к почте"
    with open('./pass.txt', 'r', encoding='utf-8') as file:
        password = file.readline().strip()

    site(fromaddr='sashanovizov@mail.ru', toaddr='sashanovizov@mail.ru', password=password)
    # api(fromaddr='sashanovizov@mail.ru', toaddr='sashanovizov@mail.ru', password=password)

def api(fromaddr: str, toaddr: str, password: str):
    date_start, date_end = calculate_date()

    usd_data = request_data(date_start.isoformat(), date_end.isoformat(), 'USD')
    jpy_data = request_data(date_start.isoformat(), date_end.isoformat(), 'JPY')

    path = Path('./result.xlsx')

    rows_count = save_to_file(path=path, usd_data=usd_data, jpy_data=jpy_data)

    body = f'Программа обработала {rows_count} {make_word(rows_count)}'

    send_mail(
        fromaddr=fromaddr,
        toaddr=toaddr,
        mypass=password,
        subj='Результат работы скрипта',
        body=body,
        file=path
    )

def site(fromaddr: str, toaddr: str, password: str):
    s = Service(executable_path='.//yandexdriver.exe')
    options = selenium.webdriver.ChromeOptions()
    driver = selenium.webdriver.Chrome(service=s, options=options)

    url = 'https://www.moex.com'
    DELAY = 10

    driver.implicitly_wait(DELAY)

    go_to_page(driver, url=url)
    check_cookies(driver=driver)
    search_page(driver=driver)
    check_agreement(driver)

    
    date_start, date_end = calculate_date()
    usd_data = request_data(driver=driver, date_start=date_start, date_end=date_end, curr='USD')
    jpy_data = request_data(driver=driver, date_start=date_start, date_end=date_end, curr='JPY')
    handle_exit(driver=driver)

    path = Path('./result.xlsx')

    rows_count = save_to_file(path=path, usd_data=usd_data, jpy_data=jpy_data)

    body = f'Программа обработала {rows_count} {make_word(rows_count)}'

    send_mail(
        fromaddr=fromaddr,
        toaddr=toaddr,
        mypass=password,
        subj='Результат работы скрипта',
        body=body,
        file=path
    )

if __name__ == '__main__':
    main()