import selenium
import selenium.webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service

from selenium.webdriver.common.by import By
import time
from datetime import date


def main():

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
    from functions import calculate_date
    date_start, date_end = calculate_date()
    driver.set_window_size(850, 400)
    jpy_data = request_data(driver=driver, date_start=date_start, date_end=date_end, curr='JPY')
    usd_data = request_data(driver=driver, date_start=date_start, date_end=date_end, curr='USD')
    print(jpy_data[0])
    handle_exit(driver=driver)

def go_to_page(driver: selenium.webdriver.Chrome, url: str):
    try:
        driver.get(url=url)
        driver.maximize_window()
    except Exception as ex:
        print(f"[ОШИБКА] Ошибка доступа к url: {ex}")
        handle_exit(driver=driver)

def search_page(driver: selenium.webdriver.Chrome):

    try:

        header = driver.find_elements(By.TAG_NAME, 'nav')[1]

        hover = header.find_element(By.LINK_TEXT, 'Срочный рынок')

        ActionChains(driver).move_to_element(hover).perform()
        nav_item =  driver.find_element(By.LINK_TEXT, 'Индикативные курсы')
        nav_item.click()
    except Exception as ex:
        print(f"[ОШИБКА] Не найден элемент на странице: {ex}")
        handle_exit(driver=driver)

def request_data(driver: selenium.webdriver.Chrome, date_start, date_end, curr: str):
    try:
        
        driver.set_window_size(850, 400)

        set_currency(driver=driver, curr=curr)

        set_dates(driver=driver, date_start=date_start, date_end=date_end)

        driver.execute_script('window.scrollTo(0, 0)')
        
        return collect_data(driver=driver)
    except Exception as ex:
        print(f"[ОШИБКА] {ex}")
        handle_exit(driver=driver)

def check_agreement(driver: selenium.webdriver.Chrome):
    try:
        button = driver.find_element(By.LINK_TEXT, "Согласен")
        if button != None: button.click()
    except Exception:
        pass

def check_cookies(driver: selenium.webdriver.Chrome):
    try:
        cookies = driver.find_element(By.CLASS_NAME, 'cookies-container-inner')
        cookies_btn = cookies.find_element(By.TAG_NAME, 'button')
        if cookies_btn != None: cookies_btn.click()
    except Exception:
        pass


def set_currency(driver: selenium.webdriver.Chrome, curr: str):
    form = driver.find_element(By.TAG_NAME, 'form')
    select = form.find_element(By.CLASS_NAME, 'ui-select__activator')
    actions = ActionChains(driver).move_to_element(form).scroll_to_element(select).click()
    select.click()
    dropdown = driver.find_element(By.CLASS_NAME, '-opened')

    opts = dropdown.find_elements(By.CLASS_NAME, 'ui-dropdown-option')
    if curr == 'USD':
        element = dropdown.find_element(By.LINK_TEXT, 'USD/RUB - Доллар США к российскому рублю')
        actions = ActionChains(driver).move_to_element(select).scroll_to_element(element).click()
    elif curr == 'JPY':
        element = dropdown.find_element(By.LINK_TEXT, 'JPY/RUB - Японская йена к российскому рублю')
        actions = ActionChains(driver).move_to_element(select).scroll_to_element(element).click()

    actions.perform()

    time.sleep(3)
    
def set_dates(driver: selenium.webdriver.Chrome, date_start, date_end):
    form = driver.find_element(By.TAG_NAME, 'form')
    date_start_input, date_end_input = form.find_elements(By.ID, 'keysParams')

    set_date(driver, date_start_input, date_start, True)
    set_date(driver, date_end_input, date_end, False)

    submit = form.find_element(By.CLASS_NAME, 'ui-button')
    submit.click()


def set_date(driver: selenium.webdriver.Chrome, input, curr_date, start: bool):
    input.click()

    calendar = driver.find_element(By.CLASS_NAME, '-opened')

    selects = calendar.find_elements(By.CLASS_NAME, 'ui-select')
    month_select, year_select = selects

    month_select.click()

    dropdown = driver.find_element(By.CLASS_NAME, '-opened')
    opts = dropdown.find_elements(By.CLASS_NAME, 'ui-dropdown-option')


    actions = ActionChains(driver).move_to_element(dropdown).scroll_to_element(opts[curr_date.month - int(start)]).click()
    actions.perform()

    input.click()    

    if date.today().month == 1:
        year_select.click()
        dropdown = driver.find_element(By.CLASS_NAME, '-opened')
        opts = dropdown.find_elements(By.CLASS_NAME, 'ui-dropdown-option')

        actions = ActionChains(driver).move_to_element(dropdown).scroll_to_element(opts[1]).click()
        actions.perform()
        time.sleep(5)

    days = calendar.find_elements(By.CLASS_NAME, '-day')
    if start: days[0].click()
    else: days[-1].click()

def collect_data(driver: selenium.webdriver.Chrome) -> list[dict]:
    table = driver.find_element(By.CLASS_NAME, 'ui-table__container')
    
    rows = table.find_elements(By.CLASS_NAME, 'ui-table-row')

    result = []

    for row in rows:
        cells = row.find_elements(By.TAG_NAME, 'td')
        result.append(
            {
                "rate": float(cells[-2].text),
                "time": cells[-1].text,
                "date": cells[0].text
            }
        )

    return result


def handle_exit(driver: selenium.webdriver.Chrome):
    driver.close()
    driver.quit()
    

if __name__ == '__main__':
    main()