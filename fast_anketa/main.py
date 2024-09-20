from selenium import webdriver
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

from collections import deque
import openpyxl
import time
from pprint import pprint
import traceback

from regions import regions


# Данные

login = '2F21R'
password = '3U41320'

# Открываем сайт
driver = webdriver.Chrome(executable_path='C:\\Users\\tkachikvv.FGUZ\\Documents\MEGA\IT\\fast_anketa\\chromedriver.exe')
driver.get('https://anket.demography.site/login')
element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, 'login-button'))
    )

# Логин
log_input = driver.find_element(By.ID, 'loginform-login')
pass_input = driver.find_element(By.ID, 'loginform-password')
confirm = driver.find_element(By.NAME, 'login-button')
log_input.send_keys(f'{login}')
pass_input.send_keys(f'{password}')
time.sleep(1)
confirm.click()


# Функция ввода данных
def data_enter(data):
    print('HERE3')
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'text-center'))
    )
    driver.get('https://anket.demography.site/deti-anket/create')
    print(f'DATA0 {data}')

    all_fields = driver.find_elements(By.CLASS_NAME, 'form-control')
    field_num = 0
    excel_num = 0
    switch_29 = False
    switch_35 = False
    skip = 0

    for el in all_fields:
        print(el.tag_name)
        print(el.get_attribute('id'))
        print(field_num)
        print(excel_num)
        print(f'{data[excel_num]}')
        print(f'SKIP {skip}')
        # if skip > 0:
        #     skip = skip - 1
        #     continue
        # else:
        if field_num < 2:
            field_num += 1
            excel_num += 1
        elif field_num == 2:
            munic_text = regions[int(data[2])]
            Select(el).select_by_visible_text(munic_text)
            field_num += 1
            excel_num += 1
        elif field_num == 4:
            time.sleep(.5)
            Select(el).select_by_value('810')
            '''Не работает'''
            field_num += 1
            excel_num += 1
            continue
        elif field_num in [18, 20, 21, 22, 23, 24, 25]:
            if not data[excel_num]:
                el.send_keys(f'1')
                field_num += 1
                excel_num += 1
            elif data[excel_num] == 97:
                el.send_keys(f'0')
                field_num += 1
                excel_num += 1
            else:
                el.send_keys(f'{data[excel_num]}')
                field_num += 1
                excel_num += 1
        elif field_num == 67 and data[excel_num] != 1:
            if not data[excel_num]:
                Select(el).select_by_value(f'98')
            else:
                Select(el).select_by_value(f'{data[excel_num]}')
            switch_29 = True
            switch_35 = True
            field_num += 1
            excel_num += 1
            continue
        elif field_num == 43 and not data[excel_num]:
            field_num += 1
            excel_num += 1
            continue
        elif field_num == 57 and data[excel_num] not in [1, 2]:
            print(data[excel_num])
            print(type(data[excel_num]))
            if not data[excel_num]:
                Select(el).select_by_value(f'98')
            else:
                Select(el).select_by_value(f'{data[excel_num]}')
            switch_35 = True
            skip = 31
            field_num += 33
            excel_num += 33
            continue
        elif field_num == 73 and switch_29:
            print(data[excel_num])
            print(type(data[excel_num]))
            field_num += 1
            excel_num += 1
            switch_29 = False
            continue
        elif field_num == 89 and not switch_35:
            switch_35 = True
            skip = 6
            field_num += 7
            excel_num += 7
            continue
        elif field_num == 99 and data[excel_num] not in [1, 2]:
            print('SWITCH 37')
            if not data[excel_num]:
                Select(el).select_by_value(f'98')
            else:
                Select(el).select_by_value(f'{data[excel_num]}')
            excel_num += 28
            field_num += 28
            skip = 27
        elif field_num == 128 and data[excel_num] != 1:
            if not data[excel_num]:
                Select(el).select_by_value(f'98')
            else:
                Select(el).select_by_value(f'{data[excel_num]}')
            skip = 1
            field_num += 2
            excel_num += 2
            print(f'SWITCH42')
        elif field_num == 156 and data[excel_num] == 0:
            el.send_keys(f'0')
            field_num += 1
            excel_num += 1
            continue
        elif field_num == 169:
            if data[excel_num]:
                Select(el).select_by_value(f'1')
                el2 = driver.find_element(By.ID, 'detiankettable4548-field46_13')
                el2.send_keys(f'{data[excel_num]}')
            else:
                Select(el).select_by_value(f'98')
            field_num += 1
            excel_num += 1
            skip = 1

        elif field_num == 172:
            if data[excel_num] != 1:
                Select(el).select_by_value(f'нет')
            else:
                Select(el).select_by_value(f'есть')
            field_num += 1
            excel_num += 1

        elif field_num == 174:
            el.send_keys(f'{data[excel_num]}')
        else:
            try:
                if el.tag_name == 'select':
                    if not data[excel_num]:
                        if field_num in range(100, 118):
                            Select(el).select_by_value(f'2')
                        else:
                            Select(el).select_by_value(f'98')
                    else:
                        if data[excel_num] == 22:
                            Select(el).select_by_value(f'2')
                        else:
                            Select(el).select_by_value(f'{data[excel_num]}')
                    print(f'EXCEL {excel_num}')
                    print(f'FIELD {field_num}')
                elif el.tag_name == 'input':
                    print(el.get_property('type'))
                    if el.get_property('type') == 'date':
                        el.send_keys(f'{data[excel_num].strftime("%d:%m:%Y")}')
                        print(data[excel_num].strftime("%d:%m:%Y"))
                    else:
                        print(f'EXCEL {excel_num}')
                        print(f'FIELD {field_num}')
                        el.send_keys(f'{data[excel_num]}')
                else:
                    if data[excel_num]:
                        el.send_keys(f'{data[excel_num]}')
                field_num += 1
                excel_num += 1
            except ElementNotInteractableException:
                print(f'EXCEPTION {switch_29, switch_35}')
                if field_num == 96:
                    continue
                else:
                    field_num += 1
                    excel_num += 1
                    continue


# Обходим данные для ввода
sheet = openpyxl.load_workbook('./Анкеты.xlsx', data_only=True).active
row_num = 0
for row in sheet.values:
    print(row)
    if row[0]:
        if row_num > 0:
            time.sleep(1)
            data_enter(deque(row))
            row_num += 1
        else:
            row_num += 1
            continue
        print(row_num)
    else:
        print('BREAK')
        break

    submit_btn = driver.find_element(By.CLASS_NAME, 'btn-outline-primary')
    driver.execute_script("arguments[0].click();", submit_btn)
    time.sleep(2)

driver.quit()

