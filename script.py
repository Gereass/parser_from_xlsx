import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Загрузка данных из .xlsx файла
def load_data_from_xlsx(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(row)  
        

    return data

# Заполнение веб-формы
def fill_web_form(name):
    driver = webdriver.Chrome()

    # Откроем веб-сайт
    driver.get('https://rpachallenge.com/?lang=EN')

    # Нажатие старт на сайте
    submit_button = driver.find_element(By.CLASS_NAME, 'waves-effect.col.s12.m12.l12.btn-large.uiColorButton')
    submit_button.click()

    for i in range(10):

        for item in data:
            # поиск полей и заполнение их
            first_name_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelFirstName"]')
            last_name_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelLastName"]')
            address_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelAddress"]')
            email_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelEmail"]')
            phone_number_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelPhone"]')
            company_name_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelCompanyName"]')
            role_in_company_field = driver.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelRole"]')
            
            first_name_field.send_keys(item[0])
            last_name_field.send_keys(item[1])
            address_field.send_keys(item[2])
            email_field.send_keys(item[3])
            phone_number_field.send_keys(item[4])
            company_name_field.send_keys(item[5])
            role_in_company_field.send_keys(item[6])

            time.sleep(2) 
            # поиск и нажатие кнопки отправки
            submit_button = driver.find_element(By.CLASS_NAME, 'btn.uiColorButton')
            submit_button.click()

            time.sleep(2)  # Ждем немного, чтобы увидеть результат (при необходимости)

    time.sleep(15) 
    driver.quit()

if __name__ == "__main__":
    # Укажите путь к вашему .xlsx файлу
    file_path = '/home/gereas/work/avto/task.xlsx'
    data = load_data_from_xlsx(file_path)
    fill_web_form(data)

