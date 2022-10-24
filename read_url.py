from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, NoSuchElementException


def parse_url(url):

    try:
        driver = webdriver.Chrome()
        driver.get(url=url)
        print(driver.title)
        # находим во всплывающем окне кнопку Астана и кликаем ее
        astana_button = WebDriverWait(driver, 10).\
            until(EC.element_to_be_clickable((By.XPATH, '//div[@class="text-no-wrap list"]/div/div')))
        astana_button.click()
        # проверяем наличие текста "Код" в элементе содержащем артикул и только после этого добываем оттуда артикул
        if WebDriverWait(driver, 10).\
                until(EC.text_to_be_present_in_element((By.XPATH, '//div[@class="flex items-center"]/div'), 'Код')):
            vendor_code = driver.find_element(By.XPATH, '//div[@class="flex items-center"]/div')
        vendor_code_value = vendor_code.text.split(' ')[1]
        # также проверяем наличие в элементе текста с символом '₸' и после этого добываем цену
        if WebDriverWait(driver, 10). \
                until(EC.text_to_be_present_in_element((By.XPATH, '//div[@class="text-bold text-ts5 text-color1"]'), '₸')):
            price = driver.find_element(By.XPATH,  '//div[@class="text-bold text-ts5 text-color1"]')
        price_value = ''.join(price.text.split(' ')[0:2])

        print(vendor_code_value, price_value)
        return vendor_code_value, price_value

    except NoSuchElementException as NEex:
        print(NEex)
        print('Element no found, check XPath')
    except WebDriverException as WDex:
        print(WDex)
        print("Make sure the webdriver is in the system PATH or read the error above")
    finally:
        driver.close()
