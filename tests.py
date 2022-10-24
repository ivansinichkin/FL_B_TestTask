import unittest
from openpyxl import load_workbook, Workbook
from read_url import parse_url
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from read_data import copy_template, search_start_point, search_data_area2, add_data, write_to_output


class TestParser(unittest.TestCase):
    def test_vendor_code(self):
        url = 'https://www.mechta.kz/product/' \
              'noutbuk-lenovo-legion-5-15imh05-82au00c2rk-156-fhdcore-i5-10300h-25-ghz8ssd512gtx1650ti4dos/'
        vendor_code = parse_url(url)[0]
        self.assertEqual(vendor_code, '03762')

    def test_price(self):
        url = 'https://www.mechta.kz/product/' \
              'noutbuk-lenovo-legion-5-15imh05-82au00c2rk-156-fhdcore-i5-10300h-25-ghz8ssd512gtx1650ti4dos/'
        price = parse_url(url)[1]
        self.assertEqual(price, '514990')

    #  проверяем, что XPATH возвращает ожидаемое значение и соответсвенно актуален
    def test_check_xpath(self):
        url = 'https://www.mechta.kz/product/' \
              'noutbuk-lenovo-legion-5-15imh05-82au00c2rk-156-fhdcore-i5-10300h-25-ghz8ssd512gtx1650ti4dos/'
        driver = webdriver.Chrome()
        driver.get(url=url)
        astana_button = WebDriverWait(driver, 10). \
            until(EC.element_to_be_clickable((By.XPATH, '//div[@class="text-no-wrap list"]/div/div')))
        astana_button.click()
        vendor_code = driver.find_element(By.XPATH, '//div[@class="flex items-center"]/div')
        price = driver.find_element(By.XPATH, '//div[@class="text-bold text-ts5 text-color1"]')
        self.assertEqual(vendor_code.text, 'Код: 03762', 'XPATH до артикула не верен либо изменился арктикул')
        self.assertEqual(price.text, '514 990 ₸', 'XPATH до цены не верен либо изменилсась цена')

    def test_search_table_header(self):
        input_file = 'TestTask_1_input_1.xlsx'
        wb_inp = load_workbook(input_file)
        sheet = wb_inp.active
        table_starts = search_start_point(sheet)
        row = table_starts[0][0]
        col = table_starts[0][1]
        self.assertEqual(sheet[row][col].value, 'Данные по товарам',
                         'Работа функции search_start_point(sht) нарушена')


if __name__ == "__main__":
    unittest.main()
