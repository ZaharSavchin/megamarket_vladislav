import time
from selenium import webdriver
from selenium.webdriver.common.by import By

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromiumService
from openpyxl import load_workbook


workbook = load_workbook(filename='megamarket.xlsx')
sheet = workbook.active
links = []
for cell in sheet['A']:
    links.append(cell.value)


price = []
cashback = []
delivery = []
pay_format = []
shop = []


with webdriver.Chrome(service=ChromiumService(ChromeDriverManager().install())) as browser:
    for link in links:
        try:
            browser.get(link)
            time.sleep(5)

            try:
                price_ = browser.find_element(By.CLASS_NAME, 'sales-block-offer-price__price-final').text
                price_ = ''.join(price_.split(' ')[:-1])
                price.append(int(price_))
            except Exception as err:
                price_ = None
                price.append(None)

            try:
                cashback_ = browser.find_element(By.CLASS_NAME, "bonus-amount").text
                cashback_ = int(''.join(cashback_.split(' ')))

                try:
                    cashback_s = browser.find_element(By.CSS_SELECTOR, 'span.bonus-amount.bonus-amount_without-percent').text
                    cashback_s = int(''.join(cashback_s.split(' '))[1:])
                    cashback_ += cashback_s
                except Exception as err:
                    print(err)

                cashback.append(cashback_)

            except Exception as err:
                print(err)
                cashback_ = None
                cashback.append(None)

            try:
                delivery_ = browser.find_element(By.CLASS_NAME, "sales-block-delivery-type__date").text
                delivery.append(delivery_)
            except Exception as err:
                delivery_ = None
                delivery.append(None)

            try:
                pay_format_ = browser.find_element(By.CLASS_NAME, "pdp-available-payment-method-block__text").text
                pay_format.append(pay_format_)
            except Exception as err:
                pay_format_ = None
                pay_format.append(None)

            try:
                shop_ = browser.find_element(By.CLASS_NAME, "pdp-merchant-rating-block__merchant-name").text
                shop.append(shop_)
            except Exception as err:
                shop_ = None
                shop.append(None)

            print(f'{link[:10]} = {price_}р, {cashback_}, {delivery_}, {pay_format_}, {shop_}')

        except Exception as err:
            price.append(None)
            cashback.append(None)
            delivery.append(None)
            pay_format.append(None)
            shop.append(None)
            print(f'{link} = товар не найден')

for i in range(len(price)):
    sheet.cell(row=i+1, column=2, value=price[i])

for i in range(len(cashback)):
    sheet.cell(row=i+1, column=3, value=cashback[i])

for i in range(len(delivery)):
    sheet.cell(row=i+1, column=4, value=delivery[i])

for i in range(len(pay_format)):
    sheet.cell(row=i+1, column=5, value=pay_format[i])

for i in range(len(shop)):
    sheet.cell(row=i+1, column=6, value=shop[i])

workbook.save(filename='megamarket.xlsx')
