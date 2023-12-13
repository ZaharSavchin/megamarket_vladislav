import time
from selenium import webdriver
from selenium.webdriver.common.by import By

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromiumService
from openpyxl import load_workbook

from other_shops import get_other_shops


workbook = load_workbook(filename='megamarket.xlsx')
sheet = workbook.active
links = []
for cell in sheet['A']:
    links.append(cell.value)


items_links = []
price = []
cashback = []
cashback_other = []
cashback_sber = []
delivery = []
pay_format = []
shop = []


with webdriver.Chrome(service=ChromiumService(ChromeDriverManager().install())) as browser:

    for link in links:
        try:
            browser.get(link)
            time.sleep(3)

            items_links.append(link)

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
                cashback.append(cashback_)
            except Exception as err:
                cashback_ = None
                cashback.append(None)

            try:
                cashback_s = browser.find_element(By.CSS_SELECTOR,
                                                  'span.bonus-amount.bonus-amount_without-percent').text
                cashback_s = int(''.join(cashback_s.split(' '))[1:])
                cashback_sber.append(cashback_s)
            except Exception as err:
                cashback_s = None
                cashback_sber.append(None)

            try:
                cashback_o = browser.find_element(By.CSS_SELECTOR,
                                                  'div.pdp-cashback-table__money-bonus.money-bonus.xs.money-bonus_loyalty.money-bonus_grey'
                                                  ).find_element(By.CLASS_NAME, 'bonus-amount').text
                cashback_o = int(''.join(cashback_o.split(' ')))
                cashback_other.append(cashback_o)
            except Exception as err:
                cashback_o = None
                cashback_other.append(None)

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

            try:
                get_other_shops(browser,
                                items_links,
                                price,
                                cashback,
                                cashback_other,
                                cashback_sber,
                                delivery,
                                pay_format,
                                shop
                                )
            except Exception as err:
                print(err)

        except Exception as err:
            items_links.append(None)
            price.append(None)
            cashback.append(None)
            cashback_other.append(None)
            cashback_sber.append(None)
            delivery.append(None)
            pay_format.append(None)
            shop.append(None)
            print(f'{link} = товар не найден')

for i in range(len(items_links)):
    sheet.cell(row=i+1, column=2, value=items_links[i])

for i in range(len(price)):
    sheet.cell(row=i+1, column=3, value=price[i])

for i in range(len(cashback)):
    sheet.cell(row=i+1, column=4, value=cashback[i])

for i in range(len(cashback_other)):
    sheet.cell(row=i + 1, column=5, value=cashback_other[i])

for i in range(len(cashback_sber)):
    sheet.cell(row=i+1, column=6, value=cashback_sber[i])

for i in range(len(delivery)):
    sheet.cell(row=i+1, column=7, value=delivery[i])

for i in range(len(pay_format)):
    sheet.cell(row=i+1, column=8, value=pay_format[i])

for i in range(len(shop)):
    sheet.cell(row=i+1, column=9, value=shop[i])

workbook.save(filename='megamarket.xlsx')
