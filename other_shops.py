from selenium import webdriver
from selenium.webdriver.common.by import By
import time


def get_other_shops(browser: webdriver.Chrome,
                    items_links,
                    price,
                    cashback,
                    cashback_other,
                    cashback_sber,
                    delivery,
                    pay_format,
                    shop):

    browser.execute_script("window.scrollBy(0, 700)")

    time.sleep(4)

    try:
        browser.find_element(By.CSS_SELECTOR, 'a.more-offers-button.offers-info__more-offers-button-wrapper').click()
        time.sleep(3)
    except Exception as err:
        print(err)

    others = browser.find_elements(By.CSS_SELECTOR, 'div.product-offer.product-offer_with-payment-method')

    for shop_ in others:
        try:
            price_ = shop_.find_element(By.CLASS_NAME, 'product-offer-price__amount').text
            cashback_ = shop_.find_element(By.CLASS_NAME, 'bonus-amount').text
            delivery_ = shop_.find_element(By.CLASS_NAME, 'offer-item-delivery-type__delivery-date').text
            pay_format_ = shop_.find_element(By.CLASS_NAME, 'pdp-available-payment-method-block__text').text
            shop_item = shop_.find_element(By.CLASS_NAME, 'pdp-merchant-rating-block__merchant-name').text
            print(f'{price_}, {cashback_}, {delivery_}, {pay_format_}, {shop_item}')

            items_links.append(None)
            price.append(price_)
            cashback.append(cashback_)
            cashback_other.append(None)
            cashback_sber.append(None)
            delivery.append(delivery_)
            pay_format.append(pay_format_)
            shop.append(shop_item)
        except Exception as err:
            print(err)