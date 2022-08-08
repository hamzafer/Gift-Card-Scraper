from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time
import os
from dotenv import load_dotenv

webURL = "https://account.microsoft.com/billing/orders"
emailID = 'i0116'
passID = 'i0118'
noButtonID = 'idBtn_Back'
timeToScrape = 100 #time in seconds

driver = webdriver.Chrome()
driver.get(webURL)

time.sleep(5)
email = driver.find_element(By.ID, emailID)
email.send_keys(os.getenv('EMAIL'))
email.send_keys(Keys.RETURN)

time.sleep(2)
password = driver.find_element(By.ID, passID)
password.send_keys(os.getenv('PASSWORD'))
password.send_keys(Keys.RETURN)

time.sleep(2)
noButton = driver.find_element(By.ID, noButtonID)
noButton.send_keys(Keys.RETURN)

time.sleep(timeToScrape)
elementsCss = '.css-233'
totalElements = driver.find_elements(By.CSS_SELECTOR, elementsCss)
count = len(totalElements)
print('Total Order Scraped: ',count)

for x in range(1,count+1):
    try:
        giftcard = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div["+str(x)+"]/div/div/div[2]/div/div[3]/div/div/div//button[@aria-label='View gift code']")
        driver.execute_script("arguments[0].click();", giftcard)
        alert = driver.switch_to.active_element
        alert.click()
        copied_data = pd.read_clipboard()
        giftCardNumber = copied_data.columns[0]
        dateShate = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div["+str(x)+"]//div/div/div[1]/span/span[1]")
        orderShorder = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div["+str(x)+"]//div/div/div[1]/span/span[3]")
        print('Gift Card Number:',giftCardNumber, ' Date: ',dateShate.text, ' Order#: ',orderShorder.text)
    except:
        continue

print('Scrapping successful!')
driver.quit()