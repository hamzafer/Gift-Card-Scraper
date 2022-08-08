from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time
import os
from os.path import join, dirname
from dotenv import load_dotenv
from threading import Thread
from xlsxwriter import Workbook

dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

webURL = "https://account.microsoft.com/billing/orders"
emailID = 'i0116'
passID = 'i0118'
noButtonID = 'idBtn_Back'
timeToScrape = 200  # time in seconds
timeToWait = 3

ENV = 'ENV'
MAC_ENV = 'MAC'
WIN_ENV = 'WIN'
EMAIL = 'EMAIL'
PASSWORD = 'PASSWORD'

class TimerScraper(Thread):
    def run(self):
        for x in range(0, timeToScrape):
            print(x+1, '...', end='')
            time.sleep(1)

def getEnvVar(varName):
    return os.environ.get(varName)

if(getEnvVar(ENV) == MAC_ENV):
    driver = webdriver.Chrome()
else:
    path = Service("C:\Program Files (x86)\chromedriver.exe")
    driver = webdriver.Chrome(service=path)

driver.get(webURL)

time.sleep(timeToWait+1)
email = driver.find_element(By.ID, emailID)
email.send_keys(getEnvVar(EMAIL))
email.send_keys(Keys.RETURN)

time.sleep(timeToWait)
password = driver.find_element(By.ID, passID)
password.send_keys(getEnvVar(PASSWORD))
password.send_keys(Keys.RETURN)

time.sleep(timeToWait)
noButton = driver.find_element(By.ID, noButtonID)
noButton.send_keys(Keys.RETURN)

print('Scrapping started!')
print('Total wait = ', timeToScrape, ' seconds')
print('\nStarting Timer in seconds...', end='')
TimerScraper().start()  # start a new timer thread
time.sleep(timeToScrape)  # sleep so the page can load - main thread

elementsCss = '.css-233'
totalElements = driver.find_elements(By.CSS_SELECTOR, elementsCss)
count = len(totalElements)
print('\n\nTotal Order Scraped: ', count)

#declaring lists
giftKeyList = list()
dateList = list()
orderList = list()
tempCount = 1;

for x in range(1, count+1):
    try:
        giftcard = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div["+str(x)+"]/div/div/div[2]/div/div[3]/div/div/div//button[@aria-label='View gift code']")
        driver.execute_script("arguments[0].click();", giftcard)
        alert = driver.switch_to.active_element
        alert.click()
        copied_data = pd.read_clipboard()
        giftCardNumber = copied_data.columns[0]
        dateShate = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div["+str(x)+"]//div/div/div[1]/span/span[1]")
        orderShorder = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div["+str(x)+"]//div/div/div[1]/span/span[3]")

        orderNumberAboutToBeStripped = orderShorder.text
        strippedOrderNumber = orderNumberAboutToBeStripped.strip("Order number ")

        giftKeyList.append(giftCardNumber)
        dateList.append(dateShate.text)
        orderList.append(strippedOrderNumber)
        tempCount += 1
    except:
        continue

print('---------------------------------------')
print(giftKeyList)
print(dateList)
print(orderList)


df = pd.DataFrame({'Date': dateList,'Order Number': orderList,'Gift Key': giftKeyList})
writer = pd.ExcelWriter('shakalakaboomboom.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()


print('Scrapping successful!')
driver.quit()
