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
import datetime


dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

webURL = "https://account.microsoft.com/billing/orders"
emailID = 'i0116'
passID = 'i0118'
noButtonID = 'idBtn_Back'
timeToScrape = 600  # time in seconds
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


#Finding All Elements Name
nameShame = driver.find_elements(By.CLASS_NAME, 'root-243')

#declaring lists
giftKeyList = list()
dateList = list()
orderList = list()
tempCount = 1;
nameList = list()

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
        stripStr = "Order number "
        strippedOrderNumber = orderNumberAboutToBeStripped.strip(stripStr)

        nameList.append(nameShame[x-1].text)
        giftKeyList.append(giftCardNumber)
        dateList.append(dateShate.text)
        orderList.append(strippedOrderNumber)
        tempCount += 1
    except:
        continue

print('---------------------------------------')
print('GK: ',*giftKeyList, sep = "\n")
print('DL: ',*dateList, sep = "\n")
print('OL: ',*orderList, sep = "\n")

country = '<br>[Country: Turkey]<br'
instructions = 'Instructions to redeem the code:<br>1. Download Nord VPN.<br>iOS link https://apps.apple.com/us/app/nordvpn-vpn-fast-secure/id905953485https://apps.apple.com/us/app/expressvpn-1-trusted-vpn/id886492891<br><br>Google PlayStore link: https://play.google.com/store/apps/details?id=com.nordvpn.android&hl=en&gl=US<br><br>2. Use these premium vpn credentials and connect to country mentioned after the code.<br>Credentials:<br><br>Email: dennenw@gmail.com<br>Pass: 1EcRU$13<br><br>**VPN is provided only to redeem the game code. Please don?t use the vpn for anything else.<br>Alternatively you can use any other vpn of your choice.<br>Go to redeem.microsoft.com and paste the code using the vpn and redeem to your account.<br>!!!SOME RUSSIAN CUSTOMERS HAVE EXPRESSED TROUBLE TO ACTIVATE KEY. IN THIS CASE	 CHANGE YOUR MICROSOFT ACCOUNT REGION BY GOING TO WWW.MICROSOFT.COM TO THE COUNTRY WRITTEN ABOVE + USE THE VPN OF THE SAME COUNTRY WRITTEN ABOVE!!!<br>FOR CUSTOMERS IN RUSSIA:<br>Sometimes Nord doesnot work in Russia use Seed4.Me:<br>Download the Seed4.MeApp<br>2. Run it and select Turkey region ??<br>3. Go to https://redeem.microsoft.com<br>4. Sign in to your account<br>5. Activate the key'

fileSuffix = '{:%Y-%b-%d %H-%M-%S}'.format(datetime.datetime.now())
folderName = 'data'
fileExtension = 'xlsx'
fileNameX = 'shakalakaboomboom'
fileName = folderName+'/'+fileNameX+'_'+fileSuffix+'.'+fileExtension

df = pd.DataFrame({'Date': dateList,'Order Number': orderList,'Name': nameList,'Gift Key': giftKeyList, 'Country': country, 'Instructions': instructions})
writer = pd.ExcelWriter(fileName, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()

print('Scrapping successful!')
driver.quit()