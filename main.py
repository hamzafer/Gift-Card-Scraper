from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
#from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

s = Service("C:\Program Files (x86)\chromedriver.exe")
driver = webdriver.Chrome(service=s)


driver.get("https://account.microsoft.com/billing/orders")
time.sleep(5)

print(driver.title)

email = driver.find_element(By.ID, 'i0116')
email.send_keys("tgd21july22@outlook.com")
email.send_keys(Keys.RETURN)

time.sleep(2)

password = driver.find_element(By.ID, 'i0118')
password.send_keys("KhurramBilal")
password.send_keys(Keys.RETURN)

time.sleep(2)

noButton = driver.find_element(By.ID, 'idBtn_Back')
noButton.send_keys(Keys.RETURN)

time.sleep(100)

#date2 = driver.find_element(By.XPATH, "/div[@class='ms-Stack css-153']/div[@id='6a96edb0-c3b4-45ad-acbf-810dd0b4c607']/div[@class='ms-Stack css-230']/div[@class='ms-Stack css-156']/div[@class='ms-Stack css-231']/span[@class='css-232']//span[1]")

#ranged = driver.find_elements(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/")
#size = len(ranged)
#print(size)

totalElements = driver.find_elements(By.CSS_SELECTOR, '.css-233')
count = len(totalElements)
print(count)

#nameShame = driver.find_element(By.XPATH, "//*[@id='6a96edb0-c3b4-45ad-acbf-810dd0b4c607']/div/div/div[2]/div/div[1]/div[2]/div/div[1]/span/a[@title]")
#nameShame = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[1]/span/a").get_attribute('title')
#print(nameShame.text)
#nameShame = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[1]/span//a")
#print(nameShame.text)

for x in range(1,count+1):
    try:
        giftcard = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div["+str(x)+"]/div/div/div[2]/div/div[3]/div/div/div//button[@aria-label='View gift code']")
        giftcard.click()
        alert = driver.switch_to.active_element
        alert.click()
        copied_data = pd.read_clipboard()
        giftCardNumber = copied_data.columns[0]
        print(giftCardNumber)
    except:
        continue





# for x in range(1,count+1):
#    dateShate = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div["+str(x)+"]//div/div/div[1]/span/span[1]")
#    print(dateShate.text)
#    orderShorder = driver.find_element(By.XPATH, "//*[@id='order-history-wrapper']/div/div[3]/div[" + str(x) + "]//div/div/div[1]/span/span[3]")
#    print(orderShorder.text)


#date1 = driver.find_element(By.XPATH("//div[@class='ms-Stack css-231']/span[@class='css-129']"))
#print(date2)

"""
try:
    mainClass = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "ms-Stack css-153"))
    )
    classss = driver.find_elements(By.TAG_NAME, "class")
finally:
    driver.quit()
"""

#element = WebDriverWait(driver, 10).until(driver.find_element(By.ID, 'idSIButton9'))

#button = driver.find_element(By.ID, 'idSIButton9')
#button.submit()



#WebDriverWait(driver, 4).until(EC.element_to_be_clickable('idSIButton9')).click()




#time.sleep(3)
#WebDriverWait(driver, 3).until(EC.presence_of_element_located('idSIButton9')).send_keys(Keys.RETURN)



"""
try:
    main = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "ms-Stack css-153"))
    )
    dateAndOrderNumbers = main.find_elements(By.CLASS_NAME, 'ms-Stack css-231')
    for x in dateAndOrderNumbers:
        header = x.find_element(By.CLASS_NAME, 'css-232')
        print(header.text)
finally:
    driver.quit()
"""

#driver.quit()
