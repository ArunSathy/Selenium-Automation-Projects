from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from Gmail_config import EMAIL,PASSWORD
import pandas as pd

#fetching the excel file

dataFrame = pd.read_excel("C:\\Users\\arhsathy\\Desktop\\Automation.xlsx")
print(dataFrame.index)

#Connecting driver and opening browser

Drive=Service(executable_path='C:\Drivers\chromedriver_win32\chromedriver.exe')

browser=webdriver.Chrome()
browser.maximize_window()
browser.get('https://www.gmail.com')
time.sleep(1)

#Entering email id & password

email=browser.find_element(By.XPATH,'//*[@id="identifierId"]')
email.send_keys(EMAIL)

email_next_button=browser.find_element(By.XPATH,'//*[@id="identifierNext"]/div/button')
email_next_button.click()

time.sleep(1)

password=browser.find_element(By.XPATH,'//*[@id="password"]/div[1]/div/div[1]/input')
password.send_keys(PASSWORD)

password_next_button=browser.find_element(By.XPATH,'//*[@id="identifierNext"]/div/button')
password_next_button.click()

#email creation and sending mails for each

for i in dataFrame.index:
    # print(dataFrame.loc[i])
    time.sleep(1)

    compose_button= browser.find_element(By.XPATH)
    compose_button.click()

    to_field=browser.find_element(By.XPATH)
    to_field.send_keys(dataFrame.loc[i]['Email']) # in the quotes we can put email

    subject_field=browser.find_element(By.XPATH)
    subject_field.send_keys('Subject')

    body_field=browser.find_element(By.XPATH)
    body_field.send_keys(f"Hi {dataFrame.loc[i]['Name']}, \nWe are learning Selenium")

    send_button=browser.find_element(By.XPATH)
    send_button.click()

    time.sleep(5)
    

time.sleep(10)

