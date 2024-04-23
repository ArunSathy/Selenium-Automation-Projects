from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyperclip
from selenium.webdriver.common.keys import Keys


driver=Service(executable_path='C:\Drivers\chromedriver_win32\chromedriver.exe')

brw=webdriver.Chrome()
brw.maximize_window()
brw.get('https://web.whatsapp.com/')

with open ('Whatsapp_Group.txt','r',encoding='utf8') as file:
    groups = [group.strip() for group in file.readlines()]   # to get names by avoiding blank space and newlines

with open('Whatsapp_Message.txt','r',encoding='utf8') as file:
    message=file.read()

for i in groups:

    # in search bar searching the group name

    search_bar_xpath = '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p'
    search_bar=WebDriverWait(brw,500).until(EC.presence_of_element_located((By.XPATH,search_bar_xpath)))
    pyperclip.copy(i)  # pyperclip : used to allow copy programmatically
    search_bar.send_keys(Keys.CONTROL+'v')  # send_keys : used to paste by using keyboard actions "ctrl+v"

    time.sleep(2)

    #entering group name and clicking

    group_title_xpath = f'//*[@title="{i}"]' # getting the xpath by title and changing it for every run by each group name
    group_title=brw.find_element(By.XPATH,group_title_xpath)
    group_title.click()

    # inputting message in the text area

    time.sleep(2)

    message_bar_xpath = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p'
    message_bar=brw.find_element(By.XPATH,message_bar_xpath)
    pyperclip.copy(message)
    message_bar.send_keys(Keys.CONTROL+'v')
    message_bar.send_keys(Keys.ENTER)

    time.sleep(3)


time.sleep(10)