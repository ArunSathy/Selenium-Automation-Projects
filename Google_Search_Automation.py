from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


chromedr=Service(executable_path='C:\Drivers\chromedriver_win32\chromedriver.exe')

#Opening the Google chrome

brw=webdriver.Chrome()
brw.maximize_window()
brw.get('https://www.google.com/')
time.sleep(1)

#Tying the thing in the Search Bar

bar=WebDriverWait(brw,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="APjFqb"]')))
bar.send_keys('Python')


button_1=brw.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[4]/center/input[1]')
button_1.click()


time.sleep(10)

