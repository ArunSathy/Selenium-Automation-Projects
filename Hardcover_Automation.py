from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

getpass=os.getlogin()
sys_username=getpass

chromedriver=Service(executable_path='C:\Drivers\chromedriver-win64\chromedriver.exe')
chrome_options=webdriver.ChromeOptions()

browser=webdriver.Chrome(service=chromedriver,options=chrome_options)
browser.maximize_window()
browser.get("https://www.nytimes.com/books/best-sellers/2024/03/31/hardcover-fiction/")

time.sleep(5)