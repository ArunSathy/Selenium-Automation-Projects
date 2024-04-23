from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from Instagram_config import *



# users=['_ameen_al','nameisnihal__']

chrome=Service(executable_path='C:\Drivers\chromedriver_win32\chromedriver.exe')

#Opening Instagram Login Page

browser=webdriver.Chrome()
browser.maximize_window()
browser.get('https://www.instagram.com')
# time.sleep(1)

#Entering Login Details

username=browser.find_element(By.XPATH,'//*[@id="loginForm"]/div/div[1]/div/label/input')
username.send_keys(login_user)

password=browser.find_element(By.XPATH,'//*[@id="loginForm"]/div/div[2]/div/label/input')
password.send_keys(login_pass)

login_button=browser.find_element(By.XPATH,'//*[@id="loginForm"]/div/div[3]')
login_button.click()

#fetching details from individual page

# for user in users:
#     browser.get(f"https://www.instagram.com/{user}")

    # post=browser.find_elements(By.XPATH,'//*[@id="mount_0_0_TD"]/div/div/div[2]/div/div/div[1]/div[1]/div[2]/div[2]/section/main/div/header/section/ul/li[1]/span/span')
    # print(str(post))

    # bio=browser.find_element(By.CLASS_NAME,'x7a106z x972fbf xcfux6l x1qhh985 xm0m39n x9f619 x78zum5 xdt5ytf x2lah0s xdj266r x11i5rnm xat24cr x1mh8g0r xexx8yu x4uap5 x18d9i69 xkhd6sd x1n2onr6 x11njtxf xwonja6 x1dyjupv x1onnzdu xwrz0qm xgmu61r x1nbz2ho xbjc6do')
    # print(bio.text)

    # with open(f"{users}.txt",'w') as file:
    #     file.write(f"Number of Posts : {post.text}\nFollowers : {followers.text}\nFollowing : {following.text}\n\nBio :\n{bio.text}")
    # time.sleep(2)

time.sleep(60)

