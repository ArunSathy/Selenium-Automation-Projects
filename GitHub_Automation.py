from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from GitHub_Config import Username,Password
import sys
import pyperclip


chromedriver=Service(executable_path='C:\Drivers\chromedriver_win32\chromedriver.exe')

# checking whether the repo name & visibility selected or not -------------------

try:
    repo_name=sys.argv[1]
except:
    print('Repo name is required ')
    sys.exit()

try:
    visibility=sys.argv[2]
except:
    visibility = 'public' # value name is from inspecting the field

chrome_options=webdriver.ChromeOptions()
chrome_options.add_argument('--headless')  # is to avoid the opening the browsser

browser=webdriver.Chrome(service=chromedriver,options=chrome_options)
browser.maximize_window()
browser.get('https://github.com/login')

# login page setup

username=browser.find_element(By.XPATH,'//*[@id="login_field"]')
username.send_keys(Username)

password=browser.find_element(By.XPATH,'//*[@id="password"]')
password.send_keys(Password)

# sign-in button ----------------------

username.submit()  # one way of setting sign-in button

# sign_in=browser.find_element(By.XPATH,'//*[@id="login"]/div[4]/form/div/input[13]')  # normal way of sign-in
# sign_in.click()

# creating a new repo ---------------

browser.get('https://github.com/new')

repo_name_field=browser.find_element(By.XPATH,'//*[@id=":r3:"]')
repo_name_field.send_keys(repo_name)

visibility_field=browser.find_element(By.XPATH,f'//*[@name="visibilityGroup" and @value=\'{visibility}\']')
visibility_field.click()

time.sleep(2)

create_button=browser.find_element(By.XPATH,'/html/body/div[1]/div[6]/main/react-app/div/form/div[5]/button')
browser.execute_script("arguments[0].scrollIntoView();",create_button)
create_button.click()

time.sleep(5)

try:
    copy_button=browser.find_element(By.XPATH,'//*[@id="repo-content-pjax-container"]/div/git-clone-help/div[2]/div[2]/div/div/clipboard-copy')
    copy_button.click()
    # print(pyperclip.paste()) # if we are using the headless method, this won't be useful, so we are just printing it
    print(f'git remote add origin https://github.com/ArunSathy/{repo_name}.git\ngit branch -M main\ngit push -u origin main')

except:
    print('Repository name already exists.....')
# time.sleep(50)
browser.quit()