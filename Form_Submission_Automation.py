from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import pandas as pd
from tkinter import filedialog, Text, Tk
from selenium.webdriver.support.ui import Select


# dataFrame = pd.read_excel("C:\\Users\\arhsathy\\Desktop\\Automation.xlsx") # file getting by excel path

#---------------  we can use either method || but better is the dynamic fetching one --------------

input_file=filedialog.askopenfilename(title="Select the Input file",filetypes=(("Excel Files", "*.xlsx"), ("CSV Files", "*.csv"), ("All Files", "*.*")))
excel_data=pd.read_excel(input_file)  # file getting by dynamically || getting from the user

# df=excel_data.iloc[1]  # Just printing the values using positional based index || iloc[] : is function used for that
# print(df)

chromedriver=Service(executable_path='C:\Drivers\chromedriver_win32\chromedriver.exe')

browser=webdriver.Chrome(service=chromedriver)
browser.maximize_window()

for i in excel_data.index:
    entry_data = excel_data.loc[i]       # loc [] : is used to identify the specific element from the data || iloc [] : is used as a positional index to find the whole row values  || both can be used

    name_field=browser.find_element(By.XPATH,'')
    name_field.send_keys(entry_data['Full Name'])  # we also use || " entry["Full Name"] " everywhere

    email_field=browser.find_element(By.XPATH,'')
    email_field.send_keys(entry_data.loc['Email'])

    gender_xpath=f"//*[@value='{entry_data.loc['Gender']}']"
    gender_selection=browser.find_element(By.XPATH,gender_xpath)
    gender_selection.click()

    # converting the string to a list to fetch multiple languages from a single cell

    languages=entry_data.loc['Interested Languages'].split(', ') # in the bracket we specify exactly what is the sepration in the excel to splited
    # print(langs)

    for j in languages:
        language_xpath=f"//*[@value=\"{j}\"]"
        language_field=browser.find_element(By.XPATH,language_xpath)
        language_field.click()

    working_status=Select(browser.find_element(By.XPATH,''))  # Select : is used to give datas for drop down
    working_status.select_by_value(entry_data.loc['Working Status'])

    time.sleep(5)

    # ------------ Submit Button ------------------------

    # name_field.submit() # one way to click "Submit" button || we can use any variable used to fetch the data before

    # submit_button_1 = browser.find_element(By.XPATH, '')
    # submit_button_1.click()            # normal way of activating "Submit" button

    submit_button=browser.find_element(By.XPATH,'')
    browser.execute_script("arugments[0].click()",submit_button)  # "Submit" button is activated by using Javascript













