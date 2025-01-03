# steps to setup
# extend monitor both, 1920  x 1080
# shift chrome to big monitor , zoom out until all words can see
# only able to copy and paste into IDE, can't run as py file
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys # enable send keys
from selenium.webdriver.support import expected_conditions as EC # enable wait-until-element appears
from selenium.webdriver.common.by import By     # enables find element "By.ID"
from selenium.webdriver.common.action_chains import ActionChains # enabled mouse Scroll
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin # enable Scroll from element
import pandas as pd
import os,time
from datetime import datetime as dt
from datetime import timedelta  # add one day for FTR


USERNAME = "chris.cheng@spirent.com"
PASSWORD = "P@ssw0rd112024"
print(f'current script path {os.path.dirname(os.path.realpath(__file__))}')
# GECKO_DRIVER_PATH = "D:\\Python363\\projects\\SFDC ticket closer\\webdriver\\geckodriver.exe"

def find_element_and_click(type, element_description, target):
    # type: ID, XPATH
    # element_description examples: "//button[@type='submit']"
    # find elements with description
    if type == 'ID':
        elements = driver.find_elements(By.ID, element_description)
    elif type == 'XPATH':
        elements = driver.find_elements(By.XPATH, element_description)
    # go through elements to find traget and click
    for i in elements:
        if i.text == target:
            i.click()


#launch browser
driver = webdriver.Chrome()
# init wait function
wait30 = WebDriverWait(driver, 30)
# browse to onelogin
driver.get('https://spirent.onelogin.com')
# wait for page to load
wait30.until(EC.element_to_be_clickable((By.ID, 'username')))

# locate username box and input username
username_box = driver.find_element(By.ID, 'username')
username_box.send_keys(USERNAME)
# click Continue button
continue_button = driver.find_element(By.XPATH, "//button[@type='submit']")
continue_button.click()

# wait for password_box to appear
wait30.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))


# input password
password_box = driver.find_element(By.XPATH, "//input[@type='password']")
password_box.send_keys(PASSWORD)
# click Continue button
continue_button = driver.find_element(By.XPATH, "//button[@type='submit']")
continue_button.click()

# key in your code from onelogin token and press continue
time.sleep(20)

# read data from csv
df1 = pd.read_csv('tickets_to_close_1.csv')

# make a url:resolution_notes dictionary
urls = df1['url'].tolist()
res = df1['resolution'].tolist()
url_res_dict = dict(zip(urls, res))

test_url_res_dict = {}
test_url_res_dict['https://spirent.lightning.force.com/lightning/r/Case/500Un00000CEfB2IAL/view'] = 'cannot replicate'

viewport = driver.get_window_size()
viewport_width = viewport['width']
viewport_height = viewport['height']
viewport_origin_width = int(viewport_width / 2)
viewport_origin_height = int(viewport_height / 13)

def click_menu_and_select_option(menu_name, option_selected, scroll_length = 100):
    # scroll to reveal target menu
    scroll_origin = ScrollOrigin.from_viewport(910, 56)
    # scroll_origin = ScrollOrigin.from_viewport(viewport_origin_width, viewport_origin_height)
    ActionChains(driver).scroll_from_origin(scroll_origin, 0, scroll_length).perform()
    # click menu to show dropdown
    menu_xpath = "//button[@aria-label='"+ menu_name + "']"
    option_xpath = "//lightning-base-combobox-item[@data-value='"+ option_selected + "']"
    wait30.until(EC.element_to_be_clickable((By.XPATH, menu_xpath)))
    status_menu = driver.find_element(By.XPATH, menu_xpath)
    status_menu.click()
    time.sleep(1)
    # click option
    close_button = driver.find_element(By.XPATH, option_xpath)
    close_button.click()
    time.sleep(1)


for url, res_note in url_res_dict.items():
    # url = 'https://spirent.lightning.force.com/lightning/r/500Un00000FYhwMIAT/view'
    # res_notes = '500 over char'
    # browse to url
    driver.get(url)
    time.sleep(5)
    # click on "edit"
    wait30.until(EC.element_to_be_clickable((By.XPATH, "//button[@name='Edit']")))
    edit_button = driver.find_element(By.XPATH, "//button[@name='Edit']")
    edit_button.click()
    # select new window
    wait30.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='slds-modal__header']")))
    new_popup_window = driver.find_element(By.XPATH, "//div[@class='slds-modal__header']")
    # Set ticket 'Status' to "Close"
    click_menu_and_select_option('Status', "Closed")
    # Set 'Case Reason' to 'Configuration'
    click_menu_and_select_option('Case Reason', "Configuration", 200)
    # Set 'Case origin' - 'Email'
    click_menu_and_select_option('Case Origin', "Email", 200)
    # Set 'Resolution' - 'Configuration'
    click_menu_and_select_option('Resolution', "Technical Support Provided")
    # Set 'Support Level' - 'Backline APAC-India'
    click_menu_and_select_option('Support Level', 'Backline APAC-India')
    # get ticket create date
    wait30.until(EC.element_to_be_clickable((By.XPATH, "//lightning-formatted-text[@data-output-element-id='output-field']")))
    ticket_created_datetime = driver.find_element(By.XPATH, "//lightning-formatted-text[@data-output-element-id='output-field']").text
    datetime_1 = dt.strptime(ticket_created_datetime, '%m/%d/%Y %I:%M %p')
    # scroll to First Tech response
    # scroll_origin = ScrollOrigin.from_viewport(910, 56)
    scroll_origin = ScrollOrigin.from_viewport(viewport_origin_width, viewport_origin_height)
    ActionChains(driver).scroll_from_origin(scroll_origin, 0, 100).perform()
    # set First technical response data and time
    # if ticket created time is 2300, then set First response to next day 9am
    if datetime_1.hour == 23:
        print('hour == 23')
        next_day_date = datetime_1 + timedelta(days = 1)    # set to next day date
        next_day_date = next_day_date.replace(hour = 9, minute = 0)     # set to 9am
        print(f'next_day_time is {next_day_date}')
        ftr_date = dt.strftime(next_day_date, '%m/%d/%Y') 
        ftr_time_1 = dt.strftime(next_day_date, '%I:%M %p')
        ftr_time_1 = ftr_time_1.replace(minute=0, second=0, microsecond=0) # round off to nearest hour
        cfs_date_1 = dt.strftime(next_day_date, '%m/%d/%Y')
        cfs_time   = next_day_date + timedelta(hours = 5)
        cfs_time_1 = dt.strftime(cfs_time, '%I:%M %p')
    else: # just add one hour to ticket-created-time   
        ftr_date = dt.strftime(datetime_1, '%m/%d/%Y') 
        ftr_time = datetime_1 + timedelta(hours=1)
        ftr_time = ftr_time.replace(minute=0, second=0, microsecond=0) # round off to nearest hour
        ftr_time_1 = dt.strftime(ftr_time, '%I:%M %p')
        # add one day to Customer Fix supplied
        cfs_date = datetime_1 + timedelta(days=1)
        cfs_date_1 = dt.strftime(cfs_date, '%m/%d/%Y')
        cfs_time = ftr_time.replace(minute=0, second=0, microsecond=0)
        cfs_time_1 = dt.strftime(cfs_time, '%I:%M %p')
    # fill in First Tech Response date and time
    wait30.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='First_Technical_Response__c']")))
    ftr = driver.find_elements(By.XPATH, "//input[@name='First_Technical_Response__c']")
    ftr[0].click()
    ftr[0].clear()
    ftr[0].send_keys(ftr_date) # fill in ftr date
    time.sleep(1)
    ftr[1].click()
    ftr[1].clear()
    ftr[1].send_keys(ftr_time_1) # fill in ftr time
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform() # Esc key-press to clear drop-down box to scroll
    time.sleep(1)
    ### fill in Customer fix supplied date and time
    scroll_origin = ScrollOrigin.from_viewport(viewport_origin_width, viewport_origin_height)
    ActionChains(driver).scroll_from_origin(scroll_origin, 0, 100).perform() # scroll down to reveal "customer fix supplied" box
    wait30.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='Customer_Fix_Supplied__c']")))
    cfs = driver.find_elements(By.XPATH, "//input[@name='Customer_Fix_Supplied__c']")
    cfs[0].click()
    cfs[0].clear()
    cfs[0].send_keys(cfs_date_1) # fill in ftr date
    time.sleep(1)
    cfs[1].click()
    cfs[1].clear()
    cfs[1].send_keys(cfs_time_1)
    time.sleep(1)
    cfs_button = driver.find_elements(By.XPATH, "//lightning-base-combobox-item[@data-value='15:00:00.000']")
    cfs_button[1].click()
    time.sleep(3)
    # scroll to Resolution notes
    scroll_origin = ScrollOrigin.from_viewport(viewport_origin_width, viewport_origin_height)
    # scroll_origin = ScrollOrigin.from_viewport(310, 83)
    ActionChains(driver).scroll_from_origin(scroll_origin, 0, 1000).perform()
    ActionChains(driver).scroll_from_origin(scroll_origin, 0, 1000).perform()
    ActionChains(driver).scroll_from_origin(scroll_origin, 0, 500).perform()
    time.sleep(5)
    # Check if res_notes are empty before filling
    wait30.until(EC.element_to_be_clickable((By.XPATH, "//textarea")))
    textboxes = driver.find_elements(By.XPATH, "//textarea")
    time.sleep(3)
    res_notes_box = textboxes[6]
    time.sleep(1)
    if res_notes_box.text == "": # if resolution notes are empty
        res_notes_box.click()
        time.sleep(1)
        res_notes_box.send_keys(res_note)
        time.sleep(1)
    # click Save
    wait30.until(EC.element_to_be_clickable((By.XPATH, "//runtime_platform_actions-action-renderer[@apiname='SaveEdit']")))
    save_button = driver.find_element(By.XPATH, "//runtime_platform_actions-action-renderer[@apiname='SaveEdit']")
    save_button.click()
    time.sleep(5)

