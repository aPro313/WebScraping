#Import libraries
import requests
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.support.select import Select
import time
from parsel import Selector
import sys
import xlsxwriter 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

# Ceating new worksheet and adding headers
workbook = xlsxwriter.Workbook('Class.xlsx') 
worksheet = workbook.add_worksheet('ClassSheet') 

worksheet.write('A1', 'course_code') 
worksheet.write('B1', 'course_title') 
worksheet.write('C1', 'couse_name') 
worksheet.write('D1', 'description') 

row = 2
# col = 1
# Finally, close the Excel file 
# via the close() method. 

#Chrome driver and getting website
driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://schedule.msu.edu/')
driver.maximize_window()
time.sleep(1)

#Selecting Term from Select options list
ddown = driver.find_element_by_name('ctl00$MainContent$ddlTerm')
ddownSel  =  Select(ddown)
ddownSel.select_by_visible_text('Spring 2020')    # ddownSel.select_by_index(2)  |  ddownSel.select_by_value('FS191194fall 2019') 

#First loop for subject iteration
for s in range(5):    #len(ddownSel.options)-1
    # Selecting first subject from Select options list
    ddown = driver.find_element_by_id('MainContent_ddlSubject')
    ddownSel  =  Select(ddown)
    ddownSel.select_by_index(s)
    #Selecting checkbox of all results
    driver.find_element_by_id('MainContent_chkAllonePg').click()
    time.sleep(1)
    #Clicking find course button
    driver.find_element_by_id('MainContent_btnSubmit').click()
    time.sleep(1)

    # Getting all courses
    courseLinks = driver.find_elements_by_xpath("//div[@id='MainContent_divHeader1_va']/h3/a")

    # storing the current window handle to get back to dashbord 
    main_page = driver.current_window_handle 

    #Second loop for courses interation
    for course in courseLinks:
        Data = []
        corCode= course.text   # course code
        course.click()
        time.sleep(3)
    
        #Switiching to course popup dialog
        driver.switch_to.frame('CourseFrame')
        # WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME,"CourseFrame")))
        
        # print(driver.page_source)
        corTitle = driver.find_element_by_xpath('//td/h3').text   # course title
        corName   = corTitle.strip(corCode)     # course name
        corDes = driver.find_elements_by_xpath("//div[@class='col-xs-6']")[2].text  # course desc

        print('Code: ',corCode)
        print('Title: ',corTitle)
        print('Name: ',corName)
        print('Desc: ',corDes)

        Data = [corCode,corTitle,corName,corDes]
        for i in range(4):
            worksheet.write(row, i, Data[i])
        #Closing popup dialog
        row+=1       
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)
        
        driver.switch_to.window(main_page)
        
        # driver.switch_to_default_content 

    
    
workbook.close()

    # sys.exit('exiting python')



    

