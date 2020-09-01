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
import csv


driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://www.sba.gov/partners/lenders/microloan-program/list-lenders')
driver.maximize_window()
time.sleep(1)

ddown = driver.find_element_by_xpath('//span[@class="Select-arrow-zone"]')

ddown.click()
states = driver.find_elements_by_xpath('//div[@class="Select-option"]')

with open('lenders.csv', 'w',newline='') as f:
    writer = csv.writer(f, delimiter=',')

    writer.writerow(['State', 'Title', 'Address','Phone'])


    for i in range(len(states)):
        
        state = driver.find_elements_by_xpath('//div[@class="Select-option"]')[i]
        stateName = state.text
        state.click()
        time.sleep(1)

        cards = driver.find_elements_by_xpath('//div[@data-testid="contact-card"]')
        
        for card in cards:

            title = card.find_element_by_xpath('.//h4').text
            address = card.find_element_by_xpath('.//div[@data-testid="contact address"]/div[2]').text
            address = address.replace('\n',' ')
            print('title: ',title)
            print("Address: ",address)
            try:   
                phone = card.find_element_by_xpath('.//div[@data-testid="contact phone"]/div[2]').text

            except:
                phone = ''
            
            print("Phone: ",phone)
            writer.writerow([stateName, title, address,phone])
            print("============ "+str(i)+" ==========")
            
        ddown.click()
        time.sleep(1)
        

f.close
    

     

    