import requests
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.support.select import Select
import time
from parsel import Selector
import sys
import xlsxwriter 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
import csv
import json
import xlrd 

avaSheens =',Matt,Low Sheen,Semi Gloss,Gloss'
Results =[]




#Loading page
driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://www.inspirationspaint.com.au/house-paint/interior-paint/walls/dulux-wash-and-wear')
driver.maximize_window()

#Reading excel file

wb = xlrd.open_workbook('Colours.xlsx') 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
# Change color range to scrape desired colors
for i in range(883,884): 
    print(sheet.cell_value(i, 0)) 
    
    

    #Finding and clicking Add Color
    wait = WebDriverWait(driver,10)
    (wait.until(ec.visibility_of_element_located((By.XPATH, "//div[@class='col-3-4']")))).click()
    
    
    #Clicking Search
    (wait.until(ec.visibility_of_element_located((By.XPATH, "//span[@class='cv-ico-general-search ico-right']")))).click()
    #Find input box and search color
    searchBox = wait.until(ec.visibility_of_element_located((By.XPATH, "//div[@class='colour-search-box']//input[@class='form-text k-input']")))
    listColor = sheet.cell_value(i, 0)   #Color from excel sheet
    searchBox.clear()
    searchBox.send_keys(listColor)  #listColor
    searchBox.send_keys(Keys.ENTER)
    #Finding searched colors list and iterating throug each color
    if not (wait.until(ec.visibility_of_element_located((By.XPATH, "//div[@class='colour-groups']//div[@class='colour-tile']/div[@class='colour-swatch-label']")))):
        continue
    colors = driver.find_elements_by_xpath("//div[@class='colour-groups']//div[@class='colour-tile']/div[@class='colour-swatch-label']")

    for color in colors:

        #Selecting searched colors
        Scolor= color.text
        color.click()
        time.sleep(1)
        #check is available sheens appears
        try:
            driver.find_element_by_xpath('//div[@data-bind="invisible: sheenMatch"][@style=""]')
            #saving all available sheen and clicking first one
            avaSheens = driver.find_elements_by_xpath("//span[@class='cart-product-title']/a")
            avaSheens = ','+','.join(avaSheen.text for avaSheen in avaSheens)
            btnAvaSheen = driver.find_element_by_xpath('//button[@class="btn cv-apply"]')
            btnAvaSheen.click()
            time.sleep(1)
        except:
            pass

        #Clicking Sheen dropdown
        (wait.until(ec.visibility_of_element_located((By.XPATH,'//span[@class="k-select"][1]')))).click()
        # Getting Sheen List items
        time.sleep(1)
        sheenList = driver.find_elements_by_xpath('//div[4]/div/div[2]/ul/li')
        # First loop to iterate throug Sheen items
        for sheen in sheenList:
            sheenName = sheen.text

            searchSheen = avaSheens.find(','+sheenName)
            if searchSheen == -1:
                continue
            sheen.click()
            time.sleep(2)
            #Checking notification box
            try:
                #Notification box click
                driver.find_element_by_xpath("//button[@class='btn small']").click()
                # Change color link click
                (wait.until(ec.visibility_of_element_located((By.XPATH, "//div[@class='change-colour-link']")))).click()
                #Finding color list and click current color
                wait.until(ec.visibility_of_element_located((By.XPATH, "//div[@class='colour-groups']//div[@class='colour-tile']/div[@class='colour-swatch-label'][1]")))
                color.click()
                time.sleep(1)
                #saving all available sheen and clicking first one
                wait.until(ec.visibility_of_element_located((By.XPATH,'//div[@id="box-slide"][@style="right: 0px;"]')))
                avaSheens = driver.find_elements_by_xpath("//span[@class='cart-product-title']/a")
                avaSheens = ','+','.join(avaSheen.text for avaSheen in avaSheens)
                btnAvaSheen = driver.find_element_by_xpath('//button[@class="btn cv-apply"]')
                btnAvaSheen.click()
                time.sleep(1)
                #Clicking Sheen dropdown
                (wait.until(ec.visibility_of_element_located((By.XPATH,'//span[@class="k-select"][1]')))).click()
                time.sleep(1)
                continue
            except:
                pass

            #Product Name
            ProName = driver.find_element_by_xpath('//h1[contains(@class,"widget-product-title")]').text
            # Clicking Size dropdown 
            (wait.until(ec.visibility_of_element_located((By.XPATH,'//span[1]/div/span/div[3]/span[1]//span[@class="k-select"]')))).click()
            # Getting Size list items
            time.sleep(1)
            sizes = driver.find_elements_by_xpath('//*[@id="body"]/div[18]/div/div[2]/ul/li')
            sizeJoin = (','.join(size.text for size in sizes)).strip('Please Select,')
            #Apending fields to list
            Results.append({
                'ColorGoogleSheet' : listColor,
                'ColorSearchResult' : Scolor,
                'Sheen' : sheenName,
                'ProductName': ProName,
                'Size': sizeJoin,
            })

            #Clicking Sheen dropdown
            (wait.until(ec.visibility_of_element_located((By.XPATH,'//span[@class="k-select"][1]')))).click()
            time.sleep(1)

        #Pressing change color link and select next color
        avaSheens =',Matt,Low Sheen,Semi Gloss,Gloss'
        #Press Esc to close Sheen dropdown
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        #Clicking add color link
        (wait.until(ec.visibility_of_element_located((By.XPATH, "//div[@class='change-colour-link']")))).click()   
        wait.until(ec.visibility_of_element_located((By.XPATH, "//div[@class='colour-groups']//div[@class='colour-tile']/div[@class='colour-swatch-label'][1]")))


    #Check if change color slide open then close it
    time.sleep(2)
    if(driver.find_elements_by_xpath('//div[@id="box-slide"][@style="right: 0px;"]')):
            print("window open")
            driver.find_element_by_xpath('//div[@id="box-slide"]/button').click()
            time.sleep(1)
            webdriver.ActionChains(driver).move_by_offset(20,20).perform()

    
#Printing results in json format
print(json.dumps(Results, indent=2))   

# Writing to CSV
with open('colors.csv', 'w', newline='') as csv_file:
    writer = csv.DictWriter(csv_file, fieldnames=Results[0].keys())
    writer.writeheader()
    
    for row in Results:
        writer.writerow(row)