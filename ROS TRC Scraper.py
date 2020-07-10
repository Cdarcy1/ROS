import os, time, openpyxl, requests, autoit
import pyinputplus as pyip
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys

#launch browser and go to Ros
browser = webdriver.Firefox()
browser.get('https://www.ros.ie')
time.sleep(5)

#upload certificate
linkElem = browser.find_element_by_link_text('Manage My Certificates') #navigate to manage certificates
linkElem.click()
time.sleep(2)
browser.find_element_by_xpath('//*[@id="file"]').send_keys(os.getcwd()+'\TaxAgent.p12.bac') #Upload TaxAgent Cert
password = pyip.inputPassword() #Enter password as ****** 
browser.find_element_by_xpath('//*[@id="certFilePassword"]').send_keys(password) 
browser.find_element_by_xpath('//*[@id="submit_importFile"]').send_keys(Keys.ENTER)
browser.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div[1]/div/div/div[2]/form/div[4]/div[1]/div/div[1]/a').send_keys(Keys.ENTER) #return to login
time.sleep(3)

#login
browser.find_element_by_xpath('//*[@id="password"]').send_keys(password)
browser.find_element_by_xpath('//*[@id="password"]').send_keys(Keys.ENTER)
time.sleep(3)

#main program loop
wb = openpyxl.load_workbook('CompaniesList.xlsx') #open list of clients per ROS
sheet = wb['Sheet1']

for x in range (1, sheet.max_row):
    time.sleep(7)
    companyName = sheet.cell(row=x,column=1).value
    browser.find_element_by_xpath('//*[@id="clientName"]').send_keys(companyName)
    browser.find_element_by_xpath('//*[@id="clientName"]').send_keys(Keys.ENTER)
    time.sleep(7)
    browser.find_element_by_xpath('//*[@id="letterOfResidenceLink"]').send_keys(Keys.ENTER)
    time.sleep(7)
    browser.find_element_by_xpath('//*[@id="submitLor"]').send_keys(Keys.ENTER)
    time.sleep(7)
    try:
        browser.find_element_by_xpath('//*[@id="confirmLor"]').send_keys(Keys.ENTER)
        time.sleep(7)
        browser.find_element_by_xpath('//*[@id="confirmLor"]').send_keys(Keys.ENTER)
        time.sleep(7)

        autoit.send('^s')
        time.sleep(7)
        autoit.send('C:\\users\\conor\\TRCs\\' + companyName)
        time.sleep(7)
        autoit.send('{ENTER}')
        time.sleep(7)
        autoit.send('^w')
        time.sleep(7)

        browser.find_element_by_xpath('//*[@id="greyBtn"]').send_keys(Keys.ENTER)
        time.sleep(7)
        browser.find_element_by_xpath('/html/body/div[7]/div[3]/div/button[2]').send_keys(Keys.ENTER)
        time.sleep(10)
        browser.find_element_by_xpath('//*[@id="link_id_AGENT_SERVICES"]').send_keys(Keys.ENTER)

    except NoSuchElementException:
        print(companyName)
        browser.find_element_by_xpath('//*[@id="greyBtn"]').send_keys(Keys.ENTER)
        time.sleep(7)
        autoit.send('{TAB}')
        time.sleep(7)
        autoit.send('{ENTER}')
        time.sleep(10)
        browser.find_element_by_xpath('//*[@id="link_id_AGENT_SERVICES"]').send_keys(Keys.ENTER)
