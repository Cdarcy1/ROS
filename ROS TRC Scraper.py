import autoit, pyinputplus, openpyxl, time, tkinter
from tkinter import filedialog
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys

#launch browser and go to Ros
browser = webdriver.Firefox()
browser.get('https://www.ros.ie')
time.sleep(5)

#prompt for and upload certificate
root = tkinter.Tk()
root.filename = filedialog.askopenfilename(initialdir = "/", title = "Select ROS Digital Certificate", filetypes = [("", "*.p12.bac")])
linkElem = browser.find_element_by_link_text('Manage My Certificates')
linkElem.click()
time.sleep(2)
browser.find_element_by_xpath('//*[@id="file"]').send_keys(root.filename.replace("/", '\\'))

#prompt for user password, program will crash if wrong password entered
password = pyinputplus.inputPassword() #keeps password input secure
browser.find_element_by_xpath('//*[@id="certFilePassword"]').send_keys(password)
browser.find_element_by_xpath('//*[@id="submit_importFile"]').send_keys(Keys.ENTER)
browser.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div[1]/div/div/div[2]/form/div[4]/div[1]/div/div[1]/a').send_keys(Keys.ENTER) #return to login
time.sleep(3)


#login
browser.find_element_by_xpath('//*[@id="password"]').send_keys(password)
browser.find_element_by_xpath('//*[@id="password"]').send_keys(Keys.ENTER)
time.sleep(3)

#prompt for client list (available as an export from ROS)
root.filename = filedialog.askopenfilename(initialdir = "/", title = "Select Client List", filetypes = [("Excel", "*.xls*")])
wb = openpyxl.load_workbook(root.filename)
sheet = wb['Sheet1']

#main program loop
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
