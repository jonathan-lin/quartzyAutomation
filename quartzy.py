#!/usr/bin/env python
# coding: utf-8

# In[3]:

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import os
from os import listdir
from os.path import isfile, join
import datetime

# In[2]:
def quartzy(group_counter):
    
	###EDIT HERE###
	
    groups = ['21333','181125'] #Group ID Numbers
    locations = ['EngineeringV','CNSI'] #Group Names (used in naming output files)
    formStrings = ['Bioengineering', '( CNSI'] #Strings that identify the requisition form that you want generated
	
	filepathStub = '/home/jon/Documents/Quartzy/downloads'
	chromeDriverLocation = '/home/jon/Documents/Quartzy/chromedriver'
	outputFilepath = '/home/jon/Documents/Paperwork/'
	
	username = raw_input('Please enter username: ')
	password = raw_intput('Please enter password: ')
	
	reqForm_name = ''
	reqForm_email = ''
	
	###END EDIT###
	
    filepath=filepathStub+str(group_counter)+'/'    
    
    os.system('find '+filepathStub+str(group_counter)+'/ -name "*.pdf" -type f -delete')
    os.system('find '+filepathStub+str(group_counter)+'/ -name "*.csv" -type f -delete')
    
    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : filepathStub+str(group_counter)+"/"}
    chrome_options.add_experimental_option("prefs",prefs)
    chrome_options.add_argument("--start-maximized")
    
    driver = webdriver.Chrome(chromeDriverLocation,chrome_options=chrome_options)
    
    #LOG IN
    driver.get("http://www.quartzy.com/login")
    quartzyWaitForLoad(driver)
    
    elem = driver.find_element_by_name("txtEmail")
    elem.send_keys(username)
    elem = driver.find_element_by_name("txtPassword")
    elem.send_keys(password)
    elem = driver.find_element_by_name('commit')
    elem.click()
    time.sleep(3)
#%%    
    #GO TO ORDER REQUESTS PAGE
    driver.get("https://app.quartzy.com/groups/"+groups[group_counter]+"/requests?status[]=APPROVED")
    
    #TEST TO SEE IF PAGE IS LOADED
    quartzyWaitForLoad(driver)
    
    #GRAB MANUFACTURERS AND GRANT NUMBERS    
    vendor = driver.find_elements_by_class_name("vendor")
    del vendor[0]
    vendorList = []
    for ven in vendor:
        venNameTemp = ven.text.split('\n')[0]
        print(venNameTemp)
        vendorList.append(venNameTemp)
    uniqueVendors = list(set(vendorList))
    if uniqueVendors == []:
        return None
    grant = driver.find_elements_by_class_name("detail-line")
    grantList = []
    for gran in grant:
        grantNameTemp = gran.text
        if 'Grant ID' in grantNameTemp:
            grantNameTemp = grantNameTemp.replace('Grant ID','')
            print grantNameTemp
            grantList.append(grantNameTemp)
    uniqueGrants = list(set(grantList))
    
    for ii in range(len(uniqueVendors)):
        currentVendor = uniqueVendors[ii]
        for jj in range(len(uniqueGrants)):
            currentGrant = uniqueGrants[jj]
            checkboxes = driver.find_elements_by_class_name('is-selected-checkbox')
            boolean=0            
            for x in range(len(vendorList)):
                if vendorList[x] == currentVendor and grantList[x] == currentGrant:
                    print(x)
                    print(vendorList[x])
                    print(grantList[x])
                    quartzyClick(checkboxes[x])
                    boolean=1
            if boolean == 0:
                continue
            time.sleep(1)

            reqButton = driver.find_elements_by_class_name('qz-button')
            for tempButton in reqButton:
                if tempButton.text == 'Create Req Form':
                    quartzyClick(tempButton)
                    break
            quartzyWaitForLoad(driver)
            pageLoaded=0
            quartzyClick(driver.find_elements_by_class_name('ember-basic-dropdown-trigger')[0])
            time.sleep(1)
            
            forms = driver.find_elements_by_class_name('ember-power-select-option')
            for tempForm in forms:
                print(tempForm.text)
                if formStrings[group_counter] in tempForm.text:
                    quartzyClick(tempForm)
                    break
            time.sleep(2)
            textFields = driver.find_elements_by_class_name('ember-text-field')
            textFields[1].clear()
            textFields[1].send_keys(reqForm_name)
            textFields[2].clear()
            textFields[2].send_keys(locations[group_counter])
            textFields[4].clear()
            textFields[4].send_keys(reqForm_email)
            
            generateFormButton = driver.find_elements_by_class_name('qz-button')
            tempFilelist = os.listdir(filepath)
            for tempButton in generateFormButton:
                if 'Create Requisition Form' in tempButton.text:
                    quartzyClick(tempButton)
                    break
            fileDownloaded=0
            while fileDownloaded == 0:
                if len(os.listdir(filepath)) > len(tempFilelist):
                    fileDownloaded=1
                time.sleep(1)
            #Return to Order Requests Page
            driver.get("https://app.quartzy.com/groups/"+groups[group_counter]+"/requests?status[]=APPROVED")
    
            #TEST TO SEE IF PAGE IS LOADED
            quartzyWaitForLoad(driver)
    checkboxes = driver.find_elements_by_class_name('is-selected-checkbox')
    for ii in checkboxes:
        ii.click()
    csvButton = driver.find_elements_by_class_name('qz-button')
    for tempButton in csvButton:
        if tempButton.text == 'Export CSV':
            quartzyClick(tempButton)
            break
    orderedButton = driver.find_elements_by_class_name('qz-button')
    for tempButton in orderedButton:
        if tempButton.text == 'Mark Ordered':
            quartzyClick(tempButton)
            time.sleep(1)
            quartzyClick(driver.find_elements_by_class_name('confirm')[0])
            time.sleep(1)
    driver.close()
    
    today = datetime.date.today()
    mypath = filepathStub+str(group_counter)+'/'
    mypath2 = outputFilePath
    files = [f for f in listdir(mypath) if isfile(join(mypath,f))]
    #print files
    pdfs=[]
    for ii in range(0,len(files)):
        if files[ii][-4:]=='.pdf':
            pdfs.append(files[ii])
    #print pdfs
    csvs=[]
    for ii in range(0,len(files)):
        if files[ii][-4:]=='.csv':
            csvs.append(files[ii])
    #print csvs
    
    import csv
    print(mypath)
    with open(join(mypath,csvs[0]),"r") as source:
        rdr = csv.reader(source)
        with open(join(mypath,locations[group_counter]+'.csv'),"w") as result:
            wtr = csv.writer(result)
            for r in rdr:
              wtr.writerow((r[0],r[6],r[8],r[12],r[23]))
    os.system('libreoffice --headless --convert-to pdf --outdir '+filepathStub+str(group_counter)+'/ '+filepathStub+str(group_counter)+'/'+locations[group_counter].replace(' ','\ ')+'.csv')        
    time.sleep(5)
    os.system('pdftk ' + mypath + '*.pdf cat output ' + mypath2 + locations[group_counter].replace(' ','') + '_order_'+today.strftime('%y%m%d')+'.pdf')
            
#%%
def quartzyClick(button):
    clickSuccess = 0
    while clickSuccess == 0:
        try:        
            button.click()
            clickSuccess = 1
        except:
            print('Click failed... Trying again...')                            
            driver.execute_script("arguments[0].scrollIntoView(true)",content[x])
            time.sleep(1)

def quartzyWaitForLoad(driver):
    loaded = 0
    while loaded == 0:
        try:
            helpButton = driver.find_element_by_id('launcher')
            loaded=1
        except:
            print('Page not loaded')
            time.sleep(1)