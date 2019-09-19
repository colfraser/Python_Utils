#!/usr/bin/env python
# coding: utf-8

# Clean version - no IP addresses, ID's or Passwords


# output to  C:\Users\colinf\OneDrive - xxx\WSL
# Setup Python and Subversion in Windows10
# sudo apt-get update
# sudo apt-get install python3-pip
# sudo pip install selenium

# Install chromedriver and include in path
# https://sites.google.com/a/chromium.org/chromedriver/downloads

# or install firefox driver as Chrome is having a problem displaying the screen and only working in headless mode: comment in this option '--headless'
# https://firefox-source-docs.mozilla.org/testing/geckodriver/#    https://github.com/mozilla/geckodriver/releases geckodriver.exe from geckodriver-v0.25.0-win64.zip
# firefox needs to have a running session.

# Once OTRS has been initiated then the user/password code can be by-passed. Comment out code with '''

# Program extracts the URLS from OTRS for each 'Maitenance Contract' and 'Newbury Data Printer' page
# Inactive entries do not show and hence are not extracted
# The URLs are then used to scrape the page details.
# This approach allows easier validation and testing


# NOTE 1. We extract the URLs first. These can not be used for very long. Certainly the day after they have been extracted they will be invalid.
# NOTE 2. Check the maximum number of maintenance and printer (2 places) pages. Could extend the code to work out when the last page has been read
# NOTE 3. item_Owner on the Maintenance page is some times problematic.
# NOTE 4. The printer details are variable after the Serial Number field. So the values are output as Other1,2,3 etc
# NOTE 5. LineFeeds and CR in details.

# Look for cwf-text to amend run time for testing. currently enabled.

# Run with python3 OTRS.py
# Output goes to the 'out' folder

import sys
import os
import datetime
import time
import selenium   # chg
from selenium import webdriver
#from selenium.webdriver.chrome.options import Options   # chrome version
from selenium.webdriver.firefox.options import Options   # firefox version
import requests
import urllib
from selenium.webdriver.common.by import By

print('OTRS Scraper - Starting')
now = datetime.datetime.now()
print ("Current date and time : ",now.strftime("%Y-%m-%d %H:%M:%S"))

os.getcwd()
print ("Path at terminal when executing this file:",os.path.realpath('.'))
print(sys.path)
print (selenium.__version__)


options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--no-proxy-server')
#options.add_argument('--headless')
options.add_argument('disable-features=NetworkService') 
options.add_argument('--disable-gpu')

#driver = webdriver.Chrome(executable_path='/mnt/c/Users/colinf/AppData/Local/Programs/Python/Python37/lib/chromedriver', chrome_options=options)
driver = webdriver.Firefox(executable_path='/mnt/c/Users/colinf/AppData/Local/Programs/Python/Python37/geckodriver.exe', options=options)



# Open OTRS page
driver.get('http://xx.xx.xx.xx/otrs/index.pl?Action=AgentITSMConfigItem;SortBy=Number;OrderBy=Down;View=;Filter=All;OTRSAgentInterface=eGsT4rd63ICgahSkT1utx1sHYi58RiBw')
driver.set_page_load_timeout(30)
time.sleep(1)     

#driver.save_screenshot('/mnt/c/Users/colinf/AppDatea/Local/test.png')  This does not work, at least with current versions of Chrome and Firefox. https://github.com/seleniumhq/selenium/issues/1141
driver.find_element_by_id('User').send_keys('user')
time.sleep(1)
driver.find_element_by_id('Password').send_keys('Password phrase')
time.sleep(1)
driver.find_element_by_id('LoginButton').click()
time.sleep(1)
print('Logged in')



###################################################################################
print('Maintenance Contract - Writing URL')

#driver.fullscreen_window()  # close button does not relocate correctly
elem = driver.find_element_by_partial_link_text("Maintenance Contract")
elem.click()
time.sleep(1)

# Not required Go to page 1  - NOTE 1 Changes every day!
#driver.get('http://10.90.100.18/otrs/index.pl?Action=AgentITSMConfigItem;SortBy=Number;OrderBy=Down;View=;Filter=137;OTRSAgentInterface=s0uSoLM0s1g5YSA3ml4KfEieiVjNQfVB');
time.sleep(1)

counterMaintenace = 0;

#link = driver.current_url
#f = urllib.request.urlopen(link)
#myfile = f.read()

writeFileObj = open('out/maintenance.txt', 'w')  # wb?


for loop1 in range(1,2,1):  #cwf-text to get the first page only
#for loop1 in range(1,41,1):     # Note 2
    genPage = '#GenericPage'+str(loop1)
    print(genPage)
    
    if  loop1 == 2:   # cwf-text
        break
    
    if  loop1 == 41:   # no skip forward at the end. To make sure we get the end of the last page
        print('41-break')
        break
    elif loop1 == 40: 
        print('40-end maitenance page')
        elem = driver.find_element_by_css_selector(genPage)
        elem.click()
    elif (loop1-1)%5 == 0 and loop1 > 2:
        elem = driver.find_element_by_partial_link_text(">>")
        elem.click()
    else:
        elem = driver.find_element_by_css_selector(genPage)
        elem.click()
    
    time.sleep(1) 
    i=0
           

    for a in driver.find_elements_by_xpath('.//a'):
        href_string = a.get_attribute('href')
        a_string = str(a.get_attribute('href')) + '\n'
        if 'AgentITSMConfigItemZoom' in str(href_string) and i==0: 
                print(a.get_attribute('href'))
                writeFileObj.write(a_string)  
                counterMaintenace += 1
                i=1
        else:
                i=0        # This is to prevent the same url being recorded twice.
                   
    time.sleep(1) 
                
        
writeFileObj.close()
print('counterMaintenace ', counterMaintenace) 

time.sleep(10) 
#driver.quit()  # Closes screen


###################################################################################
print('Printer - Writing URL')

elem = driver.find_element_by_partial_link_text("Newbury Data Printer")
elem.click()
time.sleep(1) 

#go to page 1. Changes every day!
#driver.get('http://xx.xxx.xxx.xxx/otrs/index.pl?Action=AgentITSMConfigItem;SortBy=Number;OrderBy=Down;View=;Filter=136;OTRSAgentInterface=JAjaFkyjpFJzdKbofELL411RYWvVEX2E');
#driver.get('http://xx.xx.xx.xx/otrs/index.pl?Action=AgentITSMConfigItem;Filter=136;View=;SortBy=Number;OrderBy=Down;StartWindow=60;StartHit=1576;OTRSAgentInterface=eGsT4rd63ICgahSkT1utx1sHYi58RiBw');
#time.sleep(1)  


counterPrinter = 0;

writeFilePrinter = open('out/printer.txt', 'w')   # wb


#for loop1 in range(1,71,1):   
for loop1 in range(1,2,1):    # cwf-test # Note 2
    genPage = '#GenericPage'+str(loop1)
    print(genPage)
    
    if  loop1 == 71:   # no skip forward at the end
        print('71-break')
        break
    elif loop1 == 70: 
        print('70-end')
        elem = driver.find_element_by_css_selector(genPage)
        elem.click()
    elif (loop1-1)%5 == 0 and loop1 > 2:
        #print('eq')
        elem = driver.find_element_by_partial_link_text(">>")
        elem.click()
    else:
        elem = driver.find_element_by_css_selector(genPage)
        elem.click()
    
    time.sleep(1) 
    i=0
           

    for a in driver.find_elements_by_xpath('.//a'):
        href_string = a.get_attribute('href')
        a_string = str(a.get_attribute('href')) + '\n'
        if 'AgentITSMConfigItemZoom' in str(href_string) and i==0: 
                print(a.get_attribute('href'))
                writeFilePrinter.write(a_string)  
                counterPrinter += 1
                i=1
        else:
                i=0        
    time.sleep(1) 
                
        
writeFilePrinter.close()
print('counterPrinter ',counterPrinter)

time.sleep(10) 



# read the Class: Printer Maitenance and write maitenance details and linnked printer details (also Kiosk)
# NOTE 5:  R181103  and '30741 CPS Renewal' have line feeds
# Capita Multiple Renewals - has an entry in the note field
# To format maitenance spreadsheet: Add note field column H, Add value above, Fore date fields as date, all other as text
# Renewal 33332 - bad date
# To format maintenance-printer spreadsheet - all text, one date coumn.
# 5 missing inactive maintenance contacts added manually to result set in DB SV16160, SV16998, SV17045, SV17087, SV17283
# No missing maitenance printer entries

###################################################################################
print('Maintenance Contract and linked printers - Processing')

elem = driver.find_element_by_partial_link_text("Maintenance Contract")
elem.click()
time.sleep(1) 

#go to page 1
#driver.get('http://xx.xx.xx.xx/otrs/index.pl?Action=AgentITSMConfigItem;Filter=137;View=;SortBy=Number;OrderBy=Down;StartWindow=0;StartHit=1;OTRSAgentInterface=JAjaFkyjpFJzdKbofELL411RYWvVEX2E');

CounterMaintenance=0
linkedCounter=0;
linkedHeader=0;

#readFileMaintenance = open("out/maintenance.txt", "r")
readFileMaintenance = open("out/maintenance.txt", "r")

#writeFileMaintenance = open('out/maintenance.csv', 'w') 
writeFileMaintenance = open('out/maintenance.csv', 'w') 

#writeFileLinked = open("out/maintenance_printer_linked.csv", "w")
writeFileLinked = open("out/maintenance_printer_linked.csv", "w")

a_string = 'MAINTENANCE-NAME,item_Deployment_State,item_Incident_State,item_Owner,item_Phone,item_EMail,item_Address,item_Warranty_Start_Date,item_Warranty_Expiration_Date'
print(a_string)
a_string = a_string + '\n'
writeFileMaintenance.write(a_string)  

    
while True:                            # Keep reading forever
    theline = readFileMaintenance.readline()   # Try to read next line
    if len(theline) == 0:              # If there are no more lines  
        break                          #     leave the loop

    if CounterMaintenance==3:          # cwf-test
        break                          # cwf-test
		
    # Now process the line we've just read
    print('\n'+str(theline))  #, end="")
    driver.get(str(theline))
    time.sleep(1) #might be needed to allow browser to load  
    item_Name = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[1]/td[2]')[0]
    try:
        item_Owner = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[4]/td[2]')[0]
        name_Owner = item_Owner.text
    except:
        name_Owner = 'Problem Record'     #  NOTE 3 Check output writeFileLinked for this string
        writeFileLinked.write(name_Owner)
    finally:
        time.sleep(1)
    #print(name_text)
    #if name_Owner == "Capita Travel and Events":  # if we want to select by owner
    if name_Owner is not None  or name_Owner is None :   
       #while True:
       #ry:
       #    meci = browser.find_elements_by_class_name('Content')
       #   for items in meci:
        item_Deployment_State = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[2]/td[2]')[0]
        item_Incident_State = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[3]/td[2]')[0]
        item_Phone = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[5]/td[2]')[0]
        item_EMail = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[6]/td[2]')[0]
        # NOTE 5 Addresses and other fields can contain windows line feeds
        item_Address = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[7]/td[2]')[0]
        item_Warranty_Start_Date = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[8]/td[2]')[0]
        item_Warranty_Expiration_Date = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[9]/td[2]')[0]
        #item_Attachments = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[10]/td[2]')[0]
        #item_Attachments = ' '
        
       
        #print(item_Name.text)                                         
        #print(item_Owner.text)  
        #print(item_Deployment_State.text)
        #print(item_Incident_State.text)
        #print(item_Phone.text)
        #print(item_EMail.text)
        #print(item_Address.text)
        #print(item_Warranty_Start_Date.text)
        #print(item_Warranty_Expiration_Date.text)
        
        maintenanceDetail = item_Name.text + ',' + item_Deployment_State.text + ',' + item_Incident_State.text + ',"' + item_Owner.text + '",' + item_Phone.text + ',' + item_EMail.text + ',"' + item_Address.text + '",' +  item_Warranty_Start_Date.text  + ',' + item_Warranty_Expiration_Date.text
        print('md')
        maintenanceDetail = maintenanceDetail.replace('\n', ' ').replace('\r', ' ')  # edit out line feeds and CR in address etc    
        maintenanceDetail = maintenanceDetail + '\n'   # .rstrip('\r\n')
        print(maintenanceDetail)
        writeFileMaintenance.write(maintenanceDetail)  

        CounterMaintenance += 1
        
        #noOfRows = len(driver.find_elements_by_class_name('DataTable')); # This is not a valid count as more than one datatable in structure
        #print("No of rows : ", noOfRows);      
        
        if linkedHeader==0:
            header = '"MAINTENANCE-NAME","CONFIGITEM#","PRINTER-NAME","DEPLOYMENT STATE","CREATED","LINKED AS"'
            print(header)
            header = header  + '\n'
            writeFileLinked.write(header)
            linkedHeader += 1;                   
        try:
        
            headerName = driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[5]/div[1]/div[2]/table[1]/thead[1]/tr[1]/th[3]");
            headerNameText = headerName.text
            if headerNameText == 'CONFIGITEM#':
                j1 = 0;
                while True:    
                    #allLinked.append(driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[5]/div[1]/div[2]/table[1]/thead[1]/tr[1]/th[3]").text)  #Header
                    j1 += 1
                    try:
                        allLinked = []; 
                        allLinked.append(item_Name.text);
                        allLinked.append(driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[5]/div[1]/div[2]/table[1]/tbody[1]/tr["+str(j1)+"]/td[3]/a[1]").text)
                        allLinked.append(driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[5]/div[1]/div[2]/table[1]/tbody[1]/tr["+str(j1)+"]/td[4]/span[1]").text)
                        allLinked.append(driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[5]/div[1]/div[2]/table[1]/tbody[1]/tr["+str(j1)+"]/td[5]/span[1]").text)
                        allLinked.append(driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[5]/div[1]/div[2]/table[1]/tbody[1]/tr["+str(j1)+"]/td[6]").text)
                        allLinked.append(driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[5]/div[1]/div[2]/table[1]/tbody[1]/tr["+str(j1)+"]/td[7]").text)
                    
                        b_string = '"' + '","'.join(map(str, allLinked))    # quote the output
                        b_string = b_string.replace('\n', ' ').replace('\r', ' ')  # edit out line feeds and CR in address etc
                        b_string = b_string.rstrip()
                        print(b_string)  
                        b_string = b_string + '"\n'
                        #b_string = allLinked  + '\n'
                        writeFileLinked.write(b_string)          
                        linkedCounter += 1
                    except:
                        break;
                                                            
    
        except:
            #time.sleep(1) 
            print('except-no-details-for-maintenance-contact')
    
readFileMaintenance.close()

print('CounterMaintenance ', CounterMaintenance)
print('linkedCounter', linkedCounter) 

writeFileLinked.close()
writeFileMaintenance.close()
time.sleep(10) 




###################################################################################
print('Printer details - Processing')

# Open OTRS page
driver.get('http://xx.xx.xx.xx/otrs/index.pl?Action=AgentITSMConfigItem;SortBy=Number;OrderBy=Down;View=;Filter=All;OTRSAgentInterface=eGsT4rd63ICgahSkT1utx1sHYi58RiBw')
time.sleep(1)  

driver.find_element_by_id('User').send_keys('user')
time.sleep(1)
driver.find_element_by_id('Password').send_keys('password phrase')
time.sleep(1)
driver.find_element_by_id('LoginButton').click()
time.sleep(1)

print('Keep Logged in')

elem = driver.find_element_by_partial_link_text("Newbury Data Printer")
elem.click()
time.sleep(1) 

i=0
printerCounter=0

readFilePrinter = open('out/printer.txt', "r")

writeFilePrinterFinal = open('out/printer.csv', 'w') 

while True:                                # Keep reading forever
    theline = readFilePrinter.readline()   # Try to read next line
    if len(theline) == 0:                  # If there are no more lines
        break                              #     leave the loop
    i += 1
    #if i==2:   #break out test  cwf-test
    #    break
    # Now process the line we've just read
    print(theline, end="")
    driver.get(str(theline))

    time.sleep(1)
    item_Name = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[1]/td[2]')[0]
    item_Owner = driver.find_elements_by_xpath('/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr[4]/td[2]')[0]
    Owner_text = item_Owner.text
    print(Owner_text)
    #if Owner_text == "Capita Travel and Events":
    if Owner_text is not None or Owner_text is None:
        
        #row = driver.find_elements_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]");
        #rData = [];
        #for webElement in row :
        #    rData.append(webElement.text);
        #print(rData);   
               
        noOfColumns = len(driver.find_elements_by_class_name('W25pc'));
  
        print("No of cols : ", noOfColumns);
    
        allData = [];
        allDataTitle = [];
    
        for j in range(1, noOfColumns+1) :
            # get text from the i th row and j th column
            #print(j);
            if printerCounter==0:
                allDataTitle.append(driver.find_element_by_xpath("/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr["+str(j)+"]/td[1]").text);
                
            allData.append(driver.find_element_by_xpath(     "/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr["+str(j)+"]/td[2]").text);
            
        if printerCounter==0:   # NOTE 4.  Don't know exactly what are in these columns as the details can change.
            allDataTitle[0] = 'PRINTER-NAME'
            allDataTitle.append('Other')
            allDataTitle.append('Other')
            allDataTitle.append('Other')
            allDataTitle.append('Other')
            #itemName = driver.find_element_by_xpath(     "/html[1]/body[1]/div[1]/div[3]/div[3]/div[2]/div[4]/div[1]/div[2]/table[1]/tbody[1]/tr["+str(j)+"]/td[2]").text
            print(allDataTitle)
            a_string = '"' +  '","'.join(allDataTitle) + '"' + '\n'
            writeFilePrinterFinal.write(a_string);
        
        print(allData);
        #a_string = '"' + '","'.join(allData) + '",' + ',' + ',' + '\n';
        #
        
        a_string = '"' + '","'.join(map(str, allData)) 
        a_string = a_string.replace('\n', ' ').replace('\r', '')
        a_string = a_string + '"\n'
        writeFilePrinterFinal.write(a_string)  
        
        printerCounter += 1      
    
    
writeFilePrinterFinal.close()
readFilePrinter.close()


print('printerCounter') 
print(printerCounter)

#########################################################################################



print('Script ran to end')
now = datetime.datetime.now()
print ("Current date and time : ",now.strftime("%Y-%m-%d %H:%M:%S"))

driver.quit()




