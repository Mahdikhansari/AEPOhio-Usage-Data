# Logins in AEP ohio customer website and 
# fetches interval usage data from an accounts
# Written by : Mahdi Khansari
# Date :       March 24, 2021

#---------------------------------------------------
#   Import Libraries
#---------------------------------------------------
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.remote_connection import LOGGER as seleLog
#from selenium.common.exceptions import TimeoutException, NoSuchWindowException
#from selenium.common.exceptions import WebDriverException
from datetime import date, datetime
import pandas as pd
import os, re, time
import logging as log

#---------------------------------------------------
#   AEPOhio Data Usage Fetch (IntervalData)
#---------------------------------------------------
class AEPOhioMethod:
    
    # Class Variable
    driver = None
    accountNumber = ''
    startDate, endDate = None, None
    username, password = None, None
    
    retryPageRefreshNum = 0 # retry to refresh the webpage
    retryDownloadNum = 0    # retry to download
    downloadDir = None
    downloadedfiles = []
    typeDelay = 0.5
    downloadDelay = 60
    
    bills = []              # Output bills' list 
    retryLatestBillNum = 0
    retryLatestBillMax = 4
    
    retryLogoutNum = 0
    retryLogoutMax = 4
    
    #-----------------------------------------------
    #   Initilization
    #-----------------------------------------------
    def __init__(self, usr, pwd, acc, 
                 start, end, downdir):
        self.startDate, self.endDate = start, end
        self.username, self.password = usr, pwd
        self.downloadDir = downdir
        self.downloadedfiles = acc
        
        ## Log Initiation
        logFileName = 'OEM-Log ' + date.today().strftime('%y-%m-%d') + '.log'
        log.basicConfig(format='%(asctime)s %(name)s> %(levelname)s: %(message)s', 
                        datefmt='%y-%m-%d %I:%M:%S',
                        filename=logFileName, level=log.DEBUG)
        log.info('+++++++++++++++++++++++++++++++++++++++++++')
        log.info('Method-Log Started')
        
        # Selenium -> Just warnings & errors
        seleLog.setLevel(log.WARNING)
        # urllib3 -> Just warnings & errors
        log.getLogger("urllib3").setLevel(log.WARNING)
        
    #-----------------------------------------------
    #   Create WebDriver
    #-----------------------------------------------
    def createWebDriver(self):    
        # Checks if the download dir exists or not
        # if not it makes it
        if not os.path.isdir(self.downloadDir):
            os.makedirs(self.downloadDir)
        log.info('Download Folder: ' + self.downloadDir)
        
        
        # Firefox profile
        profile = webdriver.FirefoxProfile()
        profile.set_preference('browser.download.folderList', 2) # custom location
        profile.set_preference('browser.download.manager.showWhenStarting', False)
        profile.set_preference('browser.download.dir', self.downloadDir)
        profile.set_preference("browser.download.useDownloadDir", True)
        profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')
        profile.set_preference("browser.cache.disk.enable", False)
        profile.set_preference("browser.cache.memory.enable", False)
        profile.set_preference("browser.cache.offline.enable", False)
        profile.set_preference("network.http.use-cache", False) 
        
        # Firefox driver
        self.driver = webdriver.Firefox(profile)        
        log.info('Firefox WebDriver initiated!')
        
        
    #-----------------------------------------------
    #   Close WebDriver
    #-----------------------------------------------
    def closeWebDriver(self):    
        # Firefox driver
        self.driver.close()
        log.info('Firefox WebDriver closed!')
        
        
    #-----------------------------------------------
    #   Login
    #-----------------------------------------------  
    def login(self):

        # opens AEPOhio.com
        self.driver.get("https://www.aepohio.com")
        assert "AEP Ohio" in self.driver.title

        #Login
        # clicks on login button and login form opens
        elem = self.driver.find_element_by_xpath(
            '/html/body/form/span[2]/noindexfb/nav/div/div[2]/span/button').click()
        
        # Username
        # types the username in the username field
        elem = self.driver.find_element_by_id ("ctl05_ctl04_ctl00_TbUserID")
        elem.clear()
        elem.send_keys(self.username)

        #Password
        # types the password in the password field
        elem = self.driver.find_element_by_id ("ctl05_ctl04_ctl00_TbPassword")
        elem.clear()
        elem.send_keys(self.password)

        # Login
        # clicks on login button
        elem = self.driver.find_element_by_id ("ctl05_ctl04_ctl00_BtnLogin").click()
        # the page redirects to /account
        # the usage data is in /account/usage
        log.info('Loggin in...')
        
        
    #-----------------------------------------------
    #   Data Fetch
    #-----------------------------------------------  
    def dataFetch(self, acc):
        log.info('-----------------------------')
        log.info('Account: ' + acc)
        
        #-------------------------------------------
        ## Account Selection
        # opens the account page
        self.driver.get("https://www.aepohio.com/account/?view=" + acc)
        
        # opens AEPOhio.com/Account/Usage
        self.driver.get("https://www.aepohio.com/account/usage/")
        
        """
        ## Dropdown List        
        # select the account drop list 
        select = Select(self.driver.find_element_by_xpath(
            '//*[@id="cphContentMain_ctl00_ctl02_ctl00_DdlAccounts"]'))
                
        # selects a value in the drop list
        try:
            select.select_by_value(acc)
        except:
            log.error('Account ' + acc + 'is not found!')
            return
        """
        
        ## iFrame switch
        # there is an iframe which shows the Google Analytics loaded as AJAX
        # so first the driver swtichs to the destination iframe
        # as it shows up with a delay, we used WebDriveWait
        try:
            frame = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((
                    By.XPATH,'//*[@id="energy_usage_trends-frame"]'))
            )
        except:
            # when swtiching to iframe fails, it might be a problem from AEPOhio 
            # another reason could be a failure in loading the iframe content
            log.error('iframe switch failed!')
            return
        self.driver.switch_to.frame(frame)
        
        ## Scroll Down
        # waits for the AJAX Google Analytics to load
        # scrolls down to the the three bars button
        try:
            self.driver.execute_script("arguments[0].scrollIntoView(true);",
                                  WebDriverWait(self.driver, 10).until(
                                      EC.visibility_of_element_located((
                                          By.XPATH,
                                          '/html/body/div[3]/div[2]/div/div/div[2]/button'))))
        except Exception as e:
             log.error('Source website AEPOhio.com failed to load the graph data, trying again...')
             log.error('Selenium :', str(e))
             
             self.retryPageRefreshNum = self.retryPageRefreshNum + 1      # try again for 3 times
             if self.retryPageRefreshNum < 3:
                 # let the server rest for 30 seconds
                 time.sleep(30)
                 self.dataFetch(acc)
                 return
             else:
                 log.error('All graph data retries failed!')
                 return
        
        # waits for te 3 bars button to get clickable 
        elem = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((
                By.XPATH,'/html/body/div[3]/div[2]/div/div/div[2]/button')))
        
        # after a lot of tests, it seems here we need a small delay (just 0.5 second)
        time.sleep(1)
        
        # clicks on the 3 bar button at the upper right to open the list
        try:
            elem.click()
        except:
            log.error('3 bars button is not still clickable!')
            return
        
        # clicks on the 'Export All Data CSV' item in the dropdown list
        # then it driver should wait for the dialog box 
        elem = self.driver.find_element_by_xpath(
            '/html/body/div[3]/div[2]/div/div/div[2]/ul/li[5]/a').click()
        
        # waits for the selection dialog box 'modal-Content'
        elem = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((
                By.XPATH,'/html/body/div[5]/div/div'))
        )
        
        ## Time
        # Formats dates and converts it to String
        StartDate = self.startDate
        EndDate = self.endDate
        log.info('time: [' + StartDate + '] > [' + EndDate + ']')
        
        # From Date
        # as the scipts check as you type, so I used a delayed send
        self.send_keys_delay(XPath = '//*[@id="green_button_form_start_green_button_date_range"]',
                             word = StartDate, 
                             delay = self.typeDelay)
        
        # End Date
        # as the scipts check as you type, so I used a delayed send
        self.send_keys_delay(XPath = '//*[@id="green_button_form_end_green_button_date_range"]',
                             word = EndDate, 
                             delay = self.typeDelay)
        
        ## Download
        # list the files in a dir
        from os import walk
        currFiles = []
        for (dirpath, dirnames, filenames) in walk(self.downloadDir):
            currFiles.extend(filenames)
        
        # clicks on download
        elem = self.driver.find_element_by_xpath(
            '/html/body/div[5]/div/div/div[3]/button[2]').click()
        
        isDownloaded = False
        
        # 60 * 1s wait for each file to be downloaded
        # break the for after finding a new file
        for _ in range(self.downloadDelay):
            for (dirpath, dirnames, filenames) in walk(self.downloadDir):
                for file in filenames:
                    
                    if not file in currFiles:
                        # found the new file, put it in next to the account number
                        index = [i[0] for i in self.downloadedfiles].index(self.accountNumber)
                        # index 2 is for files in config file
                        self.downloadedfiles[index][2] = file
                        isDownloaded = True
                        log.info('Download Completed! -> ' + file)

                        break
                    
            # no new file yet
            if not isDownloaded:
                time.sleep(1)  # waits 1s then checks the download folder again
            else:
                break
        
        if not isDownloaded:
             self.retryDownloadNum = self.retryDownloadNum + 1
             if self.retryDownloadNum < 3:
                 log.error('Download timed out! trying again in 30 seconds...')
                 time.sleep(30)  # let the server rest for 30 seconds
                 self.dataFetch(acc)
             else:
                 log.error('All download retries are failed!')
                 return


    #-----------------------------------------------
    #   Latest Bill Amount, Due Date
    #-----------------------------------------------  
    def latestBills(self, acc):
        
        #debug
        print('Begin', acc)
        
        #-------------------------------------------
        ## Retry
        # retry process until Max number of retry
        self.retryLatestBillNum = self.retryLatestBillNum + 1
        if self.retryLatestBillNum > self.retryLatestBillMax:
            return
        
        if self.retryLatestBillNum > self.retryLatestBillMax -1:
            self.driver.refresh()  # refresh the page in the last retry
        
        if self.retryLatestBillNum > 1:
            time.sleep(10)  # 10s sleep for each retry
        
        # switch back to default content
        self.driver.switch_to.default_content()
        
        time.sleep(1)
        
        #-------------------------------------------
        ## Account Selection
        # opens the account page
        self.driver.get("https://www.aepohio.com/account/?view=" + acc)
               
        #-------------------------------------------        
        ## Bill Amount
        try:
            elem=WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((
                    By.XPATH,'/html/body/form/div[3]/span/div/div[3]/div/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div'))
            )
        except:
            log.error(acc + ': Amount element is not found!')
            self.latestBills(acc)  # retry
            return
        
        # string bill amount to float
        billAmount = elem.text.replace(',','')
        billAmount = float(re.findall('\d+\.\d+',billAmount)[0])
        
        #-------------------------------------------          
        ## Bill Due Date
        try:
            elem=WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((
                    By.XPATH,'/html/body/form/div[3]/span/div/div[3]/div/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div'))
            )
        except:
            log.error(acc + ': Due date element is not found!')
            self.latestBills(acc)  # retry
            return 
        
        # string bill due date to date
        billDue = datetime.strptime(elem.text, '%b. %d, %Y')
        billDue = billDue.strftime('%m/%d/%Y')
        
        #-------------------------------------------  
        ## Address(Name) and SDI
        try:
            elem=WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((
                    By.XPATH,'/html/body/form/div[3]/span/div/div[1]/div/span/div/div/div/div[3]/div/div/div[1]/div/div[1]'))
            )
            
        except:
            log.error('Account Address or SDI is not found!!')
            self.latestBills(acc)  # retry
            return
        
        txt = elem.text
        # example: '1101 N CORY ST\nSDI #: 00140060724919820'
        accAddr = txt[:txt.find('\n',0)]
        SDI = txt[txt.find('SDI #: ',0)+7:]
        
        
        # Check if the values of this bill is not same with the last one
        # sometimes it happens due to the browser issues
        if len(self.bills) != 0:  # not for the first time
            if self.bills[-1][2] == '#'+SDI:
                # browser issue, retry
                log.error('Browser issue, same account values, retrying...')
                self.latestBills(acc)  # retry
                return 
            
            
        #-------------------------------------------  
        # adds this bill to the bills' list
        thisBill = ['#'+acc, accAddr, '#'+SDI, billAmount, billDue]
        #debug
        print('#'+acc, accAddr, '#'+SDI, billAmount, billDue)
        self.bills.append(thisBill)
        
        log.info(acc + ': Latest bill extracted successfully!')
        
        
    #-----------------------------------------------
    #   Bills2CSV
    #-----------------------------------------------  
    def bills2CSV(self, csvAddr):
        
        billColumns = ['AccountNo', 'AccountAddress', 'SDI', 'Bill Amount', 'Bill Due']
        
        # CSV file name => 21-03-03-Latest Bills
        billcsvFileName = csvAddr + '\\' + datetime.now().strftime('%Y-%m-%d %H%M%S') + '-Latest Bills.csv'
        
        # using Pandas to make CSV file
        bills_df = pd.DataFrame(self.bills, columns=billColumns)
        bills_df.to_csv(billcsvFileName, index=False)
        
        log.info('Bill file exported successfully in: ' + str(csvAddr))
    
    
    #-----------------------------------------------
    #   Logout
    #----------------------------------------------- 
    def logout(self):
        
        log.info('Logging out...')
        
        # retry process until Max number of retry
        self.retryLogoutNum = self.retryLogoutNum + 1
        if self.retryLogoutNum > self.retryLogoutMax:
            return
        
        if self.retryLogoutNum > self.retryLogoutMax -1:
            self.driver.refresh()  # refresh the page in the last retry
        
        if self.retryLogoutNum > 1:
            time.sleep(5)  # 10s sleep for each retry
        
        #Logout
        # switch back to default content
        self.driver.switch_to.default_content()
        
        # scroll up until logout button
        self.driver.execute_script("arguments[0].scrollIntoView(true);",
                              WebDriverWait(self.driver, 10).until(
                                  EC.visibility_of_element_located((
                                      By.XPATH,
                                      '//*[@id="ctl04_BtnGlobalLogout"]'))))
        
        # wait to get clickable
        elem = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((
                By.XPATH,'//*[@id="ctl04_BtnGlobalLogout"]')))
                
        # clicks on logout button
        elem = self.driver.find_element_by_xpath(
            '//*[@id="ctl04_BtnGlobalLogout"]')
        time.sleep(1)
        
        
        # try again to logout on failure
        try:
            elem.click()
        except:
            log.error('Logging out was failed, trying again...')
            self.logout()
            return
        
        log.info('Method-Log Ended')
        log.info('+++++++++++++++++++++++++++++++++++++++++++')
        

    #-----------------------------------------------
    #   Send Keys with delay
    #-----------------------------------------------
    def send_keys_delay(self, XPath, word, delay):   
        # locate the element
        elem = self.driver.find_element_by_xpath(XPath)
        elem.clear()
        for c in word:
            elem.send_keys(c)
            time.sleep(delay)

    #-----------------------------------------------
    #   Get\Set Functions
    #-----------------------------------------------
    def setAccount(self, acc):
        self.accountNumber = acc
        self.loadtryNum = 0
        self.downloadtryNum = 0
    
    def getAccount(self):
        return self.accountNumber
    
    