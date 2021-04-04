# Controlling the process by using methods
# Written by : Mahdi Khansari
# Date :       March 26, 2021

from pyOEMMethod import AEPOhioMethod as AEPOhio
import logging as log
from datetime import date
import json

#---------------------------------------------------
#    AEPOhio Data Usage Fetch (IntervalData)
#---------------------------------------------------
class AEPOhioProcess:
    
    config = None
    aep = None

    #-----------------------------------------------
    #   Initilization
    #-----------------------------------------------
    def __init__(self, configfile):
        ## Log Initiation
        logFileName = 'OEM-Log ' + date.today().strftime('%y-%m-%d') + '.log'
        log.basicConfig(format='%(asctime)s %(name)s> %(levelname)s: %(message)s', 
                        datefmt='%y-%m-%d %I:%M:%S',
                        filename=logFileName, level=log.DEBUG)
        log.info('+++++++++++++++++++++++++++++++++++++++++++')
        log.info('Process-Log Started')
        self.readConfig(configfile)
        
        self.aep = AEPOhio(usr = self.config['username'],
                          pwd = self.config['password'],
                          start = self.config['startDate'],
                          end = self.config['endDate'],
                          acc = self.config['accounts'],
                          downdir = self.config['downloadDir'])

        
    #-----------------------------------------------
    #   Get Interval Data
    #-----------------------------------------------
    def getData(self):
        log.info('AEPOhio interval data fetch process started!')
         
        # 1. Create a Firefox Web Driver
        self.aep.createWebDriver()
        
        # 2. Login to th AEPOhio Homepage
        self.aep.login()
        
        # 3. Fetch data for each account
        for accs in self.config['accounts']:
            self.aep.setAccount(accs[0])
            self.aep.dataFetch(accs[0])
            
        # 4. Logout and get the files 
        self.aep.logout()

        # 5. Close the WebDriver
        self.aep.closeWebDriver()
        
        # 6. returns the result dict
        downloadedfiles = self.aep.downloadedfiles
        log.info('AEPOhio interval data fetch process finished!')
        return downloadedfiles
        
    
    #-----------------------------------------------
    #   Get Latest bills
    #-----------------------------------------------
    def getLatestBills(self):
        log.info('AEPOhio latest bill process started!')
        
        # 1. Create a Firefox Web Driver
        self.aep.createWebDriver()
        
        # 2. Login to th AEPOhio Homepage
        self.aep.login()
        
        # 3. Ge latest bill for each account
        for accs in self.config['accounts']:
            self.aep.setAccount(accs[0])
            self.aep.retryLatestBillNum = 0  # retry iteration number
            self.aep.latestBills(accs[0])
        
        # 4. Logout and get the files 
        self.aep.logout()

        # 5. Close the WebDriver
        self.aep.closeWebDriver()
        
        # export the csv file including all accounts' bill
        self.aep.bills2CSV(self.config['downloadDir'])

        log.info('AEPOhio latest bill process finished!')
    
    #-----------------------------------------------
    #   read Config JSON
    #-----------------------------------------------
    def readConfig(self, jsonFile):
        # load the config from Json 
        configJson = open(jsonFile, 'r')
        self.config = json.load(configJson)
        