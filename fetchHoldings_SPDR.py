# -*- coding: utf-8 -*-
"""
Created on Mon Jan  6 21:24:29 2020

Fetch NAV and holdings for all SSGA ETFs

@author: Derek G Nokes
"""

import pandas
import requests
import datetime
from bs4 import BeautifulSoup
import xlrd
import dateutil
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def downloadFile(url,outputDirectory,outputFileName):

    # start a session
    s = requests.session()

    # set the parameters
    parameters = {'siteEntryPassthrough': True}
 
    # try to extract the data
    try:
        # fetch the data
        r = s.get(url, params=parameters)                              
        # open the output file handle                                       
        outputFileHandle = open(outputDirectory+outputFileName,'wb')
        # write the holdings
        outputFileHandle.write(str(r.content))
        # close the file handle
        outputFileHandle.close()
                
    except Exception: 
        pass    
        
    return

def fetchDownloadLinks(baseUrl):
    #
    driver = webdriver.Chrome()
    #
    driver.get(baseUrl)
    
    time.sleep(5)
    
    pageSource=driver.page_source
    
    soup=BeautifulSoup(pageSource,'html.parser')
    
    # find all of the links
    links=soup.findAll("a")
    # iterate over all of the links

    downloadLinks=list()

    for link in links:
        # if the link contains 'ajax?fileType=csv&amp;fileName'
        if 'xlsx' in str(link):
            # extract the link
            downloadLink=link.attrs['href']
            # add to list
            downloadLinks.append(downloadLink)
            
    driver.quit()
    
    return downloadLinks

def fetchXlsx(baseUrl,rawDirectory,runDate):
    # fetch download links
    downloadLinks=fetchDownloadLinks(baseUrl)
    # iterate over each download link
    for url in downloadLinks:
        downloadLink='https://www.ssga.com'+str(url)
        print(downloadLink)
        outputFileName=downloadLink.split('/')[-1]
        outputFileNameList=outputFileName.split('.')
        outputFileName2=outputFileNameList[0]+"_"+runDate.strftime('%Y%m%d')+"."+outputFileNameList[1]
        downloadFile(downloadLink,rawDirectory,outputFileName2)    
    
    return downloadLinks

# define run date
runDate=datetime.datetime.now()
# define output directory
rawDirectory="F:/marketData/global_monitoring/instrument_universe/SPDRs/All/raw/xlsx/"
# define base URL
baseUrl='https://www.ssga.com/us/en/individual/etfs/fund-finder?tab=documents'
# fetch ETF holdings, NAVs, product listing, and performance
downloadLinks=fetchXlsx(baseUrl,rawDirectory,runDate)
