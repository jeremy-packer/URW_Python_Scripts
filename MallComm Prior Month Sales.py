import os
import csv
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

#Needs to be updated
campaign_type = "2021 Revised Historical Sales"
campaign_name = "01 2021 Revised Historical Sales"

#center list by name
centers = ["Westfield Annapolis","Westfield Brandon","Westfield Broward","Westfield Century City","Westfield Citrus Park","Westfield Countryside","Westfield Culver City","Westfield Fashion Square","Westfield Galleria At Roseville","Westfield Garden State Plaza","Westfield Mission Valley","Westfield Montgomery","Westfield North County","Westfield Oakridge","Westfield Old Orchard","Westfield Palm Desert","Westfield Plaza Bonita","Westfield San Francisco Centre","Westfield Santa Anita","Westfield Sarasota Square","Westfield Siesta Key","Westfield South Shore","Westfield Southcenter","Westfield Sunrise","Topanga","Westfield Trumbull","Westfield UTC","Westfield Valencia","Valley Fair","Westfield Wheaton","World Trade"]

options = Options()
#chrome_options.add_argument("--headless")  
driver = webdriver.Chrome("C://Users//jpacker//Scripts//chromedriver")

login_url = 'https://cms.mallcommapp.com/login'
#loads url using the driver
driver.get(login_url)
#enters the username
driver.find_element_by_name('username').send_keys("jeremy.packer@urw.com")
driver.find_element_by_xpath('/html/body/div[1]/div/div/form/button').click()
#enters the password
driver.find_element_by_name('password').send_keys("S1dn2!#3G1ABC")
#clicks the submit button
driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/form/button').click()

#this switches to the sales campaigns url
driver.get('https://cms.mallcommapp.com/sales-collection/')

div_counter = 2
for center in centers:
	
	#switches center
	if(center == centers[0]):
		time.sleep(1)
		driver.find_element_by_xpath('/html/body/div[2]/div[2]/button').click()
		#driver.find_element_by_class_name("btn btn-default").click()
		time.sleep(1)
		#enters center name into search field
		driver.find_element_by_xpath('//*[@id="scSearchInput"]').send_keys(center)
		time.sleep(1)
		#selects switch center to click by xpath div
		current_xpath = '//*[@id="switchCentreForm"]/div[2]/div[2]/div[' + str(2) + ']/div/div[1]'
		driver.find_element_by_xpath(current_xpath).click()
		driver.find_element_by_name('switch-centre').click()
		time.sleep(1)
	
	else:
		if center in ["Westfield Siesta Key","Westfield Sunrise","Westfield Citrus Park","Westfield Countryside","Westfield Sarasota Square"]:
			div_counter = div_counter + 1
		else:
			time.sleep(1)
			driver.execute_script("window.scrollBy(0, -500);")
			driver.find_element_by_xpath('/html/body/div[2]/div[2]/button').click()
			time.sleep(1)
			#enters center name into search field
			driver.find_element_by_xpath('//*[@id="scSearchInput"]').send_keys(center)
			time.sleep(1)
			#selects switch center to click by xpath div
			current_xpath = '//*[@id="switchCentreForm"]/div[2]/div[2]/div[' + str(div_counter) + ']/div/div[1]'
			driver.find_element_by_xpath(current_xpath).click()
			driver.find_element_by_name('switch-centre').click()
			time.sleep(1)
			div_counter = div_counter + 1
	
	if center not in ["Westfield Siesta Key","Westfield Sunrise","Westfield Citrus Park","Westfield Countryside","Westfield Sarasota Square"]:
		#selects date range
		driver.find_element_by_name('datefilter').click()
		#li[5] selects current month and li[6] selects prior month
		driver.find_element_by_xpath('/html/body/div[15]/div[1]/ul/li[5]').click()
		#driver.find_element_by_xpath('/html/body/div[15]/div[1]/ul/li[6]').click()
		
		#select campaign type
		driver.find_element_by_id('select2-campaignid-container').click()
		driver.find_element_by_xpath('/html/body/span/span/span[1]/input').send_keys(campaign_type)
		driver.find_element_by_xpath('/html/body/span/span/span[1]/input').send_keys(u'\ue007')
		
		#selects campaign
		driver.find_element_by_id('select2-period_formid-container').click()
		driver.find_element_by_xpath('/html/body/span/span/span[1]/input').send_keys(campaign_name)
		driver.find_element_by_xpath('/html/body/span/span/span[1]/input').send_keys(u'\ue007')
		
		#presses find and confirms inputs
		driver.find_element_by_name('submit').click()
		
		#downloads the csv after waiting one second to load
		time.sleep(1)
		driver.find_element_by_xpath('//*[@id="salesData_wrapper"]/div[1]/a[2]').click()
		time.sleep(1)
	
time.sleep(10)
driver.close()


##Specifies the path where the files are saved
#folder_path = r"C:\Users\jpacker\Desktop\MallComm"
#
#consolidated_df_list = list()
#
#for file in os.listdir(folder_path):
#	print(file)
#	df = pd.read_csv(folder_path + "\\" + file)  
#	consolidated_df_list.append(df)
#	
#output_dataframe = pd.concat(consolidated_df_list, axis=0, ignore_index=True)
#output_dataframe.to_csv('MallComm Monthly Sales.csv', index=False)
#
