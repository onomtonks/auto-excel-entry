import openpyxl

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import DesiredCapabilities
from selenium.webdriver.common.proxy import Proxy, ProxyType
import time

PROXY='104.236.248.219:3128'
chrome_options=webdriver.ChromeOptions()
chrome_options.add_argument('--proxy-server=%s' % PROXY)
PAT=''
driver=webdriver.Chrome( options=chrome_options,executable_path =r'C:\Users\Ansab Waseem\Desktop\haseebproj\chromedriver.exe' )
bsObj=""
html=''
h=0
k=0
count=0

class ebayscrape(object):
	workbook=openpyxl.load_workbook(r'C:\Users\Ansab Waseem\Documents\py-list.xlsx') 
	sheet=workbook.get_sheet_by_name('Sheet2')
	for t in range(5):
		skc=t*50
		print(t)
		k=t+1
		link="https://www.ebay.com/sch/m.html?_nkw=&_armrs=1&_from=&_ssn=moaz1738&_pgn="+str(k)+'&_skc='+str(skc)+'&rt=nc'
		print(link)
		driver.get(link)
		html=driver.page_source
		bsObj = BeautifulSoup(html)
		soup=BeautifulSoup(html,"html.parser")
	
		for post in soup.findAll("li",{"class":"sresult lvresult clearfix li"}):
			p=post.findAll("a",{"class":"vip"})[0]
			h=post.findAll("a",{"class":"vip"})[0].text
			x=post.findAll("li",{"class":"lvprice prc"})[0].text
			count+=1
			print(h)
			print(x)
			print(p["href"])
			print(count)
	
			price_doll="B"+str(count+1)	
			link_ref="C"+str(count+1)
			heading="A"+str(count+1)
			sheet[heading].value=h
			sheet[price_doll].value=x
			sheet[link_ref].value=p["href"]
		'for saving in excel'
		#workbook.save(r'C:\Users\Ansab Waseem\Documents\sample.xlsx')
		for post in soup.findAll("li",{"class":"sresult lvresult clearfix li shic"}):
			p=post.findAll("a",{"class":"vip"})[0]
			h=post.findAll("a",{"class":"vip"})[0].text
			x=post.findAll("li",{"class":"lvprice prc"})[0].text
			count+=1
			print(h)
			print(x)
			print(p["href"])
			print(count)
	
			price_doll="B"+str(count+1)	
			link_ref="C"+str(count+1)
			heading="A"+str(count+1)
			sheet[heading].value=h
			sheet[price_doll].value=x
			sheet[link_ref].value=p["href"]
		'for saving in excel'
		#workbook.save(r'C:\Users\Ansab Waseem\Documents\py-list.xlsx')


