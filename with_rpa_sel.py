from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.Excel.Application import Application
import time
import re
from bs4 import BeautifulSoup
import os
import shutil
import glob
curr=os.getcwd()


browser = Selenium()
url = "https://itdashboard.gov/"
term = "python"
pdf_filename = "output/screenshot.png"
download_dir="/output"
download_dir1=curr
browser.set_download_directory(curr)
browser.open_available_browser(url)
'''
We got the xpath from the dive in button on the website
'''
btn_xpath='//*[@id="node-23"]/div/div/div/div/div/div/div/a'

browser.click_link(btn_xpath)
time.sleep(10)
page_source=browser.get_source

doc=BeautifulSoup(page_source(),'html.parser')

soup1=doc.find_all('div',{'id':'agency-tiles-2-container'})
#######
'soup1 contains the body where every department and its money is mentioned'
#######
'if you want to see just print the below statement by removing the string'
'print(soup1)'
#############
'now lets scrape through it to take what we want...'
#########
name_and_money=(soup1[0].find_all('span'))
#print(name_and_money)
#######
' our desired data lies in the above command'
name=[]
money=[]
for i in soup1[0].find_all('span',{'class':'h4 w200'}):
    name.append(i.text)
for i in soup1[0].find_all('span',{'class':'h1 w900'}):
    money.append(i.text)
data1=zip(name,money)
data=dict(data1)
print(data)
browser.click_link('view')
time.sleep(10)
browser.select_from_list_by_label('investments-table-object_length','All')
time.sleep(15)
page_source2=browser.get_source
doc2=BeautifulSoup(page_source2(),'html.parser')
soup2=doc2.find('div',{'id':'investments-table-container'})
UII=[]
for i in soup2.find_all('td',{'class':'left sorting_2'}):
    UII.append(i.text)
'Bureau=[]'
Bureau=[]
for i in soup2.find_all('td',{'class':'left select-filter'}):
	Bureau.append(i.text)
'Inv_title=[]'
Inv_title=[]
for i in soup2.find_all('td',{'class':'left'}):
	Inv_title.append(i.text)
Total_Spend=[]
for i in soup2.find_all('td',{'class':'right'}):
	Total_Spend.append(i.text)
                         

                         
CIO_Rat=[]
for i in soup2.find_all('td',{'class':'center'}):
	CIO_Rat.append(i.text)
############
'Noticed that the class of Bureau and Type are the same '
'also CIO Rating and No of proj have same class'
'We will separate them by even and and odd combination as all the odd entries'
'are from one category and the even are from other'
bureau=[]
cio=[]
                         
type1=[]
No_of_proj=[]
for count, i in enumerate(Bureau):
    if count % 2 != 1:
        bureau.append(i)
    else:
        type1.append(i)                         

for count, i in enumerate(CIO_Rat):
    if count % 2 != 1:
        cio.append(i)
    else:
        No_of_proj.append(i)
'Cleaning Investment title'
for i in UII:
	for j in Inv_title:
		if i==j:
			Inv_title.remove(j)
for i in bureau:
	for j in Inv_title:
		if i==j:
			Inv_title.remove(j)

for i in type1:
	for j in Inv_title:
		if i==j:
			Inv_title.remove(j)
"""
Now we add data into new excel files
First excel file will be called 'Agencies'
Second would be Individual Investments
"""
filename1="output/Agencies.xlsx"
filename2="output/Individual Investments.xlsx"
file=Files()
##file.create_workbook('Agencies.xlsx')
file.create_workbook()
file.set_cell_value(row=1,column=1,value='Department Name')
file.set_cell_value(row=1,column=2,value='Spendings')

for i in range(len(name)):
	file.set_cell_value(row=i+2,column=1,value=name[i])

for i in range(len(money)):
	file.set_cell_value(row=i+2,column=2,value=money[i])
file.save_workbook(filename1)
###
'''
NOw second excel file
'''

file.create_workbook()
file.set_cell_value(row=1,column=1,value='UII')
file.set_cell_value(row=1,column=2,value='Bureau')
file.set_cell_value(row=1,column=3,value='Investment Title')
file.set_cell_value(row=1,column=4,value='Total FY2021 Spending($M)')
file.set_cell_value(row=1,column=5,value='CIO Rating')
file.set_cell_value(row=1,column=6,value='# of Projects')

for i in range(len(UII)):
	file.set_cell_value(row=i+2,column=1,value=UII[i])
	
for i in range(len(bureau)):
	file.set_cell_value(row=i+2,column=2,value=bureau[i])

for i in range(len(Inv_title)):
	file.set_cell_value(row=i+2,column=3,value=Inv_title[i])

for i in range(len(Total_Spend)):
	file.set_cell_value(row=i+2,column=4,value=Total_Spend[i])

for i in range(len(cio)):
	file.set_cell_value(row=i+2,column=5,value=cio[i])

for i in range(len(No_of_proj)):
	file.set_cell_value(row=i+2,column=6,value=No_of_proj[i])
file.save_workbook(filename2)

page_source3=browser.get_source

a=browser.get_webelements('xpath://a')
l1=[]
for i in a:
    l1.append(browser.get_text(i))
for i in l1:
	if '005' not in i:
		l1.remove(i)
l2=[]
for i in l1:
    if '005-' in i:
        l2.append(i)
links=[]
for a in soup2.find_all('a',href=True):
    links.append(a['href'])
print(len(links))
print(links)
ll1=[]
for i in range(3):
    ll1.append(links[i])
for i in ll1:
    up=url+i
    browser.go_to(up)
    time.sleep(10)
    browser.click_link('Download Business Case PDF')
    time.sleep(10)
    browser.go_back()
    browser.go_back()
    time.sleep(10)
    browser.select_from_list_by_label('investments-table-object_length','All')
    time.sleep(15)
for i in glob.glob(curr+'/*.pdf'):
    shutil.move(i,curr+'/output')
    
browser.close_all_browsers()
