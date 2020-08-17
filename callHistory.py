#! python3
# -*- coding: utf-8 -*-
"""
Created on Sat Feb 29 18:02:14 2020

@author: Jono
"""
import datetime, time, os, bs4, re, openpyxl
from selenium import webdriver

os.makedirs('public_mobile',exist_ok=True) #makes public mobile folder


publicSite = 'https://selfserve.publicmobile.ca/'
browser = webdriver.Firefox()
#browser = webdriver.Chrome(executable_path='LOCATION OF WEBDRIVER')
browser.get(publicSite)
print('Please input username and password on web page then press enter')
input()
# Find End date of current period
homepage = browser.page_source
filename = os.path.join(os.getcwd(),'public_mobile',('homepage'+'.html')) # creates files for homepage
File = open(filename,'w')
File.write(homepage) #writes the content to the file
File.close()

# opens the file for parsing to extract useful data
exampleFile = open(filename,'rb')
Soup = bs4.BeautifulSoup(exampleFile,'html.parser')

payDate = Soup.select('#FullContent_DashboardContent_Overview_Prepaid_activeAccountUlPnP > li:nth-child(2)')
payDate = payDate[0].getText()

# Use regex to find the date of next payment (so the end date of current term)
dateRegex = re.compile(r'(\D{3} \d{1,2}, 20\d{2})')
mo = dateRegex.search(payDate)
payDate = datetime.datetime.strptime(mo.group(), '%b %d, %Y')

# From pay date get the end date of last period
endDate = payDate - datetime.timedelta(days=30)

# Get duration into number of days
duration = int(input('Previous 1, 2 or 3 months: '))*30
deltaTime = datetime.timedelta(days=duration)

# Get Starting Date
startDate = endDate - deltaTime

# Convert dates to readable formats to input into fields
startInput = startDate.strftime('%b %d %Y')
endInput = endDate.strftime('%b %d %Y')

# Go to Call history page
browser.get('https://selfserve.publicmobile.ca/Overview/plan-and-Add-ons/call-history/')

# Input Information in fields
startDateButton = browser.find_element_by_id('UseDateRangeRadioButton')
startDateButton.click()
startdate = browser.find_element_by_id('startdate')

startdate.send_keys(startInput)
time.sleep(.25)

# Input final date
endDate = browser.find_element_by_id('enddate').send_keys(endInput)
enterButton = browser.find_element_by_id('FullContent_DashboardContent_ViewCallHistoryButton')
enterButton.click()

print('Opening Workbook...')
wb = openpyxl.load_workbook('callHistory.xlsx')
sheetTitle = startInput[:3]+startInput[6:]
wb.create_sheet(title=sheetTitle)
sheet = wb[sheetTitle]

# Create a While loop to cycle through each page
pagenum = 0
lastFirstRow = None
firstRow = None
active = True

while active == True:
  pagenum += 1
  page = browser.page_source
  filename = os.path.join(os.getcwd(),'public_mobile',('page'+str(pagenum)+'.html')) # creates files for the page
  File = open(filename,'w')
  File.write(page) #writes the content to the file
  File.close()
  
  # opens the file for parsing to extract useful data
  print(f'Reading page {pagenum}')
  exampleFile = open(filename,'rb')
  Soup = bs4.BeautifulSoup(exampleFile,'html.parser')
  rows = Soup.select('tr')
  firstRow = rows[1]
  
  # to ensure a new chart loads if the data is the same as the
  # previous data then the program is stalling (Public Mobile Website is super slow)
  stall = 0
  while firstRow == lastFirstRow:
    stall += 1
    if stall == 8:
      browser.find_element_by_id('FullContent_DashboardContent_gvCallHistory_gvPagerTemplate_pagerNextPage').click()
      time.sleep(10)
    print('Page is loading slow... Temporarily Pausing Operation')
    time.sleep(2.5)
    page = browser.page_source
    filename = os.path.join(os.getcwd(),'public_mobile',('page'+str(pagenum)+'.html')) # creates files for the page
    File = open(filename,'w')
    File.write(page) #writes the content to the file
    File.close()
    
    # opens the file for parsing to extract useful data
    exampleFile = open(filename,'rb')
    Soup = bs4.BeautifulSoup(exampleFile,'html.parser')
    rows = Soup.select('tr')
    firstRow = rows[1]
  
  lastFirstRow = firstRow # at the end of this iteration it will autmatically 
  # trip the above while loop if new page doesnt load  
  
  # Append data below current
  xcelRow = sheet.max_row + 1
  
  print(f'Writing page {pagenum}')
  
  for row in rows[1:-2]:
    #pageTimeStart = time.perf_counter()
    xcelColumn = 1
    data = row.select('td')
    for cellItem in data:
      resultValue = cellItem.getText()
      sheet.cell(row=xcelRow,column=xcelColumn).value = resultValue
      xcelColumn += 1
    xcelRow += 1  
  
  # Find the next page button and see if its clickable
  nextButton = Soup.select('#FullContent_DashboardContent_gvCallHistory_gvPagerTemplate_pagerLastPage')
  if nextButton[0].get('href') is None:
    active = False # Next button is not clickable therefor on last page
                           
  # Go to next page
  else:
    browser.find_element_by_id('FullContent_DashboardContent_gvCallHistory_gvPagerTemplate_pagerNextPage').click()
    time.sleep(.8)
#  pageTime = time.perf_counter() - pageTimeStart
#  print(f'{pageTime} for page {pagenum}')
#  if pageTime <.2:
#    time.sleep(1.5)
  # Delete File
print('Done writing logs')
print('Tabulating final totals')  

fSheet = wb['Totals']

nextRow = fSheet.max_row
periodDate = sheetTitle
incomingDuration = datetime.timedelta()
outDuration = datetime.timedelta()
tollFree = datetime.timedelta()
dataUsage = 0
inText = 0
outText = 0
sms = 0
extraCharges = 0

for xrow in range(2,sheet.max_row + 1):
  identifier = sheet.cell(row=xrow,column=2).value
  if  identifier == 'Web': #Column B
    mb = sheet.cell(row=xrow,column=8).value
    dataUsage += float(mb[:-3])
  elif identifier == 'Incoming text':
    inText += 1
  elif identifier == 'Outgoing Text':
    outText += 1
  elif identifier == 'Data Event':
    sms += 1
  elif identifier == 'Incoming Call':
    [hr,mm,ss] = sheet.cell(row=xrow,column=7).value.split(':') #separates the time stamp
    incomingDuration += datetime.timedelta(hours=int(hr),minutes=int(mm),seconds=int(ss))
  elif identifier == 'Outgoing Call':
    phoneNum = sheet.cell(row=xrow,column=4).value
    if phoneNum.startswith('800') or phoneNum.startswith('866') or phoneNum.startswith('888'):
      tollFree += datetime.timedelta(hours=int(hr),minutes=int(mm),seconds=int(ss))
    else:
      [hr,mm,ss] = sheet.cell(row=xrow,column=7).value.split(':') #separates the time stamp
      outDuration += datetime.timedelta(hours=int(hr),minutes=int(mm),seconds=int(ss))    
  else:
    print(f'Unidentified type seen: {identifier}')
    continue
  extraCharges += float(sheet.cell(row=xrow,column=10).value[1:])

  
finalTotals = [periodDate,str(incomingDuration),str(outDuration),str(tollFree),round(dataUsage,2),
               inText,outText,sms,round(extraCharges,2)]
for i in range(1,len(finalTotals)+1):
  fSheet.cell(row=nextRow+1,column=i).value = finalTotals[i-1]    

print('Saving file')
wb.save('CallHistory_updated.xlsx') 

# TO DO: Delete folder created to recycle
