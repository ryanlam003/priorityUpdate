from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time

# load the excel spreadsheet with all the values
wb = load_workbook('Task Priority and Site List.xlsx')
sheet = wb['Sheet1']

# initialize a list of taskIDs and priorities and siteIDs
taskIDList = []
priorityList = []
siteIDList = []

# populate the taskIDList and priorityList
for columnOfCellObjects in sheet['C2':'C15476']:
    for cellObj in columnOfCellObjects:
        taskIDList.append(cellObj.value)
for columnOfCellObjects in sheet['D2':'D15476']:
    for cellObj in columnOfCellObjects:
        priorityList.append(cellObj.value)
for columnOfCellObjects in sheet['F2':'F15476']:
    for cellObj in columnOfCellObjects:
        siteIDList.append(cellObj.value)

# using chrome to access web
driver = webdriver.Chrome()

# open the website
driver.get('https://covanta-test.spherasolutions.com/essential-ehs')

# locate the id and password
id_box = driver.find_element_by_name('TextBoxUserID')
pass_box = driver.find_element_by_name('TextBoxPasswd')

# send login information
id_box.send_keys('rlam')
pass_box.send_keys('1Ringtorulethemall.*')

# click login
login_button = driver.find_element_by_name('Button1')
login_button.click()

time.sleep(2)

# prepare all of the task statement URLs
str_taskIDList = []
str_priorityList = []
str_siteIDList = []
taskStatementURLList = []
for ii in range(0,len(taskIDList)):
    str_taskIDList.append(str(taskIDList[ii]))
    str_priorityList.append(str(priorityList[ii]))
    str_siteIDList.append(str(siteIDList))
    taskStatementURLList.append('https://covanta-test.spherasolutions.com/essential-ehs/Compliance/TaskSetUpAndResult.aspx?id='
                                + str_taskIDList[ii] + '&vldsiteid=' + str_siteIDList[ii] + '&modid=52&ReqTaskIds=&ScenTaskIds=&showclose=yes')

# navigate to the first task statement URL
driver.get(taskStatementURLList[0])


