from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
from selenium.webdriver.support.select import Select


# load the excel spreadsheet with all the values
wb = load_workbook('Task, Priority, and Site List.xlsx')
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
    str_siteIDList.append(str(siteIDList[ii]))
    taskStatementURLList.append('https://covanta-test.spherasolutions.com/essential-ehs/Compliance/TaskSetUpAndResult.aspx?id='
                                + str_taskIDList[ii] + '&vldsiteid=' + str_siteIDList[ii] + '&modid=52&ReqTaskIds=&ScenTaskIds=&showclose=yes')

# loop through all tasks
for taskCounter in range(15475,15464,-1):

    # navigate to the task statement URL
    driver.get(taskStatementURLList[taskCounter-2])

    # select the Task Priority; if High->switch to Tier I,  else (it is medium or low)->switch to Tier II
    actions3 = ActionChains(driver)
    for ll in range(0,28):
        actions3.send_keys(Keys.TAB)

    if priorityList[taskCounter-2] == 'High':
        for mm in range(0,3):
            actions3.send_keys(Keys.ARROW_DOWN)
    elif priorityList[taskCounter-2] == 'Medium':
        for nn in range(0,2):
            actions3.send_keys(Keys.ARROW_DOWN)
    elif priorityList[taskCounter-2] == 'Low':
        for nn in range(0,3):
            actions3.send_keys(Keys.ARROW_DOWN)

    actions3.perform()

    # save the change
    actions4 = ActionChains(driver)
    for oo in range(0,14):
        actions4.send_keys(Keys.TAB)
    actions4.send_keys(Keys.ENTER)
    actions4.perform()
    time.sleep(2)


