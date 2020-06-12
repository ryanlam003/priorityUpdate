from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time

# load the excel spreadsheet with all the values
wb = load_workbook('Task and Priority List.xlsx')
sheet = wb['Sheet1']

# initialize a list of taskIDs and priorities
taskIDList = []
priorityList = []

# populate the taskIDList and priorityList
for columnOfCellObjects in sheet['C2':'C15476']:
    for cellObj in columnOfCellObjects:
        taskIDList.append(cellObj.value)
for columnOfCellObjects in sheet['D2':'D15476']:
    for cellObj in columnOfCellObjects:
        priorityList.append(cellObj.value)

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

# open Task List
driver.get('https://covanta-test.spherasolutions.com/essential-ehs/Compliance/TaskNav.aspx?tab=list&modid=52&objnum=20793&scid=11130')
time.sleep(2)

# CHANGE THE PRIORITY
#for ii in range(0,len(taskIDList)):
# search by ID
actions = ActionChains(driver)
for jj in range(0,8):
    actions.send_keys(Keys.TAB)

actions.send_keys(taskIDList[3])
actions.send_keys(Keys.ENTER)
actions.perform()
time.sleep(4)

# store the current window handle
window_before = driver.window_handles[0]

# select and click the task statement
driver.get('https://covanta-test.spherasolutions.com/essential-ehs/Compliance/TaskSetUpAndResult.aspx?id=14252&vldsiteid=10026&modid=52&ReqTaskIds=&ScenTaskIds=&showclose=yes')

# switch to the new window opened
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)

