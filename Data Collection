from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
import urllib
from time import sleep

username = '' # Insert Email
password = '' # Insert Password

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get('https://search.ebscohost.com/login.aspx?authtype=ip,sso&custid=s1098328&profile=ehost&defaultdb=mdc&groupid=main')

def Click(id,type):
    if type == 'xpath':
        button = driver.find_element_by_xpath(id)
    elif type == 'id':
        button = driver.find_element_by_id(id)
    elif type == 'class':
        button = driver.find_element_by_class_name(id)
    button.click()



def startUp():

    # Student Enter
    driver.implicitly_wait(20)
    Click('/html/body/div[1]/div[2]/div/div/div[2]/div/div[1]/div[3]','xpath')

    # Enter Email
    driver.implicitly_wait(20)
    usernameBar = driver.find_element_by_id('i0116')
    usernameBar.send_keys(username)

    # Enter Password
    passwordBar = driver.find_element_by_id('i0118')
    passwordBar.send_keys(password)

    sleep(3)

    # Click Next
    Click('idSIButton9','id')

    sleep(3)

    # Click Sign In
    Click('//*[@id="idSIButton9"]','xpath')

    sleep(3)

    # Click Yes
    Click('idSIButton9','id')

    sleep(10)

    # Click Advanced Search
    Click('advanced','id')

    sleep(3)

    # Enter Search Request
    search1 = driver.find_element_by_id('Searchbox1')
    search1.send_keys('') # Insert the Search you wish to request (example: ( eeg or electroencephalogram or electroencephalography ) AND An?esthe* NOT animals)

    # Click Search Button
    Click('SearchButton','id')

    sleep(4)

    # Click Page Options
    Click('//*[@id="lnkPageOptions"]','xpath')

    # Click 50 Pages
    Click('//*[@id="pageOptions"]/li[3]/ul/li[6]/a','xpath')

    sleep(1)

    print('full text')
    # Click Full Text
    Click('common_FT','id')

    sleep(2)

    print('abstract')
    # Click Abstract Available
    Click('common_AA1','id')

    sleep(2)

    print("Language")
    # Change Language
    Click('common_LA1','id')
    sleep(2)

def saving():
    sleep(1)
    Click('//*[@id="lnkFolder"]', 'xpath')

    sleep(1)
    Click('//*[@id="ctl00_ctl00_Column2_Column2_btnDelivery_lnkExport"]', 'xpath')

    sleep(1)

    Click('//*[@id="ctl00_ctl00_MainContentArea_MainContentArea_ctl00_radFormatCsv"]', 'xpath')

    sleep(1)

    Click('//*[@id="ctl00_ctl00_MainContentArea_MainContentArea_ctl00_chkRemoveFromFolder"]', 'xpath')

    Click('//*[@id="ctl00_ctl00_MainContentArea_MainContentArea_ctl00_btnSubmit"]', 'xpath')

def addFiftyItems():
    sleep(0.5)
    Click('//*[@id="ctl00_ctl00_MainContentArea_MainContentArea_lnkAlertSaveShare"]','xpath')
    sleep(2)
    Click('//*[@id="lnkAddAll"]','xpath')

startUp()

page = 1
while True:
    sleep(2)
    addFiftyItems()
    print(f'Page {page} done')

    if page%4 == 0:
        saving()
        sleep(1)
        Click('//*[@id="ctl00_ctl00_FindField_FindField_btnBack"]','xpath')
        sleep(1)
        Click('//*[@id="ctl00_ctl00_FindField_FindField_ctl00_btnBack_lnkBack"]','xpath')

    page += 1

    try:
        Click('//*[@id="ctl00_ctl00_MainContentArea_MainContentArea_bottomMultiPage_lnkNext"]','xpath')
    except selenium.common.exceptions.NoSuchElementException:
        saving()
        break
