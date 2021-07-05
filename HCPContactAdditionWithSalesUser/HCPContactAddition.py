import time
import openpyxl
import xlsxwriter
import WebElementReusability as WER
import ReadWriteDataFromExcel as RWDE
import BrowserElementProperties as BEP
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from pathlib import Path

FilePath = str(Path().resolve()) + r'\Excel Files\UrlsForProject.xlsx'
Sheet = 'Portal Urls'
Url = str(RWDE.ReadData(FilePath, Sheet, 3, 3))

option = Options()
option.add_argument("--disable-infobars")
option.add_argument("start-maximized")
option.add_argument("--disable-extensions")
option.add_experimental_option("prefs", {"profile.default_content_setting_values.notifications": 2 })

driver = webdriver.Chrome(executable_path = str(Path().resolve()) + '\Browser\chromedriver_win32\chromedriver')
driver.maximize_window()
driver.get(Url)
#print(driver.title)

FilePath = str(Path().resolve()) + '\Excel Files\HCPRegistrationFromLighteningGUI.xlsx'
Seconds = 1

#1. This is for SFDC Login

Sheet = 'Login Page Data'
RowCount = RWDE.RowCount(FilePath, Sheet)

for RowIndex in range(2, RowCount + 1):

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#username', 60)
    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 2))

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#password', 60)
    Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 3))

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#rememberUn', 60)
    Element.click()

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.CSS_SELECTOR, '#Login', 60)
    Element.click()

#1. This is for SFDC Home Page

Sheet = 'SFDC Contacts Page Data'
RowCount = RWDE.RowCount(FilePath, Sheet)

for RowIndex in range(2, RowCount + 1):
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Contacts"]', 60)
    driver.execute_script('arguments[0].click();', Element)

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "New"]', 60)
    Element.click()

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Prescribed HCP"]', 60)
    Element.click()

    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Next"]', 60)
    Element.click()

#1. This is for SFDC Registration Page

Sheet = 'SFDC Registration Page Data'
RowCount = RWDE.RowCount(FilePath, Sheet)

for RowIndex in range(2, RowCount + 1):

    time.sleep(Seconds)
    #Salutation
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "None"]', 60)
    if(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 2)) != 'None'):
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 2))
    else:
         Element.send_keys('')

    SalutationArray = ['--None--', 'Mr.', 'Ms.', 'Mrs.', 'Dr.', 'prof.']
    #print(len(SalutationArray))
    for Index in range(0, len(SalutationArray)):
        #print(SalutationArray[Index])
        if(SalutationArray[Index] == str(RWDE.ReadData(FilePath, Sheet, RowIndex, 2))):
            time.sleep(Seconds)
            SalutationElement = '//lightning-base-combobox-item[' + str(Index + 1) + ']'
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, SalutationElement, 60)
            Element.click()
            break

    # First Name
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "First Name"]', 60)
    if(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 3)) != 'None'):
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 3))
    else:
        Element.send_keys('')

    #MiddleName
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Middle Name"]', 60)
    if(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 4)) != 'None'):
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 4))
    else:
        Element.send_keys('')

    #LastName
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Last Name"]', 60)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 5)) != 'None'):
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 5))
    else:
        Element.send_keys('')

    #Suffix
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Suffix"]', 60)
    if(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 6)) != 'None' ):
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 6))
    else:
        Element.send_keys('')

    #Account Name
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//input[@placeholder = "Search Accounts..."]', 60)
    Element.click()

    #HospitalArray = ['costal hospitals', 'App Hospital', 'Standford Hospital']
    #for Index in range(0, len(HospitalArray)):
    #    print(HospitalArray[Index] + ' = ' + str(RWDE.ReadData(FilePath, Sheet, RowIndex, 7)))
    #    if (HospitalArray[Index] == str(RWDE.ReadData(FilePath, Sheet, RowIndex, 7))):
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span/span/span[. = "Standford Hospital"]', 60)
    Element.click()
    #        break

    #Title
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//force-record-layout-item[1]/div/span/slot/slot/force-record-layout-base-input//input', 60)
    driver.execute_script('arguments[0].scrollIntoView(true);', Element)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 8)) != 'None'):
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 8))
    else:
        Element.send_keys('')

    #Email
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//force-record-layout-row[4]/slot/force-record-layout-item[1]//input', 60)
    driver.execute_script('arguments[0].scrollIntoView(true);', Element)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 9)) != 'None'):
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 9))
    else:
        Element.send_keys('')

    #Phone
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//force-record-layout-row[5]/slot/force-record-layout-item[1]//input', 60)
    driver.execute_script('arguments[0].scrollIntoView(true);', Element)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 10)) != 'None'):
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 10))
    else:
        Element.send_keys('')

    #Mobile
    time.sleep(Seconds)
    Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//force-record-layout-row[6]/slot/force-record-layout-item[1]//input', 60)
    driver.execute_script('arguments[0].scrollIntoView(true);', Element)
    if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 11)) != 'None'):
        Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 11))
    else:
        Element.send_keys('')

    time.sleep(2)
    ErrorElementAvailable = WER.check_exists_by_xpath(driver, '//force-record-edit-error-header/lightning-button-icon//lightning-primitive-icon')
    if (ErrorElementAvailable == True):
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//force-dedupe-content/div', 60)
        RWDE.WriteData(FilePath, Sheet, RowIndex, 23, 'Hold')
        RWDE.WriteData(FilePath, Sheet, RowIndex, 24, Element.text)

        time.sleep(3)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//lightning-button/button[. = "Cancel"]', 60)
        Element.click()

        #print(RowCount)
        if(RowIndex != RowCount):
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "New"]', 60)
            Element.click()

            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Prescribed HCP"]', 60)
            Element.click()

            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Next"]', 60)
            Element.click()
        else:
            # SalesUser Image Button
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button/div/span//span', 60)
            Element.click()

            # LogOut
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//a[. = "Log Out"]', 60)
            Element.click()
    elif (ErrorElementAvailable == False):
        #Fax
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//force-record-layout-item[2]/div/span/slot/slot/lightning-input//input', 60)
        driver.execute_script('arguments[0].scrollIntoView(true);', Element)
        if(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 12)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 12))
        else:
            Element.send_keys('')

        #NPI Number
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//force-record-layout-item[2]/div/span/slot/slot/force-record-layout-base-input//input', 60)
        driver.execute_script('arguments[0].scrollIntoView(true);', Element)
        if(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 13)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 13))
        else:
            Element.send_keys('')

        #NPI Number Verified
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//force-form-picklist//input', 60)
        driver.execute_script('arguments[0].scrollIntoView(true);', Element)
        Element.click()
        #SelectElement = Select(Element)
        #SelectElement.select_by_visible_text(RWDE.ReadData(FilePath, Sheet, RowIndex, 14))

        #NPI Number Verified Element
        time.sleep(Seconds)
        #print('//span[2][. = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex, 14)) +'"]')
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[2][. = "' + str(RWDE.ReadData(FilePath, Sheet, RowIndex, 14)) +'"]', 60)
        driver.execute_script('arguments[0].scrollIntoView(true);', Element)
        Element.click()

        #Site Admin
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span/input', 60)
        driver.execute_script('arguments[0].scrollIntoView(true);', Element)
        if(str(RWDE.ReadData(FilePath, Sheet, RowIndex, 15)) == 'Yes'):
            Element.click()

        #Mailing Address
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//lightning-lookup-address//input', 60)
        driver.execute_script('arguments[0].scrollIntoView(true);', Element)
        Element.click()

        #Mailing Street
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//textarea', 60)
        driver.execute_script('arguments[0].scrollIntoView(true);', Element)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 18) != 'None')):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 18))
        else:
            Element.send_keys('')

        #Mailing City
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//lightning-input-address/fieldset/div/div/div[4]/lightning-input//input', 60)
        driver.execute_script('arguments[0].scrollIntoView(true);', Element)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 19)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 19))
        else:
            Element.send_keys('')

        #Mailing State/Province
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[4]/lightning-picklist//input', 60)
        driver.execute_script('arguments[0].scrollIntoView(true);', Element)
        Element.click()

        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span/span[. = "California"]', 60)
        Element.click()

        # Mailing City
        time.sleep(Seconds)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//lightning-input-address/fieldset/div/div/div[5]//input', 60)
        driver.execute_script('arguments[0].scrollIntoView(true);', Element)
        if (str(RWDE.ReadData(FilePath, Sheet, RowIndex, 21)) != 'None'):
            Element.send_keys(RWDE.ReadData(FilePath, Sheet, RowIndex, 21))
        else:
            Element.send_keys('')

        #Save & New
        time.sleep(1)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//lightning-button/button[. = "Save"]', 60)
        Element.click()

        #DownTriangle
        time.sleep(3)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//runtime_platform_actions-actions-ribbon//lightning-primitive-icon', 60)
        Element.click()

        #SetupPortalUser
        time.sleep(3)
        Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Setup Portal User"]', 60)
        Element.click()

        RWDE.WriteData(FilePath, Sheet, RowIndex, 23, 'Passed')
        if(RowIndex != RowCount):

            #Contacts
            time.sleep(20)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Contacts"]', 60)
            driver.execute_script('arguments[0].click();', Element)

            #New
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//div[. = "New"]', 60)
            Element.click()

            #PrescribedHCP
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//span[. = "Prescribed HCP"]', 60)
            Element.click()

            #Next
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button[. = "Next"]', 60)
            Element.click()
        else:
            #SalesUser Image Button
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//button/div/span//span', 60)
            Element.click()

            #LogOut
            time.sleep(Seconds)
            Element = BEP.WebElement(driver, EC.presence_of_element_located, By.XPATH, '//a[. = "Log Out"]', 60)
            Element.click()

