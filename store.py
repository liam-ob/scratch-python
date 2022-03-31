import os.path
import shutil

from selenium import webdriver
from selenium.webdriver import FirefoxOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
import time

import pandas

def empty_folder(path):
    folder = path
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


def login(driver):
    company = "atteris"
    company_element = driver.find_element(By.ID, "CompanyNameTextBox")
    company_element.send_keys(company)
    company_element.send_keys(Keys.ENTER)
    user = input("Username: ")
    user_element = driver.find_element(By.ID, "LoginNameTextBox")
    user_element.send_keys(user)
    passw = input("Password: ")
    passw_element = driver.find_element(By.ID, "PasswordTextBox")
    passw_element.send_keys(passw)
    passw_element.send_keys(Keys.ENTER)


def navigate(driver):
    time.sleep(5)
    driver.get('https://sp1-reports.replicon.com/r/reports/list')
    # reports = driver.find_element(By.LINK_TEXT, "Reports")
    # reports.click()
    time.sleep(5)
    ctr_report = driver.find_element(By.LINK_TEXT, "CTR Input Template- Current")
    ctr_report.click()
    return driver


def get_open_projects():
    data_file = pandas.read_excel(r"P:\1 - Project & Proposal Number List.xlsx", usecols="A,I")
    data_fields = pandas.DataFrame(data_file)
    open_projects = []
    try:
        data_fields = data_fields.dropna()
        for index, row in data_fields.iterrows():
            if row['Unnamed: 8'] == 'YES':
                open_projects += [row['Unnamed: 0']]
        return open_projects
    except Exception as e:
        print(e.args)

# works but very slow
def find_and_download_all_ctr_reports(driver):
    open_projects = get_open_projects()
    for i in open_projects:
        time.sleep(5)

        dropdown = driver.find_element(By.ID, "CB_FC_3Header")
        dropdown.click()
        time.sleep(2)

        dropdown2 = driver.find_element(By.ID, "select2-CB_FC_3_0_Selector-container")
        dropdown2.click()
        time.sleep(2)

        search_element = driver.find_element(By.CLASS_NAME, "select2-search__field")
        search_element.send_keys(i)
        time.sleep(2)
        search_element.send_keys(Keys.ENTER)
        time.sleep(5)

        download_button = driver.find_element(By.ID, "CB_Excel")
        download_button.click()
        time.sleep(5)

        os.rename(r'.\Downloads\CTR Input Template- Current 2022-02-16.xml', r'.\Downloads\CTR_' + str(i) + ".xml")


def select_and_download_ctr_reports(driver):
    pass


def pass_to_p_drive():

    pass


def main():
    path = str(os.getcwd()) + r"\Downloads"
    path = str(path)
    profile = FirefoxOptions()
    profile.set_preference("browser.download.panel.shown", False)
    profile.set_preference("browser.helperApps.neverAsk.openFile", 'application/vnd.ms-excel')
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", 'application/vnd.ms-excel')
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.dir", path)

    service = Service(executable_path=GeckoDriverManager().install())
    print("declared service")
    driver = webdriver.Firefox(service=service, options=profile)
    print("successfully loaded download directory")
    try:
        empty_folder(path)
        driver.get("https://login.replicon.com/DefaultV2.aspx")
        login(driver)
        navigate(driver)
        find_and_download_all_ctr_reports(driver)
        pass_to_p_drive()
        print("all done!")
        driver.quit()
    except Exception as e:
        print(e.args)
        print("invalid url")

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
