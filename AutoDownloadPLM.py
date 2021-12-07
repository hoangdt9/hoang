import sys, os, time
import ReadDataFromExcel as DataEXL
import pandas as pd
import IEHelper
from datetime import date
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoAlertPresentException
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from winreg import *


class AutoDownloadPLM:
    def __init__(self, user_name, pass_word, pass_ie):
        self.user_name = user_name
        self.pass_word = pass_word
        self.pass_word_ie = pass_ie
        self.checkLoginKnox()

        self.excel_file_name = "DEFECT_LIST_Today_Basic.xls"
        with OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
            self.dir_download = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0] + "\\"
        self.dir_sample_input = os.getcwd() + "\\sample input\\"
        self.driver = webdriver.Chrome()
        self.driver.maximize_window()

        self.runDownload()

    def loadCompleted(self, locator, timeout):
        """ check website load complete """
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, locator))
            )
            return True
        except TimeoutException:
            return False

    def clickElement(self, xpath_element):
        """ find element on website then click """
        try:
            if self.loadCompleted(xpath_element, 50):
                element = self.driver.find_element_by_xpath(xpath_element)
                element.click()
        except NoSuchElementException:
            print("can not find element:", xpath_element)
        except Exception:
            print("can not click try perform ")
            time.sleep(10)
            ex_element = self.driver.find_element_by_xpath(xpath_element)
            ActionChains(self.driver).click(ex_element).perform()

    def switchFrame(self, iframe):
        """ switch to other frame by xPath """
        try:
            frame = self.driver.find_element_by_xpath(iframe)
            self.driver.switch_to.frame(frame)
        except NoSuchElementException:
            print("can not find element:", iframe)

    def downloadTotalIssueDetail(self):
        """ download list issue detail """
        issue_open_link = '//tbody//tr/td/a[@href="javascript:lfn_popup(\'\',\'\',\'\',\'\',\'ALL\',\'\')"]'
        #'''javascript:lfn_popup('','RESOLVE','','','ALL','')'''
        if self.loadCompleted(issue_open_link, 20):
            open_issue = self.driver.find_element_by_xpath(issue_open_link)
            open_issue_download = ActionChains(self.driver).click(open_issue)
            open_issue_download.perform()
            self.clickElement(issue_open_link)
            time.sleep(5)

            self.driver.switch_to.window(self.driver.window_handles[-1])
            result_list_issue = "//div[@class='container']//iframe[@name='ResultListIframe']"
            result_load = self.loadCompleted(result_list_issue, 20)
            print("result list issue: ", result_load)

            # click to download total list issue detail
            self.switchFrame(result_list_issue)
            # delete old file download
            self.deleteAllFiles()
            time.sleep(8)
            self.download_issue_comment('issue')
            time.sleep(3)
            self.download_issue_comment('comment')

            # merge two file comment and issue
            self.mergeFileIssueComment()

    def runDownload(self):
        """ start download PLM """
        link_plm_issue = 'http://splm.sec.samsung.net//wl/tqm/statistics/getDefectByUser.do?fromPlmMainMenu=true'
        self.driver.get(link_plm_issue)
        if self.driver.current_url == link_plm_issue:
            print("Knox has been login on Chrome")
        else:
            user = self.driver.find_element_by_id('userNameInput')
            user.send_keys(self.user_name)
            password = self.driver.find_element_by_id('passwordInput')
            password.send_keys(self.pass_word_ie)
            button = self.driver.find_element_by_id('submitButton')
            button.click()
        time.sleep(5)
        self.selectOneMonth()
        self.downloadTotalIssueDetail()
        self.driver.quit()

    def mergeFileIssueComment(self):
        """ merge two file issue and comment """
        print('start merge file')

        dir_file_cmt, sheet_name_cmt = self.dir_file_download('comment')
        dir_file_issues, sheet_name_issues = self.dir_file_download('issues')

        comment_history, length = DataEXL.read_excel_file(r'%s' % dir_file_cmt, sheet_name_cmt)
        issues_data, length = DataEXL.read_excel_file(r'%s' % dir_file_issues, sheet_name_issues)
        # merge issue data and comment history
        all_data = pd.merge(issues_data, comment_history, how='left', left_on='Case Code', right_on='Case Code.')

        # delete unused or duplicated columns
        columns = ['Project Name', 'Model Name', 'Case Code.', 'Type']
        all_data.drop(columns, inplace=True, axis=1)

        # wrtie to dest file. But first remove existing file
        DEFECT_ISSUE = self.dir_sample_input + self.excel_file_name
        if os.path.exists(DEFECT_ISSUE):
            os.remove(DEFECT_ISSUE)

        # save to file
        writer = pd.ExcelWriter(DEFECT_ISSUE)
        all_data.to_excel(writer, sheet_name='DEFECT', startrow=2, index=False)
        writer.save()

    def selectOneMonth(self):
        """ select download 1 month ago """
        select = Select(self.driver.find_element_by_id("memberGroupId"))
        btn_search = "//td/span/span/button[text()='Search']"
        for option in select.options:
            if option.text == "CP":
                option.click()
        time.sleep(3)
        print("click 1 month")
        self.driver.execute_script("javascript:setRegDate(1,'MONTH');")
        time.sleep(5)
        self.clickElement(btn_search)
        time.sleep(10)

    def deleteAllFiles(self):
        """ delete all file in folder """
        for the_file in os.listdir(self.dir_download):
            file_path = os.path.join(self.dir_download, the_file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(e)

    def dir_file_download(self, file_type='issues'):
        """ return file in download folder after finish """
        today = date.today().strftime("%Y%m%d")
        file_issue = "DEFECT_LIST_" + today + "_Basic.xls"
        file_comment = "DEFECT_LIST_" + today + "_Comment_History.xls"
        sheet_name_cmt = "DEFECT_LIST_" + today + "_Comment_Hi"
        sheet_name_issues = "DEFECT"
        if file_type == 'issues':
            return self.dir_download + file_issue, sheet_name_issues
        else:
            return self.dir_download + file_comment, sheet_name_cmt

    def click_download_file(self, id_btn_download):
        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, id_btn_download)))
            download_button = self.driver.find_element_by_id(id_btn_download)
            download_button.click()
            print("click:" + id_btn_download)
        except Exception:
            print("can't find %s, try run javaScript" % id_btn_download)
            if id_btn_download == "btn_downloadExcel":
                self.driver.execute_script("downloadExcel(this);")
            else:
                self.driver.execute_script("downloadCommentExcel(this);")

    def download_issue_comment(self, file_type='issue'):
        """ start download file issue and file comment """
        if file_type == 'issue':
            file_name, sheet_name = self.dir_file_download('issues')
            self.click_download_file("btn_downloadExcel")
        else:
            file_name, sheet_name = self.dir_file_download('comment')
            self.click_download_file("btn_downloadCommentExcel")

        time.sleep(4)
        waiting_time = 15
        while not self.isAlertPresent() and waiting_time:
            waiting_time = waiting_time - 1
            if waiting_time == 10:
                print("wait alert long time, try click again ")
                if file_type == 'issue':
                    self.click_download_file("btn_downloadExcel")
                else:
                    self.click_download_file("btn_downloadCommentExcel")
            elif not waiting_time:
                print("wait allow alert timeout. EXIT !!!")
                sys.exit(1)
            time.sleep(1)

        print("waiting download done ...")
        waiting_time = 110  # waiting download in 5 minute
        while not os.path.exists(file_name):
            time.sleep(6)
            waiting_time = waiting_time - 1
            if not waiting_time:
                print("wait download time out. EXIT!")
                sys.exit(1)

        if os.path.exists(file_name):
            print("download %s done!" % file_name)

    def isAlertPresent(self):
        """ check click allow download """
        try:
            self.driver.switch_to.alert.accept()
            print("Allow download from Alert!")
            return True
        except NoAlertPresentException as e:
            print("wait utils alert present")
            return False

    def checkLoginKnox(self):
        """ auto login Know on IE """
        IEHelper.set_zoom_100()
        driverIE = webdriver.Ie()
        driverIE.minimize_window()
        url_login_done = 'http://kr2.samsung.net/portal/desktop/main.do'

        driverIE.get("http://samsung.net/")
        time.sleep(10)
        if driverIE.current_url == url_login_done:
            print("Knox has been login on IE")
            """self.loginMobihub(driverIE) - không check MRH"""
        else:
            user = driverIE.find_element_by_id('USERID')
            user.clear()
            user.send_keys(self.user_name)
            password = driverIE.find_element_by_id('USERPASSWORD')
            password.send_keys(self.pass_word)
            button = driverIE.find_element_by_class_name('btnLogin')
            button.click()
            while not driverIE.current_url == url_login_done:
                time.sleep(1)
            print("login Knox on IE done")
            """self.loginMobihub(driverIE) - không check MRH"""

    # Login to mobihub
    def loginMobihub(self, driver):
        driver.get("http://mobilerndhub.sec.samsung.net/hub/")
        try:
            user = driver.find_element_by_id('login_user_id')
            user.clear()
            user.send_keys(self.user_name)
            password = driver.find_element_by_id('login_user_password')
            password.send_keys(self.pass_word)
            button = driver.find_element_by_id('login_main_form_pw_confirm_a')
            button.click()
            time.sleep(5)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'home_main_container')))
        except TimeoutException:
            print("Login Mobihub timeout")
        except:
            print("has been login Mobihub - can't find element")
