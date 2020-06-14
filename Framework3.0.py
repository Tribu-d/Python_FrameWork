from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pandas as pd
from selenium.common.exceptions import NoSuchElementException, UnexpectedAlertPresentException, \
    ElementNotInteractableException, ElementClickInterceptedException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import os
from datetime import datetime
import time
from datetime import date
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import excel2json
import asyncio
from openpyxl import Workbook
import pyodbc
import csv
# Import smtplib for the actual sending function
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import mimetypes
import email.mime.application

os.chdir("D:")
print(os.getcwd())

'''FILE_PATH = "D:\hybridframework - PNG.xlsx"
xl = pd.ExcelFile(FILE_PATH)
df = xl.parse('Upload')'''
# print(df)
df = pd.concat(pd.read_excel('hybridframework - PNG - Copy.xlsx', sheet_name=None), ignore_index=True, sort=False)

chrome_options = webdriver.ChromeOptions()
prefs = {"profile.default_content_setting_values.notifications": 2}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(chrome_options=chrome_options,
                          executable_path="D:/Vishnu/Selenium browser drivers/chromedriver_win32/chromedriver.exe")
# database connection
conn = pyodbc.connect('Driver={ODBC Driver 13 for SQL Server};'
                      'Server= qadb2016.ivy-support.com;'
                      'Database=IvyDms_PngIndia01;'
                      'UID=IvyDMS_PngIndia_QAUser;'
                      'PWD=IvyDMS_PngIndia_QAUserPWDMJSAQYRTE;'
                      )

email_user_from = 'vishnuu.0.d@gmail.com'
email_pass = '8008609957@Dv'
email_user_to = 'd.vishnu@ivymobility.com'
smtp_ssl_host = 'smtp.gmail.com'  # smtp.mail.yahoo.com
smtp_ssl_port = 465
s = smtplib.SMTP_SSL(smtp_ssl_host, smtp_ssl_port)
s.login(email_user_from, email_pass)

msg = MIMEMultipart()
msg['Subject'] = 'Failed case'
msg['From'] = email_user_from
msg['To'] = email_user_to


def timestamp():
    now = time.time()
    localtime = time.localtime(now)
    milliseconds = '%03d' % int((now - int(now)) * 1000)
    return time.strftime('%H%M%S', localtime) + milliseconds


print(timestamp())
t1 = datetime.now().second
print(t1)
'''with open('Mycsv.csv', 'w', newline='')as f:
    fieldnames = ['column1', 'column2', 'column3', 'column4', 'column5']
    thewriters = csv.DictWriter(f, fieldnames=fieldnames)
    thewriters.writeheader()'''

# itertuples
# for row in range(0, 100):
for row in df.itertuples():
    message = row.Keyword, row.Xpath, row.Testdata, 'Fail', timestamp()
    msg = MIMEMultipart()
    msg['Subject'] = message
    msg['From'] = email_user_from
    msg['To'] = email_user_to

    txt = MIMEText('PFA')
    msg.attach(txt)
    filename = 'D:/Screenshots' + '/' + timestamp() + 'Fail' + '.' + 'png'

    if row.Keyword == 'url':
        try:
            driver.get(row.Testdata)
            driver.maximize_window()
        except:
            print(message)
            driver.save_screenshot(filename)
            # filename = 'D:/Screenshots/010052493Pass.png'  # path to file
            fo = open(filename, 'rb')
            attach = email.mime.application.MIMEApplication(fo.read(), _subtype="png")
            fo.close()
            attach.add_header('Content-Disposition', 'attachment', filename=filename)
            msg.attach(attach)
            s.send_message(msg)
            s.quit()
            print('Email sent')

        else:
            print(row.Keyword, row.Xpath, row.Testdata, 'Pass', timestamp())
            driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')


    elif row.Keyword == 'Textbox':
        time.sleep(5)
        try:
            # element = WebDriverWait(driver,20).until(EC.element_to_be_clickable(row.Xpath))
            # element.send_Keys(row.Testdata)

            driver.find_element_by_xpath(row.Xpath).send_keys(row.Testdata)
        except:
            print(message)
            driver.save_screenshot(filename)
            # filename = 'D:/Screenshots/010052493Pass.png'  # path to file
            fo = open(filename, 'rb')
            attach = email.mime.application.MIMEApplication(fo.read(), _subtype="png")
            fo.close()
            attach.add_header('Content-Disposition', 'attachment', filename=filename)
            msg.attach(attach)
            s.send_message(msg)
            s.quit()
            print('Email sent')

        else:
            print(row.Keyword, row.Xpath, row.Testdata, 'Pass', timestamp())
            driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')


    elif row.Keyword == 'button':
        time.sleep(5)
        try:
            driver.find_element_by_xpath(row.Xpath).click()
        except NoSuchElementException:
            print(message)
            driver.save_screenshot(filename)
            # filename = 'D:/Screenshots/010052493Pass.png'  # path to file
            fo = open(filename, 'rb')
            attach = email.mime.application.MIMEApplication(fo.read(), _subtype="png")
            fo.close()
            attach.add_header('Content-Disposition', 'attachment', filename=filename)
            msg.attach(attach)
            s.send_message(msg)
            s.quit()
            print('Email sent')

        else:
            print(row.Keyword, row.Xpath, row.Testdata, 'Pass', timestamp())
            driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')



    elif row.Keyword == 'Menu':
        time.sleep(4)
        try:
            scroll_Menu = driver.find_element_by_xpath(row.Xpath)
            driver.execute_script("arguments[0].click();", scroll_Menu)
        except ElementNotInteractableException:
            print(message)
            driver.save_screenshot(filename)
            # filename = 'D:/Screenshots/010052493Pass.png'  # path to file
            fo = open(filename, 'rb')
            attach = email.mime.application.MIMEApplication(fo.read(), _subtype="png")
            fo.close()
            attach.add_header('Content-Disposition', 'attachment', filename=filename)
            msg.attach(attach)
            s.send_message(msg)
            s.quit()
            print('Email sent')


        else:
            print(row.Keyword, row.Xpath, row.Testdata, 'Pass', timestamp())
            driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')



    elif row.Keyword == 'Switch_frame_in':
        time.sleep(2)
        driver.switch_to.frame(driver.find_element_by_tag_name('iframe'))
        driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')

    elif row.Keyword == 'Switch_frame_out':
        time.sleep(2)
        driver.switch_to.default_content()
        driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')

    elif row.Keyword == 'dropDown':
        time.sleep(4)
        try:
            #row.Testdata = row.Testdata
            print(row.Testdata)
            '''driver.find_element_by_xpath(
                                "//ul[@class='ui-autocomplete ui-front ui-menu ui-widget ui-widget-content']//li//a/span").click()'''
            dropD = driver.find_element_by_xpath("//ul[@class='ui-autocomplete ui-front ui-menu ui-widget ui-widget-content']//li [@class ='ui-menu-item']/a/span[contains (text(),'" + row.Testdata + "')]")
            print(dropD)
            #dropD = driver.find_element_by_xpath(row.Xpath)
            x=driver.execute_script("arguments[0].click();", dropD)
            print(x,'pass')

        except NoSuchElementException:
            print(message)
            driver.save_screenshot(filename)
            fo = open(filename, 'rb')
            attach = email.mime.application.MIMEApplication(fo.read(), _subtype="png")
            fo.close()
            attach.add_header('Content-Disposition', 'attachment', filename=filename)
            msg.attach(attach)
            s.send_message(msg)
            s.quit()
            print('Email sent')


        else:
            print(row.Keyword, ':', 'Pass')
            driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')


    elif row.Keyword == 'DatePicker':
        time.sleep(5)
        cdf = date.today().strftime(row.Testdata)  # format the date to ddmmyyyy
        print(cdf)
        try:
            # WebDriverWait(driver, 15).until_not(EC.visibility_of_element_located((By.XPATH, "//table[@class='ui-datepicker-calendar']//a[text()= '" + cdf + "']")))
            # time = driver.find_element_by_xpath("//table[@class='ui-datepicker-calendar']//a[text()= '" + cdf + "']")
            # time.click()
            element = driver.find_element_by_xpath("//table[@class='ui-datepicker-calendar']//a[text()= '" + cdf + "']")
            driver.execute_script("arguments[0].click();", element)
        except:

            print(message)
            driver.save_screenshot(filename)
            fo = open(filename, 'rb')
            attach = email.mime.application.MIMEApplication(fo.read(), _subtype="png")
            fo.close()
            attach.add_header('Content-Disposition', 'attachment', filename=filename)
            msg.attach(attach)
            s.send_message(msg)
            s.quit()
            print('Email sent')


        else:
            print(row.Keyword, ':', 'Pass')
            driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')


    elif row.Keyword == 'quit':
        time.sleep(2)
        driver.quit()
        # 3driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')
    elif row.Keyword == 'Verify_Text':
        time.sleep(2)
        try:
            vt = driver.find_element_by_xpath(row.Xpath).text

        except:
            print(row.Keyword, row.Xpath, vt, 'Fail', timestamp())
            driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Fail' + '.' + 'png')

        else:
            print(row.Keyword, row.Xpath, vt, 'Pass', timestamp())
            driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')

    elif row.Keyword == 'Upload_file':
        try:
            upload = driver.find_element_by_xpath(row.Xpath)
            upload1 = upload.send_keys(row.Testdata)
            upload1.sendKeys(Keys.ENTER)
        except:
            print(row.Keyword, row.Xpath, upload1, 'Fail', timestamp())
            driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Fail' + '.' + 'png')

        else:
            print(row.Keyword, row.Xpath, upload1, 'Pass', timestamp())
            driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')


    elif conn:
        query = row.Testdata
        df = pd.read_sql(query, conn).to_string(header=False, index=False)
        print(df)
        if row.Keyword == 'DB_Textbox_Input':
            try:
                # element = WebDriverWait(driver,20).until(EC.element_to_be_clickable(row.Xpath))
                # element.send_Keys(row.Testdata)

                driver.find_element_by_xpath(row.Xpath).send_keys(df)
            except:
                print(row.Keyword, row.Xpath, 'Testdata' + ':' + df, 'Fail', timestamp())
                driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Fail' + '.' + 'png')

            else:
                print(row.Keyword, row.Xpath, 'Testdata' + ':' + df, 'Pass', timestamp())
                driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')


        elif row.Keyword == 'DB_Text_Verify':
            try:
                # element = WebDriverWait(driver,20).until(EC.element_to_be_clickable(row.Xpath))
                # element.send_Keys(row.Testdata)

                dvt = driver.find_element_by_xpath(row.Xpath).text
            except NoSuchElementException:
                print(row.Keyword, row.Xpath, 'Testdata' + ':' + dvt, 'Fail', timestamp())
                driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Fail' + '.' + 'png')

            else:
                print(row.Keyword, row.Xpath, 'Testdata' + ':' + dvt, 'Pass', timestamp())
                driver.save_screenshot('D:/Screenshots' + '/' + timestamp() + 'Pass' + '.' + 'png')
