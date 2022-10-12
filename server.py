from flask import Flask, render_template,request
import random
import win32api
from datetime import datetime
from pytz import timezone
import requests
from bs4 import BeautifulSoup
from batch_page import login2
import multiprocessing
from datetime import date
import win32com.client as win32
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import csv
import time, sys
import urllib.request
import os
import openpyxl
from datetime import datetime
import re
import threading
import shutil
import pythoncom
import pdfkit
from PyPDF2 import PdfFileReader, PdfFileWriter
from urllib.parse import urlparse
from PIL import Image
import urllib
import re
from win32comext.shell import shell

def create_app():
    app=Flask(__name__)

    ret = {'msg': "",'email':''}
    global stopped
    stopped=False
    global processing
    processing=False

    def myfunction(dob, firstname, lastname="", student_id="" ):
        print("Wating................................")

        policies=[]
        session = requests.Session()
        payload = {'identity': '              ',
                   'password': '             '}

        s = session.post("https://claim.otcww.com/auth/login", data=payload)

        # data=session.get(f"https://claim.otcww.com/claim/examine_claim/{case_number}")

        data = session.get(
            f"https://claim.otcww.com/emergency_assistance?policy_match=&filter=policy&lastname={lastname}&firstname={firstname}&birthday={dob}&birthday2={dob}&institution=&student_id={student_id}&case_no=&apply_date=&arrival_date=&effective_date=&expiry_date=&apply_date2=&arrival_date=&effective_date2=&expiry_date2=")

        content = data.content
        soup = BeautifulSoup(content, "html.parser")

        Policy_number = soup.find_all("tr", {"class": "view-policy"})

        for links in Policy_number:
            case = []
            claim = []
            policyinfo = links['data']
            policylist = json.loads(policyinfo)
            # policies.append({"policy": policylist.get('policy', ""), "effective_date": policylist.get("effective_date",""),
            #                  "expiry_date": policylist.get("expiry_date",""), "first_name": policylist.get("firstname", ""),
            #                  "last_name": policylist.get("lastname", "")})

            link = links.find("a").get("href")
            data = session.get(link)
            soup = BeautifulSoup(data.content, "html.parser")
            claims = soup.find_all("tr", {"class": "view-policies"})

            if len(claims) < 1:
                policies.append(
                    {"policy": policylist.get('policy', ""), "effective_date": policylist.get("effective_date", ""),
                     "expiry_date": policylist.get("expiry_date", ""), "first_name": policylist.get("firstname", ""),
                     "last_name": policylist.get("lastname", "")})
                continue
            for item in claims:
                # data = item['data']
                # res = json.loads(data)
                # try:
                #     info = json.loads(res['policy_info'])
                #     policies.append({"policy":res.get('policy_no',""),"effective_date":info[0]["effective_date"],"expiry_date":info[0]["expiry_date"],"first_name": res.get("first_name","") ,"last_name":res.get("last_name","")})
                # except:
                #     pass
                claim_link = item.find("a").get("href")
                if "emergency_assistance" in claim_link:
                    case.append(claim_link)
                else:
                    claim.append(claim_link)
            policies.append({"policy": policylist.get('policy', ""), "effective_date": policylist.get("effective_date",""),
                             "expiry_date": policylist.get("expiry_date",""), "first_name": policylist.get("firstname", ""),
                             "last_name": policylist.get("lastname", ""),"claims":claim, "case":case})
        result=policies
        return(result)



    def running(batchNumber, amount, driver):
        errMsg = ""
        driver.get(
            "https://agent.jfgroup.ca/plan?product_short=0&policy=&search=Search&lastname=&firstname=&birthday=&birthday2=&uname=&student_id=&batch_number=" + str(
                batchNumber) + "&apply_date=&apply_date2=&arrival_date=&arrival_date2=&effective_date=&effective_date2=&expiry_date=&expiry_date2=&status_id=0&province2=&country2=")

        # resume from page DIY
        # batch=50
        batch = int(amount)
        if batch % 20 != 0 and batch < 20:
            pages = 1
        if batch > 20 and batch % 20 == 0:
            pages = batch // 20
        if batch > 20 and batch % 20 != 0:
            pages = batch // 20 + 1
        # pages---0,20,40

        time.sleep(2)
        for a in range(1, pages + 1):
            # resume from page DIY
            # driver.get(
            #     "https://agent.jfgroup.ca/plan?per_page=" + str(
            #         20) + "&product_short=0&policy=&search=Search&lastname=&firstname=&birthday=&birthday2=&uname=&student_id=&batch_number=" + str(
            #         batchNumber) + "&apply_date=&apply_date2=&arrival_date=&arrival_date2=&effective_date=&effective_date2=&expiry_date=&expiry_date2=&status_id=0&province2=&country2=")

            try:
                element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//tr"))
                )

            except:
                return "stuck on page " + str(a)

            elements = driver.find_elements(By.XPATH, "//tr")
            policyNumber = []

            for ele in elements[1:]:
                modified = ele.text.split(' ')
                policyNumber.append(modified[0][-6:])

            # handling batch code with no policy number
            time.sleep(3)
            if not policyNumber:
                return "No policy number under this batch number, check batch number again"

            for i in range(0, len(elements) - 1):
                if stopped==True:
                    return "Stopped"
                driver.get("https://agent.jfgroup.ca/plan/sendpackage/" + str(policyNumber[i]))
                try:
                    element = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@name='withprice']"))
                    )
                except:
                    print(policyNumber[i] + " " + "not successfully executed")

                driver.find_element(By.XPATH, "//input[@name='withprice']").click()
                time.sleep(3)

                ###email address test
                # driver.find_element(By.XPATH, "//input[@name='emailaddr']").clear()
                # driver.find_element(By.XPATH, "//input[@name='emailaddr']").send_keys("zeyu@otcww.com")

                emailAdress=driver.find_element(By.XPATH, "//input[@name='emailaddr']")

#------------------check error email and modified invalid email address-------------------------------------------------
                if emailAdress.get_attribute('value')[-1]==";" or emailAdress.get_attribute('value')[-1]==":":
                    emailModified=emailAdress.get_attribute('value')[:-1]
                    driver.find_element(By.XPATH, "//input[@name='emailaddr']").clear()
                    driver.find_element(By.XPATH, "//input[@name='emailaddr']").send_keys(emailModified)




                driver.find_element(By.XPATH, "//input[@type='submit']").click()


                time.sleep(3)
                element = driver.find_element(By.XPATH,
                                              "//body[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[3]/div[1]/div[1]/div[1]")

                if "Please input valid email address" in element.text:
                    ret['email'] =ret['email']+str(policyNumber[i]) + " on Page  " + str(a + 1) + "\n"

            # wait time
            time.sleep(3)

            # flip to next page
            driver.get(
                "https://agent.jfgroup.ca/plan?per_page=" + str(
                    a * 20) + "&product_short=0&policy=&search=Search&lastname=&firstname=&birthday=&birthday2=&uname=&student_id=&batch_number=" + str(
                    batchNumber) + "&apply_date=&apply_date2=&arrival_date=&arrival_date2=&effective_date=&effective_date2=&expiry_date=&expiry_date2=&status_id=0&province2=&country2=")
            time.sleep(3)
            sys.stdout.flush()
            print("Page " + str(a) + " Finished")

        successful = "successfully processed  " + batchNumber + "\n" + errMsg
        return successful
        # sys.stdout.flush()

    def running2(batchNumber, amount, driver,pageInput):
        errMsg = ""
        # driver.get(
        #     "https://agent.jfgroup.ca/plan?product_short=0&policy=&search=Search&lastname=&firstname=&birthday=&birthday2=&uname=&student_id=&batch_number=" + str(
        #         batchNumber) + "&apply_date=&apply_date2=&arrival_date=&arrival_date2=&effective_date=&effective_date2=&expiry_date=&expiry_date2=&status_id=0&province2=&country2=")

        # resume from page DIY
        # batch=50
        driver.get(
            "https://agent.jfgroup.ca/plan?per_page=" + str(
                (int(pageInput) - 1)*20) + "&product_short=0&policy=&search=Search&lastname=&firstname=&birthday=&birthday2=&uname=&student_id=&batch_number=" + str(
                batchNumber) + "&apply_date=&apply_date2=&arrival_date=&arrival_date2=&effective_date=&effective_date2=&expiry_date=&expiry_date2=&status_id=0&province2=&country2=")

        batch = int(amount)
        if batch % 20 != 0 and batch < 20:
            pages = 1
        if batch > 20 and batch % 20 == 0:
            pages = batch // 20
        if batch > 20 and batch % 20 != 0:
            pages = batch // 20 + 1
        # pages---0,20,40

        time.sleep(2)
        for a in range(int(pageInput), pages + 1):
            # resume from page DIY

            try:
                element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//tr"))
                )

            except:
                return "stuck on page " + str(a)

            elements = driver.find_elements(By.XPATH, "//tr")
            policyNumber = []

            for ele in elements[1:]:
                modified = ele.text.split(' ')
                policyNumber.append(modified[0][-6:])

            # handling batch code with no policy number
            time.sleep(3)
            if not policyNumber:
                return "No policy number under this batch number, check batch number again"

            for i in range(0, len(elements) - 1):
                if stopped==True:
                    return "Stopped"
                driver.get("https://agent.jfgroup.ca/plan/sendpackage/" + str(policyNumber[i]))
                try:
                    element = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@name='withprice']"))
                    )
                except:
                    print(policyNumber[i] + " " + "not successfully executed")

                driver.find_element(By.XPATH, "//input[@name='withprice']").click()
                time.sleep(3)

                ###email address test
                # driver.find_element(By.XPATH, "//input[@name='emailaddr']").clear()
                # driver.find_element(By.XPATH, "//input[@name='emailaddr']").send_keys("zeyu@otcww.com")

                driver.find_element(By.XPATH, "//input[@type='submit']").click()
                time.sleep(3)
                element = driver.find_element(By.XPATH,
                                              "//body[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[3]/div[1]/div[1]/div[1]")

                if "Please input valid email address" in element.text:
                    ret['email'] =ret['email']+str(policyNumber[i]) + " on Page  " + str(a + 1) + "\n"

            # wait time
            time.sleep(6)

            # flip to next page
            driver.get(
                "https://agent.jfgroup.ca/plan?per_page=" + str(
                    a * 20) + "&product_short=0&policy=&search=Search&lastname=&firstname=&birthday=&birthday2=&uname=&student_id=&batch_number=" + str(
                    batchNumber) + "&apply_date=&apply_date2=&arrival_date=&arrival_date2=&effective_date=&effective_date2=&expiry_date=&expiry_date2=&status_id=0&province2=&country2=")
            time.sleep(3)
            sys.stdout.flush()
            print("Page " + str(a) + " Finished")

        successful = "successfully processed  " + batchNumber + "\n" + errMsg
        return successful
        # sys.stdout.flush()

    def login(batchNumber, amount):
        global processing
        processing=True
        driver = webdriver.Chrome(r'C:/Users/Angela G.DESKTOP-A1O7G37/Desktop/chromedriver')
        driver.get("https://agent.jfgroup.ca/user/login")
        driver.find_element(By.XPATH, "//input[@name='username']").send_keys("          ")
        driver.find_element(By.XPATH, "//input[@name='password']").send_keys("          ")
        driver.find_element(By.XPATH, "//input[@type='submit']").click()
        time.sleep(2)
        message = running(batchNumber, amount, driver)
        print(message)
        processing = False

    def login2(batchNumber,amount,page):
        global processing
        processing = True
        driver = webdriver.Chrome(r'C:/Users/Angela G.DESKTOP-A1O7G37/Desktop/chromedriver')
        driver.get("https://agent.jfgroup.ca/user/login")
        driver.find_element(By.XPATH, "//input[@name='username']").send_keys("           ")
        driver.find_element(By.XPATH, "//input[@name='password']").send_keys("           ")
        driver.find_element(By.XPATH, "//input[@type='submit']").click()
        time.sleep(2)
        message = running2(batchNumber, amount, driver,page)
        print(message)
        processing = False

    def downloadPolices(username, password, batch):
        dataSet = []
        my_dict = {}
        nameIndex = 0
        empty = 0
        policyNumber = {}

        driver = webdriver.Chrome(r'C:/Users/Angela G.DESKTOP-A1O7G37/Desktop/chromedriver')
        driver.get("https://agent.jfgroup.ca/user/login")
        driver.find_element(By.XPATH, "//input[@name='username']").send_keys(username)
        driver.find_element(By.XPATH, "//input[@name='password']").send_keys(password)
        driver.find_element(By.XPATH, "//input[@type='submit']").click()
        time.sleep(2)

        driver.get(
            f"https://agent.jfgroup.ca/plan?product_short=0&policy=&search=Search&lastname=&firstname=&birthday=&birthday2=&uname=&student_id=&batch_number={batch}&apply_date=&apply_date2=&arrival_date=&arrival_date2=&effective_date=&effective_date2=&expiry_date=&expiry_date2=&status_id=0&province2=&country2=")
        time.sleep(2)
        # driver.find_element(By.CLASS_NAME,"btn-info").click()
        driver.get("https://agent.jfgroup.ca/plan/export_list")
        time.sleep(5)

        currentDate = datetime.today().strftime('%Y%m%d')

        excel = openpyxl.load_workbook(f"C:/Users/Angela G.DESKTOP-A1O7G37/Downloads/Policy{currentDate}.xlsx")

        sheet = excel.active

        for r in sheet.rows:
            for cell in r:
                if r.index(cell) == 4 and cell.value != "policy":
                    policyNumber[cell.value[-6:]] = r[42].value

        if not os.path.isdir(fr"\\jfnas\share folder\{batch}"):
            os.mkdir(fr"\\jfnas\share folder\{batch}")

        else:
            os.remove(f"C:/Users/Angela G.DESKTOP-A1O7G37/Downloads/Policy{currentDate}.xlsx")
            return fr"same {batch} folder already existing in sharefolder, please remove it"

        session = requests.Session()
        payload = {"username": username,
                   "password": password,
                   "csrf_test_name": "                              "
                   }

        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36',
            'Referer': 'https://agent.jfgroup.ca/user/login',
            "Cookie": "                                                    "
        }
        s = session.post("https://agent.jfgroup.ca/user/login/", headers=headers, data=payload)



        try:
            for key in policyNumber:

                if '"' in policyNumber[key]:
                    policyNumber[key] = policyNumber[key].replace('"', '')
                opener = urllib.request.build_opener()
                opener.addheaders = [('User-agent',
                                      'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36'),
                                     ('Cookie',
                                      '                                              '),
 
                                     ]
                urllib.request.install_opener(opener)

                urllib.request.urlretrieve("https://agent.jfgroup.ca/plan/pdf/" + key + ".pdf",
                                           fr"\\jfnas\share folder\{batch}\\" + key + " " + policyNumber[key] + ".pdf")
        except Exception as e:
            shutil.rmtree(fr"\\jfnas\share folder\{batch}")
            os.remove(f"C:/Users/Angela G.DESKTOP-A1O7G37/Downloads/Policy{currentDate}.xlsx")
            return e

        os.remove(f"C:/Users/Angela G.DESKTOP-A1O7G37/Downloads/Policy{currentDate}.xlsx")
        return f"successful downloaded, find file in sharefolder {batch}"

    def dateUnionlize(filename,row,column):
        try:
            row = int(row)
            column = int(column)
            format = '%m/%d/%Y'
            r = csv.reader(open(fr"C:\Users\Angela G.DESKTOP-A1O7G37\Downloads\{filename}"))  # Here your csv file
            lines = list(r)
            for i in range(row, len(lines)):
                if (len(lines[i][column].strip()) <= 9) and re.search('[a-zA-Z]', lines[i][column].strip()):
                    if (int(lines[i][column][-2:]) > 50):
                        lines[i][column] = lines[i][column][0:-2] + "19" + lines[i][column][-2:]
                        print(lines[i][column])

                    else:
                        lines[i][column] = lines[i][column][0:-2] + "20" + lines[i][column][-2:]

                    lines[i][column] = datetime.strptime(lines[i][column], '%d-%b-%Y').strftime('%Y-%m-%d')
                    continue
                if (len(lines[i][column].strip()) > 10):
                    lines[i][column] = lines[i][column].lstrip()
                    try:
                        lines[i][column] = datetime.strptime(lines[i][column], '%B %d, %Y').strftime('%Y-%m-%d')
                    except:
                        lines[i][column] = datetime.strptime(lines[i][column], '%A, %B %d, %Y').strftime('%Y-%m-%d')
                    continue
                if (lines[i][column].strip()):
                    if (str.isdigit(lines[i][column][:4])):
                        print(lines[i][column])
                        continue
                    if ("." in lines[i][column]):
                        lines[i][column] = lines[i][column][-4:] + "-" + lines[i][column][3:5] + "-" + lines[i][column][
                                                                                                       :2]
                        print(lines[i][column])
                        continue
                    if (not str.isdigit(lines[i][column][-4:])):
                        if (int(lines[i][column][-2:]) > 50):
                            lines[i][column] = "19" + lines[i][column][-2:] + "-" + lines[i][column][3:5] + "-" + \
                                               lines[i][column][:2]
                            continue
                        else:
                            lines[i][column] = "20" + lines[i][column][-2:] + "-" + lines[i][column][3:5] + "-" + \
                                               lines[i][column][:2]
                            continue
                    if (str.isdigit(lines[i][column][-4:])):
                        try:
                            lines[i][column] = datetime.strptime(lines[i][column], format).date()

                        except:
                            lines[i][column] = lines[i][column][-4:] + "-" + lines[i][column][3:5] + "-" + lines[i][
                                                                                                               column][
                                                                                                           :2]

            writer = csv.writer(open(fr"\\jfnas\share folder\{filename}", 'w', newline=''))

            writer.writerows(lines)

            def removeFile():
                os.remove(fr"C:\Users\Angela G.DESKTOP-A1O7G37\Downloads\{filename}")

            threading.Timer(1.5,removeFile).start()
            return "successfully correct format, find it in sharefolder"

        except Exception as e:
           os.remove(fr"C:\Users\Angela G.DESKTOP-A1O7G37\Downloads\{filename}")
           return e

    def Emailer(recipient, user, claim_number):
        outlook = win32.Dispatch('outlook.application', pythoncom.CoInitialize())
        try:
            mail = outlook.CreateItem(0)
            mail.To = recipient
            # mail.Sentonbehalfof='may@otcww.com'
            # mail. = 'may@otcww.com'
            mail.Subject = f"E-transfer update for your claim C00{claim_number}"
            mail.HtmlBody = f" <Body style=font-size:11pt;font-family:Calibri> Dear Client Jia, <br> \
                               We acknowledge received your claim documents. <br>\
                               Your claim has been processed according to the policy wording. Please kindly find the attached Letter of Explanation of Benefit. <br> \
                               The reimbursement will be sent to <br> \
                               <strong>{recipient}.</strong> It usually takes three to five business days to prepare the e-transfer. <br> \
                               <br>\
                               Best Regards and Stay Safe, </Body> <br> \
                               <span style='color: navy; font: 13px arial, sans-serif;'><strong>Ontime Care Worldwide Inc.</strong></span> <br> \
                               Richmond Hill, ON L4B 3H7 <br> \
                               Tel:905-707-3355 <br> \
                               Fax: 905-707-1513 <br> \
                               Website: www.jfgroup.ca <br> \
                               <p style='color: red; font: 13px arial, sans-serif;'><strong>This is automatically generated email, do not directly reply to this email</strong><p>\
                               <p style='color: red; font: 9px arial, sans-serif;'>This message, including any attachments, is privileged and intended only for the person(s) named above. This material may contain confidential or personal information., which may be subject to the provisions of the Freedom of Information & Protection Act. Any other distribution, copying or disclosure is strictly prohibited. If you are the intended recipient or have received this message in error, please notify us immediately by telephone, fax or email, and permanently delete the original transmission from us, including any attachments, without making a copy.</p>"
            if os.path.isfile(fr"\\jfnas\share folder\E_claim\{user}\{claim_number}.pdf"):
                mail.Attachments.Add(fr"\\jfnas\share folder\E_claim\{user}\{claim_number}.pdf")
            else:
                return f"        No claim {claim_number} attachment was found in {user} folder,EOB email not sent"


            From = None
            for myEmailAddress in outlook.Session.Accounts:
                if "Autoreply@otcww.com" in str(myEmailAddress):
                    From = myEmailAddress
                    break
            # for account in outlook.Session.Accounts:
            #     print(account)
            #     if account.DisplayName == "May@otcww.com":
            if From != None:
                # This line basically calls the "mail.SendUsingAccount = xyz@email.com" outlook VBA command
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, From))
                # mail.Display(True)
                mail.Send()
                return "successfully sent e-transfer message"
        except:
            print("no file found")
            return "something wrong"


    def payee_extraction(caseNumber,user):
        caseNumber="C00"+caseNumber
        payeeinfo = []
        case_number = caseNumber[3:]
        session = requests.Session()

        payload = {'identity': '              ',
                   'password': '              '}

        s = session.post("https://claim.otcww.com/auth/login", data=payload)

        data = session.get(f"https://claim.otcww.com/claim/examine_claim/{case_number}")

        content = data.content
        soup = BeautifulSoup(content, "html.parser")

        Case_number = soup.find("div", {"class": "title_left"}).find("h3")
        Case_number = Case_number.text.split(" ")
        casenumber = Case_number[3].replace("#", "")

        Policy_number = soup.find("div", {"class": "row policy_info"}).find_all("div", {"class": "form-group col-sm-3"})

        for o in Policy_number:
            if "Policy :" in o.text:
                policynumber = o.text.split(" ")
                policynumber = policynumber[2].replace("\xa0", "")
                break

        email_template = soup.find("table", {"class": "table table-hover table-bordered"}).find_all("div", {
            "class": "col-sm-12"})
        for o in email_template:
            if "email" in o.text:
                payeeinfo = o.text.split(":")
                break
        if payeeinfo:
            payeeinfo[2] = payeeinfo[2].replace("\xa0", "")
            payeeinfo = payeeinfo[1:]
            payeeinfo.insert(0, casenumber)
            payeeinfo.insert(0, policynumber)

            return payeeinfo
        else:
            payeeinfo = ["", "", "", "Warning: ", "By Cheque"]
            return payeeinfo

    def printer(eclaim,user,printerName):
        session = requests.Session()
        cookies = {'jf_claim': "                                        "}
        s = session.get(f"https://claim.otcww.com/eclaim/detail/{eclaim}", cookies=cookies)
        data = s.content
        soup = BeautifulSoup(data, "html.parser")
        element = soup.find_all("div", {"class": "row intake-forms-list col-sm-12"})[1]
        links = element.find_all("a", {"class": "img-responsive"})
        linklist = []
        for link in links:
            linklist.append(link['href'])
        opener = urllib.request.build_opener()
        opener.addheaders = [('User-agent',
                              'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36'),
                             ('Cookie', '                                     '),
                             ]
        urllib.request.install_opener(opener)

        if not os.path.isdir(fr"\\jfnas\share folder\print\{user}"):
            return "please create a empty folder with your name under print folder in sharefolder"
        for index, i in enumerate(linklist):
            if  re.findall(r'[\u4e00-\u9fff]+', i):
                url = urllib.parse.urlparse(i)
                url = url.scheme + "://" + url.netloc + urllib.parse.quote(url.path)
                i=url
            rnumber = str(random.randint(0, 10))
            path = urlparse(i).path
            ext = os.path.splitext(path)[1]
            urllib.request.urlretrieve(i,
                                       fr"\\jfnas\share folder\print\{user}\{rnumber + str(index)}{ext}")
            time.sleep(1)

        ##--------------------------------------------------------------print eclaim page-------------------------------------------------------------
        path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
        options = {'cookie': [('jf_claim',
                               '                                       ')]}

        config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
        pdfkit.from_url(fr"https://claim.otcww.com/eclaim/export/{eclaim}.pdf",
                        fr"\\jfnas\share folder\print\{user}\claim.pdf", configuration=config,
                        options=options)
        time.sleep(4)
        writer = PdfFileWriter()
        reader = PdfFileReader(fr"\\jfnas\share folder\print\{user}\claim.pdf")
        page = reader.getPage(0)
        writer.addPage(page)
        page = reader.getPage(1)
        writer.addPage(page)
        with open(fr"\\jfnas\share folder\print\{user}\claim.pdf", 'wb') as f:
            writer.write(f)

        time.sleep(1)

        for index, file in enumerate(os.listdir(fr"\\jfnas\share folder\print\{user}")):
            if any(x in file for x in ["jpg", "jpeg","png"]):
                image_1 = Image.open(fr'\\jfnas\share folder\print\{user}\{file}')
                im_1 = image_1.convert('RGB')
                im_1.save(fr'\\jfnas\share folder\print\{user}\{index}.pdf')
                time.sleep(1.5)

                win32api.ShellExecute(0, "printto", fr"\\jfnas\share folder\print\{user}\{index}.pdf", fr'\\jf-dc\{printerName}',
                                      ".", 0)
                # shell.ShellExecuteEx(0, "printto", fr"\\jfnas\share folder\print\{user}\{file}",
                #                      fr'\\jf-dc\{printerName}', ".", 0)
                time.sleep(1.5)
                continue
            # shell.ShellExecuteEx(0, "printto", fr"\\jfnas\share folder\print\{user}\{file}", fr'\\jf-dc\{printerName}', ".", 0)

            win32api.ShellExecute(0, "printto", fr"\\jfnas\share folder\print\{user}\{file}", fr'\\jf-dc\{printerName}', ".", 0)

            time.sleep(1.5)

        time.sleep(2.5)
        folder = fr"\\jfnas\share folder\print\{user}"
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                return('Failed to delete %s. Reason: %s' % (file_path, e))
        return f"successfully print on {printerName}"

    @app.route("/", methods=["GET"])
    def home():
            msg=""
            if processing==True:
                msg="Currently Running"
            toronto = timezone('US/Eastern')
            trt_time = datetime.now(toronto)
            formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
            return render_template("home.html",entries=(msg if msg else ""),formatted_date=formatted_date)


    @app.route("/batch", methods=["POST"])
    def batch_input():
            global stopped
            stopped = False
            errorMes=""
            # queue = multiprocessing.Queue()
            # queue.put(ret)
            batch_number = request.form.get("batch_number")
            amount = request.form.get("amount")

            if request.method=="POST":
                if batch_number.isnumeric()==True and amount.isnumeric()==True:
                    # try:
                    #    t = multiprocessing.Process(target=login(batch_number, amount,queue))
                      errorMes=login(batch_number, amount)
                       # t.start()
                       # t.join()
                       # errorMes=login(batch_number,amount,quene)
                    # except:
                    #     errorMes="something went wrong"
                else:
                    errorMes="Wrong Input"
            # if request.method == "GET":
            #     t.terminate()
            #     errorMes="stopped"
            toronto = timezone('US/Eastern')
            trt_time = datetime.now(toronto)
            formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
            return render_template("home.html",entries=errorMes,formatted_date=formatted_date)

    @app.route("/download", methods=["POST"])
    def download():
        errorMes = ""
        username = request.form.get("username")
        password = request.form.get("password")
        batch = request.form.get("batch")

        if username and password and batch:
            errorMes=downloadPolices(username,password,batch)
        else:
            errorMes="something went wrong. contact Michael"

        toronto = timezone('US/Eastern')
        trt_time = datetime.now(toronto)
        formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
        return render_template("home.html", entries=errorMes, formatted_date=formatted_date)


    # @app.route("/stop", methods=['GET'])
    # def set_stop_run():
    #     global stopped
    #     stopped=True
    #     toronto = timezone('US/Eastern')
    #     trt_time = datetime.now(toronto)
    #     formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
    #     global processing
    #     processing = False
    #     return render_template("home.html", entries="Stopped", formatted_date=formatted_date)

    @app.route('/uploader', methods=['GET', 'POST'])
    def upload_file():
        errorMes=""
        if request.method == 'POST':
            row = request.form.get("row")
            column = request.form.get("column")
            f = request.files['file']
            filename, file_extension = os.path.splitext(f.filename)
            if file_extension!=".csv":
                toronto = timezone('US/Eastern')
                trt_time = datetime.now(toronto)
                formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
                errorMes="please upload CSV file"
                return render_template("home.html", entries=errorMes, formatted_date=formatted_date)

            f.save(f"C:/Users/Angela G.DESKTOP-A1O7G37/Downloads/"+(f.filename))
            errorMes=dateUnionlize(f.filename,row,int(column)-1)
        toronto = timezone('US/Eastern')
        trt_time = datetime.now(toronto)
        formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
        return render_template("home.html", entries=errorMes, formatted_date=formatted_date)



    @app.route("/page", methods=["GET", "POST"])
    def page():
        errorMes = ""
        batch_number = request.form.get("batch_number")
        amount = request.form.get("amount")
        page=request.form.get("page")
        if request.method == "POST":

            if batch_number.isnumeric() == True and amount.isnumeric() == True:
                try:
                    msg = login2(batch_number, amount,page)
                    errorMes=msg
                except:
                    errorMes = "something went wrong"

            else:
                errorMes = "Wrong Input"
        toronto = timezone('US/Eastern')
        trt_time = datetime.now(toronto)
        formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
        return render_template("home.html", entries=errorMes, formatted_date=formatted_date)

    @app.route("/previous_policy", methods=["POST"])
    def previous_policy():
        errorMes = ""
        case=[]
        claim=[]
        policies=[]
        student_id=request.form.get("student_id").strip()
        date_birth = request.form.get("DOB").strip()
        firstname = request.form.get("firstname")
        lastname=request.form.get("lastname")
        if date_birth:
            try:
                datetime.strptime(date_birth, '%Y-%m-%d')
            except ValueError:
                errorMes="Please input valid DOB"
                return render_template("home.html",entries=errorMes)
        else:
            if firstname and lastname:
                pass
            elif student_id:
                pass
            else:
                errorMes = "Please input more details"
                return render_template("home.html", entries=errorMes)

        msg = myfunction(dob=date_birth, firstname=firstname, lastname=lastname,student_id=student_id)


        try:
            policies=msg
        except:
            errorMes = "Something went wrong"

        toronto = timezone('US/Eastern')
        trt_time = datetime.now(toronto)
        formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
        print(policies)
        return render_template("home.html",entries=errorMes,policies=policies,formatted_date=formatted_date)

    @app.route("/payee", methods=["POST"])
    def payee_extract():
        errorMes = ""
        message=""
        case_number = request.form.get("caseNumber")
        user=request.form.get("user")
        if request.method == "POST":
            if case_number.isdigit() and len(case_number)==5:

                payeeinfo=payee_extraction(case_number,user)
                print(payeeinfo[-1])
                message=Emailer(payeeinfo[-1],user,case_number)
                ##------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                errorMes = '.'.join(payeeinfo)

            else:
                errorMes = "Wrong Input"
        toronto = timezone('US/Eastern')
        trt_time = datetime.now(toronto)
        formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
        return render_template("home.html", entries=errorMes, formatted_date=formatted_date,user=user,message=message)

    @app.route("/printer", methods=["POST"])
    def e_claim():
        errorMes = ""
        eclaim = request.form.get("eclaim").strip()
        user = request.form.get("user").lower().strip()
        printerName = request.form.get("printer").upper()
        print(printerName)
        if request.method == "POST":
            if eclaim.isdigit() and len(eclaim) == 5:
                errorMes = printer(eclaim,user,printerName)
                ##------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            else:
                errorMes = "Wrong Input"
        toronto = timezone('US/Eastern')
        trt_time = datetime.now(toronto)
        formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
        return render_template("home.html", result=errorMes, formatted_date=formatted_date, user=user, printer=printerName)

    @app.errorhandler(404)
    def invalid_route():
        toronto = timezone('US/Eastern')
        trt_time = datetime.now(toronto)
        formatted_date = trt_time.strftime('%Y-%m-%d_%H:%M:%S')
        return render_template("home.html", entries="Try retype", formatted_date=formatted_date)

    return app

