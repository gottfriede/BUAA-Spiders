#!/usr/bin/python
# coding=utf-8

import os
import sys
import xlwt
from selenium import webdriver
import msvcrt
import time
import tesserocr
from PIL import Image
import requests as req
from io import BytesIO
import requests
import base64
import urllib
import urllib.request
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication 

vpnUrl = "https://e2.buaa.edu.cn/users/sign_in"
jiaowuUrl = "https://jwxt-7001.e2.buaa.edu.cn/ieas2.1"
chromeDriver_location = r"..\chromedriver_win32\chromedriver.exe"
semesterNumber = 6

def inputPassword(msg) :
    print(msg, end="", flush=True)
    password = []
    while True :
        ch = msvcrt.getch()
        if ch == b'\r' or ch == b'\n' :
            msvcrt.putch(b"\n")
            break
        elif ch == b"\b" :
            if password :
                del password[-1]
                msvcrt.putch(b"\b")
                msvcrt.putch(b" ")
                msvcrt.putch(b"\b")
        else :
            password.append(ch)
            msvcrt.putch(b"*")
    return b"".join(password).decode()

def init() :
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    # options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("User-Agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Safari/537.36 Edg/84.0.522.40")
    chrome = webdriver.Chrome(chromeDriver_location,chrome_options=options)
    return chrome

# @Description: Use Baidu Ocr Api to identify the verification code that may apper during the login process
# @Param: src (The local path of the verification code image)
# @Return: ocrResult(str) OR -1 when no response
# @Note: Please fill in your own access_token
def baiduOcr(src) :
    request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"
    access_token = '24.5415f5de64ecdfdc65398bd51f8b9e3f.2592000.1600484749.282335-22117678'     #Your own access_token here
    f = open(src, 'rb')
    img = base64.b64encode(f.read())
    params = {"image":img}
    request_url = request_url + "?access_token=" + access_token
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(request_url, data=params, headers=headers)
    if response: 
        ocrResult = ((str(response.json()).split(r"words': ' "))[1].split(r"'")[0])
        return ocrResult
    else :
        return -1

def login(username, password, url, chrome) :
    chrome.get(url)
    username_location = chrome.find_element_by_id("user_login")
    password_location = chrome.find_element_by_id("user_password")
    username_location.send_keys(username)
    password_location.send_keys(password)
    chrome.find_elements_by_name("commit")[0].click()
    time.sleep(1)
    if (chrome.current_url == url) :
        return -1
    chrome.find_elements_by_xpath('//a[@data-original-title="教务管理系统"]')[0].click()
    time.sleep(1)
    all_handles = chrome.window_handles
    for handles in all_handles:
        if chrome.current_window_handle != handles:
            chrome.switch_to_window(handles)

    chrome.switch_to.frame("loginIframe")
    time.sleep(1)
    username_location = chrome.find_element_by_id("unPassword")
    password_location = chrome.find_element_by_id("pwPassword")
    username_location.send_keys(username)
    password_location.send_keys(password)
    code_location = chrome.find_element_by_id("cptPassword")
    code = chrome.find_element_by_class_name("code")
    codeSrc = code.get_attribute("src")
    if codeSrc!=None :
        code.click()
        time.sleep(1)
        request = urllib.request.Request(codeSrc)
        responce = urllib.request.urlopen(request)
        get_img = responce.read()
        with open("code.jpg",'wb') as fp :
            fp.write(get_img)
        fp.close()
        ocrResult = baiduOcr("code.jpg")
        print("ocrResult : " + ocrResult)
        code_location.send_keys(ocrResult)
        time.sleep(1)
    chrome.find_element_by_class_name("submit-btn").click()
    time.sleep(1)

    all_handles = chrome.window_handles
    for handles in all_handles:
        if chrome.current_window_handle != handles:
            chrome.switch_to_window(handles)
    
    if (not (jiaowuUrl in chrome.current_url)) :
        return -1
    return 0

def getGrades(chrome, gradess) :
    button = chrome.find_element_by_xpath('//a[@href="/ieas2.1/cjcx/queryCjpub_ty"]')
    chrome.execute_script("arguments[0].click();", button)
    time.sleep(1)
    chrome.switch_to.frame("iframename")
    time.sleep(1)
    chrome.find_element_by_xpath('//a[@class="qmcx"]').click()
    time.sleep(1)
    for i in range(semesterNumber - 1) :
        chrome.find_element_by_xpath('//*[@id="xnxqid"]/option[' + str(i+2) + ']').click()
        chrome.find_element_by_xpath('//div[@class="addlist_button2"]/a').click()
        time.sleep(1)
        table = chrome.find_element_by_xpath("//body/div[1]/div/div[4]/table/tbody")
        rows = table.find_elements_by_tag_name("tr")
        rows.pop(0)
        grades = []
        for tr in rows :
            grade = []
            tds = tr.find_elements_by_tag_name("td")
            for td in tds :
                grade.append(td.text)
            grades.append(grade)
        gradess.append(grades)
    return 0

def geneExcel(gradess) :
    titles = ["序号","学年学期","开课院系","课程代码","课程名称","课程性质","课程类别","学分",
    "是否考试课","补考重修标记","总成绩","折算成绩","成绩备注"]
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet("grades")
    worksheet.col(5).width = 10000
    style = xlwt.XFStyle()
    al = xlwt.Alignment()
    al.horz = 0x02
    al.vert = 0x01
    style.alignment = al
    i = 1
    j = 1
    for title in titles :
        worksheet.write(0, j, label = title, style = style)
        j += 1
    j = 1
    for grades in gradess :
        credit = 0
        average = 0
        begin = i
        for grade in grades :
            creditOne = 0
            scoreOne = 0
            for text in grade :
                worksheet.write(i,j,label = text, style = style)
                if (j == 8) :
                    creditOne = float(text)
                if (j == 12) :
                    scoreOne = float(text)
                j += 1
            credit += creditOne
            average += creditOne * scoreOne
            j = 1
            i += 1
        if (credit != 0) :
            worksheet.write_merge(begin, i-1, 0, 0, grades[0][1], style)
            average /= credit
            worksheet.write(i-2,15,label = "学分", style = style)
            worksheet.write(i-2,16,label = "加权平均分", style = style)
            worksheet.write(i-1,15,label = credit, style = style)
            worksheet.write(i-1,16,label = average, style = style)
            i += 1
    workbook.save("scores.xls")
    return 0

# @Description: Send Email by QQ
# @Param: to_addrs (Email recipient`s address)
# @Note: Please fill in your own from_addr and password
def sendEmail(to_addrs) :
    from_addr = '923585159@qq.com'      #Your own from_addr here
    password = 'nzxauyexrisjbbeh'       #Your own password here
    smtp_server = 'smtp.qq.com'

    content = 'This is an automatically sent email.\n Please check the attachment.'
    textApart = MIMEText(content)
    xlsFile = 'scores.xls'
    xlsApart = MIMEImage(open(xlsFile, 'rb').read(), xlsFile.split('.')[-1])
    xlsApart.add_header('Content-Disposition', 'attachment', filename=xlsFile)
    m = MIMEMultipart()
    m.attach(textApart)
    m.attach(xlsApart)
    m['Subject'] = "[jiaowuSpider] Scores Result"
    try:
        server = smtplib.SMTP('smtp.qq.com')
        server.login(from_addr,password)
        server.sendmail(from_addr, to_addrs, m.as_string())
    except smtplib.SMTPException as e:
        server.quit()
        return -1
    server.quit()
    return 0

if __name__ == "__main__" :
    print("\n用户登录（统一认证账号及密码）：")
    username = input("Username :")
    password = inputPassword("Password :")
    emailAddress = input("Email Address :")

    attempt = 0
    while (True) :
        chrome = init()
        try :
            loginResult = login(username, password, vpnUrl, chrome)
        except Exception :
            loginResult = -1
        if (loginResult == 0) :
            break
        chrome.quit()
        attempt = attempt + 1
        print("Failed to log in! Times: " + str(attempt) + ".\nTry again")
        if (attempt >= 5) :
            print("An Error occurred during Login! It may be caused by: \n \
                1) Username or password incorrect\n \
                2) Time limit exceed\n \
                3) Cant solve the captcha(5 attempts)\n \
                4) Other reasons ")
            sys.exit(1)
    print("Login succeed!\nGetting the grades...")

    gradess = []
    if (getGrades(chrome, gradess) != 0) :
        chrome.quit()
        print("An Error occurred during getting grades! It may be caused by: \n \
            1) Time limit exceed\n \
            2) Other reasons ")
        sys.exit(1)
    chrome.quit()
    print("Getting grades succeed!\nGenerating excel...")

    if (geneExcel(gradess) != 0) :
        print("An Error occured during generating excel! It`s quite strange")
        sys.exit(1)
    print("Generating excel succeed!\nSending email... ")

    if (sendEmail(emailAddress) != 0) :
        print("An Error occured during sending email! It may be caused by: \n \
            1) I dont know ")
        sys.exit(1)
    print("Sending email succeed!\nALL IS DONE")
