# -*- coding: utf-8 -*-
import os
import shutil
import win32clipboard
import win32con
import urllib.request
import sys
import re
from time import sleep

import smtplib
import email.mime.text
from email.utils import formataddr


reportStatus = ""
reportTitle = ""
reportContent = ""


emailSender = "developer-send@qq.com"
emailReceiver = "developer-receive@qq.com"
emailServer = "smtp.qq.com"
emailPw = "pcawfmdhyfvaecea"

def GetClipboardText():
    win32clipboard.OpenClipboard()
    text = win32clipboard.GetClipboardData(win32con.CF_TEXT)
    win32clipboard.CloseClipboard()
    text = text.decode('utf-8')
    return text

def SetClipboardText(text):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_TEXT, text)
    win32clipboard.CloseClipboard()

def IsUrl(text):
    # http://
    if text[0:7] == "http://":
        return True
    elif text[0:8] == "https://":
        return True
    return False

def SendEmailAndQuit(err=False, msg=None):
    if msg == None:
        msg = "请尝试重新复制链接！"
    printHelp(msg)
    
    msg = email.mime.text.MIMEText(reportContent)
    msg['From'] = formataddr(["yif-dev-send", emailSender])
    msg['To'] = formataddr(["pasteUrl2-receiver", emailReceiver])
    msg['Subject'] = "[pasteUrl2] " + reportStatus + reportTitle

    try:
        server = smtplib.SMTP_SSL(emailServer, port=465)
        server.login(emailSender, emailPw)
        response = server.sendmail(emailSender, emailReceiver, msg.as_string())
        server.quit()
#    except Exception as e:
#        print(e)
#    print(response)
    finally:
        quit()

def printHelp(msg):
    print()
    print()
    print(msg)
    print("\n\n如果问题重复出现，请联系您的外孙")
    sleep(5)

#########################################

print("请稍等...")

clipboardText = GetClipboardText()


if not IsUrl(clipboardText):
    reportStatus += "ERR: "
    reportTitle += "非网址"
    reportContent += "剪切板内容" + '"' + clipboardText + '"'
    reportContent += "不是一个有效的网址"
    SendEmailAndQuit(err=True)

url = clipboardText

try:
    request = urllib.request.urlopen(url)
    responseContent = request.read().decode('utf-8')
except Exception as exception:
    reportStatus += "ERR: "
    reportTitle += "无法获取网页"
    reportContent += url
    reportContent += "\n"
    reportContent += str(exception)
    SendEmailAndQuit(err=True)


patternForTitle = "(?s)<title>\s*?(.+?)\s*?</title>"
patternForXIAONIANGAO = '<div class="header-album-title">\s*?(.+?)\s*?</div>'

searchPattern = ""
if "xiaoniangao.cn" in url:
    searchPattern = patternForXIAONIANGAO
else:
    searchPattern = patternForTitle

matches = re.search(searchPattern, responseContent)




title = ""
try:
    title = matches.group(1)
except Exception as exception:
    reportStatus += "ERR: "
    reportTitle += "未找到标题"
    reportContent += url
    reportContent += "\n"
    reportContent += str(exception)
    SendEmailAndQuit(err=True)
    
try:
    fileContent = "[InternetShortcut]\nURL=" + url
    fileName = title + ".url"
    try:
        existingFile = open(fileName, "r")
    except:
        pass
    else:
        reportStatus += "ERR: "
        reportTitle += "文件已经存在"
        reportContent += url
        reportContent += "\n"
        #reportContent += str(exception)
        SendEmailAndQuit(err=True, msg="文章已经保存过了，不用再保存了！")
    
    saveFile = open(fileName, "w")
    saveFile.write(fileContent)
    saveFile.close()
except Exception as exception:
    reportStatus += "ERR: "
    reportTitle += "处理文件时遇到问题"
    reportContent += url
    reportContent += "\n"
    reportContent += str(exception)
    SendEmailAndQuit(err=True, msg="程序遇到错误")

reportStatus += "成功: "
reportTitle += title
reportContent += url
SendEmailAndQuit()

##
