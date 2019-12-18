import requests
import pytesseract
from docx import Document
import time
from selenium import webdriver
import numpy as np
#读取文档模板
doc = Document('课程达标评价表2.docx')
tables = doc.tables
table = tables[0]
#用户名和密码
username = ""
password = ""
url = ''
browser = webdriver.Chrome()
#打卡网址
browser.get('https://sso.hbut.edu.cn:7002/cas/login?service=http%3A%2F%2Fportal.hbut.edu.cn%2Fportal%2Fhome%2Findex.do')
#输入用户名和密码
browser.find_element_by_id("username").send_keys(username)
browser.find_element_by_id("password").send_keys(password)
#控制开关---------------------
while True:
    print('请登录,按空格继续')
    a = input()
    if a == ' ':
        break
#------------------------------
browser.get('http://run.hbut.edu.cn/Account/sso')
browser.get('http://run.hbut.edu.cn/PutInGrade/Index/?SemesterName=20182&SemesterNameStr=2018%E5%AD%A6%E5%B9%B4%20%E7%AC%AC%E4%BA%8C%E5%AD%A6%E6%9C%9F')
#控制开关----------------------
while True:
    print('请选择学期,按空格继续')
    a = input()
    if a == ' ':
        break
#------------------------------
lss = browser.find_elements_by_tag_name('a')
hrefs = []
for ls in lss:
	if ls.text == '查看':
		print(ls.get_attribute('href'))
		hrefs.append(ls.get_attribute('href'))

xpath1 = '//*[@id="formDelete"]/table/tbody/tr/td[7]' #平时
xpath2 = '//*[@id="formDelete"]/table/tbody/tr/td[8]' #期末
xpath3 = '//*[@id="formDelete"]/table/tbody/tr/td[9]' #总评
xpath4 = '//*[@id="tableInIndex"]/table/tbody/tr/td[2]' #课名
xpath5 = '//*[@id="tableInIndex"]/table/tbody/tr/td[3]' #课程类型
xpath6 = '//*[@id="tableInIndex"]/table/tbody/tr/td[4]' #班级
xpath7 = '//*[@id="tableInIndex"]/table/tbody/tr/td[6]' #学分
cNames  = browser.find_elements_by_xpath(xpath4) #保存
cTypes = browser.find_elements_by_xpath(xpath5) #保存
classNos = browser.find_elements_by_xpath(xpath6) #保存
cName = []
cType = []
classNo = []
for href in hrefs:
    print(href)

for i in range(len(cNames)):
    print(cNames[i].text)
    cName.append(cNames[i].text)
    cType.append(cTypes[i].text)
    classNo.append(classNos[i].text)
    print(cName[i],cType[i],classNo[i])

course = []
for i in range(len(hrefs)):
    browser.get(hrefs[i])
    x1 = browser.find_elements_by_xpath(xpath1)
    x2 = browser.find_elements_by_xpath(xpath2)
    x3 = browser.find_elements_by_xpath(xpath3)
    s1 = []
    s2 = []
    s3 = []
    print(cName[i])
    for j in range(len(x3)):
        if x1[j].text=='':
            s1.append(0.0)
        else:
            s1.append(float(x1[j].text))
        if x2[j].text=='':
            s2.append(0.0)
        else:
            s2.append(float(x2[j].text))
        if x3[j].text=='':
            s3.append(0.0)
        else:
            s3.append(float(x3[j].text))
#        s2.append(float(x2[j].text))
#        s3.append(float(x3[j].text))
        print(s1[j], s2[j], s3[j])
    print('@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@')
#计算并生成文档
    #关键数据
    daily = 0.3 #平时
    final = 0.7 #期末
    total = len(s3)
    mean1 = np.mean(s1)
    mean2 = np.mean(s2)
    mean3 = np.mean(s3)
    std2 = np.std(s2)
    std3 = np.std(s3)
    gradeNo2 = [0,0,0,0,0,0]
    gradeNo3 = [0,0,0,0,0,0]
    for num2 in range(len(s2)):
        if s2[num2] >= 90:
            gradeNo2[0] = gradeNo2[0]+1
        elif s2[num2] >= 80:
            gradeNo2[1] = gradeNo2[1]+1
        elif s2[num2] >= 70:
            gradeNo2[2] = gradeNo2[2] + 1
        elif s2[num2] >= 60:
            gradeNo2[3] = gradeNo2[3]+1
        elif s2[num2] >= 30:
            gradeNo2[4] = gradeNo2[4]+1
        else:
            gradeNo2[5] = gradeNo2[5]+1
    for num3 in range(len(s3)):
        if s3[num3] >= 90:
            gradeNo3[0] = gradeNo3[0]+1
        elif s3[num3] >= 80:
            gradeNo3[1] = gradeNo3[1]+1
        elif s3[num3] >= 70:
            gradeNo3[2] = gradeNo3[2] + 1
        elif s3[num3] >= 60:
            gradeNo3[3] = gradeNo3[3]+1
        elif s3[num3] >= 30:
            gradeNo3[4] = gradeNo3[4]+1
        else:
            gradeNo3[5] = gradeNo3[5]+1

    table.cell(0,2).text = str(cName[i])
    table.cell(0,10).text = cType[i]
    table.cell(0,24).text = classNo[i]
    table.cell(1, 2).text = 'name'
    table.cell(1, 8).text = 'name'
    table.cell(1, 24).text = 'time'
    table.cell(4, 19).text = str(len(s3))
    table.cell(4, 25).text = str(len(s3))

    table.cell(7, 24).text = '平均分:' + str('%.2f' % mean2) + '\n\n标准差:' + str("%.2f" % std2)
    #目标达成度评价
    table.cell(16, 16).text = str(daily*100)
    table.cell(16, 21).text = str('%.2f' % mean2)
    table.cell(16, 25).text = str('%.2f' % (mean3 * daily / 100))
    table.cell(16, 24).text = table.cell(16, 25).text

    table.cell(11, 24).text = '平均分:' + str('%.2f' % mean3) + '\n\n标准差:' + str("%.2f" % std3)
    table.cell(17, 16).text = str(final * 100)
    table.cell(17, 21).text = str('%.2f' % mean1)
    table.cell(17, 25).text = str('%.2f' % (mean1*final/100))
    table.cell(17, 24).text = table.cell(17, 25).text

    table.cell(18, 4).text = str('%.2f' % (mean3/100))+'/0.7'

#各分数段人数
    table.cell(8, 5).text = str(gradeNo2[0])
    table.cell(8, 7).text = str(gradeNo2[1])
    table.cell(8, 9).text = str(gradeNo2[2])
    table.cell(8, 11).text = str(gradeNo2[3])
    table.cell(8, 14).text = str(gradeNo2[4])
    table.cell(8, 18).text = str(gradeNo2[5])
    if total >0:
        table.cell(9, 5).text = str("%.2f" % (gradeNo2[0] / total * 100))
        table.cell(9, 7).text = str("%.2f" % (gradeNo2[1] / total * 100))
        table.cell(9, 9).text = str("%.2f" % (gradeNo2[2] / total * 100))
        table.cell(9, 11).text = str("%.2f" % (gradeNo2[3] / total * 100))
        table.cell(9, 14).text = str("%.2f" % (gradeNo2[4] / total * 100))
        table.cell(9, 18).text = str("%.2f" % (gradeNo2[5] / total * 100))

        table.cell(13, 5).text = str("%.2f" % (gradeNo3[0] / total * 100))
        table.cell(13, 7).text = str("%.2f" % (gradeNo3[1] / total * 100))
        table.cell(13, 9).text = str("%.2f" % (gradeNo3[2] / total * 100))
        table.cell(13, 11).text = str("%.2f" % (gradeNo3[3] / total * 100))
        table.cell(13, 14).text = str("%.2f" % (gradeNo3[4] / total * 100))
        table.cell(13, 18).text = str("%.2f" % (gradeNo3[5] / total * 100))
    table.cell(12, 5).text = str(gradeNo3[0])
    table.cell(12, 7).text = str(gradeNo3[1])
    table.cell(12, 9).text = str(gradeNo3[2])
    table.cell(12, 11).text = str(gradeNo3[3])
    table.cell(12, 14).text = str(gradeNo3[4])
    table.cell(12, 18).text = str(gradeNo3[5])
    doc.save(str(classNo[i]+'_'+cName[i])+'.docx')
    time.sleep(1)



