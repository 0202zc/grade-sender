'''
大连大学成绩查询助手V1.7.190216
Code By ZC Liang
2018.6.6
Completed on 2019.2.16
'''

import getpass
import http.cookiejar
import os
import pickle
import platform
import random
import re
import smtplib
import subprocess
import sys
import time
import urllib.parse
import urllib.request
from email.mime.text import MIMEText
from email.utils import formataddr

import bs4
import numpy as np
import pandas as pd
import prettytable as pt
import pymysql
import requests
from aip import AipOcr
from bs4 import BeautifulSoup
from PIL import Image
from prettytable import PrettyTable
from requests import ReadTimeout
from requests import ConnectionError

my_sender = '发件人邮箱账号'  # 发件人邮箱账号
my_pass = '发件人邮箱密码(当时申请smtp给的口令)'  # 发件人邮箱密码(当时申请smtp给的口令)
# my_user='收件人邮箱账号'      # 收件人邮箱账号
email_send_to = ''  # 收件人邮箱账号

DstDir = os.getcwd()
searchCount = 0  # 查询次数
count = 0  # 循环计数
scorenum = 0  # 成绩条数
score = []
scorenp = np.array(score)
makeup_course_num = 0  # 重修课程数目
makeup_course_flag = -1  # 重修课程数目下标
courseList = []  # 选课情况查询列表
required_course_num = 0  # 本学期必修课总数

# 准备Cookie和opener，因为cookie存于opener中，所以以下所有网页操作全部要基于同一个opener
cookie = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(
    urllib.request.HTTPCookieProcessor(cookie))

final_url = ""  # 头 + 随机编码 + default2.aspx
final_url_head = ""
url_head = "202.199.155." + str(random.randint(33, 37))  # 随机产生网址

ddlxn = ""
ddlxq = ""

""" 你的 APPID AK SK """
APP_ID = '你的APPID'
API_KEY = '你的API_KEY'
SECRET_KEY = '你的SECRET_KEY'

client = AipOcr(APP_ID, API_KEY, SECRET_KEY)

""" 读取图片 """


def get_file_content(filePath):
    with open(filePath, 'rb') as fp:
        return fp.read()


# 判断操作系统类型


def getOpeningSystem():
    return platform.system()


# 判断操作系统类型


def getOpeningSystem():
    return platform.system()


# 判断是否联网


def isConnected():
    try:
        response = requests.get('http://' + url_head, timeout=1)
        if response.status_code == 200:
            return True
        else:
            print('网络检测错误status_code: ' + response.status_code, 'http://' + url_head)
            return False
    except (ConnectionError, ReadTimeout):
        print('无网络连接。')


# 获取重定向编码


def check_for_redirects(url):
    r = requests.head(url)
    if r.ok:
        return r.headers['location']
    else:
        return '[no redirect]'


#   图像转换并识别


def image_util(img):
    new_im = img.convert("RGB")  # 将验证码图片转换成24位图片
    new_im.save('' + DstDir + '\\ScoreHelper\\CheckCode1.jpg')  # 将24位图片保存到本地

    arr = np.array(Image.open('' + DstDir + '\\ScoreHelper\\CheckCode1.jpg').convert("L"))

    b = 255 - arr
    im = Image.fromarray(b.astype('uint8'))  # 翻转

    # d = 255 * (arr / 255) ** 2
    # im = Image.fromarray(d.astype('uint8'))  # 灰度

    #  此处验证过，翻转比灰度识别率更高
    im.save('' + DstDir + '\\ScoreHelper\\CheckCode2.jpg')


#   验证码识别
def code_recognition():
    try:
        #   调用百度云识别验证码
        result = client.basicAccurate(get_file_content('' + DstDir + '\\ScoreHelper\\CheckCode2.jpg'))
        word = result.get('words_result')
        res = ""
        if len(word):
            res = re.findall('[a-zA-Z0-9]+', word[0].get('words'))[0]
        elif len(res) > 4:  # 教务系统所有的验证码都是四位的，若大于四位，则挑选前四位
            res = res[0:4]
        return res
    except Exception as e:
        print(e)


#   登陆


def login():
    # 构造表单
    params = {
        'txtUserName': sid,
        'Textbox1': '',
        'Textbox2': spwd,
        'RadioButtonList1': '学生',
        'Button1': '',
        'lbLanguage': '',
        'hidPdrs': '',
        'hidsc': '',
    }

    #   获取验证码
    res = opener.open(final_url_head + '/checkcode.aspx').read()
    with open('' + DstDir + '\\ScoreHelper\\CheckCode.jpg', 'wb') as file:
        file.write(res)
    img = Image.open('' + DstDir + '\\ScoreHelper\\CheckCode.jpg')

    #   图片处理
    image_util(img)

    # img.show()

    print('验证码识别结果：' + code_recognition())
    vcode = code_recognition()

    # img.close()

    params['txtSecretCode'] = vcode

    #   获取ViewState
    response = urllib.request.urlopen('http://' + url_head + '/')
    html = response.read().decode('gb2312')
    viewstate = re.search(
        '<input type="hidden" name="__VIEWSTATE" value="(.+?)"', html)
    params['__VIEWSTATE'] = viewstate.group(1)
    #   尝试登陆
    loginurl = final_url
    print("\n本次登录所用网址为：" + loginurl + "\n")
    data = urllib.parse.urlencode(params).encode('gb2312')
    response = opener.open(loginurl, data)
    if response.geturl() == final_url:
        print('登陆失败，可能是姓名，学号，密码或验证码填写错误！')
        return False
    else:
        return True


#   获取本学期必修课数目


def get_RequiredCourse_num():
    global required_course_num

    print("正在查询本学期必修课数目...")
    #   构造url
    url = ''.join([
        final_url_head + '/xsxkqk.aspx',
        '?xh=',
        sid,
        '&xm=',
        urllib.parse.quote(sname),
        '&gnmkdm=N121615',
    ])
    #   构建查询学生选课情况表单
    params = {
        'ddlxn': ddlxn,
        'ddlxq': ddlxq,
    }

    #   构造Request对象，填入Header，防止302跳转，获取新的View_State
    req = urllib.request.Request(url)
    req.add_header('Referer', final_url)
    req.add_header('Origin', 'http://' + url_head + '/')
    req.add_header(
        'User-Agent',
        'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36')
    response = opener.open(req)
    html = response.read().decode('gb2312')
    viewstate = re.search(
        '<input type="hidden" name="__VIEWSTATE" value="(.+?)"', html)
    params['__VIEWSTATE'] = viewstate.group(1)

    #   查询所有成绩
    req = urllib.request.Request(
        url, urllib.parse.urlencode(params).encode('gb2312'))
    req.add_header('Referer', final_url)
    req.add_header('Origin', 'http://' + url_head + '/')
    response = opener.open(req)
    soup = BeautifulSoup(response.read().decode('gb2312'), 'html.parser')
    html = soup.find('table', class_='datelist')

    #   指定要输出的列，原网页的表格列下标从0开始
    #   用于标记是否是遍历第一行
    flag = True
    #   根据DOM解析所要数据，首位的each是NavigatableString对象，其余为Tag对象
    #   遍历行
    counter = 0
    for each in html:
        columnCounter = 0
        column = []

        if type(each) == bs4.element.NavigableString:
            pass
        else:
            #   遍历列
            for item in each.contents:
                if item != '\n':
                    if counter > 0 and columnCounter == 3:
                        courseList.append(str(item.contents[0]).strip())
                    columnCounter += 1
            if flag:
                flag = False
            counter += 1

    for each in courseList:
        if each == "必修课程":
            required_course_num += 1


#   获取成绩


def getScore():
    global searchCount
    global scorenum
    global scorenp
    global ddlxn
    global ddlxq
    score = []

    #   构造url
    url = ''.join([
        final_url_head + '/xscjcx_dq.aspx',
        '?xh=',
        sid,
        '&xm=',
        urllib.parse.quote(sname),
        '&gnmkdm=N121605',
    ])
    #   构建查询全部成绩表单
    params = {
        'ddlxn': ddlxn,  # 全部为 %C8%AB%B2%BF
        'ddlxq': ddlxq,
        'btnCx': '查询',
    }

    #   构造Request对象，填入Header，防止302跳转，获取新的View_State
    req = urllib.request.Request(url)
    req.add_header('Referer', final_url)
    req.add_header('Origin', 'http://' + url_head + '/')
    req.add_header(
        'User-Agent',
        'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36')
    response = opener.open(req)
    html = response.read().decode('gb2312')
    viewstate = re.search(
        '<input type="hidden" name="__VIEWSTATE" value="(.+?)"', html)
    params['__VIEWSTATE'] = viewstate.group(1)

    #   查询所有成绩
    req = urllib.request.Request(
        url, urllib.parse.urlencode(params).encode('gb2312'))
    req.add_header('Referer', final_url)
    req.add_header('Origin', 'http://' + url_head + '/')
    response = opener.open(req)
    soup = BeautifulSoup(response.read().decode('gb2312'), 'html.parser')
    html = soup.find('table', class_='datelist')

    print("执行第" + str(searchCount) + "次查询：")
    print('你的所有成绩如下：')

    #   指定要输出的列，原网页的表格列下标从0开始
    outColumn = [3, 4, 6, 7, 9, 11, 13]
    #   用于标记是否是遍历第一行
    flag = True
    #   根据DOM解析所要数据，首位的each是NavigatableString对象，其余为Tag对象
    #   遍历行
    for each in html:
        columnCounter = 0
        column = []

        if type(each) == bs4.element.NavigableString:
            pass
        else:
            #   遍历列
            for item in each.contents:
                if item != '\n':
                    if columnCounter in outColumn:
                        #   要使用str转换，不然陷入copy与deepcopy的无限递归
                        column.append(str(item.contents[0]).strip())
                    columnCounter += 1
            if flag:
                table = PrettyTable(column)
                flag = False
            else:
                table.add_row(column)
            score.extend([column])
    searchCount += 1
    scorenp = np.array(score)
    #   table.set_style(pt.PLAIN_COLUMNS)

    print(table)
    print("分条统计：")
    scorenum = sendScore(table)
    print("成绩数目: " + str(scorenum) + "条")


def sendScore(table):
    global scorenum
    global count
    global email_send_to
    global scorenp
    for i in table:
        print(i.get_string())
        count += 1

    if count > scorenum:
        try:
            scorenum = count

            # 文本模式
            # context = i.get_string().replace("+"," ")
            # context = context.replace("-"," ")
            # context = context.replace("2017 2018","2017-2018")
            # if(scorenum == 1):
            #     msg=MIMEText("有成绩下来了：" + context,'plain','utf-8')
            # else:
            #     msg=MIMEText("又有成绩下来了：" + context,'plain','utf-8')
            # msg = prettyScore()

            # html格式
            msg = prettyScore()

            # 括号里的对应发件人邮箱昵称、发件人邮箱账号
            msg['From'] = formataddr(["1115810371@qq.com", my_sender])
            # 括号里的对应收件人邮箱昵称、收件人邮箱账号
            msg['To'] = formataddr([email_send_to, email_send_to])

            if count == required_course_num:
                msg['Subject'] = "第" + str(count) + "次成绩推送加平均绩点"
            else:
                msg['Subject'] = "第" + str(count) + "次成绩推送"  # 邮件的主题，也可以说是标题

            # 发件人邮箱中的SMTP服务器，端口是465
            server = smtplib.SMTP_SSL("smtp.qq.com", 465)
            server.login(my_sender, my_pass)  # 括号中对应的是发件人邮箱账号、邮箱密码
            # 括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
            server.sendmail(my_sender, [email_send_to, ], msg.as_string())
            server.quit()  # 关闭连接
            print("发送成功，请注意在此邮箱查收：" + email_send_to)
        except Exception as e:
            print(e)
            print("发送失败！")
    count = 0
    if scorenum != required_course_num:
        print("程序休息中...（按'Ctrl C'结束）")
        time.sleep(1200)  # 二十分钟查一次
    return scorenum


#   接收构造成功的表格
def prettyScore():
    global scorenp
    try:
        # context = MIMEText(html,_subtype='html',_charset='utf-8')  #解决乱码
        msg = MIMEText(str(htmlText(scorenum)), "html", "gb2312")
    except Exception as e:
        print(e)
    return msg


#   构造邮件内容：成绩表格
def htmlText(scorenum):
    global required_course_num

    html = """

                        <table color="CCCC33" width="800" border="1" cellspacing="0" cellpadding="5" text-align="center">

                                <tr>

                                        <td>课程名称</td>

                                        <td>课程性质</td>

                                        <td>学分</td>

                                        <td>平时成绩</td>

                                        <td>期末成绩</td>

                                        <td>成绩</td>

                                </tr>   


                    """ + addtrs(scorenum) + """
                        </table>
                    """

    #   最后一次推送时计算GPA并与成绩表格一起推送
    if scorenum == required_course_num:
        html += """
                        <br/>
                        <div class='gpa_text' style='font-size: 25px;font-style: italic;'>-->平均绩点：%s <--</div>
                    """ % (getGPA()) + """
                        <br/>
                        <div class='end_words' style='font-size: 20px;'>本学期考试成绩查询完成！</div>
                    """
    return html


#   在发送的表格里添加成绩行
def addtrs(scorenum):
    global scorenp
    i = 1
    array = []
    while i <= scorenum:
        trs = '''
                                    <tr>   

                                            <td>%s </td>

                                            <td>%s </td>

                                            <td>%s </td>

                                            <td>%s </td>

                                            <td>%s </td>

                        ''' % (scorenp[i][0], scorenp[i][1], scorenp[i][2], scorenp[i][3], scorenp[i][4])
        if (scorenp[i][5].isalpha() and scorenp[i][5] == "A") or (scorenp[i][5].isdigit() and int(scorenp[i][5]) >= 90):
            #   等级A和90以上的成绩标记为绿色
            trs += '<td style="color:springgreen;">'

        elif (scorenp[i][5].isalpha() and scorenp[i][5] == "F") or (
                scorenp[i][5].isdigit() and int(scorenp[i][5]) < 60):
            #   不及格的成绩标记为红色
            trs += '<td style="color:red;">'

        else:
            #   普通成绩不标记
            trs += '<td>'

        trs += '''
                            %s </td>

                    </tr>
        ''' % (scorenp[i][5])
        array.append(trs)
        i += 1
    s = ""
    for x in array:
        s += str(x)
    return s


#   计算GPA
def getGPA():
    global scorenp
    global scorenum
    global makeup_course_num
    global makeup_course_flag

    sc = []
    GPAlist = []
    i = 1
    j = 0
    coursenum = 0

    while i <= scorenum:
        if scorenp[i][1] != "必修课程" or scorenp[i][6] == "是":
            #   排除非必修课以及重修课
            makeup_course_num += 1
            i += 1
            continue
        else:
            #   有些成绩是等级，需要转换为数字
            if scorenp[i][5].isalpha() and scorenp[i][5] != "F":
                sc.append(745 - 10 * ord(scorenp[i][5]))  # 计算式子：x - (x - A) + 10 * (D - x) 即 745 - 10 * x
            elif scorenp[i][5] == "F":
                sc.append(0)
            else:
                sc.append(int(scorenp[i][5]))

            if int(sc[j]) < 60:
                #   不及格的科目绩点为0
                GPAlist.append(0)
            else:
                #   计算单科绩点
                GPAlist.append((int(sc[j]) - 50) / 10 * float(scorenp[i][2]))
            i += 1
            j += 1
            coursenum += 1

    i = 1
    j = 0
    sum = 0
    scoresum = 0

    while i <= scorenum:
        if scorenp[i][1] != "必修课程" or scorenp[i][6] == "是":
            i += 1
            continue
        sum += GPAlist[j]
        scoresum += float(scorenp[i][2])
        j += 1
        i += 1
    GPA = sum / scoresum
    print("平均绩点：" + str(GPA))
    return GPA


#   根据当前日期设置查询学期
def setSemester():
    global ddlxn
    global ddlxq

    try:
        localtime = time.localtime(time.time())  # 获取当前日期

        #   第一学期是从当年9月到次年2月，第二学期则是从当年3月到8月
        if (int((localtime.tm_mon) >= 9 and int(localtime.tm_mon) <= 12) or (
                int(localtime.tm_mon) >= 1 and int(localtime.tm_mon) <= 2)):
            # if (str(localtime.tm_year) == "2020" and int((localtime.tm_mon) >= 7)):
            #     print("您已毕业，无须监控成绩！")
            #     sys.exit(0)
            if (int(localtime.tm_mon) >= 1 and int(localtime.tm_mon) <= 2):
                ddlxn = str(localtime.tm_year - 1) + '-' + str(int(localtime.tm_year))
            else:
                ddlxn = str(localtime.tm_year) + '-' + str(int(localtime.tm_year) + 1)
            ddlxq = '1'
        else:
            ddlxn = str(int(localtime.tm_year) - 1) + '-' + str(localtime.tm_year)
            ddlxq = '2'

    except Exception as e:
        print(e)


if __name__ == '__main__':
    setSemester()

    try:
        searchCount = 1
        print('欢迎使用大连大学成绩查询助手！')
        print('正在检查网络...')
        if isConnected():
            with open(r'' + DstDir + '\\ScoreHelper\\uinfo.bin', 'rb') as file:
                udick = pickle.load(file)
                sname = udick['sname']
                sid = udick['sid']
                spwd = udick['spwd']
                email_send_to = udick['email_send_to']

            #   构造登录地址
            final_url = 'http://' + url_head + \
                        check_for_redirects('http://' + url_head + '/default2.aspx')
            final_url_head = final_url[0:48]

            loginCount = 0
            while not login():
                if loginCount > 3:
                    #   超过三次未登录自动更换网址
                    url_head = "202.199.155." + str(random.randint(33, 37))
                    final_url = 'http://' + url_head + \
                                check_for_redirects('http://' + url_head + '/default2.aspx')
                    final_url_head = final_url[0:48]
                    loginCount = 0
                loginCount += 1
                print("正在等待重试...")
                time.sleep(3)
                continue

            get_RequiredCourse_num()
            getScore()
            counter = 0
            while scorenum <= required_course_num:
                counter += 1
                if scorenum == required_course_num:
                    print("本学期成绩查询完成！")
                    break
                if counter > 0:
                    getScore()
    except FileNotFoundError:
        # if os.path.exists(r'' + DstDir + '\\ScoreHelper'):
        #     os.remove(r'' + DstDir + '\\ScoreHelper')
        os.mkdir(r'' + DstDir + '\\ScoreHelper')  # 注：针对Windows目录结构
        print('这是你第一次使用，请按提示输入信息，以后可不必再次输入~')
        sid = input('请输入学号：')
        sname = input('请输入姓名：')
        # 隐藏密码
        # spwd = getpass.getpass('请输入密码：')
        spwd = input('请输入密码：')
        email_send_to = input('请输入要将成绩发送到的邮箱地址：')
        udick = {'sname': sname, 'sid': sid,
                 'spwd': spwd, 'email_send_to': email_send_to}
        file = open(r'' + DstDir + '\\ScoreHelper\\uinfo.bin', 'wb')
        pickle.dump(udick, file)
        file.close()
        final_url = 'http://' + url_head + \
                    check_for_redirects('http://' + url_head + '/default2.aspx')
        final_url_head = final_url[0:48]

        #   登录失败，重试
        while not login():
            sname = input('请输入姓名：')
            sid = input('请输入学号：')
            # spwd = getpass.getpass('请输入密码：')
            spwd = input('请输入密码：')
            email_send_to = input('请输入要将成绩发送到的邮箱地址：')
            udick = {'sname': sname, 'sid': sid,
                     'spwd': spwd, 'email_send_to': email_send_to}
            file = open(r'' + DstDir + '\\ScoreHelper\\uinfo.bin', 'wb')
            pickle.dump(udick, file)
            file.close()
            final_url = 'http://' + url_head + \
                        check_for_redirects('http://' + url_head + '/default2.aspx')
            final_url_head = final_url[0:48]
        get_RequiredCourse_num()
        getScore()
        counter = 0
        while scorenum <= required_course_num:
            counter += 1
            if scorenum == required_course_num:
                print("本学期成绩查询完成！")
                break
            if counter > 0:
                getScore()
            print(scorenum)

    except subprocess.CalledProcessError:
        print("网络连接不正常！请检查网络！")
    except Exception as e:
        print(e)
        print("失败！可能是你没有完成教学评价！没有完成教学评价则无法查看成绩！或用户中途取消或网络故障。")
    finally:
        # if os.path.exists(r'' + DstDir + '\\ScoreHelper\\CheckCode.jpg'):
        #     os.remove(r'' + DstDir + '\\ScoreHelper\\CheckCode.jpg')
        print("程序将在3秒后退出...")
        time.sleep(3)
