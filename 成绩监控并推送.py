'''
大连大学成绩查询助手V3.5
Coded By Martin Huang
Code Changed By ZC Liang
2018.6.6
'''
import re
import urllib.request
import urllib.parse
import http.cookiejar
import bs4
import getpass
import pickle
import os
import platform
import subprocess
from bs4 import BeautifulSoup
from prettytable import PrettyTable
from PIL import Image
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
import time
import prettytable as pt
import pandas as pd
import numpy as np
import requests
import random
import sys


my_sender = '发件人邮箱账号'    # 发件人邮箱账号
my_pass = '发件人邮箱密码'              # 发件人邮箱密码(当时申请smtp给的口令)
# my_user = '收件人邮箱账号'      # 收件人邮箱账号，我这边发送给自己
email_send_to = ''                # 收件人邮箱账号

DstDir = os.getcwd()
searchCount = 0  # 查询次数
count = 0  # 循环计数
scorenum = 0  # 成绩条数
score = []
scorenp = np.array(score)
makeup_course_num = 0  # 重修课程数目
makeup_course_flag = -1  # 重修课程数目下标
all_score_num = 0

# 准备Cookie和opener，因为cookie存于opener中，所以以下所有网页操作全部要基于同一个opener
cookie = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(
    urllib.request.HTTPCookieProcessor(cookie))

final_url = ""      # 头 + 随机编码 + default2.aspx
final_url_head = ""
url_head = "202.199.155." + str(random.randint(33, 37))     # 随机产生网址

ddlxn = ""
ddlxq = ""

# 判断操作系统类型


def getOpeningSystem():
    return platform.system()

# 判断是否联网


def isConnected():
    userOs = getOpeningSystem()
    if userOs == "Windows":
        subprocess.check_call(
            ["ping", "-n", "2", url_head], stdout=subprocess.PIPE)
    else:
        subprocess.check_call(
            ["ping", "-c", "2", url_head], stdout=subprocess.PIPE)

# 获取重定向编码


def check_for_redirects(url):
    r = requests.head(url)
    if r.ok:
        return r.headers['location']
    else:
        return '[no redirect]'

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
    img.show()
    vcode = input('请输入验证码：')
    img.close()
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
        print('登陆失败，可能是姓名、学号、密码、验证码填写错误！')
        return False
    else:
        return True

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
        'btnCx': '+%B2%E9++%D1%AF+',
    }
    #   构造Request对象，填入Header，防止302跳转，获取新的View_State
    req = urllib.request.Request(url)
    req.add_header('Referer', final_url)
    req.add_header('Origin', 'http://' + url_head + '/')
    req.add_header(
        'User-Agent', 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36')
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

        if(type(each) == bs4.element.NavigableString):
            pass
        else:
            #   遍历列
            for item in each.contents:
                if(item != '\n'):
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

    if(count > scorenum):
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
            if(all_score_num == 10):
                msg['Subject'] = "第" + \
                    str(count) + "次成绩推送加平均绩点"  # 邮件的主题，也可以说是标题
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
            print("发送失败")
    count = 0
    print("程序休息中...（按'Ctrl C'结束）")
    time.sleep(1200)  # 二十分钟查一次
    return scorenum


def prettyScore():
    global scorenp
    try:
        # context = MIMEText(html,_subtype='html',_charset='utf-8')  #解决乱码
        msg = MIMEText(str(htmlText(scorenum)), "html", "gb2312")
    except Exception as e:
        print(e)
    return msg


def htmlText(scorenum):
    global all_score_num

    if(scorenum == all_score_num):
        html = """

                        <table color="CCCC33" width="800" border="1" cellspacing="0" cellpadding="5" text-align="center">

                                <tr>

                                        <td text-align="center">课程名称</td>

                                        <td text-align="center">课程性质</td>

                                        <td text-align="center">学分</td>

                                        <td text-align="center">平时成绩</td>

                                        <td text-align="center">期末成绩</td>

                                        <td text-align="center">成绩</td>

                                </tr>   


                    """ + addtrs(scorenum) + """
                        </table>
                        <div><h2>-->平均绩点：%s --<</h2></div>
                    """ % (getGPA())
    else:
        html = """

                        <table color="CCCC33" width="800" border="1" cellspacing="0" cellpadding="5" text-align="center">

                                <tr>

                                        <td text-align="center">课程名称</td>

                                        <td text-align="center">课程性质</td>

                                        <td text-align="center">学分</td>

                                        <td text-align="center">平时成绩</td>

                                        <td text-align="center">期末成绩</td>

                                        <td text-align="center">成绩</td>

                                </tr>   


                    """ + addtrs(scorenum)
    return html


def addtrs(scorenum):
    global scorenp
    i = 1
    array = []
    while(i <= scorenum):
        trs = '''
                        <tr>   

                                <td text-align="center">%s </td>

                                <td>%s </td>

                                <td>%s </td>

                                <td>%s </td>

                                <td>%s </td>

                                <td>%s </td>

                        </tr>
            ''' % (scorenp[i][0], scorenp[i][1], scorenp[i][2], scorenp[i][3], scorenp[i][4], scorenp[i][5])
        array.append(trs)
        i += 1
    s = ""
    for x in array:
        s += str(x)
    return s


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

    while(i <= scorenum):
        if(scorenp[i][1] != "必修课程" or scorenp[i][6] == "是"):
            makeup_course_num += 1
            i += 1
            continue
        else:
            if(scorenp[i][5] == "F"):
                sc.append(0)
            elif(scorenp[i][5] == "A"):
                sc.append(95)
            elif(scorenp[i][5] == "B"):
                sc.append(85)
            elif(scorenp[i][5] == "C"):
                sc.append(75)
            elif(scorenp[i][5] == "D"):
                sc.append(65)
            else:
                sc.append(int(scorenp[i][5]))

            if(int(sc[j]) < 60):
                GPAlist.append(0)
            else:
                GPAlist.append((int(sc[j]) - 50)/10*float(scorenp[i][2]))
            i += 1
            j += 1
            coursenum += 1

    i = 1
    j = 0
    sum = 0
    scoresum = 0

    while(i <= scorenum):
        if(scorenp[i][1] != "必修课程" or scorenp[i][6] == "是"):
            i += 1
            continue
        sum += GPAlist[j]
        scoresum += float(scorenp[i][2])
        j += 1
        i += 1
    GPA = sum/scoresum
    print("平均绩点：" + str(GPA))
    return GPA


if __name__ == '__main__':
    try:
        localtime = time.localtime(time.time())     # 获取当前日期
        if(int(localtime.tm_mon) >= 9 and int(localtime.tm_mon) <= 12):
            if(str(localtime.tm_year) == "2020"):
                print("您已毕业，无须监控成绩！")
                sys.exit(0)
            ddlxn = str(localtime.tm_year) + '-' + str(int(localtime.tm_year) + 1)
            ddlxq = '1'
        else:
            ddlxn = str(int(localtime.tm_year) - 1) + '-' + str(localtime.tm_year)
            ddlxq = '2'

        searchCount = 1
        print('欢迎使用大连大学成绩查询助手！')
        print('正在检查网络...')
        isConnected()
        with open(r'' + DstDir + '\\ScoreHelper\\uinfo.bin', 'rb') as file:
            udick = pickle.load(file)
            sname = udick['sname']
            sid = udick['sid']
            spwd = udick['spwd']
            email_send_to = udick['email_send_to']
        all_score_num = int(input('请输入本学期所有考试科目数目（包括重修课、公选课、体育课）【用于计算绩点】：'))
        final_url = 'http://' + url_head + \
            check_for_redirects('http://' + url_head + '/default2.aspx')
        final_url_head = final_url[0:48]
        while(not login()):
            continue
        while(scorenum <= all_score_num):
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
        all_score_num = int(input('请输入本学期所有考试科目数目（包括重修课、公选课、体育课）【用于计算绩点】：'))
        final_url = 'http://' + url_head + \
            check_for_redirects('http://' + url_head + '/default2.aspx')
        final_url_head = final_url[0:48]
        while(not login()):
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
            all_score_num = int(
                input('请输入本学期所有考试科目数目（包括重修课、公选课、体育课）【用于计算绩点】：'))
            final_url = 'http://' + url_head + \
                check_for_redirects('http://' + url_head + '/default2.aspx')
            final_url_head = final_url[0:48]
        while(scorenum <= all_score_num):
            if(scorenum == all_score_num):
                getScore()
                break
            getScore()
            print(scorenum)

    except subprocess.CalledProcessError:
        print("网络连接不正常！请检查网络！")
    except:
        print("失败！可能是你没有完成教学评价！没有完成教学评价则无法查看成绩！或用户中途取消或网络故障。")
    finally:
        # if os.path.exists(r'' + DstDir + '\\ScoreHelper\\CheckCode.jpg'):
        #     os.remove(r'' + DstDir + '\\ScoreHelper\\CheckCode.jpg')
        input('Done！请按任意键退出')
