# coding:utf8
import datetime
import os
import time


def doSth():
    # 把爬虫程序放在这个类里
    print('这个程序要开始疯狂的运转啦')
    try:
        if os.path.exists(path):
            command = path + '/battle.bat'
            os.system(command)
    except (IOError, Exception) as e:
        print(e)


if __name__ == '__main__':
    path = os.getcwd()

    count = 0
    timeCount = 0
    doSth()

    while timeCount < 27:
        timeCount += 1
        count += 1
        print("程序第",end="")
        print(count,end="")
        print("次执行")
        if timeCount == 27:
            timeCount = 0
            # doSth()
        # 每隔1小时检测一次
        time.sleep(3600)
