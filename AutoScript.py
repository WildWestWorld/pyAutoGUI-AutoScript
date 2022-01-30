import os

import pyautogui
import time
import xlrd
import pyperclip
from pynput.mouse import Button, Controller

#定义鼠标事件

def mouseClick(clickTimes,operateType,imgFilePath,reTry):
    breakTry=20;
    #循环一次
    if reTry == 1:
        #一直循环直到break
        while breakTry:

            # 1.获取两张图的所有像素点的rgb
            # 2.对比两张图的全部像素点是否相同（或者偏差在固定范围内）
            # 3.符合条件返回中心坐标

            # 把我们放在目录下的图片，根据当前屏幕进行比对，把图片与屏幕比对成功的位置返回回来
            #locateCenterOnScreen()函数会返回图片在屏幕上的中心XY轴坐标值：
            # confidence=0.9 精确度设为0.9更科学，既能保证不会找错对象，又不会因为默认的精确度太苛刻，明明对象存在代码却找不到而返回None。
            location = pyautogui.locateCenterOnScreen(imgFilePath, confidence=0.9)
            print(location)
            breakTry -= 1
            if location is not None:
                # x,y为鼠标坐标，
                # lick为点击几次，
                # interval为每次点击间隔时间，
                # duration为执行此次动作设置时间，
                # 0就是立即执行，
                # button有几个选项默认是左键，- ``LEFT``, ``MIDDLE``, ``RIGHT``, ``PRIMARY``, or ``SECONDARY``.
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.15,duration=0.1,button=operateType)
                break
            print("未找到匹配图片,0.1秒后重试")
            #暂缓0.1s，为了保护系统，免得因为死循环导致内存溢出
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(imgFilePath,confidence=0.9)
            if location is not None:
                print(location)
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=operateType)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(imgFilePath,confidence=0.9)
            if location is not None:
                print(location)
                pyautogui.click(location.x,location.y,clicks=clickTimes,interval=0.2,duration=0.2,button=operateType)
                print("重复")
                i += 1
            time.sleep(0.1)




# 数据检查
def checkData(commandTable):
    #设置一个检验成功的标识allRight，如果allRight=Ture那么存在有效数据，且每行数据都无误 allRight=False意味着excel中的数据有误，或者为空
    allRight = True



    #sheet.nrows 获取该sheet中的有效行数
    #commandTable.nrows < 2 表中的有效行数小于2 ，也就是小于等于1
    #因为第一行是用于说明的，并不算做命令，所以这样用来标识excel中没有 有效命令
    if commandTable.nrows <= 1:
        print("excel中的数据为空")

        #allRight =false说明数据有误
        allRight = False
        #返回标识符AllRight
        return allRight

    #通过数据为空的检验，现在开启逐行校验
    #每行数据检查
    #值得一说的是commandTable.nrows计数是从1开始的，假设commandTable.nrows是4
    #range(1,4)就会循环4次 i也是从1开始的
    for i in range(1, commandTable.nrows):
        # 第1列 命令类型检查
        #commandTable.row(1) 打印出来是第二行的所有数据 例如#commandTable.row(1) 打印出来是第二行的所有数据 例如
        #返回回来是个数组,  commandTable.row(i)[0] 代表数组中的第一个元素也就是 第i+1行第一列中的元素 ，第一列我们一般写的都是命令的类型
        commandType = commandTable.row(i)[0]
        # ctype函数来自于xlrd模块
        #  ctype 会得到的值 : 0 empty  ,1 string  , 2 number, 3 date, 4 boolean, 5 error
        # commandType.ctype != 2 就是说commandType数据类型不是数字
        if commandType.ctype != 2 :
            print('第', i+1, "行,第1列数据有毛病")
            allRight = False
        # 第2列 内容检查
        commandTargetValue = commandTable.row(i)[1]

        # 读图点击类型指令，内容必须为字符串类型

        # ctype函数来自于xlrd模块
        #  ctype 会得到的值 : 0 empty  ,1 string  , 2 number, 3 date, 4 boolean, 5 error
        # commandType.ctype != 2 就是说commandType数据类型不是数字

        #commandType.value ==1.0 or commandType.value == 2.0 or commandType.value == 3.0 =》当命令类型是鼠标操作时
        # if commandTargetValue.ctype != 1   =》  commandTargetValue数据类型不是字符串类型的话的就报错
        #当命令类型是鼠标操作时
        if commandType.value ==1.0 or commandType.value == 2.0 or commandType.value == 3.0:
            if commandTargetValue.ctype != 1:
                print('第', i+1, "行,第2列数据有毛病")
                allRight = False

        # 输入类型，内容不能为空
        #当命令类型是键盘输入时
        # if commandTargetValue.ctype != 1   =》  commandTargetValue数据类型不是字符串类型的话的就报错
        if commandType.value == 4.0:
            if commandTargetValue.ctype == 0:
                print('第',i+1,"行,第2列数据有毛病")
                allRight = False


        # 当命令类型是等待时
        # if commandTargetValue.ctype != 2   =》  commandTargetValue数据类型不是数字类型的话的就报错
        if commandType.value == 5.0:
            if commandTargetValue.ctype != 2:
                print('第',i+1,"行,第2列数据有毛病")
                allRight = False

        # 当命令类型是滚轮时
        if commandType.value == 6.0:
            # if commandTargetValue.ctype != 2   =》  commandTargetValue数据类型不是数字类型的话的就报错
            if commandTargetValue.ctype != 2:
                print('第',i+1,"行,第2列数据有毛病")
                allRight = False

    return allRight

#任务
def mainWork(imgFilePath):

    for i in range(1, commandTable.nrows):
        #取本行指令的操作类型
        commandType = commandTable.row(i)[0]
        #如果是类型1（鼠标左键点击一次）
        if commandType.value == 1.0:
            #取图片名称
            imgFilePath = commandTable.row(i)[1].value
            #默认重复次数为1IC
            reTry = 1
            if commandTable.row(i)[2].ctype == 2 and commandTable.row(i)[2].value != 0:
                reTry = commandTable.row(i)[2].value
            #mouseClick是自定义的函数函数
            mouseClick(1,"left",imgFilePath,reTry)
            print("单击左键",imgFilePath)
        #2代表双击左键
        elif commandType.value == 2.0:
            #取图片名称
            imgFilePath = commandTable.row(i)[1].value
            #取重试次数
            reTry = 1
            if commandTable.row(i)[2].ctype == 2 and commandTable.row(i)[2].value != 0:
                reTry = commandTable.row(i)[2].value
            mouseClick(2, "left", imgFilePath, reTry)
            print("双击左键",imgFilePath)
        #3代表右键
        elif commandType.value == 3.0:
            #取图片名称
            imgFilePath = commandTable.row(i)[1].value
            #取重试次数
            reTry = 1
            if commandTable.row(i)[2].ctype == 2 and commandTable.row(i)[2].value != 0:
                reTry = commandTable.row(i)[2].value
            mouseClick(1,"right",imgFilePath,reTry)
            print("右键",imgFilePath)
        #4代表输入
        elif commandType.value == 4.0:
            inputValue = commandTable.row(i)[1].value
            #pyperclip.copy相当于复制内容
            #pyperclip python中的剪切板
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl','v')
            time.sleep(0.5)
            print("输入:",inputValue)
        #5代表等待
        elif commandType.value == 5.0:
            #取图片名称
            waitTime = commandTable.row(i)[1].value
            time.sleep(waitTime)
            print("等待",waitTime,"秒")
        #6代表滚轮
        elif commandType.value == 6.0:
            #取图片名称
            scroll = commandTable.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动", int(scroll), "距离")
        #7代表键盘输入
        elif commandType.value == 7.0:
            #提取按键名字
            keyboardName = commandTable.row(i)[1].value
            pyautogui.press(keyboardName)
            print("按下了", keyboardName,'按键')
        #8代表点击鼠标当前位置
        elif commandType.value == 8.0:
            # 获取鼠标当前位置
            location = pyautogui.position()
            print(location)
            pyautogui.click(location.x, location.y, clicks=1, interval=0.2, duration=0.1, button="left")
            print("鼠标左键单击了一下当前位置")

        #9代表手动输入（非中文）
        elif commandType.value == 9.0:
            #提取字符
            typeValue = commandTable.row(i)[1].value
            #  每次键入的时间间隔
            secs_between_keys = 0.0
            pyautogui.typewrite(typeValue, interval = secs_between_keys)
            print("手动输入了", typeValue, '按键')
        #10代表自定义热键
        elif commandType.value == 10.0:
            # 提取字符
            hotkeyValue = commandTable.row(i)[1].value

            newInput = hotkeyValue.split(',')

            # tuple( iterable )
            #例子
            # >>>list1= ['Google', 'Taobao', 'Runoob', 'Baidu']
            # >>> tuple1=tuple(list1)
            # >>> tuple1
            # ('Google', 'Taobao', 'Runoob', 'Baidu')

            pyautogui.hotkey(*tuple(newInput))
            print("执行了自定义热键")
        #11代表点击鼠标当前位置
        elif commandType.value == 11.0:
            # 提取字符
            location = commandTable.row(i)[1].value
            locationArr = location.split(',')

            print(locationArr)
            pyautogui.click(int(locationArr[0]), int(locationArr[1]), clicks=1, interval=0.2, duration=0.1, button="left")
            print("鼠标左键单击了一下"+locationArr[0] +','+ locationArr[1])

if __name__ == '__main__':
    #自定文件名，必须是xls格式的，因为后续是读取xls里面的数据当做命令来给python执行的
    filename = 'brower'
    # filename = input("请输入要使用excel\n")
    #xlrd模块，专门用于读取excel的模块
    #xlrd.open_workbook 就是要打开excel文件
    commandExcel = xlrd.open_workbook(filename+".xls")
    #sheet_by_index(x) 就是打开excel里面的第几个工作表
    #工作表就是左下角的那个Sheet
    #此外还可以用sheet_by_name(sheet_name)#通过名称获取
    #此时commandTable就是工作表1的所有内容
    commandTable = commandExcel.sheet_by_index(0)
    print('JWestWorldK自动化脚本欢迎为您服务')
    #审核数据是否规范
    #checkData是我们自定义的函数#往上面看
    #checkData就是核查excel里面的数据是否是有效操作
    allRight = checkData(commandTable)
    if allRight:
        pattern = input('选择功能: 0.清除缓存 1.做一次 2.循环到死 3.自定义循环次数\n')


        if pattern == '1':
            #循环拿出每一行指令
            mainWork(commandTable)
        #else if
        elif pattern == '2':
            #死循环
            while True:
                mainWork(commandTable)
                #暂停0.1s
                time.sleep(0.1)
                print("等待0.1秒")

        elif pattern == '3':
            print("")
            count = 0
            times = input('输入需要循环的次数，务必输入正整数。\n')
            times = int(times)
            if count < times:
                while count < times:
                    count += 1
                    print("正在执行第", count, "次", "命令")
                    print("")
                    mainWork(commandTable)
                    time.sleep(0.1)
                    print("等待0.1秒")
                    print("")
                    print("已经完成第", count, "次", "命令")
                    print("——————————————————分割线——————————————————")
                    print("")
            else:
                print('输入有误或者已经退出!')
                os.system('pause')
                print("")
                print("——————————————————————————————————————————")

        elif pattern == '0':
            print("正清理缓存文件...")
            #清理C盘缓存
            os.system('@echo off & for /d %i in (%temp%\^_MEI*) do (rd /s /q "%i")>nul')
            exit("正在退出程序...")

    else:
        print('输入有误或者已经退出!')
