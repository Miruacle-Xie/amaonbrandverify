import csv
import json
import pyautogui
import time
import cv2
import os
import win32gui
import win32con
import comparepic
import random
import sys
import msvcrt

CONFIG = dict()


def imgAutoCick(tempFile, whatDo, debug=False, ex_para=None, shotFlag=False):
    '''
        temFile :需要匹配的小图
        whatDo  :需要的操作
                pyautogui.moveTo(w/2, h/2)# 基本移动
                pyautogui.click()  # 左键单击
                pyautogui.doubleClick()  # 左键双击
                pyautogui.rightClick() # 右键单击
                pyautogui.middleClick() # 中键单击
                pyautogui.tripleClick() # 鼠标当前位置3击
                pyautogui.scroll(10) # 滚轮往上滚10， 注意方向， 负值往下滑
        更多详情：https://blog.csdn.net/weixin_43430036/article/details/84650938
        debug   :是否开启显示调试窗口
    '''
    # 读取屏幕，并保存到本地
    pyautogui.screenshot(CONFIG["tempfoler"] + 'big.png')

    # 读入背景图片
    gray = cv2.imread(CONFIG["tempfoler"] + "big.png", 0)
    # 读入需要查找的图片
    img_template = cv2.imread(tempFile, 0)

    # 得到图片的高和宽
    w, h = img_template.shape[::-1]

    # 模板匹配操作
    res = cv2.matchTemplate(gray, img_template, cv2.TM_SQDIFF)

    # 得到最大和最小值得位置
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

    top = min_loc[0]
    left = min_loc[1]
    x = [top, left, w, h]
    print("x:{}".format(x))

    top_left = min_loc  # 左上角的位置
    bottom_right = (top_left[0] + w, top_left[1] + h)  # 右下角的位

    # 先移动再操作， 进行点击动作，可以修改为其他动作
    # print(pyautogui.position())
    if "rightborder" == ex_para:
        offset_w = w - w / 2 + 10
        offset_h = 0
    else:
        offset_w = 0
        offset_h = 0
    print(top + w / 2 + offset_w, left + h / 2 + offset_h)
    duration = 0.5 + random.uniform(0.1, 1.0)
    print(duration)
    pyautogui.moveTo(top + w / 2 + offset_w, left + h / 2 + offset_h, duration, pyautogui.easeOutQuad)
    if "scroll" in str(whatDo):
        pyautogui.scroll(ex_para)
    elif "donothing" in str(whatDo):
        whatDo(ex_para)
    else:
        whatDo()
        print(str(whatDo))
        if shotFlag:
            pyautogui.screenshot(CONFIG["tempfoler"] + 'pic-shot-before.png')
    # print(whatDo)

    if debug:
        # 读取原图
        img = cv2.imread(CONFIG["tempfoler"] + "big.png", 1)
        # 在原图上画矩形
        cv2.rectangle(img, top_left, bottom_right, (0, 0, 255), 2)
        # 调试显示
        img = cv2.resize(img, (0, 0), fx=0.5, fy=0.5, interpolation=cv2.INTER_NEAREST)
        cv2.imshow("processed", img)
        cv2.waitKey(0)
        # 销毁所有窗口
        cv2.destroyAllWindows()
    os.remove(CONFIG["tempfoler"] + "big.png")


def img2AutoCick(tempFile1, tempFile2, whatDo, debug=False, ex_para=None, shotFlag=False):
    '''
        temFile :需要匹配的小图
        whatDo  :需要的操作
                pyautogui.moveTo(w/2, h/2)# 基本移动
                pyautogui.click()  # 左键单击
                pyautogui.doubleClick()  # 左键双击
                pyautogui.rightClick() # 右键单击
                pyautogui.middleClick() # 中键单击
                pyautogui.tripleClick() # 鼠标当前位置3击
                pyautogui.scroll(10) # 滚轮往上滚10， 注意方向， 负值往下滑
        更多详情：https://blog.csdn.net/weixin_43430036/article/details/84650938
        debug   :是否开启显示调试窗口
    '''
    # 读取屏幕，并保存到本地
    pyautogui.screenshot(CONFIG["tempfoler"] + 'big.png')

    # 读入背景图片
    gray = cv2.imread(CONFIG["tempfoler"] + "big.png", 0)
    # 读入需要查找的图片1
    img_template1 = cv2.imread(tempFile1, 0)
    # 读入需要查找的图片2
    img_template2 = cv2.imread(tempFile2, 0)

    # 得到图片1的高和宽
    w1, h1 = img_template1.shape[::-1]
    # 得到图片2的高和宽
    w2, h2 = img_template2.shape[::-1]

    # 图1模板匹配操作
    res1 = cv2.matchTemplate(gray, img_template1, cv2.TM_SQDIFF)
    # 图2模板匹配操作
    res2 = cv2.matchTemplate(img_template1, img_template2, cv2.TM_SQDIFF)

    # 图1得到最大和最小值得位置
    min_val1, max_val1, min_loc1, max_loc1 = cv2.minMaxLoc(res1)
    # 图2得到最大和最小值得位置
    min_val2, max_val2, min_loc2, max_loc2 = cv2.minMaxLoc(res2)

    top1 = min_loc1[0]
    left1 = min_loc1[1]
    x1 = [top1, left1, w1, h1]
    print("x1:{}".format(x1))

    top2 = min_loc2[0]
    left2 = min_loc2[1]
    x2 = [top1 + top2, left1 + left2, w2, h2]
    print("x2:{}".format(x2))

    # 图1
    top_left1 = (x1[0], x1[1])  # 左上角的位置
    bottom_right1 = (top_left1[0] + w1, top_left1[1] + h1)  # 右下角的位
    print(top_left1, bottom_right1)
    # 图2
    top_left2 = (x2[0], x2[1])
    bottom_right2 = (top_left2[0] + w2, top_left2[1] + h2)  # 右下角的位
    print(top_left2, bottom_right2)

    # 先移动再操作， 进行点击动作，可以修改为其他动作
    # print(pyautogui.position())
    print(x2[0] + x2[2] / 2, x2[1] + x2[3] / 2)
    duration = 0.5 + random.uniform(0.1, 1.0)
    print(duration)
    pyautogui.moveTo(x2[0] + x2[2] / 2, x2[1] + x2[3] / 2, duration, pyautogui.easeOutQuad)
    if "scroll" in str(whatDo):
        pyautogui.scroll(ex_para)
    elif "donothing" in str(whatDo):
        whatDo(ex_para)
    else:
        whatDo()
        print(str(whatDo))
        if shotFlag:
            pyautogui.screenshot(CONFIG["tempfoler"] + 'pic-shot-before.png')
    # print(whatDo)

    if debug:
        # 读取原图
        img = cv2.imread(CONFIG["tempfoler"] + "big.png", 1)
        # 在原图上画矩形
        cv2.rectangle(img, top_left1, bottom_right1, (0, 0, 255), 2)
        # 在原图上画矩形
        cv2.rectangle(img, top_left2, bottom_right2, (0, 0, 255), 2)
        # 调试显示
        img = cv2.resize(img, (0, 0), fx=0.5, fy=0.5, interpolation=cv2.INTER_NEAREST)
        cv2.imshow("processed", img)
        cv2.waitKey(0)
        # 销毁所有窗口
        cv2.destroyAllWindows()
    os.remove(CONFIG["tempfoler"] + "big.png")


def donothing(para=None):
    print("donothing")
    return


def ifsamepic(pic1, pic2, mode):
    list = comparepic.runAllImageSimilaryFun(pic1, pic2)
    return list[mode]


def screenshotregion(screenshotfile, regionfile=None, confidence=0.9, mode="fullscreen"):
    if regionfile is not None:
        pos = pyautogui.locateOnScreen(regionfile, confidence=confidence)
        # print(pos)
    if "save&continue" == mode:
        pyautogui.screenshot(screenshotfile, region=(pos[0], pos[1], pos[2], pos[3]))
    else:
        pyautogui.screenshot(screenshotfile)
    return


def brandverify(brand):
    result = -3
    try:
        # img2AutoCick('F:\\JetBrains\\officeTools\\brandtest\\brandinput.png',
        #              'F:\\JetBrains\\officeTools\\brandtest\\brandinputblank.png', pyautogui.click, False)
        img2AutoCick(CONFIG["brandinput"], CONFIG["brandinputblank"], pyautogui.click, False)
        pyautogui.hotkey('Ctrl', 'a')
        time.sleep(random.uniform(0.1, 0.5))
        pyautogui.press('delete')
        time.sleep(random.uniform(0.1, 0.5))
        pyautogui.typewrite(brand)
        screenshotregion(CONFIG["tempfoler"] + 'verify-before.png')
        screenshotregion(CONFIG["tempfoler"] + 'savekey-before.png', CONFIG["savekey"], mode="save&continue")
        # imgAutoCick(CONFIG["cancelkey"], pyautogui.click, debug=False, ex_para="rightborder")
        duration = random.uniform(0.1, 0.5)
        print(duration)
        pyautogui.moveTo(CONFIG["pos_x"], CONFIG["pos_y"], duration, pyautogui.easeOutQuad)
        pyautogui.click()
        tmptime = time.time()
        while True:
            if time.time() - tmptime > 30:
                print("超时")
                return
            # 读取屏幕，并保存到本地
            screenshotregion(CONFIG["tempfoler"] + 'verify-tmp.png')
            print("{}\n".format(sys._getframe().f_lineno))
            errorvalue = ifsamepic(CONFIG["tempfoler"] + 'verify-before.png', CONFIG["tempfoler"] + 'verify-tmp.png', 3)

            if errorvalue < CONFIG["errorvalue"]:
                print("error")
                if errorvalue > 0.435:
                    result = -1
                else:
                    result = -1

                duration = random.uniform(0.1, 0.5)
                print(duration)
                pyautogui.moveTo(CONFIG["pos_x"], CONFIG["pos_y"], duration, pyautogui.easeOutQuad)
                pyautogui.click()
                break
            time.sleep(1)
            screenshotregion(CONFIG["tempfoler"] + 'savekey-tmp.png', CONFIG["savekey"], mode="save&continue")
            print("{}\n".format(sys._getframe().f_lineno))
            if ifsamepic(CONFIG["succ"], CONFIG["tempfoler"] + 'savekey-tmp.png', 3) > CONFIG["succvalue"]:
                print("succ")
                result = 0
                break
            time.sleep(1)
            print("等待跳转中...")
        os.remove(CONFIG["tempfoler"] + 'savekey-before.png')
        os.remove(CONFIG["tempfoler"] + 'savekey-tmp.png')
        os.remove(CONFIG["tempfoler"] + 'verify-before.png')
        os.remove(CONFIG["tempfoler"] + 'verify-tmp.png')
    except Exception as e:
        print(repr(e))
    return result


def configfunc():
    if os.path.exists("brandverify-config.json"):
        fd = open("brandverify-config.json")
    else:
        configpath = input("输入配置文件路径:\n")
        fd = open(configpath.replace("\'", "").replace("\"", ""))
    config = fd.read()
    fd.close()
    config = json.loads(config)
    return config


def main():
    global CONFIG
    CONFIG = configfunc()
    print(CONFIG)

    filepath = input("文件路径:\n")
    filepath = filepath.replace("\'", "").replace("\"", "")
    fd = open(filepath)
    f_csv = csv.reader(fd)
    brandlist = [list(filter(None, x)) for x in f_csv]
    fd.close()
    print(brandlist)
    # return
    handle = win32gui.FindWindow("Qt5QWindowIcon", "闪店云管家")
    win32gui.ShowWindow(handle, win32con.SW_SHOWMAXIMIZED)
    win32gui.SetForegroundWindow(handle)
    time.sleep(1)
    available = []
    invalidBrand = []
    invalidEAN = []
    abnormal = []

    time_start = time.time()
    # list = ["asf", "skam", "asdf", "asdff"]
    cnt = 0
    escflag = False
    try:
        for sublist in brandlist:
            for i in sublist:
                cnt += 1
                print("{}-{}:开始验证".format(cnt, i))
                result = brandverify(i)
                if result == -2:
                    print("EAN")
                    invalidEAN.append(i)
                elif result == -1:
                    print("不可用")
                    invalidBrand.append(i)
                elif result == 0:
                    print("可用")
                    available.append(i)
                else:
                    print("异常")
                    abnormal.append(i)
                if cnt // 100 > 0 and cnt % 100 == 0:
                    print(available)
                    print(invalidBrand)
                    print(invalidEAN)
                    print(abnormal)
                    parentpath = os.path.dirname(filepath) + "\\"
                    resultreport = parentpath + "\\report.csv"
                    fp = open(resultreport, 'a')
                    fp.write("{}前可用".format(cnt) + ',')
                    fp.write(",".join(available) + '\n')
                    fp.write("{}前不可用".format(cnt) + ',')
                    fp.write(",".join(invalidBrand) + '\n')
                    # fp.write("EAN" + ',')
                    # fp.write(",".join(invalidEAN) + '\n')
                    fp.write("{}前异常".format(cnt) + ',')
                    fp.write(",".join(abnormal) + '\n')
                    fp.close()
                start_time = time.time()
                while True:
                    if msvcrt.kbhit():
                        chr = msvcrt.getche()
                        if ord(chr) == 27:  # ESC
                            print("esc:{}".format((time.time() - start_time)))
                            keyIn = input("按回车确认退出, 输入其他键按回车取消退出\n")
                            if keyIn == "":
                                escflag = True
                            else:
                                escflag = False
                            break
                    if (time.time() - start_time) > 1:
                        break
                if escflag:
                    break
            if escflag:
                break
    except Exception as e:
        print(repr(e))
    time_end = time.time()
    # '''
    print(available)
    print(invalidBrand)
    print(invalidEAN)
    print(abnormal)
    parentpath = os.path.dirname(filepath) + "\\"
    resultreport = parentpath + "\\report.csv"
    fp = open(resultreport, 'a')
    fp.write("可用" + ',')
    fp.write(",".join(available) + '\n')
    fp.write("不可用" + ',')
    fp.write(",".join(invalidBrand) + '\n')
    # fp.write("EAN" + ',')
    # fp.write(",".join(invalidEAN) + '\n')
    fp.write("异常" + ',')
    fp.write(",".join(abnormal) + '\n')
    fp.close()
    # '''
    input("已生成报告, 耗时时间:{}, 平均耗时:{}, 按回车键结束".format(time_end - time_start, (time_end - time_start) / cnt))
    return


if __name__ == "__main__":
    # qqtest()
    main()
    # time.sleep(5)
    # sys.stdin.flush()
    # start_time = time.time()
    # while True:
    #     # if msvcrt.kbhit():
    #     chr = msvcrt.getche()
    #     if ord(chr) == 27:  # ESC
    #         print("esc:{}".format((time.time() - start_time)))
    #         break
    #     if (time.time() - start_time) > 10:
    #         print("timeout")
    #         break

    # available = ['1', "2", '3']
    # invalidBrand = ['4', '5', '6', '7']
    # invalidEAN = ['8', '9', '10', '11', '12']
    # abnormal = ['13', '14', '15', '16', '17']
    # resultreport = "aa.csv"
    # for i in range(1000):
    #     if i // 100 > 0 and i % 100 == 0:
    #         print(i)
    #         fp = open(resultreport, 'a')
    #         fp.write("{}前可用".format(i) + ',')
    #         fp.write(",".join(available) + '\n')
    #         fp.write("{}前不可用".format(i) + ',')
    #         fp.write(",".join(invalidBrand) + '\n')
    #         # fp.write("EAN" + ',')
    #         # fp.write(",".join(invalidEAN) + '\n')
    #         fp.write("{}前异常".format(i) + ',')
    #         fp.write(",".join(abnormal) + '\n')
    #         fp.close()
    # fd = open("测品牌-20221109.csv")
    # f_csv = csv.reader(fd)
    # list = [list(filter(None, i)) for i in f_csv]
    # print(list)
    # fd.close()
    #
    # for sublist in list:
    #     for i in sublist:
    #         print(i)
