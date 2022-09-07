import os
import time
import datetime
from PIL import Image
import math
import operator
import functools
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import NoSuchElementException
from urllib.parse import quote
import requests
import win32gui
import win32con


def login(driver, username, password):
    iaaaUrl = 'https://iaaa.pku.edu.cn/iaaa/oauth.jsp'
    appName = quote('北京大学校内信息门户新版')
    redirectUrl = 'https://portal.pku.edu.cn/portal2017/ssoLogin.do'

    driver.get(f'{iaaaUrl}?appID=portal2017&appName={appName}&redirectUrl={redirectUrl}')
    driver.find_element(By.ID, "user_name").send_keys(username)
    time.sleep(1)
    driver.find_element(By.ID, "password").send_keys(password)
    time.sleep(1)
    driver.find_element(By.ID, "logon_button").click()
    WebDriverWait(driver, 5).until(ec.visibility_of_element_located((By.ID, 'all')))
    driver.find_element(By.CLASS_NAME, "btn").click()
    time.sleep(1)
    print("门户登录成功")


def stu_io(driver):
    driver.find_element(By.ID, "all").click()
    WebDriverWait(driver, 5).until(ec.visibility_of_element_located((By.ID, 'tag_s_stuCampusExEnReq')))
    driver.find_element(By.ID, "tag_s_stuCampusExEnReq").click()
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[-1])
    print("进入出入校填报界面")

    # 选择园区往返
    WebDriverWait(driver, 5).until_not(ec.visibility_of_element_located((By.CLASS_NAME, 'el-loading-mask')))
    time.sleep(3)
    driver.find_element(By.XPATH, "//span[text()=' 园区往返申请']").click()
    WebDriverWait(driver, 5).until_not(ec.visibility_of_element_located((By.CLASS_NAME, 'el-loading-mask')))
    time.sleep(10)
    # 点确定
    driver.find_element(By.XPATH, "//span[text()='确定']").click()
    time.sleep(1)
    # 加一步，提交图片
    def uploadPicture(pic_path):
        dialog = win32gui.FindWindow("#32770", u"文件上传")
        ComboBoxEx32 = win32gui.FindWindowEx(dialog, 0, "ComboBoxEx32", None)
        ComboBox = win32gui.FindWindowEx(ComboBoxEx32, 0, "ComboBox", None)
        Edit = win32gui.FindWindowEx(ComboBox, 0, "Edit", None)
        button = win32gui.FindWindowEx(dialog, 0, "Button", None)
        win32gui.SendMessage(Edit, win32con.WM_SETTEXT, None, pic_path)
        time.sleep(1)
        win32gui.SendMessage(dialog, win32con.WM_COMMAND, 1, button)
        time.sleep(1)
    driver.find_element(By.XPATH, "//span[text()='通信大数据行程卡']").click()
    time.sleep(1)
    xcm_path = os.path.abspath('.') + "\\" + datetime.datetime.today().__format__("%Y%m%d") + "_xcm.png"
    uploadPicture(xcm_path)
    driver.find_element(By.XPATH, "//span[text()='北京健康宝']").click()
    time.sleep(1)
    jkb_path = os.path.abspath('.') + "\\" + datetime.datetime.today().__format__("%Y%m%d") + "_jkb.png"
    uploadPicture(jkb_path)
    print("提交图片成功")
    # 如果没问题，别的都选好了
    driver.find_element(By.XPATH, "//span[text()='保存 ']").click()
    time.sleep(1)
    driver.find_element(By.XPATH, "//div[@class='el-message-box__btns']/button[2]").click()
    print("今日信息已提交")
    WebDriverWait(driver, 5).until_not(ec.visibility_of_element_located((By.CLASS_NAME, 'el-loading-mask')))
    time.sleep(1)


def check(driver):
    driver.find_element(By.XPATH, "//div[text()='返回']").click()
    WebDriverWait(driver, 5).until(ec.visibility_of_element_located((By.CLASS_NAME, 'el-card__body')))
    time.sleep(1)
    driver.find_element(By.XPATH, "//span[text()=' 申请历史']").click()
    WebDriverWait(driver, 5).until_not(ec.visibility_of_element_located((By.CLASS_NAME, 'el-loading-mask')))
    time.sleep(1)
    date_str = datetime.datetime.today().__format__("%Y%m%d")
    try:
        driver.find_element(By.XPATH, f"//div[text()=' {date_str} ']/../..//div[text()='审核通过']")
        print("检查通过")
        return True
    except NoSuchElementException:
        print("检查不通过")
        return False


def send_msg(title, content=""):
    # https://sct.ftqq.com/
    key = ""
    api = f"https://sctapi.ftqq.com/{key}.send"
    data = {"text": title, "desp": content}
    requests.post(api, data=data)

def imageContrast(img1, img2):
    image1 = Image.open(img1)
    image2 = Image.open(img2)

    h1 = image1.histogram()
    h2 = image2.histogram()

    result = math.sqrt(functools.reduce(operator.add,  list(map(lambda a, b: (a-b)**2, h1, h2)))/len(h1))
    return result


def getPicture():
    # 连接adb
    path = os.path.abspath('.')
    adb_path = path + "\\scrcpy-win64-v1.24\\adb.exe"
    res = os.popen(adb_path + " devices -l").read()
    if "手机型号" not in res:
        os.system(adb_path + " connect 192.168.100.150")
        res = os.popen(adb_path + " devices -l").read()
    if "手机型号" in res:
        print("adb连接成功")
    else:
        raise Exception("adb连接失败")

    pos_list = [(540, 1680), (250, 1100), (540, 1110), (830, 1100), (250, 1280), (540, 1280), (830, 1280), (250, 1450), (540, 1450), (830, 1450)]
    def swipe(dot1, dot2, delay, sleep_time):
        os.system(adb_path + " shell input swipe %d %d %d %d %d" % (dot1[0], dot1[1], dot2[0], dot2[1], delay))
        time.sleep(sleep_time)
    def tap(dot, sleep_time):
        os.system(adb_path + " shell input tap %d %d" % (dot[0], dot[1]))
        time.sleep(sleep_time)
    def pressHome(sleep_time):
        os.system(adb_path + " shell input keyevent 3")
        time.sleep(sleep_time)
    def pressPower(sleep_time):
        os.system(adb_path + " shell input keyevent 26")
        time.sleep(sleep_time)
    def toKeyPage():
        pressHome(sleep_time=1)
        swipe(pos_list[6], pos_list[4], delay=200, sleep_time=1)

    def getScreenShot(name):
        os.system(adb_path + " shell screencap -p /storage/emulated/0/MY/screencap/%s.png" % name)
        time.sleep(1)
        os.popen(adb_path + " pull /storage/emulated/0/MY/screencap/%s.png" % name)
        time.sleep(1)

    res = os.popen(adb_path + " shell dumpsys window policy").read()
    # 判断是否亮屏
    if "screenState=SCREEN_STATE_OFF" in res:
        pressPower(sleep_time=1)
    # 判断是否解锁
    if "showing=true" in res:
        key = [1, 2 ,3, 4, 5, 6]
        swipe(pos_list[0], pos_list[2], delay=200, sleep_time=1)
        for i in key:
            tap(pos_list[i], sleep_time=0.1)
        time.sleep(1)
    res = os.popen(adb_path + " shell dumpsys window policy").read()
    if "screenState=SCREEN_STATE_ON" in res and "showing=false" in res:
        print("亮屏、解锁成功")
    else:
        raise Exception("亮屏、解锁失败")

    date_str = datetime.datetime.today().__format__("%Y%m%d")
    fail_cnt = 0
    fail_limit = 10
    # 打开健康宝
    toKeyPage()
    jkb_pos = (160, 1200)
    zichaxun_pos = (540, 1120)
    tap(jkb_pos, sleep_time=1)
    while True:
        getScreenShot("temp")
        contrast1 = imageContrast("temp.png", "save_jkb.png")
        contrast2 = imageContrast("temp.png", "standard_jkb.png")
        print("step1 jkb_contrast1: %d jkb_contrast2: %d" % (contrast1, contrast2))
        if contrast1 > 5000 and contrast2 > 5000:
            fail_cnt += 1
            if fail_cnt > fail_limit:
                raise Exception("健康宝获取失败")
        else:
            print("step1 PASS")
            faii_cnt = 0
            break
    if contrast2 > 5000:
        tap(zichaxun_pos, sleep_time=1)
        while True:
            getScreenShot("temp")
            contrast = imageContrast("temp.png", "standard_jkb.png")
            print("step2 jkb_contrast: %d" % contrast)
            if contrast > 5000:
                fail_cnt += 1
                if fail_cnt > fail_limit:
                    raise Exception("健康宝获取失败")
            else:
                print("step2 PASS")
                fail_cnt = 0
                break
    name = date_str + "_jkb"
    getScreenShot(name)
    print("健康宝截图：%s.png" % name)

    # 打开行程码
    toKeyPage()
    xcm_pos = (410, 1200)
    confirm_pos1 = (540, 1485)
    confirm_pos2 = (540, 1680)
    tap(xcm_pos, sleep_time=1)
    while True:
        getScreenShot("temp")
        contrast = imageContrast("temp.png", "save_xcm.png")
        print("step1 xcm_contrast: %d" % contrast)
        if contrast > 5000:
            fail_cnt += 1
            if fail_cnt > fail_limit:
                raise Exception("行程码获取失败")
        else:
            fail_cnt = 0
            print("step1 PASS")
            break
    tap(confirm_pos1, sleep_time=1)
    tap(confirm_pos2, sleep_time=1)
    while True:
        getScreenShot("temp")
        contrast = imageContrast("temp.png", "standard_xcm.png")
        print("step2 xcm_contrast: %d" % contrast)
        if contrast > 5000:
            fail_cnt += 1
            if fail_cnt > fail_limit:
                raise Exception("行程码获取失败")
        else:
            fail_cnt = 0
            print("step2 PASS")
            break
    name = date_str + "_xcm"
    getScreenShot(name)
    print("行程码截图：%s.png" % name)

    # 返回HOME并锁屏
    pressHome(sleep_time=1)
    pressPower(sleep_time=1)

if __name__ == "__main__":
    try:
        getPicture()
        driver = webdriver.Firefox()
        login(driver, 学号, "密码")
        stu_io(driver)
        if not check(driver):
            driver.quit()
            raise Exception("检查不通过")
        driver.quit()
    except Exception as e:
        print(e)
        print("消息推送已发出")
        send_msg("PKUAutoSubmit填报失败！", str(e))
        exit(0)

    print("消息推送已发出")
    send_msg("PKUAutoSubmit填报成功！")
