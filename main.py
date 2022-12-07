import time
import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import NoSuchElementException
from urllib.parse import quote
import requests


def easy_find_elements(driver, by_what, s):
    cnt = 0
    while(cnt < 10):
        try:
            return driver.find_elements(by_what, s)
        except:
            cnt += 1
            time.sleep(1)
    raise Exception("%s无法找到" % s)


def easy_click_1st_displayed_element(elements):
    for i in elements:
        if i.is_displayed():
            i.click()
            time.sleep(0.5)
            break


def easy_sendkey(driver, by_what, s, key):
    cnt = 0
    while(cnt < 10):
        try:
            driver.find_element(by_what, s).send_keys(key)
            time.sleep(0.5)
            return True
        except:
            cnt += 1
            time.sleep(1)
    raise Exception("%s无法找到" % s)


def easy_click(driver, by_what, s):
    cnt = 0
    while(cnt < 10):
        try:
            driver.find_element(by_what, s).click()
            time.sleep(0.5)
            return True
        except:
            cnt += 1
            time.sleep(1)
    raise Exception("%s无法找到" % s)


def is_exist(driver, by_what, s):
    try:
        driver.find_element(by_what, s)
        return True
    except:
        return False


def login(driver, username, password):
    iaaaUrl = 'https://iaaa.pku.edu.cn/iaaa/oauth.jsp'
    appName = quote('北京大学校内信息门户新版')
    redirectUrl = 'https://portal.pku.edu.cn/portal2017/ssoLogin.do'
    driver.get(f'{iaaaUrl}?appID=portal2017&appName={appName}&redirectUrl={redirectUrl}')

    easy_sendkey(driver, By.ID, "user_name", username)
    easy_sendkey(driver, By.ID, "password", password)
    easy_click(driver, By.ID, "logon_button")
    easy_click(driver, By.CLASS_NAME, "btn")
    print("门户登录成功")


def stu_io(driver):
    easy_click(driver, By.ID, "all")
    easy_click(driver, By.ID, "tag_s_stuCampusExEnReq")
    driver.switch_to.window(driver.window_handles[-1])
    print("进入出入校填报界面")

    # 选择园区往返
    easy_click(driver, By.XPATH, "//span[text()=' 园区往返申请']")
    # 不用点确定了
    # time.sleep(3)
    # easy_click(driver, By.XPATH, "//span[text()='确定']")
    # 默认当天
    places = ["燕园", "物理学院"]
    easy_click(driver, By.XPATH, "//label[text()='园区（出）']/..//div")
    for i in places:
        if not is_exist(driver, By.XPATH, "//label[text()='园区（出）']/..//div[@class='el-select__tags']//span[text()='%s']" % i):
            elements = easy_find_elements(driver, By.XPATH, "//div[@class='el-scrollbar']//span[text()='%s']" % i)
            easy_click_1st_displayed_element(elements)
    easy_click(driver, By.XPATH, "//label[text()='园区（入）']/..//div")
    for i in places:
        if not is_exist(driver, By.XPATH, "//label[text()='园区（入）']/..//div[@class='el-select__tags']//span[text()='%s']" % i):
            elements = easy_find_elements(driver, By.XPATH, "//div[@class='el-scrollbar']//span[text()='%s']" % i)
            easy_click_1st_displayed_element(elements)
    
    # 别的都选好了
    easy_click(driver, By.XPATH, "//span[text()='保存 ']")
    easy_click(driver, By.XPATH, "//div[@class='el-message-box__btns']/button[2]")
    print("今日信息已提交")


def check(driver):
    easy_click(driver, By.XPATH, "//div[text()='返回']")
    easy_click(driver, By.XPATH, "//span[text()=' 申请历史']")
    date_str = datetime.datetime.today().__format__("%Y%m%d")
    try:
        easy_click(driver, By.XPATH, f"//div[text()=' {date_str} ']/../..//div[text()='审核通过']")
        print("检查通过")
        return True
    except:
        print("检查不通过")
        return False


def send_msg(title, content=""):
    # https://sct.ftqq.com/
    key = ""
    api = f"https://sctapi.ftqq.com/{key}.send"
    data = {"text": title, "desp": content}
    requests.post(api, data=data)


if __name__ == "__main__":
    try:
        options = webdriver.FirefoxOptions()
        options.add_argument('-headless')
        driver = webdriver.Firefox(options=options)
        login(driver, 123456, "123456")
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
