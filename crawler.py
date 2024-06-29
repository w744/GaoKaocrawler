'''
执行抓取功能的应用程序
By Aquamarine & w744

所需环境：
pip install selenium
pip install openpyxl
pip install bs4
pip install pandas
pip install PIL
下载 Chrome 浏览器及相应 chromedriver
Windows执行脚本权限：set-executionpolicy remotesigned
'''

excel_file = "701-1100.xlsx" # input
output_file = "score.csv"
chromedriver_path = r"D:/programfiles/chromedriver/chromedriver.exe"

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import time
import csv
import pandas as pd
import sys
import os
import base64
import random
import base64
from PIL import Image, ImageChops
from io import BytesIO


# 请忽略终端中输出的日志信息
service = Service(chromedriver_path)
options = Options()
options.add_argument("--disable-logging") # 禁用日志记录
options.add_argument("--ignore-certificate-errors") # 忽略SSL证书错误
options.add_argument("--disable-web-security") # 禁用Web安全功能
options.add_argument("--no-sandbox") # 禁用沙箱
options.add_argument("--log-level=3")  # 设置日志级别为静默模式
options.add_experimental_option('excludeSwitches', ['enable-automation'])
browser = webdriver.Chrome(service=service, options=options)

base_image = None

def getVerticalLineOffsetX(bgImage):
    diff = ImageChops.difference(bgImage, base_image)
    width, height = bgImage.size
    threshold = 10
    def is_column_colored(img, col):
        for row in range(height):
            r, g, b = img.getpixel((col, row))
            if r > threshold or g > threshold or b > threshold:
                return True
        return False
    colored_column = None
    for col in range(width):
        if is_column_colored(diff, col):
            colored_column = col
            break
    return colored_column

class DragUtil():
    def __init__(self, driver):
        self.driver = driver

    def simulateDragX(self, source, targetOffsetX):
        """
        模仿人的拖拽动作：快速沿着X轴拖动（存在误差），再暂停，然后修正误差
        防止被检测为机器人，出现“图片被怪物吃掉了”等验证失败的情况
        :param source:要拖拽的html元素
        :param targetOffsetX: 拖拽目标x轴距离
        :return: None
        """
        action_chains = webdriver.ActionChains(self.driver)
        # 点击，准备拖拽
        action_chains.click_and_hold(source)
        offset = 0
        s = 20
        while targetOffsetX - offset >= 25:
            s = random.randint(15, 25)
            offset += s
            action_chains.move_by_offset(s, 0)
        action_chains.move_by_offset(targetOffsetX - offset, 0)
        action_chains.release()
        action_chains.perform()

def checkVeriImage(driver):
    # 获取特定ID的<img>元素的src属性
    img_element = driver.find_element(By.ID, "tianai-captcha-slider-bg-img")
    img_src = img_element.get_attribute("src")
    # 下载图像数据
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[-1])  # 切换到新标签页
    driver.get(img_src)  # 在新标签页中加载图片
    img_bytes = driver.find_element(By.TAG_NAME, "img").screenshot_as_png
    # 将图像数据转换为Base64编码
    img_base64 = base64.b64encode(img_bytes).decode('utf-8')
    im_bytes = base64.b64decode(img_base64)
    image_data = BytesIO(im_bytes)
    bgImage = Image.open(image_data)
    # 切换回原始标签页
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

    offsetX = getVerticalLineOffsetX(bgImage)
    print("offsetX: {}".format(offsetX))
    dragVeriImage(driver, offsetX)

def dragVeriImage(driver, offsetX):
    # 可能产生检测到右边缘的情况
    # 拖拽
    eleDrag = driver.find_element(By.CLASS_NAME, "slider-move-btn")
    dragUtil = DragUtil(driver)
    dragUtil.simulateDragX(eleDrag, offsetX / 2)

# 别问为什么用拼音，问就是和查分网站保持一致
def read_credentials(excel_file, sheet_name, sfzh_row, sfzh_col, zkzh_row, zkzh_col, bmh_row, bmh_col):
    """
    从指定的Excel文件中读取身份证号码和准考证号
    """
    # 读取Excel文件
    dtype_spec = {sfzh_col: str, zkzh_col: str, bmh_col: str}
    data = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=dtype_spec)

    # 从指定行和列读取用户名和密码
    sfzh = data.iloc[sfzh_row, sfzh_col]
    zkzh = data.iloc[zkzh_row, zkzh_col]
    bmh = data.iloc[bmh_row, bmh_col]

    if pd.isna(sfzh):
        print("无身份证号，已终止本次查询")
        os._exit(1)
    else:
        # 如果获得的内容在格式上存在问题，需要进行处理
        if sfzh[0] == '\t':
            sfzh = sfzh[1:]
        assert len(sfzh) == 18, "身份证号长度不正确"
    if pd.isna(zkzh):
        zkzh = None
    if pd.isna(bmh):
        bmh = None

    return sfzh, zkzh, bmh

# 函数：登录网站
def login(sfzh, zkzh, bmh):  # 戴默认值的参数需要在不带默认值的参数之后
        browser.get('http://119.96.209.228:82/n_score/')
        WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.NAME, "gkbmh")))

        browser.find_element(By.NAME, "sfzh").send_keys(sfzh)
        if (zkzh != None):
            browser.find_element(By.NAME, "zkzh").send_keys(zkzh)
        elif (bmh != None):
            browser.find_element(By.NAME, "gkbmh").send_keys(bmh)
        else:
            print("没有正确的准考证号或报名号，查个屁！")
            return

        checkVeriImage(browser)
        browser.find_element(By.ID, 'cx').click()
        time.sleep(0.5)

        while not check_login():
            try:
                browser.find_element(By.CLASS_NAME, "popup-close").click()
            except:
                pass
            browser.find_element(By.ID, 'cx').click()
            if not check_login():
                try:
                    browser.find_element(By.CLASS_NAME, "popup-close").click()
                except:
                    pass
                checkVeriImage(browser)
                browser.find_element(By.ID, 'cx').click()
                time.sleep(0.5)

        while not check_login():
            time.sleep(0.5)

# 函数：检查是否登录成功
def check_login():
    try:
        # 通过寻找一个只有在登录之后才能看到的元素来实现登录检查
        WebDriverWait(browser, 1).until(
            EC.visibility_of_element_located((By.ID, "result_mc1")))
        return True
    except:
        return False

# 函数：写入对应科目
def into(text, score, data):
    if (text == '化学'):
        data['化学'] = score

    if (text == '生物学'):
        data['生物'] = score

    if (text == '地理'):
        data['地理'] = score

    if (text == '思想政治'):
        data['政治'] = score

# 函数：抓取内容
def fetch(web, num):
    # 检查是否处于正常的登录状态
    if not check_login():
        print("未登录")
        return

    # 从网页中解析内容
    html = web.page_source
    soup = BeautifulSoup(html, 'html.parser')

    # 将抓取到的内容写入文件当中
    with open(output_file, 'a', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['语文', '数学', '英语', '物理',
                      '化学', '生物', '历史', '地理', '政治', '姓名']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        if csvfile.tell() == 0:
            writer.writeheader()

        score_data = {}  # Initialize dictionary to store all score data for this row

        base_tables = soup.find_all('table', class_='base-t')
        for base_table in base_tables:
            score_data['姓名'] = base_table.find('td', id='result_xm').text

        # 从网页中抓取具体的内容
        tables = soup.find_all('table', class_='score-t')
        for table in tables:
            score_data['语文'] = table.find('td', id='result_score1').text
            score_data['数学'] = table.find('td', id='result_score2').text
            score_data['英语'] = table.find('td', id='result_score3').text

            text4 = table.find('td', id='result_mc4').text
            score4 = table.find('td', id='result_score4').text
            if (text4 == '物理'):
                score_data['物理'] = score4
            else:
                score_data['历史'] = score4

            text5 = table.find('td', id='result_mc5').text
            score5 = table.find('td', id='result_score5').text
            into(text5, score5, score_data)

            text6 = table.find('td', id='result_mc6').text
            score6 = table.find('td', id='result_score6').text
            into(text6, score6, score_data)

            writer.writerow(score_data)
    print("row", num, "succeed!")
    web.quit()

def run():
    with open('base.png','rb') as f:
        image_data = f.read()
    global base_image
    base_image = Image.open(BytesIO(image_data)).convert('RGB')

    if len(sys.argv) > 1:  # 确保至少有一个参数被传入
        input_value = sys.argv[1]  # 获取第一个命令行参数
        print("Received input:", input_value)
        i = int(input_value)
        sfzh, zkzh, bmh = read_credentials(
            excel_file, sheet_name='Sheet1', sfzh_row=i, sfzh_col=14, zkzh_row=i, zkzh_col=15, bmh_row=i, bmh_col=16)
        login(sfzh, zkzh, bmh)
        fetch(browser, i)

    else:
        print("Warning! No input received.")

if __name__ == "__main__":
    run()