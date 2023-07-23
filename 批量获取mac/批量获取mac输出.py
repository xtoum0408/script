######################################################################
# 需登记相机的mac地址，通过网页登入，在相机设置中获取mac地址并输出
########################################################################

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
from openpyxl import Workbook
import openpyxl as xl
import os

key = "xinkaifa9"
user = "admin"
mac_add = str()

book = xl.load_workbook("(总)各楼层摄像机编号和IP.xlsx")
chromedriver = "chromedriver.exe"  # 这里写本地的chromedriver 的所在路径

for ip in range(58, 60):    # 使用前需先修改ip范围
    address = f'http://10.11.3.{ip}'
    os.environ["webdriver.Chrome.driver"] = chromedriver  # 调用chrome浏览器
    driver = webdriver.Chrome(chromedriver)
    driver.get(address)
    driver.refresh()  # 刷新页面

    # 登入
    driver.find_element_by_id("login_user").send_keys(user)
    driver.find_element_by_id("login_psw").send_keys(key)
    sleep(0.5)
    driver.find_element_by_xpath("//div[@id='login']/div[1]/div/div[2]/form/div[6]/a[1]").click()
    sleep(2)

    # 设置
    driver.find_element_by_xpath("//div[@id='main']/ul/li[7]/span").click()

    # 网络设置
    driver.find_element_by_xpath("//*[@id='set-menu']").is_enabled()
    driver.find_element_by_xpath("//ul[@id='set-menu']/li[2]/a/span").click()
    driver.find_element_by_xpath("//ul[@id='set-menu']/li[2]/a/span").is_enabled()

    # tcp/ip
    driver.find_element_by_xpath("//ul[@id='set-menu']/li[2]/ul/li[1]/span").click()
    driver.find_element_by_xpath("//*[@id='page_networkConfig']").is_enabled()
    sleep(3)

    # 获得mac地址
    driver.find_element_by_xpath("//*[@id='NN_macadd']").is_enabled()
    for n in range(1, 7):
        driver.find_element_by_xpath(f"//div[@id='NN_macadd']/input[{n}]").is_enabled()
        mac = driver.find_element_by_xpath(f"//div[@id='NN_macadd']/input[{n}]").get_attribute('value')
        mac_add = mac_add + mac + '.'

    print(mac_add[:-1])
    sleep(2)
    driver.close()
    sleep(2)
