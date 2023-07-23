# 导入库
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
from openpyxl import Workbook
import openpyxl as xl
import os

# 定义变量
user = 'admin'
key = 'xinkaifa9'
file_name = '(总)各楼层摄像机编号和IP.xlsx'
ip_list = []  # 机ip地址list
serial_list = []
i = 0
index = 0

# 遍历表格，确定人脸识别相机的ip地址，并赋值给一个list
book = xl.load_workbook(file_name)
floor = input('请输入楼层：')
# print(floors)

sheet = book[floor]
max_r = sheet.max_row + 1

for r in range(2, max_r):
    ip_list.append(sheet[f'A{r}'].value)
    serial_list.append(sheet[f'B{r}'].value)

print(ip_list)
print(serial_list)
'''
# 循环list进行人脸设置
for ip in ip_list:
    serial = serial_list[index]
    print('当前IP地址为：', ip)
    print('当前相机编号为：', serial)
    index += 1

    driver = webdriver.Firefox()
    address = f'http://{ip}'
    driver.get(address)
    driver.refresh()
    driver.maximize_window()

    # 登录
    driver.find_element_by_id("login_user").send_keys(user)
    driver.find_element_by_id("login_psw").send_keys(key)
    sleep(0.5)
    driver.find_element_by_xpath("//div[@id='login']/div[1]/div/div[2]/form/div[6]/a[1]").click()
    sleep(2)

    # 设置
    driver.find_element_by_xpath("//div[@id='main']/ul/li[7]/span").click()
    # driver.find_element_by_xpath("//div[@id='main']/ul/li[6]/span").click()   # 人脸识别枪机
    sleep(1)

    # 更改通道名称
    # 视频
    driver.find_element_by_xpath("//*[@id='set-menu']/li[1]/ul/li[2]/span").click()
    sleep(2)

    # 视频叠加
    driver.find_element_by_xpath("//*[@id='page_encodeConfig']").is_enabled()
    driver.find_element_by_xpath("//*[@id='page_encodeConfig']/ul").is_enabled()
    driver.find_element_by_xpath("//*[@id='page_encodeConfig']/ul/li[3]").click()
    sleep(1)

    # 视频通道
    driver.find_element_by_xpath("//*[@id='page_encodeConfig']/div/div[3]/div[2]").is_enabled()
    driver.find_element_by_xpath("//*[@id='osd_content']/ul").is_enabled()
    driver.find_element_by_xpath("//*[@id='osd_content']/ul/li[2]/span").click()  # 通道标题
    driver.find_element_by_xpath("//*[@id='osd_content']/ul/li[2]/span").click()
    driver.find_element_by_xpath("//*[@id='osd_channel_title']/ul/li/input").clear()
    driver.find_element_by_xpath("//*[@id='osd_channel_title']/ul/li/input").send_keys(serial)
    driver.find_element_by_id("video_OSD_confirm").click()
    sleep(1)
'''
