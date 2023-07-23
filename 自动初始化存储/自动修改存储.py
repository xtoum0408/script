#####################################
# 初始化存储设备，并填入设置信息
#####################################

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from time import sleep
import os

ip_4 = str(input("请输入ip地址最后一位："))
ip = ['10', '11', '11', ip_4]
wangguan = ['10', '11', '11', '254']
address = 'http://192.168.2.108'
i = 0
user = 'admin'
key = 'xinkaifa9'
phone_number = '17700000000'

driver = webdriver.Firefox()

driver.get(address)
driver.refresh()  # 刷新页面
driver.maximize_window() #浏览器最大

# 初始化流程
def init():
    # 勾选已阅读
    driver.find_element_by_xpath("//div[@id='devinit_dialog_software']/div[2]/div[2]/input").click()
    # 点击下一步
    driver.find_element_by_id("devinit_software_confirm").click()
    # 设置密码
    driver.find_element_by_id("devinit_pwd_ipt").send_keys(key)
    driver.find_element_by_id("devinit_pwd2_ipt").send_keys(key)
    driver.find_element_by_xpath("//div[@id='devinit_dialog_init']/div[3]/button[1]").click()
    sleep(1)
    # 密码保护
    driver.find_element_by_id("devinit_phone_chk").click()
    driver.find_element_by_xpath("//div[@id='devinit_dialog_init']/div[3]/button[1]").click()
    sleep(1)
    # 完成设置
    driver.find_element_by_xpath("//div[@id='devinit_dialog_init']/div[3]/button[2]").click()
    return


def login():
    # 登录
    user_0 = driver.find_element_by_id("login_user")
    ActionChains(driver).double_click(user_0).perform()
    user_0.send_keys(user)
    sleep(1)
    driver.find_element_by_id("login_psw").send_keys(key)
    sleep(1)
    log = driver.find_element_by_xpath("//div[@id='login']/div[1]/div[3]/div/button")
    ActionChains(driver).click(log).perform()
    sleep(5)
    return


init()
sleep(1)
# 选择不安装控件
cancel = driver.find_element_by_xpath("/html/body/div[5]/div/div[2]/a[2]")
ActionChains(driver).click(cancel).perform()
sleep(1)

login()

# 系统配置
driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/div[1]/ul/li[5]").is_enabled()
driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/div[1]/ul/li[5]/ul").is_enabled()
sys = driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/div[1]/ul/li[5]")
ActionChains(driver).double_click(sys).perform()
ActionChains(driver).click(sys).perform()
# 网卡设置
wangkashezhi = driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/div[1]/ul/li[5]/ul/li[1]")
ActionChains(driver).click(wangkashezhi).perform()
sleep(15)
# 网卡1
# 编辑
# driver.find_element_by_xpath("//div[@id='tcpip_table']/div/div[2]/table/tbody/tr/td[6]").click()
wangka1 = driver.find_element_by_xpath("//div[@id='tcpip_table']/div/div[2]/table/tbody/tr/td[6]")
wangka1.is_enabled()
ActionChains(driver).double_click(wangka1).perform()
# 读取mac地址
driver.find_element_by_id('tcp_mac').is_enabled()
mac = driver.find_element_by_id('tcp_mac').get_attribute('value')
print(mac)

# 修改ip
for ip00 in ip:
    ip_option = driver.find_element_by_xpath(f"//div[@id='tcp_ipv4']/input[{i+1}]")
    ActionChains(driver).double_click(ip_option).perform()
    ip_option.send_keys(ip[i])
    i = i + 1

# 修改网关
wang_option = driver.find_element_by_xpath("//div[@id='tcp_gateway']/input[1]")
ActionChains(driver).double_click(wang_option).perform()
wang_option.send_keys(wangguan[0])

wang_option = driver.find_element_by_xpath("//div[@id='tcp_gateway']/input[2]")
ActionChains(driver).double_click(wang_option).perform()
wang_option.send_keys(wangguan[1])

wang_option = driver.find_element_by_xpath("//div[@id='tcp_gateway']/input[3]")
ActionChains(driver).double_click(wang_option).perform()
wang_option.send_keys(wangguan[2])

wang_option = driver.find_element_by_xpath("//div[@id='tcp_gateway']/input[4]")
ActionChains(driver).double_click(wang_option).perform()
wang_option.send_keys(wangguan[3])

# 确定
confirm = driver.find_element_by_xpath("//div[@id='tcpip_modDialog']/div[3]/button[1]")
ActionChains(driver).click(confirm).perform()
sleep(2)

# set_confirm = driver.find_element_by_xpath("//div[@id='info_content']/div[3]/div/div[2]/div[2]/div/button[2]")
# set_confirm.is_enabled()
# ActionChains(driver).click(set_confirm).perform()
#sleep(1)

#driver.close()

