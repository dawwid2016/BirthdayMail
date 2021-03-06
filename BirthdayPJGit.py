#! -*- coding utf-8 -*-
#! @Time  :2021/1/19 15:05
#! Author :Jiacheng Zhang
#! @File  :BirthdayPJ.py
#! Python Version 3.9.1

"""
模块功能：读取当前文件夹里的Excel文件，发送某旦学邮
"""

# 导入所需包
import pandas as pd
import datetime
from selenium import webdriver
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# 从excel文件中读取姓名-生日-邮箱信息
def readList():
    bookDir = './BirthdayList.xlsx'
    df = pd.read_excel(bookDir, sheet_name='Sheet1')
    ncols=df.columns.size
    if ncols != 3:
        return 1, {}
    else:
        return 0, df

# 发送到草稿箱
def send(df):
    #Open the Mail Website
    chrome_driver = './chromedriver.exe'
    driver = webdriver.Chrome(executable_path = chrome_driver)
    driver.get("https://mail.fudan.edu.cn/")
    # 登录
    time.sleep(2)
    driver.find_element_by_id("uid").click()    # 点选账号栏
    driver.find_element_by_id("uid").clear()    # 清空账号栏
    studID = '1630068XXXX' #change to your student id
    driver.find_element_by_id("uid").send_keys("%s")    # 输入账号
    driver.find_element_by_id("password").click()    # 点选密码栏
    driver.find_element_by_id("password").clear()    # 清空密码栏
    driver.find_element_by_id("password").send_keys("PASSWD")    # 输入密码
    driver.find_element_by_xpath('//*[@id="logArea"]/div[6]/button').click()  # 登录按钮
    time.sleep(2)
    # 找到发送按钮并进入发送界面
    driver.find_element_by_xpath('//*[@id="navContainer"]/div[1]/a[2]').click()
    time.sleep(2)
    # 进入写信框架
    driver.switch_to.frame("compose1")
    # 读取列表
    nrows = df.shape[0]
    count = 0
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 计数器控制发送
    while count < nrows:        
        # 提取姓名-生日-邮箱
        name = df.iloc[count, 0]
        birthday = df.iloc[count, 1]
        mailAddress = df.iloc[count, 2]     
        # 编辑标题和内容
        title = "Happy Birthday!"
        content = "Hi %s,\n     Happy Birthday to you! \n XXX office"%(name)
        # 收件人填写
        address = driver.find_element_by_id("InputEx_To")
        ActionChains(driver).move_to_element(address).send_keys(mailAddress).send_keys(Keys.RETURN).perform()
        # 标题填写
        driver.find_element_by_id("subject").click()
        driver.find_element_by_id("subject").clear()
        driver.find_element_by_id("subject").send_keys(title)
        # 内容填写
        driver.switch_to.frame("htmleditor")
        driver.switch_to.frame("HtmlEditor")
        text = driver.find_element_by_xpath('/html/body')
        ActionChains(driver).move_to_element(text).click().send_keys(content).send_keys(Keys.RETURN).perform()
        time.sleep(0.5)
        # 勾选定时发送
        driver.switch_to.default_content()
        driver.switch_to.frame("compose1")
        driver.find_element_by_id("chkTimeSet").click()
        # 选定时间
        date = str(birthday)
        month = date[4:6]
        day = date[6:8]
        if (int(month)>int(now[5:7])) | (int(month) == int(now[5:7]) & (int(day)>int(now[8:10]))):
            year = now[0:4]
        else:
            year = str(int(now[0:4])+1)
        print(year, month, day)
        driver.find_element_by_xpath('//*[@id="timeSetContainer"]/p[1]/select[1]').find_element_by_xpath("//option[@value='%s']"%(year)).click()
        driver.find_element_by_xpath('//*[@id="timeSetContainer"]/p[1]/select[2]').find_element_by_xpath("//option[@value='%s']"%(str(int(month)))).click()        
        Select(driver.find_element_by_xpath('//*[@id="timeSetContainer"]/p[1]/select[3]')).select_by_visible_text("%s"%(str(int(day))))
        driver.find_element_by_xpath('//*[@id="timeSetContainer"]/p[1]/select[4]').find_element_by_xpath("//option[@value='0']").click()        
        Select(driver.find_element_by_xpath('//*[@id="timeSetContainer"]/p[1]/select[5]')).select_by_visible_text("0")
        # 点击定时发送
        driver.find_element_by_xpath('//*[@id="btnSchedule"]/span').click()
        WebDriverWait(driver,60).until(EC.presence_of_element_located((By.NAME,'sigbtn_1')))
        driver.find_element_by_xpath('//*[@id="body_area"]/table/tbody/tr[3]/td/div[2]/span').click()
        time.sleep(1)
        # 点击再写一封
        count += 1
    #完成发送，退出
    print("done")
    time.sleep(5)
    driver.close()

if __name__ == '__main__':
    flag, df = readList()
    if flag == 1:
        print("The excel file has more or less than 3 columns.") # 格式出错的返回
    if flag == 0:
        send(df)
