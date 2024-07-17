# -*- coding: utf-8 -*-
"""
Created on Tue Jul 16 11:09:08 2024

@author: PingEn
"""

import pygsheets
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
import datetime
import os

#service_account_key = os.getenv('googleServiceAccountKey')
# 使用 Google 服务账户密钥内容进行授权
#gc = pygsheets.authorize(service_account_file=service_account_key)

gc = pygsheets.authorize(service_account_file = 'service_account.json')
sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1607r51uPLPASt07lVIDP8-xTi6Z3D4LjEzVZwQoilW4/edit?gid=0#gid=0')
wks = sht[0]

# 初始化Chrome浏览器
#driver = webdriver.Chrome(executable_path='D:\\Users\\e55677\\Downloads\\chromedriver-win64\\chromedriver.exe')

# 使用 WebDriver Manager 自動管理並下載適合您系統的 ChromeDriver
#driver = webdriver.Chrome(ChromeDriverManager().install())

driver = webdriver.Chrome()

def fetch_latest_earthquake_link():

    try:
        url = 'https://www.cwa.gov.tw/V8/C/E/index.html'
        driver.get(url)
        
        # 等待页面加载
        time.sleep(5)

        # 找到地震信息表格的第一行
        latest_earthquake_row = driver.find_element(By.ID, 'eq-1')

        if latest_earthquake_row:
            # 获取href链接
            link_element = latest_earthquake_row.find_element(By.TAG_NAME, 'a')
            href = link_element.get_attribute('href')
            return href
        else:
            return "找不到地震資訊"
    finally:
        driver.quit()

def fetch_max_intensity(url):
    try:
        driver.get(url)
        
        time.sleep(5)  # 等待页面加载

        # 找到包含宜蘭縣地區最大震度信息的 <a> 標籤元素
        earthquake_info_element = driver.find_element(By.XPATH, '//a[contains(text(), "宜蘭縣地區最大震度")]')
        # 獲取 <a> 標籤的文字內容
        earthquake_info_text = earthquake_info_element.text
        earthquake_time_element = driver.find_element(By.XPATH, '//ul[@class="list-unstyled quake_info"]/li[contains(text(), "發震時間")]')

        # 獲取 <li> 元素的文字內容
        earthquake_time_text = earthquake_time_element.text
        # 從文字內容中提取時間部分
        time_str = earthquake_time_text.split("：")[1].strip()
        
        return earthquake_info_text, time_str
        
    finally:
        driver.quit()

def write_to_google_sheet(data):
    # 清除工作表中現有的內容
    wks.clear()
    # 将 DataFrame 写入 Google 表格
    #wks.set_dataframe(df, (1, 1))
    wks.update_values(crange='A1', values=data)

def time_get():
    loc_dt = datetime.datetime.today() 
    loc_dt_format = loc_dt.strftime("%Y/%m/%d %H:%M:%S")
    return loc_dt_format

def job():
    # 抓取地震資訊
    time_now = time_get()
    link = fetch_latest_earthquake_link()
    driver = webdriver.Chrome()
    max_intensity = fetch_max_intensity(link)
    # 構造數據元組
    data = [['地震資訊：', max_intensity[0]],
            ['發生時間：', max_intensity[1]],
            ['資料抓取時間：', time_now]]

    # 將數據寫入 Google 表格
    write_to_google_sheet(data)


if __name__ == '__main__':
    job()
