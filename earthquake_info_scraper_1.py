# -*- coding: utf-8 -*-
"""
Created on Tue Jul 16 11:09:08 2024

@author: e55677
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
import os, json

options = Options()
options.add_argument("--headless")  # 无头模式运行，适合无界面环境如CI/CD
options.add_argument("--disable-dev-shm-usage")  # 禁用/dev/shm使用，适用于虚拟化环境
options.add_argument("--no-sandbox")  # 禁用沙盒模式，适用于虚拟化环境
driver = webdriver.Chrome(options=options)

def authorize_google_sheets():
    # Load Google service account credentials from service_account.json
    with open('service_account.json') as f:
        credentials = json.load(f)
    gc = pygsheets.authorize(service_account_file=credentials)
    return gc

def fetch_latest_earthquake_link():
    driver = webdriver.Chrome(options=options)
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
    driver = webdriver.Chrome(options=options)
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
    gc = authorize_google_sheets()
    sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1607r51uPLPASt07lVIDP8-xTi6Z3D4LjEzVZwQoilW4/edit?gid=0#gid=0')
    wks = sht[0]
    
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
    max_intensity = fetch_max_intensity(link)
    # 構造數據元組
    data = [['地震資訊：', max_intensity[0]],
            ['發生時間：', max_intensity[1]],
            ['資料抓取時間：', time_now]]

    # 將數據寫入 Google 表格
    write_to_google_sheet(data)


if __name__ == '__main__':
    job()
