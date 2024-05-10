'''
文件介绍：
        数据切分之后放入缓存CSV
'''
# 版本号：
# 版本新增：
# 开发时间：2023/3/5 16:41

import random

# from bs4 import BeautifulSoup
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
import schedule
import time
import csv

# 创建 ChromeOptions 对象
options = webdriver.ChromeOptions()
# 设置为无头浏览器
options.add_argument('--headless')
# 启动Chrome浏览器
driver = webdriver.Chrome(options=options)
# driver = webdriver.Chrome()

def write_to_csv(data):
    with open('data_AQI.csv', 'a', newline='') as csvfile:
        fieldnames = ['City', 'AQI', 'WRdegree', 'firstp']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        if csvfile.tell() == 0:
            writer.writeheader()
        for row in data:
            writer.writerow(row)

def crawl_data():
    # 为防止中间某一次抓取失败（或抓取行为被服务器识别到），导致停机，添加 try-except 块捕获抓取数据时可能出现的异常。
    # 如果抓取数据失败，将打印错误信息，并等待10秒后再次执行 crawl_data 函数。
    try:
        # 访问包含城市按钮的页面
        driver.get('http://1.192.88.18:8088/TodayMonitor')
        s = 0
        # 定位到城市的按钮并点击
        for i in range(1,10):
            s += 1
            # 点击速度过快，被服务器发现，随机添加0-10s等待
            time.sleep(random.random()*10)
            btn = driver.find_element(by=By.ID,value=f'410{i}00')
            btn.click()
            city_name = driver.find_element(by=By.CLASS_NAME, value='i1').text.strip()
            AQI_val = driver.find_element(by=By.ID,value='spanAQI').text.strip()
            WRdegree = driver.find_element(by=By.ID, value='airlevel').text.strip()
            firstp = driver.find_element(by=By.ID, value='firstp').text.strip()
            print(s,':',city_name,'-',AQI_val,'-',WRdegree,'-',firstp)
            data = []
            data.append({'City': city_name, 'AQI': AQI_val, 'WRdegree': WRdegree, 'firstp': firstp})
            # 将数据转存至 CSV 文件
            write_to_csv(data)
        for k in range(9):
            s += 1
            # 点击速度过快，被服务器发现，随机添加0-10s等待
            time.sleep(random.random()*10)
            btn = driver.find_element(by=By.ID,value=f'411{k}00')
            btn.click()
            city_name = driver.find_element(by=By.CLASS_NAME, value='i1').text.strip()
            AQI_val = driver.find_element(by=By.ID, value='spanAQI').text.strip()
            WRdegree = driver.find_element(by=By.ID, value='airlevel').text.strip()
            firstp = driver.find_element(by=By.ID, value='firstp').text.strip()
            print(s, ':', city_name, '-', AQI_val, '-', WRdegree, '-', firstp)
            data = []
            data.append({'City': city_name, 'AQI': AQI_val, 'WRdegree': WRdegree, 'firstp': firstp})
            # 将数据转存至 CSV 文件
            write_to_csv(data)
        print('OK')
    except Exception as e:
        # 打印错误信息并等待10秒后再次执行任务
        print('抓取数据失败：', e)
        time.sleep(10)
        crawl_data()

# 每隔一个小时调用一次 crawl_data 函数
schedule.every(1).minutes.do(crawl_data)

while True:
    schedule.run_pending()
    time.sleep(1)

