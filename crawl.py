import requests
from selenium import webdriver
from selenium.webdriver.support.ui import Select

from bs4 import BeautifulSoup
import json
import re
import openpyxl
import xml.etree.ElementTree as ET
import time
import pickle
headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.139 Safari/537.36"
}

def get_school_info(schoolname):
    Referer = "https://gkcx.eol.cn/school/search?schoolflag=&argschtype=&province=&recomschprop=&keyWord1=%s".format(schoolname)
    data_url = "https://data-gkcx.eol.cn/soudaxue/queryschool.html"
    params = {
        "messtype" : "jsonp",
        "_":"1530074932531",
        "callback" : "",
        "keyWord1" : schoolname
    }
    headers["Referer"] = Referer.encode('utf-8')

    response = requests.request("GET", data_url,headers=headers,params=params)

    # print(response)
    text = ((response.text).split(');',1)[0]).split('(',1)[1]
    j = json.loads(text)
    school_list = j['school']
    school_info = [{'schoolid': i['schoolid'], 'schoolname':i['schoolname']} for i in school_list]
    return school_info

def shift_to_yy(driver):
    driver.get('https://gkcx.eol.cn')
    shift_city = driver.find_element_by_xpath('//*[@id="root"]/div/div/div/div/div[1]/div/div[1]/div[3]/div/p[1]')
    shift_city.click()
    time.sleep(2)
    xiala = driver.find_element_by_class_name('ant-select-arrow')
    xiala.click()
    time.sleep(1)
    pros = driver.find_elements_by_class_name('ant-select-dropdown-menu-item')
    pros[17].click()
    time.sleep(1)
    xiala_city = driver.find_element_by_xpath('//*[@id="root"]/div/div/div/div/div[1]/div/form/div/div/div/div[2]/div/span/div[2]/div/div/div')
    xiala_city.click()
    time.sleep(1)
    yy = driver.find_elements_by_class_name('ant-select-dropdown-menu-item')
    try:
        for i in yy:
            if '岳阳' in i.text:
                i.click()
    except:
        pass

def get_data(driver, schoolname, schoolid):
    school_url = 'https://gkcx.eol.cn/school/%s/specialtyline?cid=0' % schoolid
    print(school_url)
    driver.get(school_url)

    time.sleep(2)
    try:
        select = driver.find_elements_by_class_name('ant-select-selection-selected-value')
        for i in select:
            if '2018' in i.text:
                year_xiala = i
        year_xiala.click()
        time.sleep(1)
        years = driver.find_elements_by_class_name('ant-select-dropdown-menu-item')
        length = len(years)
        print('find %d years' % length)
        for i in range(length):
            driver.get(school_url)
            time.sleep(2)
            select = driver.find_elements_by_class_name('ant-select-selection-selected-value')
            for x in select:
                if '201' in x.text:
                    year_xiala = x
            year_xiala.click()
            time.sleep(1)
            years = driver.find_elements_by_class_name('ant-select-dropdown-menu-item')
            # print(len(years))
            years[i].click()
            time.sleep(1)
            pages = driver.find_elements_by_class_name('none')
            with open('data/%s_%s_%d.txt'%(schoolname, schoolid, 2018-i), 'w') as file:
                tbody = driver.find_element_by_tag_name('tbody')
                file.write(tbody.text)
            print('year %d find %d pages' % (2018-i, len(pages)+1))
            for page in pages:
                page.click()
                time.sleep(1)
                with open('data/%s_%s_%d.txt'%(schoolname, schoolid, 2018-i), 'a') as file:
                    tbody = driver.find_element_by_tag_name('tbody')
                    file.write(tbody.text)
    except:
        pass

driver = webdriver.Chrome()
shift_to_yy(driver)
with open('湖南高校列表.txt', 'r') as file:
    for i in file:
        school_info = get_school_info(i)
        for school in school_info:
            print(school)
            get_data(driver, school['schoolname'], school['schoolid'])