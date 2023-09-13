##########################################################################################
# Main function of crwaler
##########################################################################################
import os

from datetime import datetime
import re
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.keys import Keys
import time
import random

def __create_browser():    
    # 使用 Chrome 的 WebDriver
    options = Options()
    options.headless = False
    browser = webdriver.Firefox(executable_path = GeckoDriverManager().install(), options=options)

    return browser


####################################################################################
# 查詢台灣 Quadro 網路價格
####################################################################################
def __check_if_keywords_in_patent(patent_number, keywords):
    
    _browser = __create_browser()
    _browser.get("https://gpss3.tipo.gov.tw/gpsskmc/gpssbkm?@@0.13291281677003863")

    time.sleep(random.random() * 5)
        
    # 搜尋專利號
    input_element = _browser.find_element_by_xpath("//input[@name='_21_1_T']")
    input_element.send_keys(patent_number)
    time.sleep(random.random() * 1)
    input_element.send_keys(Keys.ENTER)
    time.sleep(random.random() * 10)

    #查詢所有連結
    link_elements = _browser.find_elements_by_css_selector('a.link02')

    check_if_get_right_element = False
    for link_element in link_elements:
        href = link_element.get_attribute('href')
        text = link_element.text
        if text == patent_number:
            link_element.click()
            time.sleep(random.random() * 10)
            check_if_get_right_element = True
            break

    #沒有查到準確連結，則終止程序
    if not check_if_get_right_element:
        _browser.quit()
        return
    
    divs_with_class_morelink = _browser.find_elements_by_css_selector('div.moreLink')

    # 底選所有 more link
    for div in divs_with_class_morelink:
        if div.text == "more":
            div.click()
            time.sleep(random.random() * 5)

    # 取得正確結果，進行文字檢查
    html = _browser.page_source 
    soup = BeautifulSoup(html, features="html.parser")
    
    contents = soup.find_all("div", class_="panel-body")

    result = dict()
    for keyword in keywords:
        result[keyword] = 0

    for content in contents:               
        if result[keyword] == 1:
            break

        for keyword in keywords:
            if keyword in content.text:
                result[keyword] = 1
                break

    _browser.quit()

    return result



def main():
    keyword_list = list()
    patent_number_list = list()
    with open("keyword_list.txt", "r", encoding="utf-8") as file:
        for line in file:
            if line.strip() != "":
                keyword_list.append(line.strip()) 
    with open("patent_number_list.txt", "r", encoding="utf-8") as file:
        for line in file:
            if line.strip() != "":
                patent_number_list.append(line.strip()) 

    final_reports = list()
    for patent_number in patent_number_list:
        item = __check_if_keywords_in_patent(patent_number = patent_number, keywords=keyword_list)
        if item != None:
            report = patent_number + "," + ",".join([str(num) for num in list(item.values())])
            final_reports.append(report)

    with open("report.csv", "w", encoding="utf-8") as file:
        file.write("專利號," + ",".join(keyword_list) + "\n")
        for report in final_reports:
            file.write(report + "\n")

if __name__ == '__main__':
    main()
