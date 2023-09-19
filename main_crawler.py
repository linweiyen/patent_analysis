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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

def __create_browser():    
    # 使用 Chrome 的 WebDriver
    options = Options()
    options.headless = False
    browser = webdriver.Firefox(executable_path = GeckoDriverManager().install(), options=options)

    return browser


####################################################################################
# 查詢台灣 Quadro 網路價格
####################################################################################
def __check_if_keywords_in_patent(patent_number, knowledge_dict):
    
    _browser = __create_browser()
    _browser.get("https://gpss3.tipo.gov.tw/gpsskmc/gpssbkm?@@0.13291281677003863")

    time.sleep(2.5 + random.random() * 2.5)
        
    # 搜尋專利號
    input_element = _browser.find_element_by_xpath("//input[@name='_21_1_T']")
    input_element.send_keys(patent_number)
    time.sleep(random.random() * 1)
    input_element.send_keys(Keys.ENTER)
    time.sleep(5 + random.random() * 5)

    #查詢所有連結
    link_elements = _browser.find_elements_by_css_selector('a.link02')

    check_if_get_right_element = False
    for link_element in link_elements:
        href = link_element.get_attribute('href')
        text = link_element.text
        if text == patent_number:
            _browser.get(href)
            time.sleep(5 + random.random() * 5)
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
            _browser.execute_script("arguments[0].scrollIntoView();", div)
            div.click()
            time.sleep(2.5 + random.random() * 2.5)

    # 取得正確結果，進行文字檢查
    html = _browser.page_source 
    soup = BeautifulSoup(html, features="html.parser")
    
    contents = soup.find_all("div", class_="panel-body")

    result = dict()
    for category_name in knowledge_dict.keys():
        result[category_name] = 0

    for content in contents:
        # 透過分類字典檢查是否所有關鍵字都存在
        for category_name in knowledge_dict.keys():
            # 取得 OR 關鍵字組合
            keywords_list = knowledge_dict[category_name]
            for keywords in keywords_list:
                # 單一 Keywords 要完全出現在 content 之中
                if all(keyword in content.text for keyword in keywords):
                    result[category_name] = 1
                    break

    _browser.quit()

    return result



def main():
    knowledge_dict = dict()
    patent_number_list = list()
    with open("keyword_list.txt", "r", encoding="utf-8") as file:
        for line in file:
            if line.strip() != "":
                items = line.strip().split(":")
                # 處理關鍵字
                keyword_token_list = items[1].replace("][", "&").replace("[", "").replace("]", "").split("&")
                knowledge_dict[items[0]] = [keyword_token.split(";") for keyword_token in keyword_token_list]
    with open("patent_number_list.txt", "r", encoding="utf-8") as file:
        for line in file:
            if line.strip() != "":
                patent_number_list.append(line.strip()) 

    final_reports = dict()
    for patent_number in patent_number_list:
        item = __check_if_keywords_in_patent(patent_number = patent_number, knowledge_dict=knowledge_dict)
        print("Complete " + patent_number)
        if item != None:
            final_reports[patent_number] = item

    data = {
            "公開/公告號": patent_number_list,
        }
    
    for keyword in knowledge_dict.keys():
        column_item = list()
        for patent_number in patent_number_list:
            column_item.append(final_reports[patent_number][keyword])
        data[keyword] = column_item

    df = pd.DataFrame(data)

    # 指定要保存的Excel文件名和工作表名
    excel_file = "report.xlsx"
    sheet_name = "數據表"

    # 使用to_excel方法将数据写入Excel文件
    df.to_excel(excel_file, sheet_name=sheet_name, index=False, engine='openpyxl')

if __name__ == '__main__':
    main()
