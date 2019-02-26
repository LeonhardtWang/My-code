from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import re
import csv

def get_HTML(pro_sequence, browser):
    element = browser.find_element_by_name("sequence")
    element.send_keys(pro_sequence)
    try: element.submit()
    except: browser.refresh()
    return browser.page_source

def page_parse(source):
    pro_info_list = []
    soup = BeautifulSoup(source, 'html.parser')
    text_list = soup.text.split(':')
    regex = re.compile(r'[-]{0,1}\d+\.{0,1}\d*')
    for i in range(2, 17):
        if i == 12 or i == 13 or i == 9:continue
        elif i == 5:
            for j in range(1, 50, 2):
                pro_info_list.append(float(regex.findall(text_list[i])[j]))
        elif i == 8:
            for j in range(5):
                pro_info_list.append(int(regex.findall(text_list[i])[j]))
        elif i == 11:
            pro_info_list.append((float(regex.findall(text_list[i])[3]) +
                                 float(regex.findall(text_list[i])[7])) / 2)
            pro_info_list.append((float(regex.findall(text_list[i])[-5]) +
                                 float(regex.findall(text_list[i])[-1])) / 2)
        else:
            pro_info_list.append(float(regex.findall(text_list[i])[0]))
    return pro_info_list

def main(filename):
    web_path = "F:\Google\Chrome\Application\chromedriver.exe"
    data = pd.read_csv(filename)
    data_arr = np.array(data)
    m = data_arr.shape[0]
    browser = webdriver.Chrome(web_path)
    browser.implicitly_wait(30)
    
    for i in range(m):
        try:
            try: browser.get("https://web.expasy.org/protparam/")
            except: browser.refresh()
            source = get_HTML(data_arr[i,1], browser)
            pro_info_list = page_parse(source)
            pro_info_list.insert(0, data_arr[i,0])
            with open("pro_train_info.csv", "a", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(pro_info_list)
            print("\r当前速度：{:.2f}%".format((i+1)*100/m), end='')
        except:
            print("\r当前速度：{:.2f}%".format((i+1)*100/m), end='')
            continue
        f.close()
        
    browser.close()

main("df_protein_train.csv")
                
