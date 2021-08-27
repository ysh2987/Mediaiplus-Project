import requests
import re
from bs4 import BeautifulSoup
import openpyxl
from selenium import webdriver
import time
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
driver = webdriver.Chrome(executable_path='chromedriver',options = chrome_options)




dept_link = [
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1896',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=142',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1868',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1869',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1870',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1871',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1872',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1873',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1874',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1875',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1876',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1877',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=3680',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1878',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1879',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1880',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1881',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1882',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1883',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1855',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1856',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1886',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1887',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1888',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1889',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1890',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1891',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1892',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1893',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1894',
    'https://www.schmc.ac.kr/bucheon/dept/doctr.do?key=1895'
]


real_link = []
for i in dept_link :
    driver.get(i)
    time.sleep(2)
    htmlsrc  = driver.page_source
    soup = BeautifulSoup(htmlsrc, "html.parser", from_encoding='utf=8')
    a_hre = soup.select('a.introduce')
    for r_link in a_hre :
        r_link = "https://www.schmc.ac.kr" + r_link["href"]
        if r_link != 'https://www.schmc.ac.kr#' :
            real_link.append(r_link)

doc_num = 0 
for doc_one in real_link :
    doc_num += 1 # id 값
    driver.get(doc_one)
    htmlsrc  = driver.page_source
    soup = BeautifulSoup(htmlsrc, "html.parser", from_encoding='utf=8')
    wb = openpyxl.load_workbook('순천향.xlsx')
    time.sleep(2)
    # 기본정보
    ws = wb.worksheets[0]
    dept = soup.select_one('p.subj').text.strip()
    name = soup.select_one('div.doc_txt_area h5').text.strip().replace('교수','')
    array = []
    array.append('doc' + str(doc_num).zfill(8))
    array.append('순천향대학 부천병원')
    array.append(name)
    array.append(dept)
    ws.append(array)


    # 전문분야
    ws = wb.worksheets[1]  
    array.append('doc' + str(doc_num).zfill(8))
    specil_one = soup.select_one('p.txt').text.strip()
    specialty = re.split(r',(?![^()]*\))', specil_one)
    for zz in specialty :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        zz = zz.lstrip()
        array.append(zz)
        ws.append(array)
    # 111
    test = soup.select('div._careerContainer')

    for i in test :
        w = i.select_one('h3.cont_tit').text.strip()
        all_data = i.select('td')
        if w in '학력' : 
            ws = wb.worksheets[2] # 학력
            for data in all_data :
                data = data.text.strip().replace('-','~').split('~')
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                for data_one in data :
                    array.append(data_one)
                ws.append(array)
        if w in '약력' or w in '경력' : 
            ws = wb.worksheets[3] # 경력
            for data in all_data :
                data = data.text.strip().replace('-','~').split('~')
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                for data_one in data :
                    array.append(data_one)
                ws.append(array)
        if w in '수상' or w in '연구·수상' or w in '수상내역': 
            ws = wb.worksheets[4] # 수상
            for data in all_data :
                data = data.text.strip().replace('-','~').split('~')
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                for data_one in data :
                    array.append(data_one)
                ws.append(array)
        if w in '학회활동' : 
            ws = wb.worksheets[5] # 수상
            for data in all_data :
                data = data.text.strip().replace('-','~').split('~')
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                for data_one in data :
                    array.append(data_one)
                ws.append(array)

    ws = wb.worksheets[6] # 논문
    books = soup.select('li._thesisIem')
    for book in books :
        book = book.text.strip()
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(book)
        ws.append(array)
            
    wb.save('순천향.xlsx')
    wb.close()


