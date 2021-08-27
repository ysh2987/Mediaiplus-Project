import requests
import re
from bs4 import BeautifulSoup
import time
import openpyxl
from selenium import webdriver

dept_link = [
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=24&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=11&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=7&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=6&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=29&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=27&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=31&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=22&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=19&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=18&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=9&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=1&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=10&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=17&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=5&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=3&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=20&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=28&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=14&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=25&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=21&doctorMode=true#tab-3',
    # 'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=8&doctorMode=true#tab-3', 일반 내과 의료진 없음
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=23&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=12&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=16&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=30&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=26&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=13&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=32&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=4&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=2&doctorMode=true#tab-3',
    'https://www.gnah.co.kr/kor/CMS/DeptMgr/intro.do?mCode=MN021&dept_seq=15&doctorMode=true#tab-3'
]
doc_link = []
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
driver = webdriver.Chrome(executable_path='chromedriver',options = chrome_options)
for i in dept_link :
    driver.get(i)
    time.sleep(2)
    genre_ul = driver.find_elements_by_class_name('detail')
    for q in genre_ul :
        q = q.get_attribute('href')
        doc_link.append(q)

doc_num = 1162
for i in doc_link :
    doc_num += 1
    res = requests.get(i) # URL
    soup = BeautifulSoup(res.content, "html.parser", from_encoding='utf=8')
    wb = openpyxl.load_workbook('강릉아산병원.xlsx')

    # 기본정보
    ws = wb.worksheets[0]
    name = soup.select_one('span.doctName').text.strip() # 이름
    dept = soup.select_one('span.deptName').text.strip() # 학과
    array = [] 
    array.append('doc' + str(doc_num).zfill(8))
    array.append(name)
    array.append(dept)
    ws.append(array)

    # 전문분야
    specil_one = soup.select_one('dl.doctSect dd').text.strip() 
    specialty = re.split(r',(?![^()]*\))', specil_one) 
    ws = wb.worksheets[1]
    for zz in specialty :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        zz = zz.lstrip()
        array.append(zz)
        ws.append(array)

    # 학력/경력
    all_unit = soup.select('div.unit')
    if len(all_unit) == 1 :
        ws = wb.worksheets[2]
        cont_list = soup.select('div.unit')[0] 
        all_li = cont_list.select('li')
        docter_arr = []
        docter_date = []
        docter_content = []
        for i in all_li :
            date = i.select_one('span.tDate').text.strip().replace('-','~').split('~')
            content = i.select_one('span.tText').text.strip()
            docter_content.append(content)
            if len(date) == 2 :
                docter_arr.append(date[0])
                docter_date.append(date[1])
            elif len(date) == 1 :
                docter_arr.append(date[0])
                docter_date.append(date[0])
            elif len(date) == 0 :
                docter_arr.append('')
                docter_date.append('')
        
        for tax in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(docter_arr[tax])
            array.append(docter_date[tax])
            array.append(docter_content[tax])
            ws.append(array)

    if len(all_unit) > 1 :
        ws = wb.worksheets[2]
        cont_list = soup.select('div.unit')[0] # 학력/경력
        all_li = cont_list.select('li')
        docter_arr = []
        docter_date = []
        docter_content = []
        
        for i in all_li :
            date = i.select_one('span.tDate').text.strip().replace('-','~').split('~')
            content = i.select_one('span.tText').text.strip()
            docter_content.append(content)
            if len(date) == 2 :
                docter_arr.append(date[0])
                docter_date.append(date[1])
            elif len(date) == 1 :
                docter_arr.append(date[0])
                docter_date.append(date[0])
            elif len(date) == 0 :
                docter_arr.append('')
                docter_date.append('')
        for tax in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(docter_arr[tax])
            array.append(docter_date[tax])
            array.append(docter_content[tax])
            ws.append(array)

        # 학회
        ws = wb.worksheets[3]
        cont_list = soup.select('div.unit')[1] 
        all_li = cont_list.select('li')
        docter_arr = []
        docter_date = []
        docter_content = []
        for i in all_li :
            if len(i) > 1 :
                date = i.select_one('span.tDate').text.strip().replace('-','~').split('~')
                content = i.select_one('span.tText').text.strip()
                docter_content.append(content)
                if len(date) == 2 :
                    docter_arr.append(date[0])
                    docter_date.append(date[1])
                elif len(date) == 1 :
                    docter_arr.append(date[0])
                    docter_date.append(date[0])
                elif len(date) == 0 :
                    docter_arr.append('')
                    docter_date.append('')
            else :
                content = i.text.strip()
                docter_arr.append('')
                docter_date.append('')
                docter_content.append(content)
        for tax in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(docter_arr[tax])
            array.append(docter_date[tax])
            array.append(docter_content[tax])
            ws.append(array)
    print(name)
    print(dept)        
    wb.save('강릉아산병원.xlsx')
    wb.close()

    


   