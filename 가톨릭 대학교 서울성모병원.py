import requests
import re
from bs4 import BeautifulSoup
import openpyxl
from selenium import webdriver
import time
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
driver = webdriver.Chrome(executable_path='chromedriver',options = chrome_options)


driver.get('https://www.cmcseoul.or.kr/page/department/A')
htmlsrc  = driver.page_source
soup = BeautifulSoup(htmlsrc, "html.parser", from_encoding='utf=8')

all_a = soup.select('a.btn-bak1')
dept_arr = []
for i in all_a :
    href_div = i["href"]
    dept_link = 'https://www.cmcseoul.or.kr' + href_div[0:-1] + '2'
    dept_arr.append(dept_link)
    print(dept_link)
dept_arr.remove('https://www.cmcseoul.or.krhttp://pipet.or.kr2')  # 임상 약리과 이상한 페이지로 감
doc_link =[]
for i in dept_arr :
    driver.get(i)
    time.sleep(2)
    htmlsrc  = driver.page_source
    soup = BeautifulSoup(htmlsrc, "html.parser", from_encoding='utf=8')
    
    doc_href = soup.select('a.btn_doc_info')
    for i in doc_href :
        real_link = 'https://www.cmcseoul.or.kr' + i['href']
        doc_link.append(real_link)
print(len(doc_link))
doc_num = 3695
for doc_one in doc_link :
    doc_num += 1
    driver.get(doc_one)
    time.sleep(2)
    htmlsrc  = driver.page_source
    soup = BeautifulSoup(htmlsrc, "html.parser", from_encoding='utf=8')
    wb = openpyxl.load_workbook('가톨릭대 서울성모병원.xlsx')

    # 기본정보
    ws = wb.worksheets[0]
    
    info = soup.select_one('div.doc_name')
    name = soup.select_one('div.doc_name strong').text.strip()
    dept = soup.select_one('div.doc_name em')
    if dept != None :
        dept = dept.text.strip()
    array = []
    array.append('doc' + str(doc_num).zfill(8))
    array.append('가톨릭대 서울성모병원')
    array.append(name)
    array.append(dept)
    ws.append(array)
    # 전문분야
    ws = wb.worksheets[1]
    specil_one = soup.select_one('a.medical_part_btn').text.strip()
    specialty = re.split(r',(?![^()]*\))', specil_one)
    for sp in specialty :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        sp = sp.lstrip()
        array.append(sp)
        ws.append(array)


    all_list = soup.select('div.cont_main_profile div')
    for i in all_list :
        title = i.select_one('strong').text.strip()

        # 학력
        if title in '학력' :
            ws = wb.worksheets[2]
            docter_arr = []
            docter_date = []
            docter_content = []
            edu_date = i.select('li')
            for i in edu_date :
                date = i.select_one('dt')
                content = i.select_one('dd')

                if content == None :
                    docter_content.append('')
                if content != None : 
                    content = i.select_one('dd').text.strip()
                    docter_content.append(content)
                if date == None :
                    docter_date.append('')
                    docter_arr.append('')
                elif date != None:
                    divison = date.text.strip().split('~')
                    if len(divison) == 2 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[1])
                    elif len(divison) == 1 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[0])
                    elif len(divison) == 0 :
                        docter_arr.append('')
                        docter_date.append('')
            for tax in range(len(docter_content)):
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                array.append(docter_arr[tax])
                array.append(docter_date[tax])
                array.append(docter_content[tax])
                ws.append(array)

        # 경력
        if title in '경력' :
            ws = wb.worksheets[3]
            docter_arr = []
            docter_date = []
            docter_content = []
            edu_date = i.select('li')
            for i in edu_date :
                date = i.select_one('dt')
                content = i.select_one('dd')

                if content == None :
                    docter_content.append('')
                if content != None : 
                    content = i.select_one('dd').text.strip()
                    docter_content.append(content)
                if date == None :
                    docter_date.append('')
                    docter_arr.append('')
                elif date != None:
                    divison = date.text.strip().split('~')
                    if len(divison) == 2 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[1])
                    elif len(divison) == 1 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[0])
                    elif len(divison) == 0 :
                        docter_arr.append('')
                        docter_date.append('')
            for tax in range(len(docter_content)):
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                array.append(docter_arr[tax])
                array.append(docter_date[tax])
                array.append(docter_content[tax])
                ws.append(array)

        # 학회활동
        if title in '학회활동' :
            ws = wb.worksheets[4]
            docter_arr = []
            docter_date = []
            docter_content = []
            edu_date = i.select('li')
            for i in edu_date :
                date = i.select_one('dt')
                content = i.select_one('dd')

                if content == None :
                    docter_content.append('')
                if content != None : 
                    content = i.select_one('dd').text.strip()
                    docter_content.append(content)
                if date == None :
                    docter_date.append('')
                    docter_arr.append('')
                elif date != None:
                    divison = date.text.strip().split('~')
                    if len(divison) == 2 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[1])
                    elif len(divison) == 1 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[0])
                    elif len(divison) == 0 :
                        docter_arr.append('')
                        docter_date.append('')
            for tax in range(len(docter_content)):
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                array.append(docter_arr[tax])
                array.append(docter_date[tax])
                array.append(docter_content[tax])
                ws.append(array)

        # 수상이력
        if title in '수상이력' :
            ws = wb.worksheets[5]
            docter_arr = []
            docter_date = []
            docter_content = []
            edu_date = i.select('li')
            for i in edu_date :
                date = i.select_one('dt')
                content = i.select_one('dd')

                if content == None :
                    docter_content.append('')
                if content != None : 
                    content = i.select_one('dd').text.strip()
                    docter_content.append(content)
                if date == None :
                    docter_date.append('')
                    docter_arr.append('')
                elif date != None:
                    divison = date.text.strip().split('~')
                    if len(divison) == 2 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[1])
                    elif len(divison) == 1 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[0])
                    elif len(divison) == 0 :
                        docter_arr.append('')
                        docter_date.append('')
            for tax in range(len(docter_content)):
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                array.append(docter_arr[tax])
                array.append(docter_date[tax])
                array.append(docter_content[tax])
                ws.append(array)
               

    # 논문
    book_if = soup.select('div.thesis_list')
    if book_if != None :
        ws = wb.worksheets[6]
        books = soup.select('div.thesis_list li')
        date = []
        content = []
        for i in books: 
            book_date = i.select_one('em.f_eng').text.strip()
            book_cotent = i.select_one('div.title p').text.strip()
            date.append(book_date)
            content.append(book_cotent)
        for i in range(len(content)):
            try:
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                array.append(date[i])
                array.append('')
                array.append(content[i])
                ws.append(array)
            except :
                print(name, '오류데이터')
            
        print(name,dept)
    wb.save('가톨릭대 서울성모병원.xlsx')
    wb.close()


