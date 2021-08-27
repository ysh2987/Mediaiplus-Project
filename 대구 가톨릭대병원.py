import requests
import re
from bs4 import BeautifulSoup
import openpyxl

res = requests.get('https://www.dcmc.co.kr/content/01reserv/01_01.asp', verify = False) # URL
soup = BeautifulSoup(res.text, 'html.parser')
dept_list = soup.select('div.ov')
dept_link = []
doc_num = 2875
for i in dept_list :
    i = i.select('a')[1]
    link = 'https://www.dcmc.co.kr' + i["href"]
    dept_link.append(link)
doc_link = []
for i in dept_link :
    res = requests.get(i, verify = False)
    soup = BeautifulSoup(res.text, 'html.parser')
    real_link = soup.select('div.tar')
    for i in real_link : 
        i = i.select('a')[0]
        link = 'https://www.dcmc.co.kr' + i["href"] 
        doc_link.append(link)


for docter in doc_link :
    doc_num += 1
    res = requests.get(docter, verify = False) # URL
    soup = BeautifulSoup(res.content, "html.parser", from_encoding='utf=8')
    wb = openpyxl.load_workbook('대구.xlsx')
    
    
    # 기본정보
    ws = wb.worksheets[0]
    name = soup.select_one('p.name').text.strip() 
    dept = soup.select_one('p.team').text.strip() 
    array = []
    array.append('doc' + str(doc_num).zfill(8))
    array.append('대구 가톨릭대학병원')
    array.append(name)
    array.append(dept)
    ws.append(array)
    print(name,dept)

    # 전문분야
    ws = wb.worksheets[1]
    specil_one = soup.select_one('dl.part dd').text.strip() 
    specialty = re.split(r',(?![^()]*\))', specil_one) 
    for sp in specialty :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        sp = sp.lstrip()
        array.append(sp)
        ws.append(array)

    # 학력
    ws = wb.worksheets[2]
    ul_list = soup.select('ul.list')[0] 
    strong_del = ul_list.select('li strong')
    edu_content = ul_list.select('li')
    docter_arr = []
    docter_date = []
    docter_content = []

    for i in edu_content:
        date = i.select_one('strong')
        if date == None :
            docter_date.append('')
            docter_arr.append('')
        elif date != None:
            cnt = date.text.strip().replace(' ','').replace('-', '~').replace('년','.').replace('월','.').replace('/','.').split('~')
            if len(cnt) == 2 :
                docter_arr.append(cnt[0])
                docter_date.append(cnt[1])
            elif len(cnt) == 1 :
                docter_arr.append(cnt[0])
                docter_date.append(cnt[0])
            elif len(cnt) == 0 :
                docter_arr.append('')
                docter_date.append('')
    for i in strong_del : # strong 태그내용 제거
        i.extract()
    for i in edu_content :  
        i = i.text.strip().replace('\t','')
        docter_content.append(i)
    for tex in range(len(docter_content)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter_content[tex])
        ws.append(array)

    # 경력
    ws = wb.worksheets[3]
    ul_list = soup.select('ul.list')[1] 
    strong_del = ul_list.select('li strong')
    edu_content = ul_list.select('li')
    docter_arr = []
    docter_date = []
    docter_content = []

    for i in edu_content:
        date = i.select_one('strong')
        if date == None :
            docter_date.append('')
            docter_arr.append('')
        elif date != None:
            cnt = date.text.strip().replace(' ','').replace('-', '~').replace('년','.').replace('월','.').replace('/','.').split('~')
            if len(cnt) == 2 :
                docter_arr.append(cnt[0])
                docter_date.append(cnt[1])
            elif len(cnt) == 1 :
                docter_arr.append(cnt[0])
                docter_date.append(cnt[0])
            elif len(cnt) == 0 :
                docter_arr.append('')
                docter_date.append('')
    for i in strong_del : # strong 태그내용 제거
        i.extract()
    for i in edu_content :  
        i = i.text.strip().replace('\t','')
        docter_content.append(i)
    for tex in range(len(docter_content)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter_content[tex])
        ws.append(array)



    # 학회
    ws = wb.worksheets[4]
    ul_list = soup.select('ul.list')[2] 
    strong_del = ul_list.select('li strong')
    edu_content = ul_list.select('li')
    docter_arr = []
    docter_date = []
    docter_content = []

    for i in edu_content:
        date = i.select_one('strong')
        if date == None :
            docter_date.append('')
            docter_arr.append('')
        elif date != None:
            cnt = date.text.strip().replace(' ','').replace('-', '~').replace('년','.').replace('월','.').replace('/','.').split('~')
            if len(cnt) == 2 :
                docter_arr.append(cnt[0])
                docter_date.append(cnt[1])
            elif len(cnt) == 1 :
                docter_arr.append(cnt[0])
                docter_date.append(cnt[0])
            elif len(cnt) == 0 :
                docter_arr.append('')
                docter_date.append('')
            else :
                docter_arr.append('에러 data')
                docter_date.append('에러 data')

    for i in strong_del : # strong 태그내용 제거
        i.extract()
    for i in edu_content :  
        i = i.text.strip().replace('\t','')
        docter_content.append(i)
    for tex in range(len(docter_content)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter_content[tex])
        ws.append(array)

    # 수상
    ws = wb.worksheets[5]
    ul_list = soup.select('ul.list')[4] 
    strong_del = ul_list.select('li strong')
    edu_content = ul_list.select('li')
    docter_arr = []
    docter_date = []
    docter_content = []

    for i in edu_content:
        date = i.select_one('strong')
        if date == None :
            docter_date.append('')
            docter_arr.append('')
        elif date != None:
            cnt = date.text.strip().replace(' ','').replace('-', '~').replace('년','.').replace('월','.').replace('/','.').split('~')
            if len(cnt) == 2 :
                docter_arr.append(cnt[0])
                docter_date.append(cnt[1])
            elif len(cnt) == 1 :
                docter_arr.append(cnt[0])
                docter_date.append(cnt[0])
            elif len(cnt) == 0 :
                docter_arr.append('')
                docter_date.append('')
            else :
                docter_arr.append('에러 data')
                docter_date.append('에러 data')

    for i in strong_del : # strong 태그내용 제거
        i.extract()
    for i in edu_content :  
        i = i.text.strip().replace('\t','')
        docter_content.append(i)
    for tex in range(len(docter_content)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter_content[tex])
        ws.append(array)

    #  논문
    ws = wb.worksheets[6]
    book = soup.select_one('ul.book_list')
    book_tit = book.select('p.tit')
    book_date = soup.select('div.cont2')
    docter_content = []
    docter_date = []

    for i in book_tit :
        i = i.text.strip()
        docter_content.append(i)

    if len(book_date) == 0 :
        docter_date.append('')
    else :
        for i in book_date :
            i = i.select_one('dl dd').text.strip().replace('/','.')
            docter_date.append(i)
    for tex in range(len(docter_content)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_date[tex])
        array.append(docter_content[tex])
        ws.append(array)

    wb.save('대구.xlsx')
    wb.close()

