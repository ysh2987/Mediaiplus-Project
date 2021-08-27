import requests
import re
from bs4 import BeautifulSoup
import openpyxl

res = requests.get('https://www.cnuh.co.kr/prog/cnuhTreatment/list.do', verify = False) # URL
soup = BeautifulSoup(res.text, 'html.parser')
www = soup.select('div.photos')
dept_link = []
for i in www :
    ww_list = i.select('li')[1]
    for i in ww_list :
        link = 'https://www.cnuh.co.kr' + i["href"]
        dept_link.append(link)

doc_link = []
for i in dept_link :
    res = requests.get(i, verify = False)
    soup = BeautifulSoup(res.text, 'html.parser')
    real_link = soup.select('div.photos')
    for j in real_link : 
        a_link = j.select_one('a')
        link = 'https://www.cnuh.co.kr' + a_link["href"] 
        doc_link.append(link)

doc_num = 1923
for doc_one in doc_link :
    doc_num += 1 # id 값
    res = requests.get(doc_one, verify = False) # URL
    soup = BeautifulSoup(res.content, "html.parser", from_encoding='utf=8')
    wb = openpyxl.load_workbook('충남대.xlsx')

    # 기본정보
    ws = wb.worksheets[0]
    name = soup.select_one('strong.title span').text.strip().replace(' 교수', '')
    dept = soup.select_one('strong.title em').text.strip() 
    array = []
    array.append('doc' + str(doc_num).zfill(8))
    array.append('충남대병원')
    array.append(name)
    array.append(dept)
    ws.append(array)

    # 전문 분야
    ws = wb.worksheets[1]
    specil_one = soup.select_one('ul.data div').text.strip() 
    specialty = re.split(r',(?![^()]*\))', specil_one) # 전문분야
    for sp in specialty :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        sp = sp.lstrip()
        array.append(sp)
        ws.append(array)

    # 학력
    ws = wb.worksheets[2]
    cont_list = soup.select('div.updown_list')[0] 
    cont_div = cont_list.select('div')[0]
    all_li = cont_div.select('li')
    span_del = cont_div.select('li span')
    docter_arr = []
    docter_date = []
    docter_content = []
    
    for i in all_li :
        date = i.select_one('span')
        if date == None :
            docter_date.append('')
            docter_arr.append('')
        elif date != None:
            divison = date.text.strip().replace(' ','').split('~')
            if len(divison) == 2 :
                docter_arr.append(divison[0])
                docter_date.append(divison[1])
            elif len(divison) == 1 :
                docter_arr.append(divison[0])
                docter_date.append(divison[0])
            elif len(divison) == 0 :
                docter_arr.append('')
                docter_date.append('')
    for i in span_del : # date 제거
        i.extract()
    for i in all_li : # date제거 한 텍스트  
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
    cont_list = soup.select('div.updown_list')[0] 
    cont_div = cont_list.select('div')[1]
    all_li = cont_div.select('li')
    span_del = cont_div.select('li span')
    docter_arr = []
    docter_date = []
    docter_content = []

    for i in all_li :
        date = i.select_one('span')
        if date == None :
            docter_date.append('')
            docter_arr.append('')
        elif date != None:
            divison = date.text.strip().replace(' ','').split('~')
            if len(divison) == 2 :
                docter_arr.append(divison[0])
                docter_date.append(divison[1])
            elif len(divison) == 1 :
                docter_arr.append(divison[0])
                docter_date.append(divison[0])
            elif len(divison) == 0 :
                docter_arr.append('')
                docter_date.append('')
            else :
                docter_arr.append('에러 데이터')
                docter_date.append('에러 데이터')
    for i in span_del : # date 제거
        i.extract()
    for i in all_li : # date제거 한 텍스트  
        i = i.text.strip().replace('\t','')
        docter_content.append(i)
    for tex in range(len(docter_content)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter_content[tex])
        ws.append(array)



    # 학회활동
    ws = wb.worksheets[4]
    cont_list = soup.select('div.updown_list')[1] 
    cont_div = cont_list.select('div')[0]
    all_li = cont_div.select('li')
    span_del = cont_div.select('li span')
    docter_arr = []
    docter_date = []
    docter_content = []
    for i in all_li :
        date = i.select_one('span')
        if date == None :
            docter_date.append('')
            docter_arr.append('')
        elif date != None:
            divison = date.text.strip().replace(' ','').split('~')
            if len(divison) == 2 :
                docter_arr.append(divison[0])
                docter_date.append(divison[1])
            elif len(divison) == 1 :
                docter_arr.append(divison[0])
                docter_date.append(divison[0])
            elif len(divison) == 0 :
                docter_arr.append('')
                docter_date.append('')
    for i in span_del : # date 제거
        i.extract()
    for i in all_li : # date제거 한 텍스트  
        i = i.text.strip().replace('\t','')
        docter_content.append(i)
    for tex in range(len(docter_content)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter_content[tex])
        ws.append(array)

    
    test = soup.select('div.updown_list')
    if len(test) == 3 :
        categorys = soup.select('ul.data strong')[3].text.strip()
        category_aw = categorys.find('수상')
        category_book = categorys.find('논문')

        # 수상내역
        if category_aw >= 0 :
            ws = wb.worksheets[5]
            cont_list = soup.select('div.updown_list')[2] 
            cont_div = cont_list.select('div')[0]
            all_li = cont_div.select('li')
            span_del = cont_div.select('li span')
            docter_arr = []
            docter_date = []
            docter_content = []
            for i in all_li :
                date = i.select_one('span')
                if date == None :
                    docter_date.append('')
                    docter_arr.append('')
                elif date != None:
                    divison = date.text.strip().replace(' ','').split('~')
                    if len(divison) == 2 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[1])
                    elif len(divison) == 1 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[0])
                    elif len(divison) == 0 :
                        docter_arr.append('')
                        docter_date.append('')
            for i in span_del : # date 제거
                i.extract()
            for i in all_li : # date제거 한 텍스트  
                i = i.text.strip().replace('\t','')
                docter_content.append(i)
            for tex in range(len(docter_content)):
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                array.append(docter_arr[tex])
                array.append(docter_date[tex])
                array.append(docter_content[tex])
                ws.append(array)

        # 논문
        if category_book >= 0 :
            ws = wb.worksheets[6]
            cont_list = soup.select('div.updown_list')[2] 
            cont_div = cont_list.select('div')[0]
            all_li = cont_div.select('li')
            span_del = cont_div.select('li span')
            docter_arr = []
            docter_date = []
            docter_content = []
            for i in all_li :
                date = i.select_one('span')
                if date == None :
                    docter_date.append('')
                    docter_arr.append('')
                elif date != None:
                    divison = date.text.strip().replace(' ','').split('~')
                    if len(divison) == 2 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[1])
                    elif len(divison) == 1 :
                        docter_arr.append(divison[0])
                        docter_date.append(divison[0])
                    elif len(divison) == 0 :
                        docter_arr.append('')
                        docter_date.append('')
            for i in span_del : # date 제거
                i.extract()
            for i in all_li : # date제거 한 텍스트  
                i = i.text.strip().replace('\t','')
                docter_content.append(i)
            for tex in range(len(docter_content)):
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                array.append(docter_arr[tex])
                array.append(docter_date[tex])
                array.append(docter_content[tex])
                ws.append(array)

    wb.save('충남대.xlsx')
    wb.close()
