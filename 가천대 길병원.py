import requests
import re
from bs4 import BeautifulSoup
import openpyxl

page_nubers = ['https://www.gilhospital.com/web/www/-153',
                'https://www.gilhospital.com/web/www/-156',
                'https://www.gilhospital.com/web/www/-159',
                'https://www.gilhospital.com/web/www/-162',
                'https://www.gilhospital.com/web/www/-168',
                'https://www.gilhospital.com/web/www/-174',
                'https://www.gilhospital.com/web/www/-177',
                'https://www.gilhospital.com/web/www/-180',
                'https://www.gilhospital.com/web/www/-183',
                'https://www.gilhospital.com/web/www/-186',
                'https://www.gilhospital.com/web/www/-189',
                'https://www.gilhospital.com/web/www/-192',
                'https://www.gilhospital.com/web/www/-195 ',  
                'https://www.gilhospital.com/web/www/-198',
                'https://www.gilhospital.com/web/www/-201' , 
                'https://www.gilhospital.com/web/www/-204',
                'https://www.gilhospital.com/web/www/-207',
                'https://www.gilhospital.com/web/www/-213',
                'https://www.gilhospital.com/web/www/-216',
                'https://www.gilhospital.com/web/www/-222',
                'https://www.gilhospital.com/web/www/-225',
                'https://www.gilhospital.com/web/www/-228',
                'https://www.gilhospital.com/web/www/-231',
                'https://www.gilhospital.com/web/www/-234',
                'https://www.gilhospital.com/web/www/-237',
                'https://www.gilhospital.com/web/www/-240',
                'https://www.gilhospital.com/web/www/-243',
                'https://www.gilhospital.com/web/www/-246',
                'https://www.gilhospital.com/web/www/-249',
                'https://www.gilhospital.com/web/www/-165',
                'https://www.gilhospital.com/web/www/-171',
                'https://www.gilhospital.com/web/www/-210',
                'https://www.gilhospital.com/web/www/-219',
                'https://www.gilhospital.com/web/www/-252',
                'https://www.gilhospital.com/web/www/-255',
                'https://www.gilhospital.com/web/www/-485',
                'https://www.gilhospital.com/web/www/-258'
                ]
                

docter_site = []

for get_site in page_nubers :
    site_html = requests.get(get_site)
    site_html_list = BeautifulSoup(site_html.text, 'html.parser')
    docter_box = site_html_list.select('div.thumb')
    for docter_one in docter_box :
        link =  docter_one.a["href"] 
        docter_site.append(link)


doc_num = 694
for docter in docter_site :
    doc_num += 1 # id 값
    res = requests.get(docter) # 진짜 URL
    soup = BeautifulSoup(res.text, 'html.parser')          
    wb = openpyxl.load_workbook('가천대 길병원.xlsx') 


    # 기본정보
    ws = wb.worksheets[0] 
    name = soup.select_one("div.doctor-name") 
    dept = name.select("div.doctor-name > span")
    span_del = name.select("span")
    for script in span_del: # span 제거(담당 과 제거) 
        script.extract()
    name = name.text.strip()
    array = []
    array.append('doc' + str(doc_num).zfill(8))
    array.append('가천대 길병원')
    array.append(name)
    for qe in dept :
        qe = qe.text.strip()
        array.append(qe) # 담당 과
    ws.append(array)
    # 담당과만 나와있는 부분이 없음

    # 담당 클리닉
    ws = wb.worksheets[1] 
    for clinic in dept :
        clinic = clinic.text.strip()
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(clinic)
        ws.append(array)

    # 전문분야
    ws = wb.worksheets[2]
    infom = soup.select("ul.table-list li")[0] 
    span_del = infom.select("span")
    for script in span_del:
        script.extract()
    specil_one = infom.text.strip().replace('‚',',') # , 특수문자 들어가 있음
    specialty = re.split(r',(?![^()]*\))', specil_one)
    for sp in specialty :
        sp = sp.lstrip()
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(sp)
        ws.append(array)



    # 학력
    ws = wb.worksheets[3]
    content = soup.select("div.content-section")[1] 
    docter_content = content.select('tr')
    one_date = []
    one_date2 = []
    one_contet = []
    for docter_one in docter_content :
        D_date = docter_one.select('th')
        D_content = docter_one.select('td')
        for date in D_date :
            date = date.text.strip().split('~')
            if len(date) == 2 :
                one_date.append(date[0])
                one_date2.append(date[1])
            if len(date) == 1 :
                one_date.append(date[0])
                one_date2.append('')
            if len(date) == 0 :
                one_date.append('')
                one_date2.append('')
        for content in D_content:
            content = content.text.strip()
            one_contet.append(content)
    for tax in range(len(one_contet)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(one_date[tax])
        array.append(one_date2[tax])
        array.append(one_contet[tax])
        ws.append(array)


    # 경력
    ws = wb.worksheets[3]
    content = soup.select("div.content-section")[2] 
    docter_content = content.select('tr')
    one_date = []
    one_date2 = []
    one_contet = []
    for docter_one in docter_content :
        D_date = docter_one.select('th')
        D_content = docter_one.select('td')
        for date in D_date :
            date = date.text.strip().split('~')
            if len(date) == 2 :
                one_date.append(date[0])
                one_date2.append(date[1])
            if len(date) == 1 :
                one_date.append(date[0])
                one_date2.append('')
            if len(date) == 0 :
                one_date.append('')
                one_date2.append('')
        for content in D_content:
            content = content.text.strip()
            one_contet.append(content)
    for tax in range(len(one_contet)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(one_date[tax])
        array.append(one_date2[tax])
        array.append(one_contet[tax])
        ws.append(array)



    ws = wb.worksheets[4]
    content = soup.select("div.content-section")[3] # 학회
    docter_content = content.select('tr')
    one_date = []
    one_date2 = []
    one_contet = []
    for docter_one in docter_content :
        D_date = docter_one.select('th')
        D_content = docter_one.select('td')
        for date in D_date :
            date = date.text.strip().split('~')
            if len(date) == 2 :
                one_date.append(date[0])
                one_date2.append(date[1])
            if len(date) == 1 :
                one_date.append(date[0])
                one_date2.append('')
            if len(date) == 0 :
                one_date.append('')
                one_date2.append('')
        for content in D_content:
            content = content.text.strip()
            one_contet.append(content)
    for tax in range(len(one_contet)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(one_date[tax])
        array.append(one_date2[tax])
        array.append(one_contet[tax])
        ws.append(array)
        
    books = soup.select("div.section-thesis table tr") # 논문
    book_date = []
    book_content = []
    ws = wb.worksheets[5]
    for i in books :
        date = i.select_one('th').text.strip()
        content = i.select_one('td').text.strip()
        book_date.append(date)
        book_content.append(content)
    for i in range(len(book_content)) :
        try:
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(book_date[i])
            array.append('')
            array.append(book_content[i])
            ws.append(array)
        except :
            print(name, '오류데이터')
            continue    # 문정근
    wb.save('가천대 길병원.xlsx')
    wb.close()
