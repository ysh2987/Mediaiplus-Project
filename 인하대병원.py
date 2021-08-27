import requests
import re
from bs4 import BeautifulSoup
import openpyxl



sites = []
for i in range(1, 13): # 1페이지 부터 12페이지
    base_site = 'https://www.inha.com/page/department/medicine/doctor?&currentPage={}&dataPerPage=20'.format(i)
    sites.append(base_site)
real_link = []
for site in sites :
    site_html = requests.get(site)
    site_html_list = BeautifulSoup(site_html.text, 'html.parser')
    dept_box = site_html_list.find_all('li', attrs={'class':"doc-box"})
    for i in dept_box :
        link = "https://www.inha.com"+ i.a["href"]
        real_link.append(link)
doc_num = 236
for doc_link in real_link :
    doc_num += 1 
    res = requests.get(doc_link) 
    soup = BeautifulSoup(res.text, 'html.parser') 
    wb = openpyxl.load_workbook('인하대.xlsx')
    

    # 기본정보
    ws = wb.worksheets[0]
    name = soup.select_one("p.name").text.strip() # 이름
    dept = soup.select_one("p.dept").text.strip() # 학과
    array = []
    array.append('doc' + str(doc_num).zfill(8))
    array.append(name)
    array.append(dept)
    ws.append(array)

    # 전문 분야
    ws = wb.worksheets[1]
    specil_one = soup.select_one("div.prg-cont p").text.strip() 
    specialty = re.split(r',(?![^()]*\))', specil_one)
    for sp in specialty :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        sp = sp.lstrip()
        array.append(sp)
        ws.append(array)

    # 담당 센터
    ws = wb.worksheets[2]
    cnt = soup.select("ul.list-type1 li")
    for i in cnt : 
        i = i.text.strip()
        blank_del = ' '.join(i.split())
        cnt_arr = []
        cnt_arr.append('doc' + str(doc_num).zfill(8))
        cnt_arr.append(blank_del)
        ws.append(cnt_arr)

    # 학력
    ws = wb.worksheets[3]
    sel = soup.select("div.profile")[0]
    p_del = sel.select("li p")
    all_content = sel.select("li")
    docter_arr = []
    docter_date = []
    docter2_arr = []

    for i in p_del :
        i = i.text.strip().replace('\n', '').replace(' ', '').split('~')
        if len(i) == 2 :
            docter_arr.append(i[0])
            docter_date.append(i[1])
        elif len(i) == 1 :
            docter_arr.append(i[0])
            docter_date.append('')
        elif len(i) == 0 :
            docter_arr.append('')
            docter_date.append('')
    for script in p_del:
        script.extract()
    for divison in all_content:
        w = divison.get_text(strip=True)
        docter2_arr.append(w)
    for tex in range(len(all_content)) :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter2_arr[tex])
        ws.append(array)


    # 주요 경력
    ws = wb.worksheets[4]
    sel = soup.select("div.profile")[1]
    p_del = sel.select("li p")
    all_content = sel.select("li")
    docter_arr = []
    docter_date = []
    docter2_arr = []

    for i in p_del :
        i = i.text.strip().replace('\n', '').replace(' ', '').split('~')
        if len(i) == 2 :
            docter_arr.append(i[0])
            docter_date.append(i[1])
        elif len(i) == 1 :
            docter_arr.append(i[0])
            docter_date.append('')
        elif len(i) == 0 :
            docter_arr.append('')
            docter_date.append('')
    for script in p_del:
        script.extract()
    for divison in all_content:
        w = divison.get_text(strip=True)
        docter2_arr.append(w)
    for tex in range(len(all_content)) :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter2_arr[tex])
        ws.append(array)

    #  학회 활동
    ws = wb.worksheets[5]
    sel = soup.select("div.profile")[2]
    p_del = sel.select("li p")
    all_content = sel.select("li")
    docter_arr = []
    docter_date = []
    docter2_arr = []

    for i in p_del :
        i = i.text.strip().replace('\n', '').replace(' ', '').split('~')
        if len(i) == 2 :
            docter_arr.append(i[0])
            docter_date.append(i[1])
        elif len(i) == 1 :
            docter_arr.append(i[0])
            docter_date.append('')
        elif len(i) == 0 :
            docter_arr.append('')
            docter_date.append('')
    for script in p_del:
        script.extract()
    for divison in all_content:
        w = divison.get_text(strip=True)
        docter2_arr.append(w)
    for tex in range(len(all_content)) :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter2_arr[tex])
        ws.append(array)

    # 수상경력
    ws = wb.worksheets[6]
    sel = soup.select("div.profile")[3]
    p_del = sel.select("li p")
    all_content = sel.select("li")
    docter_arr = []
    docter_date = []
    docter2_arr = []

    for i in p_del :
        i = i.text.strip().replace('\n', '').replace(' ', '').split('~')
        if len(i) == 2 :
            docter_arr.append(i[0])
            docter_date.append(i[1])
        elif len(i) == 1 :
            docter_arr.append(i[0])
            docter_date.append('')
        elif len(i) == 0 :
            docter_arr.append('')
            docter_date.append('')
    for script in p_del:
        script.extract()
    for divison in all_content:
        w = divison.get_text(strip=True)
        docter2_arr.append(w)
    for tex in range(len(all_content)) :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter2_arr[tex])
        ws.append(array)

    #  논문
    ws = wb.worksheets[7]
    sel = soup.select("div.profile")[4]
    p_del = sel.select("li p")
    all_content = sel.select("li")
    docter_arr = []
    docter_date = []
    docter2_arr = []

    for i in p_del :
        i = i.text.strip().replace('\n', '').replace(' ', '').split('~')
        if len(i) == 2 :
            docter_arr.append(i[0])
            docter_date.append(i[1])
        elif len(i) == 1 :
            docter_arr.append(i[0])
            docter_date.append('')
        elif len(i) == 0 :
            docter_arr.append('')
            docter_date.append('')
    for script in p_del:
        script.extract()
    for divison in all_content:
        w = divison.get_text(strip=True)
        docter2_arr.append(w)
    for tex in range(len(all_content)) :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[tex])
        array.append(docter_date[tex])
        array.append(docter2_arr[tex])
        ws.append(array)

    wb.save('인하대.xlsx')
    wb.close()

