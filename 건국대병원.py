import requests
import re
from bs4 import BeautifulSoup
import openpyxl

res = requests.get('https://www.kuh.ac.kr/medical/dept/deptList.do') # URL
soup = BeautifulSoup(res.text, 'html.parser')
all_ul = soup.select('ul.bullet-col02')
dept_link = []
for i in all_ul :
    www = i.select('li')[1]
    for i in www :
        link = 'https://www.kuh.ac.kr/medical/dept/' + i["href"]
        dept_link.append(link)
dr_sid = []
dept_cd = []
for i in dept_link :
    res = requests.get(i)
    soup = BeautifulSoup(res.text, 'html.parser')
    a_hre = soup.select('div.docSchedule div a')

    for i in a_hre :
        i = i["onclick"][11:-2].replace("\'",'').replace(' ','').split(',')
        dr_sid.append(i[0])
        dept_cd.append(i[1])
doc_num = 3263
all_list = []
for docter in range(len(dept_cd)):
    doc_num += 1
    res = requests.get('https://www.kuh.ac.kr/doctor/basicInfo.do?dr_sid={}&dept_cd={}'.format(dr_sid[docter],dept_cd[docter])) # URL
    soup = BeautifulSoup(res.content, "html.parser", from_encoding='utf=8')
    wb = openpyxl.load_workbook('건국.xlsx') # 엑셀 쓰기

    # 기본정보
    ws = wb.worksheets[0]
    name = soup.select_one('div.desc strong').text.strip() # 이름
    dpt_one = soup.select_one('div.desc span').text.strip().replace('\r', '').replace('\n', '').replace('\t', '').replace('[', '').replace(']', '')
    dpt_lstrip = re.split(r',(?![^()]*\))', dpt_one) # 담당부서
    dpt_arr =  []
    for i in dpt_lstrip :
        i = i.lstrip()
        dpt_arr.append(i)
    dept = dpt_arr[0] # 학과
    array = []
    array.append('doc' + str(doc_num).zfill(8))
    array.append('건국 대학교병원')
    array.append(name)
    array.append(dept)
    ws.append(array)

    # 담당부서 
    ws = wb.worksheets[1]
    for clinic in dpt_arr :
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(clinic)
        ws.append(array)

    # 전문분야
    ws = wb.worksheets[2]
    specil_one = soup.select_one('div.desc p').text.strip() 
    specialty = re.split(r',(?![^()]*\))', specil_one) 
    for sp in specialty : 
        array = [] 
        sp = sp.lstrip()
        array.append('doc' + str(doc_num).zfill(8))
        array.append(sp)
        ws.append(array)

    # 학력
    ws = wb.worksheets[3]
    docter_arr = []
    docter_date = []
    docter_content = []
    item_list = soup.select('div.icon01 li') 
    for i in item_list :
        date = i.text.strip()
        if '|' in date :
            i = i.text.strip().replace('\xa0', '').split('|')
            docter_content.append(i[1])
            divison = i[0].split('-')
            if len(divison) == 2 :
                docter_arr.append(divison[0])
                docter_date.append(divison[1])
            elif len(divison) == 1 :
                docter_arr.append(divison[0])
                docter_date.append(divison[0])

        elif '|' not in date :
            docter_content.append(date)
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
    item_list = soup.select('div.icon02 ul')
    if len(item_list) == 1 :
        ws = wb.worksheets[4]
        docter_arr = []
        docter_date = []
        docter_content = []
        item_award = soup.select('div.icon02 ul')[0] 
        item_list = item_award.select('li')
        for i in item_list :
            date = i.text.strip()
            if '|' in date :
                i = i.text.strip().replace('\xa0', '').split('|')
                docter_content.append(i[1])
                divison = i[0].split('-')
                if len(divison) == 2 :
                    docter_arr.append(divison[0])
                    docter_date.append(divison[1])
                elif len(divison) == 1 :
                    docter_arr.append(divison[0])
                    docter_date.append(divison[0])
                
            elif '|' not in date :
                docter_content.append(date)
                docter_arr.append('')
                docter_date.append('')
        for qe in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(docter_arr[qe])
            array.append(docter_date[qe])
            array.append(docter_content[qe])
            ws.append(array)   

    # 수상
    item_list = soup.select('div.icon02 ul')
    if len(item_list) == 2 :
        ws = wb.worksheets[5]
        docter_arr = []
        docter_date = []
        docter_content = []
        item_award = soup.select('div.icon02 ul')[1] 
        item_list = item_award.select('li')
        for i in item_list :
            date = i.text.strip()
            if '|' in date :
                i = i.text.strip().replace('\xa0', '').split('|')
                docter_content.append(i[1])
                divison = i[0].split('-')
                if len(divison) == 2 :
                    docter_arr.append(divison[0])
                    docter_date.append(divison[1])
                elif len(divison) == 1 :
                    docter_arr.append(divison[0])
                    docter_date.append(divison[0])
                
            elif '|' not in date :
                docter_content.append(date)
                docter_arr.append('')
                docter_date.append('')
        for qe in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(docter_arr[qe])
            array.append(docter_date[qe])
            array.append(docter_content[qe])
            ws.append(array)     

    

    # 학회
    ws = wb.worksheets[6]
    docter_arr = []
    docter_date = []
    docter_content = []
    item_list = soup.select('div.icon03 li') 
    for i in item_list :
        date = i.text.strip()
        if '|' in date :
            i = i.text.strip().replace('\xa0', '').split('|')
            docter_content.append(i[1])
            divison = i[0].split('-')
            if len(divison) == 2 :
                docter_arr.append(divison[0])
                docter_date.append(divison[1])
            elif len(divison) == 1 :
                docter_arr.append(divison[0])
                docter_date.append(divison[0])

        elif '|' not in date :
            docter_content.append(date)
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
    ws = wb.worksheets[7]
    qq = requests.get('https://www.kuh.ac.kr/doctor/paper.do?dr_sid={}&dept_cd={}'.format(dr_sid[docter],dept_cd[docter]))
    ww = BeautifulSoup(qq.content, "html.parser", from_encoding='utf=8')
    books = ww.select('ul.paper li')
    docter_arr = []
    docter_date = []
    docter_content = []
    for i in books :    
        content = i.select_one('strong').text.strip()
        docter_arr.append('')
        docter_date.append('')
        docter_content.append(content)
    for qe in range(len(docter_content)):
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append(docter_arr[qe])
        array.append(docter_date[qe])
        array.append(docter_content[qe])
        ws.append(array)     

    wb.save('건국.xlsx')
    wb.close()