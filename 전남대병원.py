import requests
import re
from bs4 import BeautifulSoup
import openpyxl
# 전남대병원
res = requests.get('https://www.cnuh.com/medical/info/dept.cs?m=42', verify=False)
soup = BeautifulSoup(res.content, "html.parser", from_encoding='utf=8')
dept_list = soup.select('li.sectionList ul')
dept_link = []
doc_num = 2186
for i in dept_list :
    i = i.select('li')[1]
    doct_dept = i.select_one('a')
    link = 'https://www.cnuh.com/medical/info/dept.cs' + doct_dept['href']
    dept_link.append(link)
for i in dept_link :
    res = requests.get(i, verify=False) 
    soup = BeautifulSoup(res.text, 'html.parser')
    real_link = soup.select('li.intro a')

    for doc_link in real_link :
        doc_num += 1
        doc_link = 'https://www.cnuh.com/' + doc_link["href"]
        res = requests.get(doc_link, verify=False)
        soup = BeautifulSoup(res.content, "html.parser", from_encoding='utf=8')
        wb = openpyxl.load_workbook('전남대.xlsx') # 엑셀 쓰기

        # 기본정보
        ws = wb.worksheets[0]
        span_del = soup.select_one('div.introHeader dt span')
        dept = soup.select_one('div.introHeader dt span').text.strip()
        for i in span_del :
            i.extract()
        name = soup.select_one('div.introHeader dt').text
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append('전남대병원')
        array.append(name)
        array.append(dept)
        ws.append(array)
        print(name,dept)

        # 전문분야
        ws = wb.worksheets[1]
        specil_one = soup.select_one('div.introHeader dd').text.strip()
        specialty = re.split(r',(?![^()]*\))', specil_one)
        for sp in specialty :
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            sp = sp.lstrip()
            array.append(sp)
            ws.append(array)

        # 학력
        ws = wb.worksheets[2]
        tab = soup.select('div.introTabCont')[0]
        education = tab.select('div.viewArea')[0] 
        edu_all = education.select('dd')
        date_start = []
        date_end = []
        docter_content = []
        
        for i in edu_all :
            edu_date = i.select_one('span.date')
            edu_content = i.select_one('span.txt')
            if edu_content == None :
                docter_content.append('')
            if  edu_content != None : 
                edu_content = i.select_one('span.txt').text.strip()
                docter_content.append(edu_content)
            if edu_date == None :
                date_start.append('')
                date_end.append('')
            elif edu_date != None:

                division = edu_date.text.strip().replace(' ','').replace('-', '~').split('~')
                if len(division) == 2 :
                    date_start.append(division[0])
                    date_end.append(division[1])
                elif len(division) == 1 :
                    date_start.append(division[0])
                    date_end.append(division[0])
                elif len(division) == 0 :
                    date_start.append('')
                    date_end.append('')

        print(date_start)
        print(docter_content)
        for i in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(date_start[i])
            array.append(date_end[i])
            array.append(docter_content[i])
            ws.append(array)


        # 경력
        ws = wb.worksheets[3]
        education = tab.select('div.viewArea')[1] 
        edu_all = education.select('dd')
        date_start = []
        date_end = []
        docter_content = []

        for i in edu_all :
            edu_date = i.select_one('span.date')
            edu_content = i.select_one('span.txt')
            if edu_content == None :
                docter_content.append('')
            if edu_content != None : 
                edu_content = i.select_one('span.txt').text.strip()
                docter_content.append(edu_content)
            if edu_date == None :
                date_start.append('')
                date_end.append('')
            elif edu_date != None:
                division = edu_date.text.strip().replace(' ','').replace('-', '~').split('~')
                if len(division) == 2 :
                    date_start.append(division[0])
                    date_end.append(division[1])
                elif len(division) == 1 :
                    date_start.append(division[0])
                    date_end.append(division[0])
                elif len(division) == 0 :
                    date_start.append('')
                    date_end.append('')

        for i in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(date_start[i])
            array.append(date_end[i])
            array.append(docter_content[i])
            ws.append(array)

         # 경력 중 수상경력
        ws = wb.worksheets[4]
        education = tab.select('div.viewArea')[2] 
        edu_all = education.select('dd')
        date_start = []
        date_end = []
        docter_content = []

        for i in edu_all :
            edu_date = i.select_one('span.date')
            edu_content = i.select_one('span.txt')
            if edu_content == None :
                docter_content.append('')
            if edu_content != None : 
                edu_content = i.select_one('span.txt').text.strip()
                docter_content.append(edu_content)
            if edu_date == None :
                date_start.append('')
                date_end.append('')
            elif edu_date != None:
                division = edu_date.text.strip().replace(' ','').replace('-', '~').split('~')
                if len(division) == 2 :
                    date_start.append(division[0])
                    date_end.append(division[1])
                elif len(division) == 1 :
                    date_start.append(division[0])
                    date_end.append(division[0])
                elif len(division) == 0 :
                    date_start.append('')
                    date_end.append('')

        for i in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(date_start[i])
            array.append(date_end[i])
            array.append(docter_content[i])
            ws.append(array)

        # 논문
        ws = wb.worksheets[5] 
        tab = soup.select('div.introTabCont')[1]
        education = tab.select('div.viewArea')[0]
        edu_all = education.select('dd')
        date_start = []
        date_end = []
        docter_content = []

        for i in edu_all :
            edu_date = i.select_one('span.date')
            edu_content = i.select_one('span.txt')
            if edu_content == None :
                docter_content.append('')
            if edu_content != None : 
                edu_content = i.select_one('span.txt').text.strip()
                docter_content.append(edu_content)
            if edu_date == None :
                date_start.append('')
                date_end.append('')
            elif edu_date != None:
                division = edu_date.text.strip().replace(' ','').replace('-', '~').split('~')
                if len(division) == 2 :
                    date_start.append(division[0])
                    date_end.append(division[1])
                elif len(division) == 1 :
                    date_start.append(division[0])
                    date_end.append(division[0])
                elif len(division) == 0 :
                    date_start.append('')
                    date_end.append('')
                elif len(division) > 2 : # 정인석 오류 데이터 처리
                    date_start.append('')
                    date_end.append('')

        for i in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(date_start[i])
            array.append(date_end[i])
            array.append(docter_content[i])
            ws.append(array)

        # 학회활동
        ws = wb.worksheets[6]
        tab = soup.select('div.introTabCont')[2]
        education = tab.select('div.viewArea')[0] 
        edu_all = education.select('dd')
        date_start = []
        date_end = []
        docter_content = []

        for i in edu_all :
            edu_date = i.select_one('span.date')
            edu_content = i.select_one('span.txt')
            if edu_content == None :
                docter_content.append('')
            if edu_content != None : 
                edu_content = i.select_one('span.txt').text.strip()
                docter_content.append(edu_content)
            if edu_date == None :
                date_start.append('')
                date_end.append('')
            if edu_date != None:
                division = edu_date.text.strip().replace(' ','').replace('-', '~').split('~')
                if len(division) == 2 :
                    date_start.append(division[0])
                    date_end.append(division[1])
                elif len(division) == 1 :
                    date_start.append(division[0])
                    date_end.append(division[0])
                elif len(division) == 0 :
                    date_start.append('')
                    date_end.append('')

        for i in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(date_start[i])
            array.append(date_end[i])
            array.append(docter_content[i])
            ws.append(array)

        wb.save('전남대.xlsx')
        wb.close()