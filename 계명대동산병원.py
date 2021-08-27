import requests
import re
from bs4 import BeautifulSoup
import openpyxl



res = requests.get('https://dongsan.dsmc.or.kr:49870/content/02health/01_01.php') # URL
soup = BeautifulSoup(res.text, 'html.parser')
ww = soup.select('div.treat_list li')
dept_link = []
for i in ww :
    i = i.select_one('a')
    link = 'https://dongsan.dsmc.or.kr:49870' + i["href"]
    dept_link.append(link)

all_list = []
doc_num = 2615
for i in dept_link :
    
    res = requests.get(i)
    soup = BeautifulSoup(res.text, 'html.parser')
    real_link = soup.select('a.navy')
    for i in real_link : 
        a = i["href"].split('pop')[1]
        doc_num += 1
        res = requests.get('https://dongsan.dsmc.or.kr:49870/content/02health/01_0102pop'+a)
        soup = BeautifulSoup(res.content, "html.parser", from_encoding='utf=8')
        wb = openpyxl.load_workbook('계명대.xlsx') # 엑셀 쓰기
        

        # 기본정보
        name = soup.select_one('p.btxt').text 
        dept = soup.select_one('p.stxt').text.replace('[','').replace(']','')
        ws = wb.worksheets[0]
        array = []
        array.append('doc' + str(doc_num).zfill(8))
        array.append('계명대학교 동산병원')
        array.append(name) 
        array.append(dept)
        ws.append(array)
        print(name,dept)
        

        # 전문분야
        ws = wb.worksheets[1]
        specil = soup.select_one('div.box dd')
        if specil != None:
            specil_one = soup.select_one('div.box dl dd').text.strip()
            specialty = re.split(r',(?![^()]*\))', specil_one)
            for sp in specialty :
                array = []
                array.append('doc' + str(doc_num).zfill(8))
                sp = sp.lstrip()
                array.append(sp)
                ws.append(array)

        all_list = soup.select('div.box1')
        for i in all_list :
            title = i.select_one('h3.tit').text.strip()
            
            # 학력/경력
            if title in '학력/경력' :
                docter_arr = []
                docter_date = []
                docter_content = []
                edu_date = i.select('tr')
                for i in edu_date :
                    date = i.select_one('th')
                    content = i.select_one('td')
                    if content == None :
                        docter_content.append('')
                    if content != None : 
                        content = i.select_one('td').text.strip()
                        docter_content.append(content)
                    if date == None :
                        docter_date.append('')
                        docter_arr.append('')
                    elif date != None:
                        divison = date.text.strip().replace(' ','').replace('-', '~').split('~')
                        if len(divison) == 2 :
                            docter_arr.append(divison[0])
                            docter_date.append(divison[1])
                        elif len(divison) == 1 :
                            docter_arr.append(divison[0])
                            docter_date.append(divison[0])
                        elif len(divison) == 0 :
                            docter_arr.append('')
                            docter_date.append('')
                ws = wb.worksheets[2]
                for tax in range(len(docter_content)):
                    array = []
                    array.append('doc' + str(doc_num).zfill(8))
                    array.append(docter_arr[tax])
                    array.append(docter_date[tax])
                    array.append(docter_content[tax])
                    ws.append(array)

            # 학회활동
            if title in '학회활동' : 
                docter_arr = []
                docter_date = []
                docter_content = []
                edu_date = i.select('tr')
                for i in edu_date :
                    date = i.select_one('th')
                    content = i.select_one('td')
                    if content == None :
                        docter_content.append('')
                    if content != None : 
                        content = i.select_one('td').text.strip()
                        docter_content.append(content)
                    if date == None :
                        docter_date.append('')
                        docter_arr.append('')
                    elif date != None:
                        divison = date.text.strip().replace(' ','').replace('-', '~').split('~')
                        if len(divison) == 2 :
                            docter_arr.append(divison[0])
                            docter_date.append(divison[1])
                        elif len(divison) == 1 :
                            docter_arr.append(divison[0])
                            docter_date.append(divison[0])
                        elif len(divison) == 0 :
                            docter_arr.append('')
                            docter_date.append('')
                ws = wb.worksheets[3]
                for tax in range(len(docter_content)):
                    array = []
                    array.append('doc' + str(doc_num).zfill(8))
                    array.append(docter_arr[tax])
                    array.append(docter_date[tax])
                    array.append(docter_content[tax])
                    ws.append(array)

            # 수상내역
            if title in '수상' :
                docter_arr = []
                docter_date = []
                docter_content = []
                edu_date = i.select('tr')
                for i in edu_date :
                    date = i.select_one('th')
                    content = i.select_one('td')
                    if content == None :
                        docter_content.append('')
                    if content != None : 
                        content = i.select_one('td').text.strip()
                        docter_content.append(content)
                    if date == None :
                        docter_date.append('')
                        docter_arr.append('')
                    elif date != None:
                        divison = date.text.strip().replace(' ','').replace('-', '~').split('~')
                        if len(divison) == 2 :
                            docter_arr.append(divison[0])
                            docter_date.append(divison[1])
                        elif len(divison) == 1 :
                            docter_arr.append(divison[0])
                            docter_date.append(divison[0])
                        elif len(divison) == 0 :
                            docter_arr.append('')
                            docter_date.append('')
                ws = wb.worksheets[4]
                for tax in range(len(docter_content)):
                    array = []
                    array.append('doc' + str(doc_num).zfill(8))
                    array.append(docter_arr[tax])
                    array.append(docter_date[tax])
                    array.append(docter_content[tax])
                    ws.append(array)

        # 논문
        req = requests.get('https://dongsan.dsmc.or.kr:49870/content/02health/01_0102pop2'+a)
        soup = BeautifulSoup(req.content, "html.parser", from_encoding='utf=8')
        books = soup.select('div.box1 table')
        docter_arr = []
        docter_date = []
        docter_content = []
        for i in books : 
            table_date = i.select('tr')[-1].select('td')
            table_title = i.select('tr')[0].select('td')
            if len(table_date) == 0 :
                docter_date.append('')
                docter_arr.append('')
            for i in table_date : 
                i = i.text.strip()
                docter_date.append('')
                docter_arr.append(i)
            for i in table_title : 
                j = i.text.strip()
                docter_content.append(j)

        ws = wb.worksheets[5]
        for qe in range(len(docter_content)):
            array = []
            array.append('doc' + str(doc_num).zfill(8))
            array.append(docter_arr[qe])
            array.append(docter_date[qe])
            array.append(docter_content[qe])
            ws.append(array)

        wb.save('계명대.xlsx')
        wb.close()
