# Mediaipuls-Project
## - 프로젝트 개요
- 전국 상급병원에 있는 의사 정보를 제공해주는 서비스를 위한 데이터 베이스 구축
- 파이썬 웹 크롤링으로 데이터를 수집하였으며, 요류데이터 수정을 위해 엑셀로 출력
## - 동작 원리
1. 해당 병원 접속 후 진료과 리스트 페이지 접속
```python
res = requests.get('https://www.dcmc.co.kr/content/01reserv/01_01.asp', verify = False) # URL
soup = BeautifulSoup(res.text, 'html.parser')
dept_list = soup.select('div.ov')
dept_link = []
for i in dept_list :
    i = i.select('a')[1]
    link = 'https://www.dcmc.co.kr' + i["href"]
    dept_link.append(link)
```
- div 태그에 ov 클래스를 가지는 요소중에 a태그에 2번째 요소에 href 주소를 찾습니다.
- 'https://www.dcmc.co.kr' + 찾은 href를 합쳐 주어 dept_link에 배열에 삽입해 줍니다.


<img src = "./img/dept.png">
<br>
<br>

2. 해당 과 의사 리스트 페이지 접속
<img src = "./img/doc_list.png">
<br>
<br>

```python
doc_link = []
for i in dept_link :
    res = requests.get(i, verify = False)
    soup = BeautifulSoup(res.text, 'html.parser')
    real_link = soup.select('div.tar')
    for i in real_link : 
        i = i.select('a')[0]
        link = 'https://www.dcmc.co.kr' + i["href"] 
        doc_link.append(link)
```

- 찾은 학과 list 배열을 for문으로 구현해 의료진소개 href를 다시 doc_link에 삽입해줍니다.

3. 해당 의사에 원하는 데이터 추출

<img src = "./img/doc.png">
<br>
<br>

```python
for docter in doc_link :
    res = requests.get(docter, verify = False) # URL
    soup = BeautifulSoup(res.content, "html.parser", from_encoding='utf=8')
```    
- 찾은 의사 한명에 주소 값 리스트를 for문으로 돌려주면 전체 의사 데이터를 추출할수 있습니다.

<!-- 4. 
<img src = "./img/doc.png">

```python
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

``` -->