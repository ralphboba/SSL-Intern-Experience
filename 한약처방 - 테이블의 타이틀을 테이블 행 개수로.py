#안녕하세요! 해당 코드를 작성한 SSL 두은혁 인턴입니다. 해당 코드 사용방법을 알려드리겠습니다. 원하시는 테이블의 타이틀의 CSS selector 코드를 찾으신 다음에
#19번째 줄 괄호 안에 있는 정보를 해당 CSS selector 코드로 바꿔주시면 됩니다! 20번째 줄에는 예시가 있습니다.


import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

def process_url(url):
    # 웹 페이지 요청
    response = requests.get(url)
    response.raise_for_status()  # 요청이 성공했는지 확인

    # BeautifulSoup을 사용하여 HTML 파싱
    soup = BeautifulSoup(response.content, 'html.parser')

    # 모든 H2[class="depth1_title"] 요소 찾기
    h2_elements = soup.select('H4[class="depth2_title"]')
                                #view01 .contents H2[class="depth1_title"]
    # 각 H2 요소의 텍스트 추출 및 테이블 첫 번째 열의 행 개수에 따라 반복
    data = []
    if not h2_elements:
        data.append(("정보 없음", url))
    else:
        for h2 in h2_elements:
            text = h2.get_text(strip=True)
            
            # 해당 제목 다음에 나오는 테이블 찾기
            table = h2.find_next('table')
            if table:
                first_column_cells = table.find_all('tr')[1:]  # 첫 번째 행은 헤더로 가정하고 제외
                row_count = len(first_column_cells)
                data.extend([(text, url)] * (row_count + 1))
            else:
                data.append(("정보 없음", url))
    
    return data




# 여러 URL 설정
urls = [
"https://oasis.kiom.re.kr/oasis/herb/monoDetailView_M05.jsp?idx=1&tab=5#view02",
"https://oasis.kiom.re.kr/oasis/herb/monoDetailView_M05.jsp?idx=2&tab=5#view02",
"https://oasis.kiom.re.kr/oasis/herb/monoDetailView_M05.jsp?idx=3&tab=5#view02"



]
# 모든 URL 처리
all_data = []
for url in urls:
    result = process_url(url)
    all_data.extend(result)

# 데이터프레임 생성
df = pd.DataFrame(all_data, columns=['Title', 'URL'])

# 엑셀 파일로 저장
file_name = 'output.xlsx'
df.to_excel(file_name, index=False)

# 엑셀 파일 열기 (Windows의 경우)
os.startfile(file_name)

print(f"{file_name} 파일이 성공적으로 생성되고 열렸습니다.")
