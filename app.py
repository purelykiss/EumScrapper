import tkinter as tk
from tkinter import filedialog
import requests
from bs4 import BeautifulSoup
import pdfkit
import os
import json
import re
import pandas as pd
from datetime import datetime

# 기본 경로 설정
current_directory = os.getcwd()
option_file_path = os.path.join(current_directory, 'option.txt')

# 옵션 파일 읽기 및 초기화
def read_option_file():
    global path_to_wkhtmltopdf, rest_api_key, save_path

    # 기본 값 설정
    path_to_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
    rest_api_key = ''
    save_path = r'.\result'

    try:
        with open(option_file_path, 'r') as file:
            lines = file.readlines()
            for line in lines:
                if line.startswith('wkhtmltopdf_path='):
                    path_to_wkhtmltopdf = line.split('=')[1].strip()
                elif line.startswith('kakao_rest_api_key='):
                    rest_api_key = line.split('=')[1].strip()
                elif line.startswith('save_pdf_path='):
                    save_path = line.split('=')[1].strip()

        # 경로가 유효한지 확인
        if not os.path.isfile(path_to_wkhtmltopdf) or not os.path.isdir(os.path.dirname(save_path)):
            raise ValueError("파일 경로나 디렉토리가 유효하지 않습니다.")

    except (FileNotFoundError, ValueError):
        # 기본 옵션 파일 생성
        create_default_option_file()

# 기본 옵션 파일 생성
def create_default_option_file():
    default_content = [
        r"wkhtmltopdf_path=.\wkhtmltopdf\bin\wkhtmltopdf.exe",
        r"kakao_rest_api_key=",
        r"save_pdf_path=.\result"
    ]

    with open(option_file_path, "w") as file:
        for line in default_content:
            file.write(line + "\n")

# wkhtmltopdf 경로 설정
read_option_file()
config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)

# URL 템플릿 (pnu 값만 바뀌는 형태)
BASE_URL = "https://www.eum.go.kr/web/ar/lu/luLandDetPrintPop.jsp?isNoScr=script&s_type=1&p_type=select&p_type1=true&p_type2=true&p_type3=true&p_type4=true&p_type5=false&mode=search&pnu={}"

# GET 요청을 보내는 함수 (PNU 기반)
def get_request(pnu, save_directory):
    url = BASE_URL.format(pnu)  # PNU 값을 URL에 삽입
    try:
        response = requests.get(url)  # GET 요청
        check_location(response.text, url, save_directory)  # 소재지 체크 함수 호출
    except requests.exceptions.RequestException as e:
        raise Exception(f"GET 요청 오류: {str(e)}")

# 소재지를 확인하고 PDF로 저장하는 함수
def check_location(html, url, save_directory):
    soup = BeautifulSoup(html, 'html.parser')

    # <th scope="row">소재지</th> 태그 찾기
    location_header = soup.find('th', scope="row", string="소재지")
    if location_header:
        # 다음 <td> 태그 찾기
        location_td = location_header.find_next('td')
        if location_td and location_td.text.strip():  # <td>에 텍스트가 있는지 확인
            location = location_td.text.strip()  # 공백 제거
            save_as_pdf(location, url, save_directory)  # PDF 저장 함수 호출
        else:
            raise Exception("소재지 값이 없음, 저장하지 않음.")  # 오류 발생 시 raise
    else:
        raise Exception("소재지 항목이 없음, 저장하지 않음.")  # 오류 발생 시 raise

# PDF로 저장하는 함수
def save_as_pdf(location, url, save_directory):
    # 공백을 _로 변경한 파일 이름
    filename = location.replace(' ', '_') + ".pdf"

    if save_directory:  # 사용자가 경로를 선택했을 때
        file_path = os.path.join(save_directory, filename)  # 파일 전체 경로

        try:
            # pdfkit을 이용해 HTML 페이지를 PDF로 변환, configuration 인자 추가
            pdfkit.from_url(url, file_path, configuration=config)
        except Exception as e:
            raise Exception(f"PDF 저장 실패: {str(e)}")
    else:
        raise Exception("저장 경로가 선택되지 않음.")

# XLS 파일에서 주소와 지번을 읽어오는 함수
def read_addresses_from_xls():
    # XLS 파일 경로를 선택
    xls_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx")])

    if xls_file_path:  # 파일이 선택된 경우
        try:
            # 저장 경로 선택 (XLS 파일마다 경로 선택)
            save_directory = filedialog.askdirectory(title="저장 경로 선택")

            if not save_directory:
                status_label.config(text="저장 경로가 선택되지 않았습니다.", fg="red")
                return

            # XLS 파일 읽기
            df = pd.read_excel(xls_file_path)

            # 원래 행 번호 열 추가
            df['원래 행 번호'] = df.index + 1  # 1부터 시작하는 인덱스

            # 필요한 열이 비어있는 경우 해당 행 삭제
            df = df.dropna(subset=['읍면동', '지번'], how='any')  # '읍면동' 또는 '지번'이 비어 있는 행 삭제

            # 행 수가 0인 경우 경고 메시지
            if df.empty:
                status_label.config(text="모든 행이 삭제되었습니다. 유효한 데이터를 제공해주세요.", fg="red")
                return

            total_rows = len(df)
            success_count = 0
            error_count = 0
            errors = []  # 오류 정보를 저장할 리스트

            # 현재 시각을 기반으로 오류 로그 파일명 생성
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            log_filename = f"eum_error_log_{timestamp}.txt"
            log_file_path = os.path.join(save_directory, log_filename)

            # 주소와 지번 정보 읽기
            for index, row in df.iterrows():
                addr = row['읍면동']
                detail_addr = row['지번']
                original_row_number = row['원래 행 번호']

                try:
                    # 각 주소와 지번에 대해 PDF 요청
                    if validate_detail_address(detail_addr):  # 지번 유효성 검사
                        b_code = kakao_request(addr, detail_addr)  # b_code 요청
                        if b_code:
                            pnu_code = create_pnu_code(b_code, detail_addr)  # PNU 코드 생성
                            get_request(pnu_code, save_directory)  # PDF 요청
                            success_count += 1  # 성공 카운트 증가
                except Exception as e:
                    error_count += 1  # 오류 카운트 증가
                    errors.append(f"원래 행 번호: {original_row_number}, 메서드: {e.__class__.__name__}, 메시지: {str(e)}")

            # 오류가 발생한 경우 로그 파일에 기록
            if errors:
                with open(log_file_path, 'w') as log_file:
                    for error in errors:
                        log_file.write(error + '\n')

            # 성공 및 실패 결과 출력
            status_label.config(text=f"{total_rows}개 시도 중 {success_count}개 성공, {error_count}개 실패", fg="green" if error_count == 0 else "red")

        except Exception as e:
            status_label.config(text=f"파일 처리 오류: {str(e)}", fg="red")  # 오류 메시지 표시


# 카카오 API를 통해 b_code 값 요청
def kakao_request(addr, detail_addr):
    global rest_api_key  # API 키 가져오기
    # 카카오 API URL 및 헤더
    kakao_url = "https://dapi.kakao.com/v2/local/search/address"
    headers = {"Authorization": f"KakaoAK {rest_api_key}"}

    # POST 데이터
    query = f"{addr} {detail_addr}"
    data = {"query": query}

    try:
        response = requests.post(kakao_url, headers=headers, data=data)
        response_json = response.json()

        # b_code 값 추출 및 주소 비교
        matching_addresses = []  # 일치하는 주소를 저장할 리스트

        for document in response_json['documents']:
            address_name = document['address']['address_name']
            b_code = document['address']['b_code']
            if b_code:  # b_code가 존재하는 경우
                matching_addresses.append(b_code)  # 일치하는 b_code 추가

        # 유일한 b_code 선택
        if len(matching_addresses) == 1:
            return matching_addresses[0]
        else:
            return None  # b_code가 없거나 모호한 경우

    except requests.exceptions.RequestException as e:
        status_label.config(text=f"API 요청 오류: {str(e)}", fg="red")  # 오류 메시지 표시
        return None


# PNU 코드 생성 함수
def create_pnu_code(b_code, detail_addr):
    # b_code에서 행정구역 코드 부분 추출
    admin_code = b_code[:10]  # 행정구역 코드

    # land_type 결정: detail_addr의 첫 글자가 "산"인지 여부
    land_type = '2' if detail_addr.startswith('산') else '1'  # "산"이면 2, 아니면 1

    if(detail_addr.startswith('산')):
        temp_detail_addr = detail_addr[1:]
    else:
        temp_detail_addr = detail_addr

    main_number = temp_detail_addr.split('-')[0]  # 본번
    sub_number = temp_detail_addr.split('-')[1] if '-' in detail_addr else '0'  # 부번 (없으면 0으로 설정)

    # PNU 코드 구성 (패딩 추가)
    pnu_code = f"{admin_code}{land_type}{main_number.zfill(4)}{sub_number.zfill(4)}"
    return pnu_code


# 유효한 지번인지 확인하는 함수
def validate_detail_address(detail_addr):
    # 지번의 형식이 유효한지 확인 (정규 표현식 예: 숫자-숫자)
    pattern = r'^\d{1,5}(-\d{1,5})?$'
    return bool(re.match(pattern, detail_addr))


# Tkinter GUI 설정
root = tk.Tk()
root.title("토지이음 PDF 변환기")
root.geometry("600x400")

# 파일 선택 버튼
select_file_button = tk.Button(root, text="XLS 파일 선택", command=read_addresses_from_xls)
select_file_button.pack(pady=20)

# 상태 라벨
status_label = tk.Label(root, text="")
status_label.pack(pady=10)

# PDF 오류 라벨
pdf_error_label = tk.Label(root, text="")
pdf_error_label.pack(pady=10)

root.mainloop()
