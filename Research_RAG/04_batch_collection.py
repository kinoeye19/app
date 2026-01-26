import os
import sys
import time
import json
import re
import requests
import pickle
import gspread
from urllib.parse import urlparse, parse_qs
from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 환경변수 로드
load_dotenv()

# ==========================================
# ⚙️ 설정 및 경로
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)

CLIENT_SECRET_PATH = os.path.join(PROJECT_ROOT, "mail_auto", "client_secret.json")
TOKEN_PATH = os.path.join(PROJECT_ROOT, "mail_auto", "token.json")

if not os.path.exists(CLIENT_SECRET_PATH):
    CLIENT_SECRET_PATH = os.path.join(PROJECT_ROOT, "client_secret.json")
    TOKEN_PATH = os.path.join(PROJECT_ROOT, "token.json")

SHEET_ID = os.getenv("TARGET_SHEET_ID") 
SERPER_API_KEY = os.getenv("SERPER_API_KEY")

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

# ==========================================
# 🛠️ 유틸리티 함수
# ==========================================

def get_gspread_client():
    creds = None
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, 'rb') as token:
            try: creds = pickle.load(token)
            except: pass
            
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CLIENT_SECRET_PATH):
                print(f"❌ 인증 파일 없음: {CLIENT_SECRET_PATH}")
                return None
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, 'wb') as token:
            pickle.dump(creds, token)

    return gspread.authorize(creds)

def clean_brackets(text):
    """괄호()와 대괄호[] 및 그 안의 내용을 제거"""
    # 1. 괄호 내용 제거
    cleaned = re.sub(r'\([^)]*\)', '', text)
    cleaned = re.sub(r'\[[^\]]*\]', '', cleaned)
    # 2. 다중 공백 제거
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    # 만약 괄호 제거 후 내용이 너무 짧아지면(예: 전체가 괄호였음) 원본 반환
    if len(cleaned) < 2:
        return text
    return cleaned

def extract_main_title(title):
    """부제 구분자(:, -, = 등) 앞쪽만 추출"""
    main_title = re.split(r'[:\-\=]', title)[0]
    return main_title.strip()

def get_riss_id_from_url(url):
    try:
        parsed = urlparse(url)
        qs = parse_qs(parsed.query)
        if 'control_no' in qs:
            return qs['control_no'][0]
    except:
        pass
    return ""

# ==========================================
# 🔍 4단계 스마트 검색 로직
# ==========================================

def search_riss_link_smart(title, author):
    url = "https://google.serper.dev/search"
    headers = {'X-API-KEY': SERPER_API_KEY, 'Content-Type': 'application/json'}

    # 검색 후보군 생성
    candidates = []
    
    # 1단계: 원본 엄격 검색 ("제목")
    candidates.append({"q": f'site:riss.kr "{title}" {author}', "type": "1.엄격(원본)"})
    
    # 2단계: 원본 유연 검색 (제목 - 따옴표 제거) -> 특수문자/띄어쓰기 무시
    candidates.append({"q": f'site:riss.kr {title} {author}', "type": "2.유연(원본)"})
    
    # 3단계: 괄호 청소 검색 (한자 병기 제거)
    cleaned_title = clean_brackets(title)
    if cleaned_title != title:
        candidates.append({"q": f'site:riss.kr {cleaned_title} {author}', "type": "3.유연(괄호제거)"})
    
    # 4단계: 메인 제목 검색 (부제 제거)
    main_title = extract_main_title(title)
    # 메인 제목이 원본/청소본과 다르고, 2글자 이상일 때만
    if main_title != title and main_title != cleaned_title and len(main_title) >= 2:
        candidates.append({"q": f'site:riss.kr {main_title} {author}', "type": "4.유연(부제제거)"})

    # 순차 실행
    for item in candidates:
        query = item['q']
        q_type = item['type']
        
        # 쿼리 길이 제한 (Serper 오류 방지)
        if len(query) > 300: query = query[:300]

        print(f"   🔎 시도 [{q_type}]: {query.replace('site:riss.kr', '').strip()[:40]}...")
        
        try:
            payload = json.dumps({"q": query, "num": 3, "gl": "kr", "hl": "ko"})
            resp = requests.post(url, headers=headers, data=payload).json()
            
            for res in resp.get("organic", []):
                link = res.get("link", "")
                if "riss.kr" in link and "DetailView" in link:
                    print(f"   ✨ 발견 성공! ({q_type})")
                    return link
            time.sleep(0.5) # API 속도 조절
        except Exception as e:
            print(f"   ⚠️ 검색 API 에러: {e}")

    return None

def scrape_riss_details(driver, url):
    data = {"abstract": "", "keywords": "", "id": ""}
    data["id"] = get_riss_id_from_url(url)

    try:
        driver.get(url)
        wait = WebDriverWait(driver, 5)
        # 본문 로딩 대기
        try: wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.wrapper")))
        except: pass

        # '더보기' 버튼들 클릭
        try:
            buttons = driver.find_elements(By.CSS_SELECTOR, "a.moreView, a.btn_more")
            for btn in buttons:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.1)
        except: pass
        
        full_text = driver.find_element(By.TAG_NAME, "body").text
        
        # 초록 추출
        if "국문초록" in full_text:
            temp = full_text.split("국문초록")[1]
            data["abstract"] = temp.split("목차")[0] if "목차" in temp else temp[:1500]
        elif "Abstract" in full_text:
            temp = full_text.split("Abstract")[1]
            data["abstract"] = temp.split("Table of Contents")[0] if "Table of Contents" in temp else temp[:1500]
        else:
            try: data["abstract"] = driver.find_element(By.CSS_SELECTOR, "div.additionalInfo").text
            except: data["abstract"] = "초록 없음"

        # 주제어 추출
        try:
            lines = full_text.split('\n')
            for line in lines:
                if "주제어" in line and len(line) < 300:
                    data["keywords"] = line.replace("주제어", "").strip()
                    break
                if "Keywords" in line and len(line) < 300:
                    data["keywords"] = line.replace("Keywords", "").strip()
                    break
        except: pass

        data["abstract"] = data["abstract"].strip()
        data["keywords"] = data["keywords"].strip()

    except Exception as e:
        print(f"   ⚠️ 스크래핑 오류: {e}")

    return data

# ==========================================
# 🚀 메인 실행
# ==========================================
def main():
    if not SHEET_ID:
        print("❌ 오류: .env 파일에 'TARGET_SHEET_ID'가 없습니다.")
        return

    client = get_gspread_client()
    if not client: return

    try:
        doc = client.open_by_key(SHEET_ID)
        worksheet = doc.worksheet("논문")
        print(f"✅ 타겟 시트 연결: {doc.title}")
    except Exception as e:
        print(f"❌ 시트 연결 실패: {e}")
        return

    # 헤더 설정
    headers = worksheet.row_values(1)
    new_cols = ["논문ID", "RISS_링크", "초록", "주제어"]
    for col_name in new_cols:
        if col_name not in headers:
            worksheet.update_cell(1, len(headers) + 1, col_name)
            headers.append(col_name)

    idx_id = headers.index("논문ID") + 1
    idx_link = headers.index("RISS_링크") + 1
    idx_abs = headers.index("초록") + 1
    idx_kw = headers.index("주제어") + 1

    rows = worksheet.get_all_records()
    print(f"📊 총 {len(rows)}건 작업 시작...\n")

    # 브라우저 옵션
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    for i, row in enumerate(rows):
        row_num = i + 2
        
        title = str(row.get("논문명", "")).strip()
        author = str(row.get("이름", "")).strip()
        existing_link = str(row.get("RISS_링크", ""))

        # 이미 링크가 있으면 건너뜀 (단, '검색실패'라고 적힌 건 다시 시도)
        if title and existing_link and "http" in existing_link:
            continue
        if not title:
            continue

        print(f"[{i+1}/{len(rows)}] 🔍 {title} ({author})")
        
        # 스마트 검색 실행
        link = search_riss_link_smart(title, author)
        
        if link:
            details = scrape_riss_details(driver, link)
            print(f"   ✅ 수집 완료: ID({details['id']}) / 키워드({details['keywords'][:10]}...)")
            
            try:
                worksheet.update_cell(row_num, idx_id, details['id'])
                worksheet.update_cell(row_num, idx_link, link)
                worksheet.update_cell(row_num, idx_abs, details['abstract'][:4000])
                worksheet.update_cell(row_num, idx_kw, details['keywords'])
            except Exception as e:
                print(f"   ❌ 저장 실패: {e}")
        else:
            print("   ⚠️ 모든 검색 시도 실패")
            # 확실히 실패했을 때만 기록
            if not existing_link:
                try: worksheet.update_cell(row_num, idx_link, "검색실패")
                except: pass
            
        time.sleep(2) # 차단 방지

    driver.quit()
    print("\n🎉 모든 작업 완료!")

if __name__ == "__main__":
    main()