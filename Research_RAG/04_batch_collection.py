import os
import sys
import time
import re
import pickle
import gspread
import difflib
import urllib.parse
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

def clean_text_for_compare(text):
    # 한글, 영문, 숫자만 남기고 나머지 제거
    text = re.sub(r'[^\w\s]', '', text)
    return text.replace(" ", "").lower()

def calculate_similarity(s1, s2):
    c1 = clean_text_for_compare(s1)
    c2 = clean_text_for_compare(s2)
    if not c1 or not c2: return 0.0
    return difflib.SequenceMatcher(None, c1, c2).ratio()

def get_riss_id_from_url(url):
    try:
        parsed = urllib.parse.urlparse(url)
        qs = urllib.parse.parse_qs(parsed.query)
        if 'control_no' in qs: return qs['control_no'][0]
    except: pass
    return ""

# ==========================================
# 🔍 RISS 검색 로직 (전수 조사 방식)
# ==========================================

def search_riss_direct(driver, user_title, author):
    # 검색 URL
    encoded_query = urllib.parse.quote(user_title)
    search_url = f"https://www.riss.kr/search/Search.do?isDetailSearch=N&searchGubun=true&strQuery={encoded_query}&query={encoded_query}&colName=all"
    
    driver.get(search_url)
    
    # [중요] 페이지 로딩 대기 (3초)
    time.sleep(3) 

    candidates = []
    
    try:
        # [핵심 변경] 화면의 "모든 링크(a tag)"를 싹 다 긁어옵니다.
        # CSS 선택자에 의존하지 않습니다.
        all_links = driver.find_elements(By.TAG_NAME, "a")
        
        # 긁어온 수백 개의 링크 중 '제목'일 것 같은 놈만 골라냅니다.
        for el in all_links:
            try:
                text = el.text.strip()
                link = el.get_attribute("href")
                
                # 1. 텍스트가 너무 짧거나(메뉴바 등) 없으면 패스
                if not text or len(text) < 5: continue
                
                # 2. 링크가 없거나 자바스크립트면 패스 (단, RISS는 상세페이지에 javascript를 쓰지 않음)
                if not link or "javascript" in link: continue
                
                # 3. RISS 상세페이지 URL 특징 확인 (DetailView)
                if "DetailView" not in link: continue

                # 4. 유사도 검사
                score = calculate_similarity(user_title, text)
                
                # 유사도가 일정 수준 이상인 것만 후보 등록
                if score > 0.3:
                     # URL 절대경로 보정
                    if not link.startswith("http"):
                        link = "https://www.riss.kr" + link
                        
                    candidates.append({
                        "link": link,
                        "title": text,
                        "score": score
                    })
            except:
                continue

    except Exception as e:
        pass

    if not candidates:
        return None

    # 점수순 정렬
    candidates.sort(key=lambda x: x['score'], reverse=True)
    best = candidates[0]
    
    # 가장 높은 점수가 40% 이상이면 채택
    if best['score'] >= 0.4: 
        print(f"   🎯 RISS 발견: {best['title'][:15]}... ({int(best['score']*100)}%)")
        return best['link']
    else:
        # 디버깅: 가장 비슷했던 게 뭐였는지 출력
        print(f"   💨 유사도 낮음 (최고: {int(best['score']*100)}% - '{best['title']}')")
        return None

def scrape_riss_details(driver, url):
    data = {"abstract": "", "keywords": "", "id": ""}
    data["id"] = get_riss_id_from_url(url)

    try:
        driver.get(url)
        time.sleep(2)
        
        try:
            buttons = driver.find_elements(By.CSS_SELECTOR, "a.moreView, a.btn_more")
            for btn in buttons:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.5)
        except: pass
        
        full_text = driver.find_element(By.TAG_NAME, "body").text
        
        if "국문초록" in full_text:
            temp = full_text.split("국문초록")[1]
            data["abstract"] = temp.split("목차")[0] if "목차" in temp else temp[:1500]
        elif "Abstract" in full_text:
            temp = full_text.split("Abstract")[1]
            data["abstract"] = temp.split("Table of Contents")[0] if "Table of Contents" in temp else temp[:1500]
        else:
            try: data["abstract"] = driver.find_element(By.CSS_SELECTOR, "div.additionalInfo").text
            except: data["abstract"] = "초록 없음"

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
    print(f"📊 총 {len(rows)}건 작업 시작 (전수 조사 모드)...\n")

    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_window_size(1200, 900)

    consecutive_failures = 0

    for i, row in enumerate(rows):
        row_num = i + 2
        
        title = str(row.get("논문명", "")).strip()
        author = str(row.get("이름", "")).strip()
        existing_link = str(row.get("RISS_링크", ""))

        if title and existing_link and "http" in existing_link:
            continue
        if not title:
            continue

        print(f"[{i+1}/{len(rows)}] 🔍 {title[:20]}... ({author})")
        
        link = search_riss_direct(driver, title, author)
        
        if link:
            consecutive_failures = 0
            details = scrape_riss_details(driver, link)
            print(f"   ✅ 수집: ID({details['id']}) / 주제어({details['keywords'][:10]}...)")
            
            try:
                worksheet.update_cell(row_num, idx_id, details['id'])
                worksheet.update_cell(row_num, idx_link, link)
                worksheet.update_cell(row_num, idx_abs, details['abstract'][:4000])
                worksheet.update_cell(row_num, idx_kw, details['keywords'])
            except Exception as e:
                print(f"   ❌ 저장 실패: {e}")
        else:
            consecutive_failures += 1
            print(f"   ⚠️ 검색 실패 (연속 {consecutive_failures}회)")
            if not existing_link:
                try: worksheet.update_cell(row_num, idx_link, "검색실패")
                except: pass
        
        if consecutive_failures >= 3: # 3회로 완화
            print("\n" + "="*50)
            print("🚨 [중단] 연속 3회 실패. RISS 접근이 차단되었거나 페이지 구조가 완전히 다릅니다.")
            print("="*50)
            break
        
        time.sleep(2)

    driver.quit()

if __name__ == "__main__":
    main()