import os
import sys
import time
import pandas as pd
import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from gspread.exceptions import APIError

# --- [설정 영역] ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CLIENT_SECRET_FILE = os.path.join(BASE_DIR, 'client_secret.json')
TOKEN_FILE = os.path.join(BASE_DIR, 'token.json')

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing"

SHEET_STUDENT_LIST = "mail_list"
SHEET_PAPER = "논문"
SHEET_BOOK = "저서"
SHEET_CONF = "학술대회"

SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

# --- [인증 함수] ---
def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds

# --- [스마트 너비 조절 - 작동 확인된 방식] ---
def smart_resize_columns(worksheet, df):
    """
    01_create_personal_sheets.py에서 검증된 방식 그대로 적용
    - WRAP 모드로 줄바꿈 허용
    - 문자 수 × 12px로 너비 계산
    - 최소 50px, 최대 350px 제한
    """
    if df.empty:
        return
    
    row_count = len(df) + 1
    
    # 1. 줄바꿈 허용 설정 (WRAP)
    worksheet.format(f"A1:Z{row_count+20}", {"wrapStrategy": "WRAP"})

    MAX_WIDTH = 350
    MIN_WIDTH = 50
    requests = []
    
    for i, col in enumerate(df.columns):
        # 2. 헤더 길이 × 1.5로 시작
        max_len = len(str(col)) * 1.5 
        
        # 3. 데이터 중 최대 길이 찾기 (상위 50행만)
        column_data = df[col].astype(str).head(50)
        for val in column_data:
            length = len(val)
            if length > max_len:
                max_len = length
        
        # 4. 문자 수 × 12 = 픽셀 너비
        pixel_width = int(max_len * 12) 
        
        # 5. 최소/최대 제한
        if pixel_width > MAX_WIDTH: 
            pixel_width = MAX_WIDTH
        elif pixel_width < MIN_WIDTH: 
            pixel_width = MIN_WIDTH
            
        # 6. 너비 설정 요청 추가
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": i,
                    "endIndex": i + 1
                },
                "properties": {"pixelSize": pixel_width},
                "fields": "pixelSize"
            }
        })
    
    # 7. 헤더 스타일 (진하게, 배경색)
    header_request = {
        "repeatCell": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 0,
                "endRowIndex": 1
            },
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {"bold": True},
                    "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE"
                }
            },
            "fields": "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment)"
        }
    }
    requests.append(header_request)
    
    # 8. 모든 요청 한 번에 실행
    if requests:
        try:
            worksheet.spreadsheet.batch_update({"requests": requests})
        except Exception as e:
            print(f"      (서식 적용 중 경고: {e})")


# --- [탭 업데이트 함수] ---
def update_tab_safe(sheet_obj, title, df_data):
    max_retries = 3
    retry_delay = 30 

    for attempt in range(max_retries):
        try:
            # 1. 시트 초기화 및 데이터 쓰기
            try:
                ws = sheet_obj.worksheet(title)
                ws.clear() 
            except gspread.WorksheetNotFound:
                ws = sheet_obj.add_worksheet(title=title, rows=100, cols=20)
            
            if df_data.empty:
                data = [df_data.columns.tolist()]
            else:
                data = [df_data.columns.tolist()] + df_data.values.tolist()
            
            # 데이터 입력
            ws.update(range_name='A1', values=data)
            
            # 2. 스마트 너비 조절 적용
            if not df_data.empty:
                smart_resize_columns(ws, df_data)
            
            return True

        except APIError as e:
            if '429' in str(e):
                print(f"\n      ⚠️ [API 과부하] {retry_delay}초 대기 후 재시도... ({attempt+1}/{max_retries})")
                time.sleep(retry_delay)
                retry_delay *= 2 
            else:
                raise e 
        except Exception as e:
            raise e 

    raise Exception(f"API 한도 초과로 '{title}' 탭 업데이트 실패")

# --- [학생 1명 전체 처리] ---
def process_student(gc, target_url, df_paper, df_book, df_conf, student_id):
    try:
        sh = gc.open_by_url(target_url)
    except Exception as e:
        print(f"      ❌ 시트 접속 불가: {e}")
        return False

    my_paper = df_paper[df_paper['학번'] == student_id]
    my_book = df_book[df_book['학번'] == student_id]
    my_conf = df_conf[df_conf['학번'] == student_id]

    try:
        update_tab_safe(sh, "논문", my_paper)
        time.sleep(2)  # 탭 간 딜레이
        update_tab_safe(sh, "저서", my_book)
        time.sleep(2) 
        update_tab_safe(sh, "학술대회", my_conf)
        
        try:
            intro = sh.sheet1
            now_str = time.strftime("%Y-%m-%d %H:%M:%S")
            intro.update_cell(6, 1, f"✅ 업데이트 완료: {now_str}")
        except: 
            pass
        
        return True
    except Exception as e:
        print(f"      ❌ 처리 중단: {e}")
        return False

def main():
    print("🚀 [스마트 너비 조절] 데이터 입력 + WRAP 모드 + 자동 너비 계산")
    
    creds = get_credentials()
    gc = gspread.authorize(creds)
    print("✅ 인증 완료")

    try:
        master_doc = gc.open_by_url(SPREADSHEET_URL)
        df_list = pd.DataFrame(master_doc.worksheet(SHEET_STUDENT_LIST).get_all_records())
        df_paper = pd.DataFrame(master_doc.worksheet(SHEET_PAPER).get_all_records())
        df_book = pd.DataFrame(master_doc.worksheet(SHEET_BOOK).get_all_records())
        df_conf = pd.DataFrame(master_doc.worksheet(SHEET_CONF).get_all_records())
        
        for df in [df_list, df_paper, df_book, df_conf]:
            df.columns = [c.strip() for c in df.columns]
            if '학번' in df.columns: 
                df['학번'] = df['학번'].astype(str).str.strip()
            if 'Student_No' in df.columns: 
                df['Student_No'] = df['Student_No'].astype(str).str.strip()
            
    except Exception as e:
        print(f"❌ 데이터 로드 실패: {e}")
        return

    total_target = sum(1 for _, r in df_list.iterrows() if str(r.get('개별시트링크', '')).startswith('http'))
    update_count = 0
    
    print(f"📋 총 {total_target}명의 시트를 최신화합니다.\n")

    for idx, row in df_list.iterrows():
        name = str(row.get('Name_2', '')).strip()
        student_id = str(row.get('Student_No', '')).strip()
        link = str(row.get('개별시트링크', '')).strip()

        if not link.startswith('http'):
            continue

        print(f"🔄 [{update_count + 1}/{total_target}] {name} ...", end=" ", flush=True)
        
        if process_student(gc, link, df_paper, df_book, df_conf, student_id):
            print("성공 ✅")
            update_count += 1
            time.sleep(1.5) 
        else:
            print("실패 ❌")
            time.sleep(5)

    print(f"\n🎉 모든 작업이 완료되었습니다. (성공: {update_count}/{total_target})")

if __name__ == "__main__":
    main()