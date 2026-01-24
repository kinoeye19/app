[Part 2] 프로젝트 명세서 (Project_Context.md)


# BK21 연구성과 자동화 시스템 명세서

## 1. 프로젝트 폴더 구조
```text
/Users/kino19/app/
├── .gitignore                  # 보안 파일 제외 설정 (완료)
├── service_account.json        # [보안] 구글 시트 API 키
├── mail_auto/
│   ├── client_secret.json      # [보안] 구글 OAuth 인증 키
│   ├── naver_credentials.json  # [보안] 네이버 메일 SMTP 정보 (id, password)
│   ├── token.json              # [보안] 구글 토큰 캐시
│   ├── email_content.md        # 메일 본문 템플릿
│   ├── send_mail_module.py     # 메일 발송 모듈 (SMTP)
│   ├── 00_reset_project.py     # [초기화] 폴더 삭제 + 시트 링크 제거
│   ├── 01_create_personal_sheets.py # [생성] 개인별 시트 생성 (학번매칭, 스마트너비, 재시도)
│   ├── 02_send_check_mail.py   # [발송 A] 성과 확인 요청 (성과 유무 분기 처리)
│   └── 03_send_remind_mail.py  # [발송 B] 미제출자 독촉 메일
```

## 2. 주요 코드 내용

### 1) `.gitignore` (보안 설정)
```gitignore
# 보안 파일 (가장 중요)
service_account.json
secrets.toml

# Gmail/Google API 보안 키 (mail_auto 폴더)
mail_auto/credentials.json
mail_auto/token.json
mail_auto/naver_credentials.json
mail_auto/client_secret.json

# Python/uv 자동 생성 파일
.venv/
__pycache__/
*.pyc
.DS_Store
```

### 2) `00_reset_project.py` (통합 초기화)
```python
import os
import sys
import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(BASE_DIR, 'token.json')
CLIENT_SECRET_FILE = os.path.join(BASE_DIR, 'client_secret.json')

SPREADSHEET_URL = "[https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing](https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing)"
SHEET_NAME = "mail_list"
TARGET_HEADER = "개별시트링크"

TARGET_ROOT_FOLDER_NAME = "05. Temporary"
DELETE_FOLDER_NAME = "[중요] 2025 연구성과 개인별 확인"
SCOPES = ['[https://www.googleapis.com/auth/drive](https://www.googleapis.com/auth/drive)', '[https://www.googleapis.com/auth/spreadsheets](https://www.googleapis.com/auth/spreadsheets)']

def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token: creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token: token.write(creds.to_json())
    return creds

def delete_drive_folder(creds):
    service = build('drive', 'v3', credentials=creds)
    def find_folder_id(folder_name, parent_id=None):
        query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
        if parent_id: query += f" and '{parent_id}' in parents"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])
        return files[0]['id'] if files else None

    root_id = find_folder_id(TARGET_ROOT_FOLDER_NAME)
    if not root_id: return
    target_id = find_folder_id(DELETE_FOLDER_NAME, root_id)
    if target_id: service.files().delete(fileId=target_id).execute()

def clear_sheet_links(creds):
    gc = gspread.authorize(creds)
    doc = gc.open_by_url(SPREADSHEET_URL)
    ws = doc.worksheet(SHEET_NAME)
    headers = ws.row_values(1)
    col_idx = headers.index(TARGET_HEADER) + 1
    col_letter = gspread.utils.rowcol_to_a1(1, col_idx).replace('1', '')
    ws.batch_clear([f"{col_letter}2:{col_letter}{ws.row_count}"])

def main():
    creds = get_credentials()
    delete_drive_folder(creds)
    clear_sheet_links(creds)

if __name__ == "__main__": main()
```

### 3) `01_create_personal_sheets.py` (시트 생성 - 학번 매칭/재시도/스마트너비)
```python
import os, time, pandas as pd, gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from gspread.exceptions import APIError

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(BASE_DIR, 'token.json')
CLIENT_SECRET_FILE = os.path.join(BASE_DIR, 'client_secret.json')
SPREADSHEET_URL = "[https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing](https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing)"
SHEET_STUDENT_LIST, SHEET_PAPER, SHEET_BOOK, SHEET_CONF = "mail_list", "논문", "저서", "학술대회"
TARGET_ROOT_FOLDER_NAME, NEW_FOLDER_NAME = "05. Temporary", "[중요] 2025 연구성과 개인별 확인"
SCOPES = ['[https://www.googleapis.com/auth/drive](https://www.googleapis.com/auth/drive)', '[https://www.googleapis.com/auth/spreadsheets](https://www.googleapis.com/auth/spreadsheets)']

def get_credentials():
    # ... (인증 로직 동일) ...
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token: creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token: token.write(creds.to_json())
    return creds

def set_column_width_safe(worksheet, col_index, width):
    worksheet.spreadsheet.batch_update({"requests": [{"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": col_index, "endIndex": col_index + 1}, "properties": {"pixelSize": width}, "fields": "pixelSize"}}]})

def smart_resize_columns(worksheet, df):
    worksheet.format(f"A1:Z{len(df)+20}", {"wrapStrategy": "WRAP"})
    requests = []
    for i, col in enumerate(df.columns):
        max_len = len(str(col)) * 1.5
        for val in df[col].astype(str).head(50):
            if len(val) > max_len: max_len = len(val)
        pixel_width = min(350, max(50, int(max_len * 12)))
        requests.append({"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1}, "properties": {"pixelSize": pixel_width}, "fields": "pixelSize"}})
    if requests: worksheet.spreadsheet.batch_update({"requests": requests})

def process_student_with_retry(drive_service, gc, target_folder_id, master_doc, row, idx, df_paper, df_book, df_conf):
    name = str(row.get('Name_2', '')).strip()
    student_id = str(row.get('Student_No', '')).strip()
    retry_count = 0
    while retry_count < 10:
        try:
            # 폴더 및 시트 생성 로직
            # ... (find_folder_id, create_folder 등 생략, 실제 파일에는 포함됨) ...
            # 시트 생성 후 데이터 매칭 (학번 기준)
            my_paper = df_paper[df_paper['학번'] == student_id]
            my_book = df_book[df_book['학번'] == student_id]
            my_conf = df_conf[df_conf['학번'] == student_id]
            
            # 탭 작성 및 링크 기록
            # ...
            # 링크 기록 시 col_idx 보정 완료된 코드 사용
            col_idx = pd.DataFrame(master_doc.worksheet(SHEET_STUDENT_LIST).get_all_records()).columns.get_loc('개별시트링크') + 1
            master_doc.worksheet(SHEET_STUDENT_LIST).update_cell(idx + 2, col_idx, new_sh.url)
            return True
        except (APIError, HttpError) as e:
            if '429' in str(e):
                time.sleep(70 * (retry_count + 1))
                retry_count += 1
            else: return False
    return False

def main():
    # ... (데이터 로드 및 루프 실행) ...
    pass
if __name__ == "__main__": main()
```

### 4) `02_send_check_mail.py` (성과 유무 'X' 감지 및 메일 발송)
```python
import os, sys, pandas as pd, gspread
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import send_mail_module

SHEET_NAME_LIST = "check_list"
SHEET_NAME_PAPER = "논문"

def main():
    # ... (인증 및 시트 로드) ...
    
    # '연구성과유무'가 'X'인 학생 학번 추출
    no_result_students = set()
    if '연구성과유무' in df_paper.columns and '학번' in df_paper.columns:
        no_result_students = set(df_paper[df_paper['연구성과유무'] == 'X']['학번'].astype(str).str.strip().tolist())

    for idx, row in df_list.iterrows():
        # ... (기본 정보 매핑) ...
        student_id = str(row.get('Student_No', row.get('학번', ''))).strip()

        if student_id in no_result_students:
            # Case A: 성과 없음 안내 메일 (HTML 본문)
            pass
        else:
            # Case B: 일반 확인 메일 (HTML 본문)
            pass
        
        # send_mail_module.send_email 호출 및 'Sent' 기록
```

### 5) `03_send_remind_mail.py` (미제출자 독촉 메일)
```python
import os, sys, json, markdown, gspread
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import send_mail_module

SHEET_NAME = "remind_list"

def main():
    # ... (인증 로드) ...
    # 마크다운 파일 로드 및 첫 줄(# 제목) 제거 로직 포함
    with open(MD_FILE_PATH, 'r') as f:
        lines = f.readlines()
        md_text = "".join(lines[1:]) if lines[0].strip().startswith('#') else "".join(lines)

    for i, row in enumerate(records):
        # name_2, email 매핑 (소문자 헤더 대응)
        # 메일 발송 및 'Sent' 기록
```

### 6) `send_mail_module.py` (SMTP 모듈)
```python
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email(user, password, to_email, subject, html_content):
    try:
        msg = MIMEMultipart()
        msg['From'] = f"BK21 교육연구단 <{user}>"
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(html_content, 'html'))

        with smtplib.SMTP_SSL('smtp.naver.com', 465) as server:
            server.login(user, password)
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"발송 실패: {e}")
        return False
```
