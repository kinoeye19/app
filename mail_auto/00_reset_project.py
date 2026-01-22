import os
import sys
import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# --- [설정 영역] ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(BASE_DIR, 'token.json')
CLIENT_SECRET_FILE = os.path.join(BASE_DIR, 'client_secret.json')

# 구글 시트 및 드라이브 설정
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing"
SHEET_NAME = "mail_list"  # 링크를 지울 명단 시트
TARGET_HEADER = "개별시트링크"  # 지울 컬럼명

# 삭제할 드라이브 폴더명 (상위 -> 하위)
TARGET_ROOT_FOLDER_NAME = "05. Temporary"
DELETE_FOLDER_NAME = "[중요] 2025 연구성과 개인별 확인"

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

# --- [기능 1: 폴더 삭제] ---
def delete_drive_folder(creds):
    print("\n🗑️  [1단계] 드라이브 폴더 삭제 중...")
    service = build('drive', 'v3', credentials=creds)

    def find_folder_id(folder_name, parent_id=None):
        query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
        if parent_id:
            query += f" and '{parent_id}' in parents"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])
        return files[0]['id'] if files else None

    # 상위 폴더 찾기
    root_id = find_folder_id(TARGET_ROOT_FOLDER_NAME)
    if not root_id:
        print(f"   ❌ '{TARGET_ROOT_FOLDER_NAME}' 폴더를 찾을 수 없습니다.")
        return

    # 삭제 대상 폴더 찾기
    target_id = find_folder_id(DELETE_FOLDER_NAME, root_id)
    
    if target_id:
        try:
            service.files().delete(fileId=target_id).execute()
            print(f"   🔥 폴더 삭제 완료: [{DELETE_FOLDER_NAME}]")
        except Exception as e:
            print(f"   ❌ 삭제 실패: {e}")
    else:
        print(f"   ✅ 삭제할 폴더가 없습니다. (이미 삭제됨)")

# --- [기능 2: 시트 링크 초기화] ---
def clear_sheet_links(creds):
    print("\n🧹 [2단계] 시트 링크 데이터 초기화 중...")
    gc = gspread.authorize(creds)
    
    try:
        doc = gc.open_by_url(SPREADSHEET_URL)
        ws = doc.worksheet(SHEET_NAME)
    except Exception as e:
        print(f"   ❌ 시트 접속 실패: {e}")
        return

    # 헤더 위치 찾기
    headers = ws.row_values(1)
    try:
        col_idx = headers.index(TARGET_HEADER) + 1
    except ValueError:
        print(f"   ❌ '{TARGET_HEADER}' 헤더를 찾을 수 없습니다.")
        return

    # 데이터 지우기 (2행부터 끝까지)
    row_count = ws.row_count
    col_letter = gspread.utils.rowcol_to_a1(1, col_idx).replace('1', '')
    range_to_clear = f"{col_letter}2:{col_letter}{row_count}"
    
    try:
        ws.batch_clear([range_to_clear])
        print(f"   ✨ 링크 데이터 삭제 완료 (범위: {range_to_clear})")
    except Exception as e:
        print(f"   ❌ 초기화 실패: {e}")

# --- [메인 실행] ---
def main():
    print("🚀 [프로젝트 초기화] 작업을 시작합니다.")
    print("   이 작업은 생성된 폴더를 삭제하고, 시트의 링크 정보를 지웁니다.")
    
    creds = get_credentials()
    
    # 1. 드라이브 폴더 삭제
    delete_drive_folder(creds)
    
    # 2. 시트 링크 지우기
    clear_sheet_links(creds)
    
    print("\n🎉 프로젝트 초기화 완료! 이제 '01_create_sheets.py'를 실행할 준비가 되었습니다.")

if __name__ == "__main__":
    main()