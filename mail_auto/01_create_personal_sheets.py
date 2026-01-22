import os
import sys
import time
import pandas as pd
import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
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

TARGET_ROOT_FOLDER_NAME = "05. Temporary"
NEW_FOLDER_NAME = "[중요] 2025 연구성과 개인별 확인"

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

# --- [드라이브 함수] ---
def find_folder_id(service, folder_name, parent_id=None):
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    return files[0]['id'] if files else None

def create_folder(service, folder_name, parent_id):
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    file = service.files().create(body=file_metadata, fields='id').execute()
    return file.get('id')

def make_folder_public(service, folder_id):
    permission = {'type': 'anyone', 'role': 'reader'}
    try:
        service.permissions().create(fileId=folder_id, body=permission, fields='id').execute()
        return True
    except HttpError as e:
        print(f"   ⚠️ 권한 설정 실패: {e}")
        return False

def move_file_to_folder(service, file_id, folder_id):
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    service.files().update(fileId=file_id, addParents=folder_id, removeParents=previous_parents).execute()

# --- [안전한 너비 조절 함수] ---
def set_column_width_safe(worksheet, col_index, width):
    body = {
        "requests": [{
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": col_index,
                    "endIndex": col_index + 1
                },
                "properties": {"pixelSize": width},
                "fields": "pixelSize"
            }
        }]
    }
    worksheet.spreadsheet.batch_update(body)

# --- [스마트 너비 조절] ---
def smart_resize_columns(worksheet, df):
    row_count = len(df) + 1
    worksheet.format(f"A1:Z{row_count+20}", {"wrapStrategy": "WRAP"})

    MAX_WIDTH = 350
    MIN_WIDTH = 50
    requests = []
    
    for i, col in enumerate(df.columns):
        max_len = len(str(col)) * 1.5 
        column_data = df[col].astype(str).head(50)
        for val in column_data:
            length = len(val)
            if length > max_len:
                max_len = length
        
        pixel_width = int(max_len * 12) 
        if pixel_width > MAX_WIDTH: pixel_width = MAX_WIDTH
        elif pixel_width < MIN_WIDTH: pixel_width = MIN_WIDTH
            
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
    
    if requests:
        worksheet.spreadsheet.batch_update({"requests": requests})

# --- [핵심: 작업 재시도 처리 함수] ---
def process_student_with_retry(drive_service, gc, target_folder_id, master_doc, row, idx, df_paper, df_book, df_conf):
    name = str(row.get('Name_2', '')).strip()
    student_id = str(row.get('Student_No', '')).strip()
    
    retry_count = 0
    max_retries = 10 
    
    while retry_count < max_retries:
        try:
            # 1. 폴더 생성 (이름은 보기 좋게 Name_2 사용)
            folder_name = f"{name}_{student_id}"
            student_folder_id = find_folder_id(drive_service, folder_name, target_folder_id)
            if not student_folder_id:
                student_folder_id = create_folder(drive_service, folder_name, target_folder_id)
            
            make_folder_public(drive_service, student_folder_id)

            # 2. 시트 생성
            sheet_title = f"[성과확인] {name}_{student_id}"
            new_sh = gc.create(sheet_title)
            move_file_to_folder(drive_service, new_sh.id, student_folder_id)

            # 3. 데이터 기입 함수
            def write_tab(sh, title, df_data):
                if df_data.empty:
                    data = [df_data.columns.tolist()]
                else:
                    data = [df_data.columns.tolist()] + df_data.values.tolist()
                
                # 워크시트 추가
                ws = sh.add_worksheet(title=title, rows=len(data)+20, cols=len(data[0])+5)
                # 에러 발생 시 즉시 멈추도록 try-except 제거
                ws.update(range_name='A1', values=data)
                
                if not df_data.empty:
                    smart_resize_columns(ws, df_data)
                    
                ws.format("A1:Z1", {
                    "textFormat": {"bold": True},
                    "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}
                })

            # -----------------------------------------------------------------
            # [변경됨] 데이터 매칭 로직: 이름 무시, 오직 '학번'으로만 필터링
            # -----------------------------------------------------------------
            my_paper = df_paper[df_paper['학번'] == student_id]
            my_book = df_book[df_book['학번'] == student_id]
            my_conf = df_conf[df_conf['학번'] == student_id]
            # -----------------------------------------------------------------

            # 안내 시트
            intro = new_sh.sheet1
            intro.update_title("안내")
            intro.update(range_name='A1', values=[[f"안녕하세요 {name}님,"],
                                ["이 시트는 본인이 앱을 통해 입력한 연구성과를 확인하는 페이지입니다."],
                                ["각 탭(논문, 저서, 학술대회)을 눌러 입력 내용을 확인해 주세요."],
                                ["⚠️ 내용이 길어 잘린 부분은 자동으로 줄바꿈 되어 표시됩니다."],
                                ["수정 요청은 회신 메일로 주시면 반영하겠습니다."]])
            set_column_width_safe(intro, 0, 500)

            # 탭 작성 (실패 시 여기서 에러 발생 -> catch 블록으로 이동)
            write_tab(new_sh, "논문", my_paper)
            write_tab(new_sh, "저서", my_book)
            write_tab(new_sh, "학술대회", my_conf)

            # 4. 링크 기록
            try:
                col_idx = pd.DataFrame(master_doc.worksheet(SHEET_STUDENT_LIST).get_all_records()).columns.get_loc('개별시트링크') + 1
                master_doc.worksheet(SHEET_STUDENT_LIST).update_cell(idx + 2, col_idx, new_sh.url)
            except: pass

            return True # 성공

        except (APIError, HttpError) as e:
            # 에러 감지 (429 Quota Exceeded 등)
            is_quota_error = False
            if isinstance(e, APIError) and '429' in str(e): is_quota_error = True
            if isinstance(e, HttpError) and e.resp.status == 429: is_quota_error = True

            if is_quota_error:
                wait_time = 70 * (retry_count + 1)
                print(f"\n   ⏳ 과부하 감지 (Quota Exceeded)! {wait_time}초 동안 대기 후 재시도합니다... ({retry_count+1}/{max_retries})")
                time.sleep(wait_time)
                retry_count += 1
            else:
                print(f"   ❌ 치명적 오류 발생: {e}")
                return False
        except Exception as e:
             print(f"   ❌ 알 수 없는 오류: {e}")
             return False
    
    return False

def main():
    print("🚀 [전체 학생] 폴더 및 시트 생성 (학번 매칭 + 재시도 버전) 시작...")

    creds = get_credentials()
    gc = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)
    print("✅ 인증 완료")

    print("📊 데이터 로드 중...", end=" ")
    try:
        master_doc = gc.open_by_url(SPREADSHEET_URL)
        df_list = pd.DataFrame(master_doc.worksheet(SHEET_STUDENT_LIST).get_all_records())
        df_paper = pd.DataFrame(master_doc.worksheet(SHEET_PAPER).get_all_records())
        df_book = pd.DataFrame(master_doc.worksheet(SHEET_BOOK).get_all_records())
        df_conf = pd.DataFrame(master_doc.worksheet(SHEET_CONF).get_all_records())
        
        for df in [df_list, df_paper, df_book, df_conf]:
            df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        print(f"\n❌ 시트 로드 실패: {e}")
        return
    
    # 학번을 문자열로 통일 (매칭 정확도 향상)
    for df in [df_list, df_paper, df_book, df_conf]:
        if '학번' in df.columns: df['학번'] = df['학번'].astype(str).str.strip()
        if 'Student_No' in df.columns: df['Student_No'] = df['Student_No'].astype(str).str.strip()
        if '이름' in df.columns: df['이름'] = df['이름'].astype(str)
    print("완료!")

    # 폴더 준비
    root_id = find_folder_id(drive_service, TARGET_ROOT_FOLDER_NAME)
    target_folder_id = find_folder_id(drive_service, NEW_FOLDER_NAME, root_id)
    if not target_folder_id:
        target_folder_id = create_folder(drive_service, NEW_FOLDER_NAME, root_id)

    created_count = 0
    
    for idx, row in df_list.iterrows():
        name = str(row.get('Name_2', '')).strip()
        student_id = str(row.get('Student_No', '')).strip()

        if not name or not student_id:
            continue
            
        # 이미 링크가 있으면 패스
        if str(row.get('개별시트링크', '')).startswith('http'):
             continue

        print(f"🔨 작업 중: {name} ({student_id})...", end=" ")
        
        success = process_student_with_retry(
            drive_service, gc, target_folder_id, master_doc, 
            row, idx, df_paper, df_book, df_conf
        )

        if success:
            print("완료! ✅")
            created_count += 1
            time.sleep(5)
        else:
            print("최종 실패 ❌")
            time.sleep(5)

    print(f"\n🎉 총 {created_count}명의 시트 생성/수정 완료!")

if __name__ == "__main__":
    main()