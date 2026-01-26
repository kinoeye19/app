import os
import sys
import pickle
from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# 환경변수 로드
load_dotenv()

# ==========================================
# ⚙️ 설정 및 경로 (참조 파일 방식 적용)
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) # Research_RAG
PROJECT_ROOT = os.path.dirname(BASE_DIR)              # app

# 인증 파일 찾기 (mail_auto 폴더 우선 탐색)
CLIENT_SECRET_PATH = os.path.join(PROJECT_ROOT, "mail_auto", "client_secret.json")
TOKEN_PATH = os.path.join(PROJECT_ROOT, "mail_auto", "token.json")

# 만약 mail_auto에 없으면 현재 폴더나 상위 폴더도 확인
if not os.path.exists(CLIENT_SECRET_PATH):
    # 백업 경로 확인
    CLIENT_SECRET_PATH = os.path.join(PROJECT_ROOT, "client_secret.json")
    TOKEN_PATH = os.path.join(PROJECT_ROOT, "token.json")

# 원본 시트 ID
SOURCE_FILE_ID = os.getenv("GOOGLE_SHEET_ID")

# 프로젝트 폴더/파일 명칭
PROJECT_FOLDER_NAME = "[Project] R-E_Network_DB (Research-Education Linkage)"
NEW_FILE_NAME = "MASTER_DATASET_v1 (Do Not Delete)"

# 권한 범위 (참조 파일과 동일)
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

def get_credentials():
    """사용자 계정(OAuth)으로 인증 정보를 가져옵니다."""
    creds = None
    
    # 1. 기존 토큰 파일이 있으면 로드
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, 'rb') as token:
            try:
                # pickle 방식 (참조 파일 방식이 pickle일 경우 대비)
                # 하지만 json 방식일 수도 있으므로 예외처리 필요
                creds = pickle.load(token)
            except:
                pass
                
    # 토큰이 없거나 만료되었으면 새로 로그인
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # client_secret.json 필수
            if not os.path.exists(CLIENT_SECRET_PATH):
                print(f"❌ 'client_secret.json' 파일을 찾을 수 없습니다.")
                print(f"   경로: {CLIENT_SECRET_PATH}")
                sys.exit(1)
                
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
            
        # 새로운 토큰 저장 (다음엔 로그인 안 해도 되게)
        with open(TOKEN_PATH, 'wb') as token:
            pickle.dump(creds, token)
            
    return creds

def find_or_create_folder(service, folder_name):
    """폴더가 있으면 ID 반환, 없으면 만들고 ID 반환"""
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
    results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    files = results.get('files', [])

    if files:
        print(f"📂 기존 프로젝트 폴더를 찾았습니다: {files[0]['name']}")
        return files[0]['id']
    else:
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        folder = service.files().create(body=file_metadata, fields='id').execute()
        print(f"✨ 새 프로젝트 폴더를 생성했습니다: {folder_name}")
        return folder.get('id')

def copy_file_to_folder(service, file_id, folder_id, new_name):
    """파일 복사 (사용자 계정 용량 사용)"""
    if not file_id:
        print("❌ .env 파일에 GOOGLE_SHEET_ID가 없습니다.")
        sys.exit(1)

    # 원본 확인
    try:
        origin = service.files().get(fileId=file_id).execute()
        print(f"📄 원본 파일 확인됨: {origin.get('name')}")
    except Exception as e:
        print(f"❌ 원본 파일 접근 불가 (ID 확인 필요): {e}")
        sys.exit(1)

    file_metadata = {
        'name': new_name,
        'parents': [folder_id]
    }
    
    try:
        copied_file = service.files().copy(
            fileId=file_id,
            body=file_metadata,
            fields='id, name, webViewLink'
        ).execute()
        
        print(f"\n✅ 데이터 복제 성공! (사용자 계정 용량 사용)")
        print(f"   - 파일명: {copied_file.get('name')}")
        print(f"   - 링크: {copied_file.get('webViewLink')}")
        return copied_file.get('id')
        
    except Exception as e:
        print(f"❌ 파일 복사 실패: {e}")
        sys.exit(1)

def main():
    print("🚀 [초기화] 연구-교육 네트워크 DB 구축 (OAuth 모드)...")
    
    # 1. 사용자 인증 (브라우저 로그인 or 토큰)
    creds = get_credentials()
    service = build('drive', 'v3', credentials=creds)
    
    # 2. 프로젝트 폴더 확보
    folder_id = find_or_create_folder(service, PROJECT_FOLDER_NAME)
    
    # 3. 데이터셋 안전 복제
    new_file_id = copy_file_to_folder(service, SOURCE_FILE_ID, folder_id, NEW_FILE_NAME)
    
    # 4. 결과 안내
    print("\n" + "="*60)
    print("📌 [중요] .env 파일 업데이트")
    print("-" * 60)
    print(f"GOOGLE_SHEET_ID={new_file_id}")
    print("-" * 60)
    print("위 ID를 .env 파일에 붙여넣어 주세요.")
    print("="*60)

if __name__ == "__main__":
    main()