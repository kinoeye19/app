import os
import sys
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build

# 환경변수 로드
load_dotenv()

# ==========================================
# ⚙️ 설정 및 경로 (사용자 환경 맞춤)
# ==========================================
# 1. 인증 키 경로 (mail_auto_agent 폴더 참조)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)
JSON_KEY_PATH = os.path.join(PROJECT_ROOT, "mail_auto_agent", "service_account.json")

# 2. 원본 파일 ID (.env에서 가져옴)
SOURCE_FILE_ID = os.getenv("GOOGLE_SHEET_ID")

# 3. 생성할 프로젝트 폴더 및 파일 명칭
PROJECT_FOLDER_NAME = "[Project] R-E_Network_DB (Research-Education Linkage)"
NEW_FILE_NAME = "MASTER_DATASET_v1 (Do Not Delete)"

# 4. 권한 범위 (드라이브 전체 제어 + 스프레드시트)
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

def authenticate_drive_api():
    """구글 드라이브 API 인증 및 서비스 빌드"""
    if not os.path.exists(JSON_KEY_PATH):
        print(f"❌ 인증 키 파일이 없습니다: {JSON_KEY_PATH}")
        sys.exit(1)
        
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_PATH, SCOPES)
    service = build('drive', 'v3', credentials=creds)
    return service

def find_or_create_folder(service, folder_name):
    """폴더가 있으면 ID 반환, 없으면 만들고 ID 반환"""
    # 1. 폴더 검색 (삭제되지 않은(trashed=false) 폴더 중 이름 일치)
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and trashed=false"
    results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    files = results.get('files', [])

    if files:
        print(f"📂 기존 프로젝트 폴더를 찾았습니다: {files[0]['name']} (ID: {files[0]['id']})")
        return files[0]['id']
    else:
        # 2. 없으면 생성
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        folder = service.files().create(body=file_metadata, fields='id').execute()
        print(f"✨ 새 프로젝트 폴더를 생성했습니다: {folder_name} (ID: {folder.get('id')})")
        return folder.get('id')

def copy_file_to_folder(service, file_id, folder_id, new_name):
    """파일을 특정 폴더로 복사하고 이름 변경"""
    # 원본 파일 정보 확인
    try:
        origin = service.files().get(fileId=file_id).execute()
        print(f"📄 원본 파일 확인됨: {origin.get('name')}")
    except Exception as e:
        print(f"❌ 원본 파일을 찾을 수 없습니다. .env의 GOOGLE_SHEET_ID를 확인하세요.\n오류: {e}")
        sys.exit(1)

    # 복사 메타데이터 설정 (부모 폴더 지정, 이름 변경)
    file_metadata = {
        'name': new_name,
        'parents': [folder_id]
    }
    
    try:
        # 파일 복사 실행
        copied_file = service.files().copy(
            fileId=file_id,
            body=file_metadata,
            fields='id, name, webViewLink'
        ).execute()
        
        print(f"\n✅ 데이터 복제 성공!")
        print(f"   - 파일명: {copied_file.get('name')}")
        print(f"   - 위치: {PROJECT_FOLDER_NAME} 폴더 내부")
        print(f"   - 링크: {copied_file.get('webViewLink')}")
        return copied_file.get('id')
        
    except Exception as e:
        print(f"❌ 파일 복사 실패: {e}")
        sys.exit(1)

def main():
    print("🚀 [초기화] 연구-교육 네트워크 DB 구축을 시작합니다...")
    
    # 1. API 연결
    service = authenticate_drive_api()
    
    # 2. 프로젝트 폴더 확보
    folder_id = find_or_create_folder(service, PROJECT_FOLDER_NAME)
    
    # 3. 데이터셋 안전 복제
    new_file_id = copy_file_to_folder(service, SOURCE_FILE_ID, folder_id, NEW_FILE_NAME)
    
    # 4. 다음 단계를 위한 안내
    print("\n" + "="*50)
    print("📌 [중요] 다음 단계(데이터 수집)를 위해 아래 내용을 참고하세요.")
    print(f"새로 생성된 마스터 데이터 시트 ID: {new_file_id}")
    print("👉 .env 파일의 'GOOGLE_SHEET_ID'를 위 ID로 변경하면,")
    print("   원본 손상 없이 안전하게 스크래핑 작업을 진행할 수 있습니다.")
    print("="*50)

if __name__ == "__main__":
    main()