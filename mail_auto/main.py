import os
import markdown
import base64
import gspread  # 구글 시트 제어 라이브러리
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# --- [설정 영역] ---
# 1. 파일 경로 설정 (상대 경로 활용)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # mail_auto 폴더
PARENT_DIR = os.path.dirname(BASE_DIR)                 # app 폴더

# 인증 키 경로
SHEET_KEY_PATH = os.path.join(PARENT_DIR, 'service_account.json') # 상위 폴더
GMAIL_KEY_PATH = os.path.join(BASE_DIR, 'credentials.json')       # 현재 폴더
GMAIL_TOKEN_PATH = os.path.join(BASE_DIR, 'token.json')           # 자동 생성됨
MD_FILE_PATH = os.path.join(BASE_DIR, 'email_content.md')         # 메일 본문

# 2. 구글 시트 설정
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing"  # <--- 수정 필요
SHEET_NAME = "test"  # 하단 탭 이름

# 3. Gmail API 권한 범위
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def get_gmail_service():
    """Gmail API 서비스 객체를 생성합니다."""
    creds = None
    # 토큰이 이미 있으면 로드
    if os.path.exists(GMAIL_TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(GMAIL_TOKEN_PATH, SCOPES)
    
    # 토큰이 없거나 유효하지 않으면 새로 로그인
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(GMAIL_KEY_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        # 다음 실행을 위해 토큰 저장
        with open(GMAIL_TOKEN_PATH, 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

def send_email(service, to_email, subject, html_content):
    """이메일을 전송합니다."""
    message = MIMEMultipart()
    message['to'] = to_email
    message['subject'] = subject

    # HTML 본문 첨부
    msg = MIMEText(html_content, 'html')
    message.attach(msg)

    # Base64 인코딩
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    
    try:
        service.users().messages().send(userId="me", body={'raw': raw_message}).execute()
        return True
    except Exception as e:
        print(f"❌ 전송 실패 ({to_email}): {e}")
        return False

def main():
    print("🚀 메일 자동화 프로그램을 시작합니다...")

    # 1. 마크다운 파일 읽기
    try:
        with open(MD_FILE_PATH, 'r', encoding='utf-8') as f:
            md_text = f.read()
    except FileNotFoundError:
        print("⚠️ email_content.md 파일이 없습니다.")
        return

    # 2. 구글 시트 연결
    print("📊 구글 시트에 연결 중...")
    try:
        gc = gspread.service_account(filename=SHEET_KEY_PATH)
        doc = gc.open_by_url(SPREADSHEET_URL)
        worksheet = doc.worksheet(SHEET_NAME)
    except Exception as e:
        print(f"⚠️ 시트 연결 오류: {e}")
        return

    # 데이터 가져오기 (헤더 포함)
    records = worksheet.get_all_records()
    print(f"📋 총 {len(records)}개의 데이터를 가져왔습니다.")

    # 3. Gmail 서비스 연결
    print("📧 Gmail API 인증 중...")
    gmail_service = get_gmail_service()

    # 4. 발송 루프 시작
    success_count = 0
    
    # 헤더 위치 찾기 (행 번호 계산을 위해 필요, 1행은 헤더이므로 데이터는 2행부터 시작)
    # gspread는 1-based index 사용
    
    for i, row in enumerate(records):
        row_num = i + 2  # 실제 시트상의 행 번호 (헤더가 1행이므로)
        
        name = row.get('Name_2')  # 시트의 '이름' 컬럼
        email = row.get('E-mail') # 시트의 '이메일' 컬럼
        status = row.get('발송여부') # 시트의 '발송여부' 컬럼

        # 필수 정보 체크
        if not email or not name:
            continue

        # 이미 보낸 사람은 패스
        if status == 'Sent':
            print(f"⏭️  [Skip] {name}님은 이미 발송 완료.")
            continue

        print(f"📩 발송 시도: {name} ({email}) ...", end=" ")

        # 마크다운 -> HTML 변환 (치환 기능 포함)
        # {{이름}}을 실제 이름으로 변경
        personalized_md = md_text.replace("{{이름}}", str(name))
        html_content = markdown.markdown(personalized_md)

        # 메일 제목 설정
        subject = f"[안내] {name}님, 요청하신 자료입니다."

        # 전송
        if send_email(gmail_service, email, subject, html_content):
            print("성공! ✅")
            # 시트에 'Sent' 기록
            worksheet.update_cell(row_num, list(row.keys()).index('발송여부') + 1, 'Sent')
            success_count += 1
        else:
            print("실패 ❌")

    print(f"\n🎉 작업 완료! 총 {success_count}건의 메일을 새로 발송했습니다.")

if __name__ == "__main__":
    main()