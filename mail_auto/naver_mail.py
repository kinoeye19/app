import os
import json
import smtplib
import markdown
import gspread
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- [설정 영역] ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)

# 인증 키 경로
SHEET_KEY_PATH = os.path.join(PARENT_DIR, 'service_account.json')
NAVER_KEY_PATH = os.path.join(BASE_DIR, 'naver_credentials.json') # 네이버 키 파일
MD_FILE_PATH = os.path.join(BASE_DIR, 'email_content.md')

# 구글 시트 설정
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing"  # <--- 수정 필요
SHEET_NAME = "test"  # 하단 탭 이름

# 네이버 SMTP 설정
SMTP_SERVER = "smtp.naver.com"
SMTP_PORT = 465

def get_naver_credentials():
    """JSON 파일에서 네이버 아이디/비번을 가져옵니다."""
    with open(NAVER_KEY_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def send_email_naver(user_id, user_pwd, to_email, subject, html_content):
    """네이버 SMTP를 통해 이메일을 전송합니다."""
    
    # 메일 객체 생성
    msg = MIMEMultipart()
    msg['From'] = f"{user_id}@naver.com"
    msg['To'] = to_email
    msg['Subject'] = subject

    # 본문 추가
    msg.attach(MIMEText(html_content, 'html'))

    try:
        # SMTP 서버 연결 (SSL 보안 연결)
        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        server.login(user_id, user_pwd)
        
        # 메일 전송
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"❌ 전송 실패 ({to_email}): {e}")
        return False

def main():
    print("🚀 [네이버] 메일 자동화 프로그램을 시작합니다...")

    # 1. 네이버 계정 정보 로드
    try:
        creds = get_naver_credentials()
        NAVER_ID = creds['id']
        NAVER_PWD = creds['password']
    except FileNotFoundError:
        print("⚠️ naver_credentials.json 파일이 없습니다.")
        return

    # 2. 마크다운 파일 읽기
    try:
        with open(MD_FILE_PATH, 'r', encoding='utf-8') as f:
            md_text = f.read()
    except FileNotFoundError:
        print("⚠️ email_content.md 파일이 없습니다.")
        return

    # 3. 구글 시트 연결
    print("📊 구글 시트에 연결 중...")
    try:
        gc = gspread.service_account(filename=SHEET_KEY_PATH)
        doc = gc.open_by_url(SPREADSHEET_URL)
        worksheet = doc.worksheet(SHEET_NAME)
        records = worksheet.get_all_records()
    except Exception as e:
        print(f"⚠️ 시트 연결 오류: {e}")
        return

    print(f"📋 총 {len(records)}개의 데이터를 가져왔습니다.")

    # 4. 발송 루프 시작
    success_count = 0
    
    for i, row in enumerate(records):
        row_num = i + 2
        
        name = row.get('Name_2')
        email = row.get('E-mail')
        status = row.get('발송여부')

        if not email or not name:
            continue

        if status == 'Sent':
            print(f"⏭️  [Skip] {name}님은 이미 발송 완료.")
            continue

        print(f"📩 발송 시도: {name} ({email}) ...", end=" ")

        # 내용 치환 및 변환
        # 1. 내용 치환 및 HTML 변환
        personalized_md = md_text.replace("{{이름}}", str(name))
        raw_html = markdown.markdown(personalized_md, extensions=['nl2br'])

        # 2. 한국형 메일 스타일(맑은고딕/애플고딕, 적당한 크기) 적용
        # 네이버 메일의 기본 포맷과 유사하게 설정했습니다.
        styled_html = f"""
        <div style="font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
            {raw_html}
        </div>
        """
        
        subject = f"[BK21] 2025학년도 연구실적 입력 요청 (1/23 마감)" # 제목은 고정하거나 필요시 수정
        
        # 3. 전송 (styled_html을 보냄)
        if send_email_naver(NAVER_ID, NAVER_PWD, email, subject, styled_html):
            print("성공! ✅")
            worksheet.update_cell(row_num, list(row.keys()).index('발송여부') + 1, 'Sent')
            success_count += 1
        else:
            print("실패 ❌")
        # [수정된 부분 끝] -------------------------------------------

    print(f"\n🎉 작업 완료! 총 {success_count}건의 메일을 네이버 계정으로 발송했습니다.")

if __name__ == "__main__":
    main()