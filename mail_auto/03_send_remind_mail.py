import os
import sys
import json
import markdown
import gspread

# --- [모듈 경로 설정] ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(BASE_DIR)
import send_mail_module

# --- [설정 영역] ---
PARENT_DIR = os.path.dirname(BASE_DIR)

# 파일 경로
SHEET_KEY_PATH = os.path.join(PARENT_DIR, 'service_account.json')
NAVER_KEY_PATH = os.path.join(BASE_DIR, 'naver_credentials.json')
MD_FILE_PATH = os.path.join(BASE_DIR, 'email_content.md')

# 구글 시트 설정
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing"
SHEET_NAME = "remind_list"  # [수정] 실제 발송 대상 시트 이름

def get_naver_credentials():
    with open(NAVER_KEY_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def main():
    print("🚀 [리마인드] 미제출자 독촉 메일 발송을 시작합니다...")

    # 1. 인증 정보 로드
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
            lines = f.readlines()
            
        # [수정] 제목 중복 방지: 첫 줄이 '#'으로 시작하면 제목으로 간주하고 제거
        if lines and lines[0].strip().startswith('#'):
            md_text = "".join(lines[1:]) # 첫 줄 빼고 나머지 합치기
        else:
            md_text = "".join(lines)
            
    except FileNotFoundError:
        print("⚠️ email_content.md 파일이 없습니다.")
        return

    # 3. 구글 시트 연결
    print("📊 구글 시트(remind_list)에 연결 중...")
    try:
        gc = gspread.service_account(filename=SHEET_KEY_PATH)
        doc = gc.open_by_url(SPREADSHEET_URL)
        worksheet = doc.worksheet(SHEET_NAME)
        records = worksheet.get_all_records()
    except Exception as e:
        print(f"⚠️ 시트 연결 오류: {e}")
        return

    print(f"📋 총 {len(records)}명의 데이터를 가져왔습니다.")

    # 4. 발송 루프 시작
    success_count = 0
    
    for i, row in enumerate(records):
        row_num = i + 2 # 헤더가 1행이므로 데이터는 2행부터 시작
        
        # [수정] remind_list의 헤더는 소문자입니다 (name_2, email)
        name = str(row.get('name_2', '')).strip()
        email = str(row.get('email', '')).strip()
        status = str(row.get('발송여부', '')).strip()

        if not email or not name:
            continue

        if status == 'Sent':
            print(f"⏭️  [Skip] {name} - 이미 발송 완료.")
            continue

        print(f"📩 발송 시도: {name} ({email}) ...", end=" ")

        # 1. 내용 치환 (이름 등)
        personalized_md = md_text.replace("{{이름}}", name)
        
        # 2. HTML 변환
        raw_html = markdown.markdown(personalized_md, extensions=['nl2br'])

        # 3. 디자인 적용
        styled_html = f"""
        <div style="font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
            {raw_html}
        </div>
        """
        
        # 메일 제목 설정
        subject = f"[긴급] {name} 학생, 2025학년도 BK21 참여학생 연구실적 유/무를 입력해주세요 (1/25 마감)" 

        # 4. 전송
        if send_mail_module.send_email(NAVER_ID, NAVER_PWD, email, subject, styled_html):
            print("성공! ✅")
            # 발송여부 기록
            try:
                # 안전하게 컬럼 위치 찾기 (소문자/대소문자 이슈 방지 위해 다시 로드하지 않고 인덱스 계산)
                # get_all_records()의 키 리스트에서 찾음
                header_keys = list(row.keys())
                col_idx = header_keys.index('발송여부') + 1
                worksheet.update_cell(row_num, col_idx, 'Sent')
                success_count += 1
            except:
                pass 
        else:
            print("실패 ❌")

    print(f"\n🎉 [리마인드] 작업 완료! 총 {success_count}건 발송.")

if __name__ == "__main__":
    main()