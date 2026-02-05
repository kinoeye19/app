import os
import sys
import time
import platform # OS 감지용
import pandas as pd
import gspread
from dotenv import load_dotenv

# -----------------------------------------------------------
# [기본 설정] 경로 및 모듈 임포트
# -----------------------------------------------------------
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(CURRENT_DIR)

import send_mail_module as send_mail_module

def main():
    # 1. 환경 변수 로드
    load_dotenv()
    print("🚀 [성과확인] 메일 발송 시스템을 시작합니다...")

    # -----------------------------------------------------------
    # [OS 자동 감지 및 경로 할당]
    # -----------------------------------------------------------
    current_os = platform.system()
    
    service_account_file = None
    naver_cred_file = None
    
    if current_os == 'Windows':
        print(f">> 감지된 OS: Windows (윈도우 설정을 사용합니다)")
        service_account_file = os.getenv("GOOGLE_JSON_KEY_WIN")
        naver_cred_file = os.getenv("NAVER_CRED_PATH_WIN")
        
    elif current_os == 'Darwin': # Mac
        print(f">> 감지된 OS: macOS (맥 설정을 사용합니다)")
        service_account_file = os.getenv("GOOGLE_JSON_KEY_MAC")
        naver_cred_file = os.getenv("NAVER_CRED_PATH_MAC")
        
    else:
        print(f"[경고] 알 수 없는 운영체제입니다: {current_os}. 윈도우 설정을 시도합니다.")
        service_account_file = os.getenv("GOOGLE_JSON_KEY_WIN")
        naver_cred_file = os.getenv("NAVER_CRED_PATH_WIN")

    # [공통 설정 로드]
    spreadsheet_url = os.getenv("MAIL_SHEET_URL")

    # [필수 파일 검증] - 구글 키
    if not service_account_file or not os.path.exists(service_account_file):
        print(f"\n[치명적 오류] 구글 서비스 계정 키 파일을 찾을 수 없습니다.")
        print(f"- 경로: {service_account_file}")
        print(">> Dropbox 동기화 여부 및 .env 경로를 확인하세요.")
        return

    # [필수 파일 검증] - 네이버 키
    if not naver_cred_file or not os.path.exists(naver_cred_file):
        print(f"\n[치명적 오류] 네이버 인증 파일을 찾을 수 없습니다.")
        print(f"- 경로: {naver_cred_file}")
        print(">> Dropbox 동기화 여부 및 .env 경로를 확인하세요.")
        return

    # -----------------------------------------------------------
    # [내부 함수] 네이버 계정 로드
    # -----------------------------------------------------------
    def get_naver_credentials():
        import json
        with open(naver_cred_file, 'r', encoding='utf-8') as f:
            return json.load(f)

    # 시트/컬럼 명칭 정의
    SHEET_MAIL_LIST = "mail_list"
    SHEET_CHECK_LIST = "check_list"
    SHEET_PAPER = "논문"
    COL_SENT = "발송여부"
    COL_ID = "Student_No"

    # -----------------------------------------------------------
    # [로직 시작]
    # -----------------------------------------------------------
    
    # 2. 구글 시트 연결
    try:
        gc = gspread.service_account(filename=service_account_file)
        doc = gc.open_by_url(spreadsheet_url)
        
        ws_mail = doc.worksheet(SHEET_MAIL_LIST)
        ws_check = doc.worksheet(SHEET_CHECK_LIST)
        ws_paper = doc.worksheet(SHEET_PAPER)
        print(">> 구글 스프레드시트 접속 성공!")
        
    except Exception as e:
        print(f"[오류] 시트 접속 실패: {e}")
        return
    
    # 3. 데이터 로딩
    print(">> 데이터 로딩 중...")
    try:
        data_mail = ws_mail.get_all_records()
        data_check = ws_check.get_all_records()
        data_paper = ws_paper.get_all_records()
    except Exception as e:
        print(f"[오류] 데이터 로딩 실패: {e}")
        return
    
    df_mail = pd.DataFrame(data_mail)
    df_check = pd.DataFrame(data_check)
    df_paper = pd.DataFrame(data_paper)
    
    # 컬럼 공백 제거
    df_mail.columns = [str(c).strip() for c in df_mail.columns]
    df_check.columns = [str(c).strip() for c in df_check.columns]
    df_paper.columns = [str(c).strip() for c in df_paper.columns]
    
    # 4. '연구성과유무'가 'X'인 학생 추출
    no_result_students = set()
    if '연구성과유무' in df_paper.columns and '학번' in df_paper.columns:
        target_rows = df_paper[df_paper['연구성과유무'] == 'X']
        no_result_students = set(target_rows['학번'].astype(str).str.strip().tolist())
        print(f"ℹ️  연구성과 '없음(X)' 제출자 수: {len(no_result_students)}명")

    # 5. 발송 기록 매핑
    mail_map = {}
    for idx, row in df_mail.iterrows():
        s_id = str(row.get(COL_ID, row.get('학번', ''))).strip()
        status = str(row.get(COL_SENT, '')).strip()
        if s_id:
            mail_map[s_id] = {'row_idx': idx + 2, 'status': status}

    # '발송여부' 컬럼 인덱스 찾기
    header_values = ws_mail.row_values(1)
    try:
        sent_col_idx = header_values.index(COL_SENT) + 1
    except ValueError:
        print(f"[오류] '{SHEET_MAIL_LIST}' 시트에 '{COL_SENT}' 열이 없습니다.")
        return

    # 6. 네이버 메일 정보 로드
    naver_info = get_naver_credentials()
    if not naver_info: return
    
    smtp_user = naver_info['id']
    smtp_password = naver_info['password']
    
    print(f"📋 총 {len(df_check)}명의 명단을 확인합니다.")
    
    success_count = 0
    count_skip = 0

    # 7. 발송 루프
    for idx, row in df_check.iterrows():
        name = str(row.get('name_2', row.get('Name', ''))).strip()
        email = str(row.get('email', row.get('Email', ''))).strip()
        link = str(row.get('개별시트링크', '')).strip()
        student_id = str(row.get('Student_No', row.get('학번', ''))).strip()

        if not name or not email: continue
        if student_id not in mail_map: continue

        if mail_map[student_id]['status'] == 'Sent':
            print(f"⏭️  [Skip] {name} - 이미 발송 완료")
            count_skip += 1
            continue
            
        if not link.startswith('http'): continue

        # 메일 내용 작성
        subject = ""
        html_content = ""

        # A. 성과 없음(X)
        if student_id in no_result_students:
            print(f"📩 [성과없음] 발송: {name} ({email}) ...", end=" ")
            subject = f"[중요] {name} 학생에게, 2025학년도 BK21 참여학생 연구실적 입력 결과 확인 요청"
            html_content = f"""
            <div style="font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
                <p><strong>{name}</strong> 학생에게,</p>
                <br>
                <p>안녕하세요. 국어국문학과 BK21 교육연구단 연구교수 유승진입니다.</p>
                <div style="background-color: #fff3cd; padding: 15px; border-left: 5px solid #ffc107;">
                    <p style="margin: 0;"><strong>📢 확인 사항</strong></p>
                    <p style="margin-top: 5px;">현재 <strong>'연구성과 없음'</strong>으로 제출되었습니다. 본인이 제출한 내용이 맞는지 확인 부탁드립니다.</p>
                </div>
                <br>
                <p><strong>🔗 내 성과 확인하기:</strong> <a href="{link}" target="_blank">{link}</a></p>
                <br>
                <p>BK21 교육연구단 유승진 드림</p>
            </div>
            """
        # B. 일반 (성과 있음)
        else:
            print(f"📩 [일반] 발송: {name} ({email}) ...", end=" ")
            subject = f"[BK21] 2025학년도 연구실적 입력 결과 확인 요청 ({name} 학생)"
            html_content = f"""
            <div style="font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
                <p><strong>{name}</strong> 학생에게,</p>
                <br>
                <p>안녕하세요. 국어국문학과 BK21 교육연구단 연구교수 유승진입니다.</p>
                <p>제출해주신 2025학년도 연구실적 데이터를 공유합니다. 누락이나 오타가 없는지 확인 바랍니다.</p>
                <br>
                <div style="background-color: #f0f8ff; padding: 20px; border-left: 5px solid #007bff; margin: 10px 0;">
                    <h3 style="margin-top: 0; color: #0056b3;">✅ 내 성과 확인하기</h3>
                    <p><strong>확인 링크:</strong> <a href="{link}" target="_blank">{link}</a></p>
                </div>
                <br>
                <p>BK21 교육연구단 유승진 드림</p>
            </div>
            """

        # 전송 실행
        if send_mail_module.send_email(smtp_user, smtp_password, email, subject, html_content):
            print("성공! ✅")
            try:
                ws_mail.update_cell(mail_map[student_id]['row_idx'], sent_col_idx, 'Sent')
                success_count += 1
                time.sleep(1.5)
            except Exception as e:
                print(f" (기록 실패: {e})")
        else:
            print("실패 ❌")

    print(f"\n🎉 총 {success_count}명 발송 완료! (이미 발송됨: {count_skip}명)")

if __name__ == "__main__":
    main()