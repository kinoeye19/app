import os
import sys
import time
import pandas as pd
import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from gspread.exceptions import APIError

# -----------------------------------------------------------
# [수정 1] 경로 설정 및 모듈 임포트 수정
# -----------------------------------------------------------
# 현재 파일(02_send_check_mail.py)이 있는 폴더 경로 구하기
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

# 현재 폴더를 파이썬 검색 경로(sys.path)에 추가 (모듈 인식을 위해)
sys.path.append(CURRENT_DIR)

# [핵심 수정] 'mail_auto.' 제거 -> 같은 폴더에 있으므로 바로 import
import send_mail_module as send_mail_module

# -----------------------------------------------------------
# [수정 2] 인증 파일 경로 재설정 (현재 폴더 기준)
# -----------------------------------------------------------
# 파일들이 모두 02번 파일과 같은 'mail_auto' 폴더 안에 있다고 가정
TOKEN_FILE = os.path.join(CURRENT_DIR, 'token.json')
CLIENT_SECRET_FILE = os.path.join(CURRENT_DIR, 'client_secret.json')
NAVER_CRED_FILE = os.path.join(CURRENT_DIR, 'naver_credentials.json')

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing"

# 시트 명칭 정의 (원문 코드 기준 + 로직 개선용)
SHEET_MAIL_LIST = "mail_list"   # [Master] 발송 상태 기록 (Sent 저장소)
SHEET_CHECK_LIST = "check_list" # [Target] 발송 대상 목록 (순회 대상)
SHEET_PAPER = "논문"            # [Ref] 성과 유무 확인용

# 컬럼명 정의 (원문 코드 기준)
COL_SENT = "발송여부"
COL_ID = "Student_No" # 또는 '학번'

SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

# -----------------------------------------------------------
# [인증] 구글 API 인증 (OAuth 방식 유지)
# -----------------------------------------------------------
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

# -----------------------------------------------------------
# [기능] 네이버 계정 로드
# -----------------------------------------------------------
def get_naver_credentials():
    import json
    with open(NAVER_CRED_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

# -----------------------------------------------------------
# [메인] 실행 로직
# -----------------------------------------------------------
def main():
    print("🚀 [성과확인] 메일 발송 (Windows 경로 수정됨) 시작...")
    
    # 1. 구글 시트 연결
    creds = get_credentials()
    gc = gspread.authorize(creds)
    
    try:
        doc = gc.open_by_url(SPREADSHEET_URL)
        ws_mail = doc.worksheet(SHEET_MAIL_LIST)   # 기록용
        ws_check = doc.worksheet(SHEET_CHECK_LIST) # 타겟용
        ws_paper = doc.worksheet(SHEET_PAPER)      # 참조용
    except Exception as e:
        print(f"⚠️ 시트 로드 오류: {e}")
        return
    
    # 2. 데이터프레임 변환
    data_mail = ws_mail.get_all_records()
    data_check = ws_check.get_all_records()
    data_paper = ws_paper.get_all_records()
    
    df_mail = pd.DataFrame(data_mail)
    df_check = pd.DataFrame(data_check)
    df_paper = pd.DataFrame(data_paper)
    
    # 컬럼 공백 제거
    df_mail.columns = [str(c).strip() for c in df_mail.columns]
    df_check.columns = [str(c).strip() for c in df_check.columns]
    df_paper.columns = [str(c).strip() for c in df_paper.columns]
    
    # 3. '연구성과유무'가 'X'인 학생 학번 추출
    no_result_students = set()
    if '연구성과유무' in df_paper.columns and '학번' in df_paper.columns:
        target_rows = df_paper[df_paper['연구성과유무'] == 'X']
        no_result_students = set(target_rows['학번'].astype(str).str.strip().tolist())
        print(f"ℹ️  연구성과 '없음(X)' 제출자 수: {len(no_result_students)}명")
    else:
        print("⚠️ 주의: '논문' 시트에서 컬럼을 찾을 수 없어 전원 일반 메일로 분류될 수 있습니다.")

    # 4. [핵심 로직] mail_list 매핑 (학번 -> {행번호, 현재상태})
    mail_map = {}
    for idx, row in df_mail.iterrows():
        # mail_list에서는 'Student_No'를 우선 찾고 없으면 '학번'
        s_id = str(row.get(COL_ID, row.get('학번', ''))).strip()
        status = str(row.get(COL_SENT, '')).strip()
        
        if s_id:
            mail_map[s_id] = {
                'row_idx': idx + 2,  # gspread 1-based index (헤더 제외)
                'status': status
            }

    # '발송여부' 컬럼 인덱스 찾기 (mail_list 기준)
    header_values = ws_mail.row_values(1)
    try:
        sent_col_idx = header_values.index(COL_SENT) + 1
    except ValueError:
        print(f"!!! [오류] '{SHEET_MAIL_LIST}' 시트에 '{COL_SENT}' 열이 없습니다.")
        return

    # 5. 네이버 SMTP 정보 로드
    naver_info = get_naver_credentials()
    smtp_user = naver_info['id']
    smtp_password = naver_info['password']
    
    print(f"📋 총 {len(df_check)}명의 명단을 확인합니다.")
    
    success_count = 0
    count_skip = 0

    # 6. 발송 루프 (check_list 순회)
    for idx, row in df_check.iterrows():
        name = str(row.get('name_2', row.get('Name', ''))).strip()
        email = str(row.get('email', row.get('Email', ''))).strip()
        link = str(row.get('개별시트링크', '')).strip()
        student_id = str(row.get('Student_No', row.get('학번', ''))).strip()

        if not name or not email:
            continue
            
        if student_id not in mail_map:
            continue

        # [검증 B] 이미 발송되었는가? (mail_list의 상태값 참조)
        current_status = mail_map[student_id]['status']
        
        if current_status == 'Sent':
            print(f"⏭️  [Skip] {name} - 이미 발송 완료 (mail_list 기준)")
            count_skip += 1
            continue
            
        if not link.startswith('http'):
            continue

        # --- [메일 내용 분기 처리: 사용자 원본 그대로 적용] ---
        subject = ""
        html_content = ""

        # A. 성과 없음(X) 학생일 경우
        if student_id in no_result_students:
            print(f"📩 [성과없음] 발송: {name} ({email}) ...", end=" ")
            
            subject = f"[중요] {name} 학생에게, 2025학년도 BK21 참여학생 연구실적 입력 결과 확인 요청"
            
            html_content = f"""
            <div style="font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
                <p><strong>{name}</strong> 학생에게,</p>
                <br>
                <p>안녕하세요. 국어국문학과 BK21 교육연구단 연구교수 유승진입니다.</p>
                <p>2025학년도 연구실적 입력 기간 동안 앱을 통해 제출해주신 내용을 확인하였습니다.</p>
                <br>
                <div style="background-color: #fff3cd; padding: 15px; border-left: 5px solid #ffc107;">
                    <p style="margin: 0;"><strong>📢 확인 사항</strong></p>
                    <p style="margin-top: 5px;">
                        현재 {name} 학생은 현재 <strong>'연구성과 없음(실적 없음)'</strong>으로 제출된 상태입니다.<br>
                        혹시 <strong>성과가 있는데 실수로 '없음'을 선택했거나, 누락된 내용이 있는지</strong> 다시 한번 확인을 부탁드립니다.
                    </p>
                </div>
                <br>
                <p>아래 링크를 클릭하시면 현재 제출된 상태를 확인하실 수 있습니다.</p>
                <p><strong>🔗 내 성과 확인하기:</strong> <a href="{link}" target="_blank">{link}</a></p>
                <br>
                <p>만약 성과가 있는데 잘못 기입된 경우, <strong>1월 26일(월) 오전까지</strong> 해당 메일로 회신해 주시기 바랍니다.<br>
                (실적이 없는 것이 맞다면, 별도로 회신하지 않으셔도 됩니다.)</p>
                <br>
                <p>감사합니다.<br>
                BK21 교육연구단 유승진 드림</p>
            </div>
            """

        # B. 일반(성과 있음) 학생일 경우
        else:
            print(f"📩 [일반] 발송: {name} ({email}) ...", end=" ")
            
            subject = f"[BK21] 2025학년도 연구실적 입력 결과 확인 요청 ({name} 학생)"
            
            html_content = f"""
            <div style="font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
                <p><strong>{name}</strong> 학생에게,</p>
                <br>
                <p>안녕하세요. 국어국문학과 BK21 교육연구단 연구교수 유승진입니다.</p>
                <p>지난 1년 동안 교육연구단의 일원으로서 학업과 연구에 매진하느라 고생이 많으셨습니다.<br>
                그간의 성실함으로 방학 중에도 목표한 연구 활동을 성실히 이어가고 있을 것이라 생각합니다.</p>
                <br>
                <p>다름이 아니라, <strong>{name}</strong> 학생이 <strong>'연구성과 입력 앱'</strong>을 통해 제출해 준 
                2025학년도 연구실적(논문/저서/학술대회) 데이터를 정리하여 공유하니, 확인을 부탁드립니다.</p>
                <p>해당 내용은 한국연구재단에 우리 사업단의 성과로 최종 보고될 예정이므로, 
                보고에 앞서 <strong>누락되거나 잘못 기입된 부분은 없는지 학생 본인의 확인</strong>이 필요합니다.</p>
                <br>
                <div style="background-color: #f0f8ff; padding: 20px; border-left: 5px solid #007bff; margin: 10px 0;">
                    <h3 style="margin-top: 0; color: #0056b3;">✅ 내 성과 확인하기</h3>
                    <p><strong>확인 링크:</strong> <a href="{link}" target="_blank" style="font-weight: bold; color: #007bff;">{link}</a></p>
                    <p><strong>확인 방법:</strong> 위 링크를 클릭하여 내용 확인 (읽기 전용)</p>
                </div>
                <br>
                <h3 style="color: #d9534f;">⚠️ 확인 및 수정 요청</h3>
                <ul>
                    <li><strong>내용 확인:</strong> 본인이 수행한 실적 중 빠진 내용이나 오타가 없는지 꼼꼼히 살펴봐 주기 바랍니다.</li>
                    <li><strong>수정 요청:</strong> 파일은 직접 수정할 수 없습니다. 수정이 필요한 경우, <strong>이 메일로 회신</strong>하여 내용을 알려주면 반영하겠습니다.</li>
                    <li><strong>회신 기한:</strong> <span style="background-color: #ffffcc; font-weight: bold;">1월 26일(월) 오전까지</span> 확인 후 회신해 주기 바랍니다. (수정 사항이 없다면 회신하지 않아도 됩니다.)</li>
                </ul>
                <br>
                <p>연구단의 원활한 성과 관리를 위해 협조해 주셔서 감사합니다.<br>
                남은 방학 기간에도 계획하고 있는 연구 활동이 좋은 결실을 맺기를 바라며, 늘 건승하길 응원합니다.</p>
                <br>
                <p><strong>BK21 교육연구단 유승진 드림</strong></p>
            </div>
            """

        # --- [전송 및 기록] ---
        if send_mail_module.send_email(smtp_user, smtp_password, email, subject, html_content):
            print("성공! ✅")
            
            # [기록 위치] mail_list 시트에 기록
            target_row_idx = mail_map[student_id]['row_idx']
            try:
                ws_mail.update_cell(target_row_idx, sent_col_idx, 'Sent')
                success_count += 1
                time.sleep(1.5) # API 부하 방지
            except Exception as e:
                print(f" (발송은 됐으나 기록 실패: {e})")
        else:
            print("실패 ❌")

    print(f"\n🎉 총 {success_count}명에게 확인 메일 발송 완료! (스킵: {count_skip}명)")

if __name__ == "__main__":
    main()