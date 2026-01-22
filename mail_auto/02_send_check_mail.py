import os
import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# 메일 모듈 경로
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import send_mail_module

# --- [설정 영역] ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)

# 파일 경로
SHEET_KEY_PATH = os.path.join(PARENT_DIR, 'service_account.json')
NAVER_KEY_PATH = os.path.join(BASE_DIR, 'naver_credentials.json')

# 구글 시트 주소
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing"
SHEET_NAME_LIST = "check_list"   # 메일 보낼 명단
SHEET_NAME_PAPER = "논문"        # 성과 유무(X/O) 확인할 시트

def get_naver_credentials():
    import json
    with open(NAVER_KEY_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def main():
    print("🚀 [성과확인] 메일 발송 (성과 유무 구분 발송) 시작...")

    # 1. 인증 및 데이터 로드
    try:
        creds_naver = get_naver_credentials()
        NAVER_ID = creds_naver['id']
        NAVER_PWD = creds_naver['password']
        
        gc = gspread.service_account(filename=SHEET_KEY_PATH)
        doc = gc.open_by_url(SPREADSHEET_URL)
        
        # (1) 명단 로드
        ws_list = doc.worksheet(SHEET_NAME_LIST)
        df_list = pd.DataFrame(ws_list.get_all_records())
        
        # (2) 논문 시트 로드 (성과 유무 확인용)
        ws_paper = doc.worksheet(SHEET_NAME_PAPER)
        df_paper = pd.DataFrame(ws_paper.get_all_records())
        
        # 컬럼명 공백 제거
        df_list.columns = [c.strip() for c in df_list.columns]
        df_paper.columns = [c.strip() for c in df_paper.columns]
        
    except Exception as e:
        print(f"⚠️ 설정/접속 오류: {e}")
        return

    # 2. '연구성과 없음(X)' 학생 리스트업 (학번 기준)
    # 논문 시트의 '연구성과유무' 컬럼이 'X'인 학생의 학번을 추출
    no_result_students = set()
    
    if '연구성과유무' in df_paper.columns and '학번' in df_paper.columns:
        # 학번을 문자열로 변환하여 매칭 준비
        target_rows = df_paper[df_paper['연구성과유무'] == 'X']
        no_result_students = set(target_rows['학번'].astype(str).str.strip().tolist())
        print(f"ℹ️  연구성과 '없음(X)' 제출자 수: {len(no_result_students)}명")
    else:
        print("⚠️ 주의: '논문' 시트에서 '연구성과유무' 또는 '학번' 컬럼을 찾을 수 없습니다. 전원 일반 메일로 발송합니다.")

    print(f"📋 총 {len(df_list)}명의 명단을 확인합니다.")

    # 3. 발송 루프
    success_count = 0
    
    for idx, row in df_list.iterrows():
        # 데이터 매핑
        name = str(row.get('name_2', '')).strip()
        email = str(row.get('email', '')).strip()
        link = str(row.get('개별시트링크', '')).strip()
        status = str(row.get('발송여부', '')).strip()
        
        # 학번 (매칭용)
        # check_list에 'Student_No' 혹은 '학번' 컬럼이 있어야 함
        student_id = str(row.get('Student_No', row.get('학번', ''))).strip()

        # 필수 정보 체크
        if not name or not email:
            continue
        
        # 링크 없으면 스킵
        if not link.startswith('http'):
            # (로그 너무 많으면 주석 처리)
            # print(f"⏭️  [Skip] {name} - 시트 링크 없음")
            continue

        # 이미 보냈으면 스킵
        if status == 'Sent':
            print(f"⏭️  [Skip] {name} - 이미 발송 완료")
            continue

        # --- [메일 내용 분기 처리] ---
        
        # A. 성과 없음(X) 학생일 경우
        if student_id in no_result_students:
            print(f"📩 [성과없음] 발송: {name} ({email}) ...", end=" ")
            
            subject = f"[중요] {name} 학생에게, 2025학년도 BK21 참여학생 연구실적 입력 결과 확인 요청"
            
            # (중요) 성과 없음 대상 메일 본문
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
                <p>만약 성과가 있는데 잘못 기입된 경우, <strong>1월 24일(토) 오전까지</strong> 앱을 통해 재입력하시거나 이 메일로 회신 주시기 바랍니다.<br>
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
            
            # 일반 메일 본문
            html_content = f"""
            <div style="font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; font-size: 11pt; line-height: 1.6; color: #333;">
                <p><strong>{name}</strong> 학생에게,</p>
                <br>
                <p>안녕하세요. 국어국문학과 BK21 교육연구단 연구교수 유승진입니다.</p>
                <p>지난 1년 동안 교육연구단의 일원으로서 학업과 연구에 매진하느라 고생이 많으셨습니다.<br>
                방학 중에도 목표한 연구 활동을 성실히 이어가고 있을 것이라 생각합니다.</p>
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
                    <li><strong>회신 기한:</strong> <span style="background-color: #ffffcc; font-weight: bold;">1월 24일(토) 오전까지</span> 확인 후 회신해 주기 바랍니다. (수정 사항이 없다면 회신하지 않아도 됩니다.)</li>
                </ul>
                <br>
                <p>연구단의 원활한 성과 관리를 위해 협조해 주셔서 감사합니다.<br>
                남은 방학 기간에도 계획하고 있는 연구 활동이 좋은 결실을 맺기를 바라며, 늘 건승하길 응원합니다.</p>
                <br>
                <p><strong>BK21 교육연구단 유승진 드림</strong></p>
            </div>
            """

        # --- [전송] ---
        if send_mail_module.send_email(NAVER_ID, NAVER_PWD, email, subject, html_content):
            print("성공! ✅")
            # 발송여부 기록
            try:
                # header에서 '발송여부' 위치 찾기
                col_idx = list(df_list.columns).index('발송여부') + 1
                ws_list.update_cell(idx + 2, col_idx, 'Sent')
                success_count += 1
            except Exception as e:
                pass
        else:
            print("실패 ❌")

    print(f"\n🎉 총 {success_count}명에게 확인 메일 발송 완료!")

if __name__ == "__main__":
    main()