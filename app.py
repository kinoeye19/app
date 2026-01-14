import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import datetime

# ---------------------------------------------------------
# 1. Google Sheets 인증 및 연결 설정
# ---------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_connection():
    # Streamlit Cloud 배포 환경 (Secrets 사용)
    if "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        # 로컬 테스트 환경 (json 파일 사용)
        # 로컬에서 실행할 때는 'service_account.json' 파일이 같은 폴더에 있어야 합니다.
        try:
            creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
        except FileNotFoundError:
            st.error("로컬 인증 파일(service_account.json)을 찾을 수 없습니다.")
            st.stop()
    
    client = gspread.authorize(creds)
    return client

# ---------------------------------------------------------
# 2. 화면 구성 (UI)
# ---------------------------------------------------------

st.set_page_config(page_title="연구 성과 제출", page_icon="🎓")

st.title("🎓 연구 결과물 제출 시스템")
st.markdown("""
연구실 성과 취합을 위한 페이지입니다.  
**논문, 저서, 학술대회 발표** 실적을 정확하게 입력해 주세요.
""")

# 세션 상태 초기화 (성과 항목 추가/삭제 관리)
if 'research_items' not in st.session_state:
    st.session_state.research_items = []

# --- A. 학생 기본 정보 (상단 고정) ---
with st.container():
    st.subheader("1. 기본 정보")
    col1, col2 = st.columns(2)
    with col1:
        student_name = st.text_input("이름", placeholder="홍길동")
    with col2:
        student_id = st.text_input("학번", placeholder="20241234")

st.divider()

# --- B. 연구 성과 입력 (동적 폼) ---
st.subheader("2. 연구 성과 입력")
st.info("여러 건의 성과가 있다면 '➕ 성과 추가하기' 버튼을 눌러 계속 추가할 수 있습니다.")

def add_item():
    st.session_state.research_items.append({
        "type": "논문",  # 기본값
        "role": "",
        "authors_all": "",
        "author_count": 1,
        "title": "",
        "journal": "",
        "details": "",
        "date": ""
    })

def remove_item(index):
    st.session_state.research_items.pop(index)

if st.button("➕ 성과 추가하기"):
    add_item()

# 입력 폼 생성 루프
for i, item in enumerate(st.session_state.research_items):
    with st.expander(f"📝 성과 #{i+1} 입력", expanded=True):
        # 1. 삭제 버튼
        if st.button("이 항목 삭제", key=f"del_{i}"):
            remove_item(i)
            st.rerun()

        # 2. 구분 선택 (논문/저서/발표)
        type_options = ["논문", "저서", "학술대회 발표"]
        selected_type = st.selectbox(
            "구분", 
            type_options, 
            key=f"type_{i}",
            index=type_options.index(item["type"]) if item["type"] in type_options else 0
        )
        
        # 3. 참여 역할 선택 (구분에 따라 선택지 변경)
        role_options = []
        if selected_type == "논문":
            role_options = ["단독", "제1저자", "공동저자", "교신저자"]
        elif selected_type == "저서":
            role_options = ["단독저자", "공동저자(챕터 집필)", "공동저자(전체 공저)", "대표저자/에디터"]
        else: # 학술대회 발표
            role_options = ["발표자", "공동연구자(발표안함)"]
        
        selected_role = st.selectbox("참여 역할", role_options, key=f"role_{i}")

        # 4. 저자 정보 입력
        c1, c2 = st.columns([3, 1])
        with c1:
            authors_all = st.text_input(
                "전체 저자 명단 (순서대로 기입)", 
                placeholder="예: 홍길동(본인), 김철수, 이영희", 
                key=f"auth_all_{i}",
                help="논문/저서에 기재된 순서대로 모든 저자를 적어주세요."
            )
        with c2:
            author_count = st.number_input(
                "전체 저자 수", 
                min_value=1, 
                value=item.get("author_count", 1), 
                key=f"auth_cnt_{i}"
            )

        # 5. 상세 정보 입력
        title_label = "논문 제목"
        journal_label = "저널명 (Journal Name)"
        detail_label = "세부정보 (Vol, No, page)"
        date_label = "게재일자 (년월)"

        # 라벨 동적 변경
        if selected_type == "저서":
            title_label = "저서명 (책 제목)"
            journal_label = "출판사"
            detail_label = "ISBN 혹은 개정판 정보"
            date_label = "출판일자"
        elif selected_type == "학술대회 발표":
            title_label = "발표 제목"
            journal_label = "학술대회명"
            detail_label = "개최 장소"
            date_label = "발표일자"

        title = st.text_input(title_label, key=f"title_{i}")
        
        cc1, cc2 = st.columns(2)
        with cc1:
            journal = st.text_input(journal_label, key=f"journal_{i}")
        with cc2:
            details = st.text_input(detail_label, placeholder="예: Vol.10, No.2, pp.10-20", key=f"detail_{i}")
            
        date_val = st.text_input(date_label, placeholder="예: 2024-05", key=f"date_{i}")

        # 입력값 세션에 업데이트
        st.session_state.research_items[i].update({
            "type": selected_type,
            "role": selected_role,
            "authors_all": authors_all,
            "author_count": author_count,
            "title": title,
            "journal": journal,
            "details": details,
            "date": date_val
        })

st.divider()

# --- C. 제출 버튼 및 구글 시트 전송 ---
if st.button("📤 제출하기", type="primary"):
    # 유효성 검사
    if not student_name or not student_id:
        st.error("맨 위의 '이름'과 '학번'을 반드시 입력해주세요.")
    elif len(st.session_state.research_items) == 0:
        st.warning("입력된 성과가 없습니다. '성과 추가하기'를 눌러 내용을 작성해주세요.")
    else:
        try:
            with st.spinner("데이터를 저장 중입니다..."):
                client = get_connection()
                
                # *** 아래 URL을 선생님의 실제 구글 시트 URL로 변경하세요 ***
                SHEET_URL = "https://docs.google.com/spreadsheets/d/여기에_구글시트_ID_입력"
                
                sheet = client.open_by_url(SHEET_URL).sheet1
                
                rows_to_add = []
                now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                for item in st.session_state.research_items:
                    # 헤더 순서: 타임스탬프, 이름, 학번, 구분, 참여역할, 전체저자, 저자수, 제목, 게재지, 세부정보, 게재일자, 비고
                    row = [
                        now_str,
                        student_name,
                        student_id,
                        item["type"],
                        item["role"],           # 드롭다운 선택값
                        item["authors_all"],    # 전체 저자 텍스트
                        item["author_count"],   # 저자 수
                        item["title"],
                        item["journal"],
                        item["details"],
                        item["date"],
                        ""                      # 비고 (공란)
                    ]
                    rows_to_add.append(row)
                
                sheet.append_rows(rows_to_add)
                
            st.success("✅ 제출이 완료되었습니다! 수고하셨습니다.")
            
            # (선택) 제출 후 폼 초기화하고 싶으면 아래 주석 해제
            # st.session_state.research_items = []
            # st.rerun()
            
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
            st.warning("구글 시트 URL이 정확한지, 공유 설정(서비스 계정 이메일 추가)이 되었는지 확인해주세요.")