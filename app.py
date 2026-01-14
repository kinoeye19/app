import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import datetime

# ---------------------------------------------------------
# 1. Google Sheets 인증 설정
# ---------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_connection():
    if "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        try:
            creds = Credentials.from_service_account_file("service_account.json", scopes=SCOPES)
        except FileNotFoundError:
            st.error("로컬 인증 파일(service_account.json)을 찾을 수 없습니다.")
            st.stop()
    
    client = gspread.authorize(creds)
    return client

# ---------------------------------------------------------
# 2. 화면 구성
# ---------------------------------------------------------
st.set_page_config(page_title="연구 성과 제출", page_icon="🎓")

st.title("🎓 연구 결과물 제출 시스템")
st.markdown("연구 성과(논문, 저서, 학술대회)를 **유형별로 분류하여** 입력해 주세요.")

if 'research_items' not in st.session_state:
    st.session_state.research_items = []

# A. 학생 기본 정보
with st.container():
    st.subheader("1. 기본 정보")
    col1, col2 = st.columns(2)
    with col1:
        student_name = st.text_input("이름", placeholder="홍길동")
    with col2:
        student_id = st.text_input("학번", placeholder="20241234")

st.divider()

# B. 성과 입력
st.subheader("2. 연구 성과 입력")

def add_item():
    st.session_state.research_items.append({
        "type": "논문",
        "role": "",
        "authors_all": "",
        "author_count": 1,
        "title": "",
        "journal": "",   
        "details": "",   
        "date": datetime.date.today() # 기본값은 오늘 날짜
    })

def remove_item(index):
    st.session_state.research_items.pop(index)

if st.button("➕ 성과 추가하기"):
    add_item()

for i, item in enumerate(st.session_state.research_items):
    with st.expander(f"📝 성과 #{i+1} 입력", expanded=True):
        if st.button("삭제", key=f"del_{i}"):
            remove_item(i)
            st.rerun()

        # 구분 선택
        type_options = ["논문", "저서", "학술대회 발표"]
        selected_type = st.selectbox(
            "구분 (선택하면 시트가 자동으로 분류됩니다)", 
            type_options, 
            key=f"type_{i}",
            index=type_options.index(item["type"]) if isinstance(item["type"], str) and item["type"] in type_options else 0
        )
        
        # 역할 선택
        role_options = []
        if selected_type == "논문":
            role_options = ["단독", "제1저자", "공동저자", "교신저자"]
        elif selected_type == "저서":
            role_options = ["단독저자", "공동저자(챕터)", "공동저자(전체)", "대표저자"]
        else: 
            role_options = ["발표자", "공동연구자(발표안함)"]
        
        selected_role = st.selectbox("참여 역할", role_options, key=f"role_{i}")

        # 저자 정보
        c1, c2 = st.columns([3, 1])
        with c1:
            authors_all = st.text_input("전체 저자/발표자 명단", placeholder="홍길동(본인), 김철수", key=f"auth_all_{i}")
        with c2:
            author_count = st.number_input("전체 인원 수", min_value=1, value=item.get("author_count", 1), key=f"auth_cnt_{i}")

        # 상세 정보 라벨링
        if selected_type == "논문":
            lbl_title = "논문 제목"
            lbl_journal = "저널명 (Journal)"
            lbl_detail = "권호 (Vol, No)"
            lbl_date = "게재일자 (YYYY-MM-DD)"
        elif selected_type == "저서":
            lbl_title = "저서명 (Book Title)"
            lbl_journal = "출판사"
            lbl_detail = "ISBN / 개정판 정보"
            lbl_date = "출판일자 (YYYY-MM-DD)"
        else:
            lbl_title = "발표 제목"
            lbl_journal = "학술대회명"
            lbl_detail = "개최 장소"
            lbl_date = "발표일자 (YYYY-MM-DD)"

        title = st.text_input(lbl_title, key=f"title_{i}")
        cc1, cc2 = st.columns(2)
        with cc1:
            journal = st.text_input(lbl_journal, key=f"journal_{i}")
        with cc2:
            details = st.text_input(lbl_detail, key=f"detail_{i}")
        
        # [변경점] 텍스트 입력 대신 날짜 선택기(Calendar) 사용
        # 입력받은 값을 YYYY-MM-DD 문자열로 변환하여 저장
        date_pick = st.date_input(
            lbl_date, 
            value=datetime.date.today(), # 기본값
            key=f"date_pick_{i}",
            help="달력을 클릭하여 날짜를 선택하세요."
        )
        date_str = date_pick.strftime("%Y-%m-%d") # 구글 시트가 좋아하는 형식으로 변환

        st.session_state.research_items[i].update({
            "type": selected_type,
            "role": selected_role,
            "authors_all": authors_all,
            "author_count": author_count,
            "title": title,
            "journal": journal,
            "details": details,
            "date": date_str
        })

st.divider()

# C. 제출 로직
if st.button("📤 제출하기", type="primary"):
    if not student_name or not student_id:
        st.error("이름과 학번을 입력해주세요.")
    elif len(st.session_state.research_items) == 0:
        st.warning("입력된 성과가 없습니다.")
    else:
        try:
            with st.spinner("구글 시트에 저장 중..."):
                client = get_connection()
                
                # *** [중요] 여기에 선생님의 실제 구글 시트 URL을 넣어주세요 ***
                # 아까 에러의 원인이 이 부분이 수정되지 않아서일 확률이 높습니다!
                SHEET_URL = "https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing" 
                
                doc = client.open_by_url(SHEET_URL)

                # 유형별 데이터 리스트
                rows_paper = []
                rows_book = []
                rows_conf = []
                
                now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for item in st.session_state.research_items:
                    row = [
                        now_str,
                        student_name,
                        student_id,
                        item["type"],
                        item["role"],
                        item["authors_all"],
                        item["author_count"],
                        item["title"],
                        item["journal"],
                        item["details"],
                        item["date"], # YYYY-MM-DD 형식의 문자열
                        ""
                    ]
                    
                    if item["type"] == "논문":
                        rows_paper.append(row)
                    elif item["type"] == "저서":
                        rows_book.append(row)
                    else: # 학술대회 발표
                        rows_conf.append(row)

                # 각 시트에 저장
                if rows_paper:
                    doc.worksheet("논문").append_rows(rows_paper)
                if rows_book:
                    doc.worksheet("저서").append_rows(rows_book)
                if rows_conf:
                    doc.worksheet("학술대회").append_rows(rows_conf)

            st.success("✅ 제출 완료! 날짜 형식이 정확하게(YYYY-MM-DD) 저장되었습니다.")
            st.session_state.research_items = []
            st.rerun()

        except gspread.WorksheetNotFound:
            st.error("오류: 구글 시트에 '논문', '저서', '학술대회' 탭이 만들어져 있는지 확인해주세요.")
        except Exception as e:
            st.error(f"저장 중 오류 발생: {e}")
            st.warning("팁: 코드 안의 'SHEET_URL'에 실제 구글 시트 주소를 넣으셨는지 꼭 확인해주세요!")