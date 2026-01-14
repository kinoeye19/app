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
# 2. 화면 구성 및 스타일
# ---------------------------------------------------------
st.set_page_config(page_title="연구 성과 제출", page_icon="🎓", layout="centered")

st.title("🎓 연구 결과물 제출 시스템")
st.markdown("""
연구 성과를 입력하는 페이지입니다.  
모든 항목(이름 포함)은 특별한 사유가 없다면 **한글로 입력**해 주세요.  
항목명 옆에 **빨간색 별표(*)**가 있는 것은 **필수 입력 항목**입니다.
""")

if 'research_items' not in st.session_state:
    st.session_state.research_items = []

# A. 학생 기본 정보
with st.container():
    st.subheader("1. 기본 정보")
    col1, col2 = st.columns(2)
    with col1:
        # [수정] 한글 입력 유도
        student_name = st.text_input("이름 *", placeholder="예: 홍길동 (한글 입력)")
    with col2:
        student_id = st.text_input("학번 *", placeholder="20241234")

st.divider()

# B. 성과 입력 로직
st.subheader("2. 연구 성과 입력")

def add_item():
    st.session_state.research_items.append({
        "type": "논문",
        "class_name": "", "prof_name": "",
        "p_type_code": "국외전문학술지(01)", "p_sci": "SCI/SSCI/A&HCI(01)", 
        "p_journal": "", "p_title": "", "p_issn": "", "p_doi": "", 
        "p_first_auth": "", "p_contrib": 0, "p_co_auth": "", 
        "p_vol": "", "p_page_start": "", "p_page_end": "", 
        "p_impact": 0.0, "p_date": datetime.date.today(), "p_abstract": "",
        "o_role": "", "o_authors_all": "", "o_author_count": 1,
        "o_title": "", "o_journal": "", "o_details": "", "o_date": datetime.date.today()
    })

def remove_item(index):
    st.session_state.research_items.pop(index)

if st.button("➕ 성과 추가하기"):
    add_item()

# ---------------------------------------------------------
# 입력 폼 생성 루프
# ---------------------------------------------------------
for i, item in enumerate(st.session_state.research_items):
    with st.expander(f"📝 성과 #{i+1} 입력 (클릭하여 열기/접기)", expanded=True):
        if st.button("🗑️ 이 항목 삭제", key=f"del_{i}"):
            remove_item(i)
            st.rerun()

        # 1. 성과 구분
        type_options = ["논문", "저서", "학술대회 발표"]
        selected_type = st.selectbox(
            "성과 구분", 
            type_options, 
            key=f"type_{i}",
            index=type_options.index(item["type"]) if item["type"] in type_options else 0
        )
        st.session_state.research_items[i]["type"] = selected_type

        # -----------------------------------------------------
        # CASE 1: 논문
        # -----------------------------------------------------
        if selected_type == "논문":
            st.markdown("##### 📄 논문 상세 정보")
            
            c1, c2 = st.columns(2)
            with c1:
                p_type_code = st.selectbox("논문구분 *", ["국외전문학술지(01)", "국내전문학술지(03)"], key=f"p_type_{i}")
            with c2:
                p_sci = st.selectbox("SCI(E)구분 *", ["SCI/SSCI/A&HCI(01)", "비SCI(02)"], help="01: SCI급, 02: 비SCI", key=f"p_sci_{i}")

            p_journal = st.text_input("학술지명(Full Name) *", placeholder="예: 네이처 (Nature)", help="가능하면 정식 명칭 기재", key=f"p_jour_{i}")
            p_title = st.text_input("논문명(Full Name) *", placeholder="학술지에 게재된 제목 그대로", key=f"p_tit_{i}")

            c1, c2 = st.columns(2)
            with c1:
                p_issn = st.text_input("ISSN *", placeholder="1234-5678", help="하이픈(-) 포함. 모르면 0000-0000", key=f"p_issn_{i}")
            with c2:
                p_doi = st.text_input("DOI *", placeholder="10.1038/xxx", help="10.으로 시작. 모르면 0", key=f"p_doi_{i}")

            # 저자 정보 (한글 입력 유도)
            c1, c2 = st.columns([2, 1])
            with c1:
                # [수정] 한글 이름 예시 제공
                p_first_auth = st.text_input("주저자명(제1저자) *", placeholder="예: 홍길동 (한글 권장)", help="한글 입력을 원칙으로 합니다.", key=f"p_fa_{i}")
            with c2:
                p_contrib = st.number_input("기여율(%)", min_value=0, max_value=100, value=item.get("p_contrib", 0), key=f"p_con_{i}")
            
            # [수정] 한글 이름 예시 제공
            p_co_auth = st.text_input("공동저자명", placeholder="예: 김철수; 이영희", help="2인 이상은 세미콜론(;)으로 구분하여 한글로 기재", key=f"p_co_{i}")

            # 게재 정보
            c1, c2 = st.columns(2)
            with c1:
                p_vol = st.text_input("볼륨번호, 권(호) *", placeholder="예: 12(3)", help="단위 없이 숫자만. 모르면 N", key=f"p_vol_{i}")
            with c2:
                p_impact = st.number_input("임팩트팩터(IF)", format="%.5f", step=0.01, value=float(item.get("p_impact", 0.0)), key=f"p_if_{i}")

            c1, c2 = st.columns(2)
            with c1:
                p_page_start = st.text_input("시작 페이지 *", placeholder="예: 151 (모르면 0)", key=f"p_ps_{i}")
            with c2:
                p_page_end = st.text_input("끝 페이지", placeholder="예: 157 (모르면 0)", key=f"p_pe_{i}")

            p_date_pick = st.date_input("학술지 출판일자 *", value=item.get("p_date", datetime.date.today()), key=f"p_d_{i}")
            p_abstract = st.text_area("초록 *", placeholder="논문의 초록 내용을 입력하세요.", height=100, key=f"p_abs_{i}")

            st.session_state.research_items[i].update({
                "p_type_code": p_type_code, "p_sci": p_sci, "p_journal": p_journal,
                "p_title": p_title, "p_issn": p_issn, "p_doi": p_doi,
                "p_first_auth": p_first_auth, "p_contrib": int(p_contrib), "p_co_auth": p_co_auth,
                "p_vol": p_vol, "p_page_start": p_page_start, "p_page_end": p_page_end,
                "p_date": p_date_pick, "p_impact": p_impact, "p_abstract": p_abstract
            })

        # -----------------------------------------------------
        # CASE 2 & 3: 저서 / 학술대회
        # -----------------------------------------------------
        else:
            st.markdown(f"##### 📘 {selected_type} 상세 정보")
            role_options = ["단독저자", "공동저자(챕터)", "공동저자(전체)", "대표저자"] if selected_type == "저서" else ["발표자", "공동연구자(발표안함)"]
            o_role = st.selectbox("참여 역할 *", role_options, key=f"o_r_{i}")

            # [수정] 한글 이름 예시
            o_authors_all = st.text_input("전체 저자/발표자 명단 *", placeholder="예: 홍길동, 김철수 (한글 입력)", key=f"o_aa_{i}")
            o_author_count = st.number_input("전체 인원 수", min_value=1, value=item.get("o_author_count", 1), key=f"o_ac_{i}")

            lbl_title = "저서명 *" if selected_type == "저서" else "발표 제목 *"
            lbl_journal = "출판사 *" if selected_type == "저서" else "학술대회명 *"
            lbl_detail = "ISBN / 개정판 정보" if selected_type == "저서" else "개최 장소"

            o_title = st.text_input(lbl_title, key=f"o_t_{i}")
            o_journal = st.text_input(lbl_journal, key=f"o_j_{i}")
            o_details = st.text_input(lbl_detail, key=f"o_dt_{i}")
            o_date_pick = st.date_input("출판/발표 일자 *", value=item.get("o_date", datetime.date.today()), key=f"o_d_{i}")

            st.session_state.research_items[i].update({
                "o_role": o_role, "o_authors_all": o_authors_all, "o_author_count": o_author_count,
                "o_title": o_title, "o_journal": o_journal, "o_details": o_details, "o_date": o_date_pick
            })

        # -----------------------------------------------------
        # 공통: 수업 연계 정보
        # -----------------------------------------------------
        st.markdown("---")
        st.info("💡 **연구성과물과 연계된 교과명 및 담당 교수자 정보를 입력해주세요.**")
        
        c1, c2 = st.columns(2)
        with c1:
            class_name = st.text_input("연계 교과목명", placeholder="예: 디지털인문학", key=f"cl_{i}")
        with c2:
            prof_name = st.text_input("담당 교수", placeholder="예: 김철수 교수", key=f"pr_{i}")

        st.session_state.research_items[i].update({
            "class_name": class_name,
            "prof_name": prof_name
        })

st.divider()

# ---------------------------------------------------------
# 3. 제출 및 저장 로직 (필수항목 차단 기능 포함)
# ---------------------------------------------------------
if st.button("📤 제출하기", type="primary"):
    # 1. 기본정보 검사
    if not student_name or not student_id:
        st.error("❌ 맨 위의 [이름]과 [학번]을 반드시 입력해주세요.")
    elif len(st.session_state.research_items) == 0:
        st.warning("⚠️ 입력된 성과가 없습니다. [성과 추가하기]를 눌러 내용을 작성해주세요.")
    else:
        # 2. 필수항목 유효성 검사 (하나라도 비면 validation_error = True)
        validation_error = False
        for idx, item in enumerate(st.session_state.research_items):
            missing = []
            if item["type"] == "논문":
                # 논문 필수값 리스트
                if not item["p_journal"]: missing.append("학술지명")
                if not item["p_title"]: missing.append("논문명")
                if not item["p_issn"]: missing.append("ISSN")
                if not item["p_doi"]: missing.append("DOI")
                if not item["p_first_auth"]: missing.append("주저자명")
                if not item["p_vol"]: missing.append("볼륨번호")
                if not item["p_page_start"]: missing.append("시작페이지")
                if not item["p_abstract"]: missing.append("초록")
            else:
                # 저서/학술대회 필수값 리스트
                if not item["o_title"]: missing.append("제목")
                if not item["o_journal"]: missing.append("출판사/학술대회명")
                if not item["o_authors_all"]: missing.append("저자/발표자 명단")

            # 누락된 항목이 있으면 에러 메시지 출력
            if missing:
                st.error(f"❌ [성과 #{idx+1} - {item['type']}] 다음 필수 항목이 비어있습니다: {', '.join(missing)}")
                validation_error = True

        # 3. 에러가 하나도 없을 때만 저장 실행 (차단 장치)
        if not validation_error:
            try:
                with st.spinner("구글 시트에 저장 중입니다..."):
                    client = get_connection()
                    
                    # [주소 반영]
                    SHEET_URL = "https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing"
                    doc = client.open_by_url(SHEET_URL)

                    rows_paper = []
                    rows_book = []
                    rows_conf = []
                    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    for item in st.session_state.research_items:
                        common_front = [now_str, student_name, student_id]
                        common_back = [item["class_name"], item["prof_name"], ""]

                        if item["type"] == "논문":
                            t_code = "01" if "01" in item["p_type_code"] else "03"
                            s_code = "01" if "01" in item["p_sci"] else "02"
                            date_str = item["p_date"].strftime("%Y%m%d")

                            row = common_front + [
                                t_code, item["p_journal"], item["p_title"], item["p_issn"], item["p_doi"],
                                item["p_contrib"], item["p_first_auth"], item["p_co_auth"], item["p_vol"],
                                s_code, item["p_page_start"], item["p_page_end"], item["p_impact"],
                                date_str, item["p_abstract"]
                            ] + common_back
                            rows_paper.append(row)
                        else:
                            date_std = item["o_date"].strftime("%Y-%m-%d")
                            row = common_front + [
                                item["type"], item["o_role"], item["o_authors_all"], item["o_author_count"],
                                item["o_title"], item["o_journal"], item["o_details"], date_std
                            ] + common_back
                            
                            if item["type"] == "저서":
                                rows_book.append(row)
                            else:
                                rows_conf.append(row)

                    if rows_paper: doc.worksheet("논문").append_rows(rows_paper)
                    if rows_book: doc.worksheet("저서").append_rows(rows_book)
                    if rows_conf: doc.worksheet("학술대회").append_rows(rows_conf)

                st.success("✅ 제출이 성공적으로 완료되었습니다!")
                st.session_state.research_items = []
                st.rerun()

            except gspread.WorksheetNotFound:
                st.error("오류: 구글 시트 탭 이름('논문', '저서', '학술대회')을 확인해주세요.")
            except Exception as e:
                st.error(f"저장 중 오류 발생: {e}")