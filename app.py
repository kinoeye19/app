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
st.set_page_config(page_title="연구 성과 제출", page_icon="🎓", layout="wide")

st.title("🎓 연구 결과물 제출 시스템")
st.markdown("""
연구 성과를 입력하는 페이지입니다.  
항목명 옆에 **빨간색 별표(*)**가 있는 것은 **필수 입력 항목**입니다.  
물음표 아이콘(?)에 마우스를 올리면 상세 작성 요령을 볼 수 있습니다.
""")

if 'research_items' not in st.session_state:
    st.session_state.research_items = []

# A. 학생 기본 정보
with st.container():
    st.subheader("1. 기본 정보")
    col1, col2 = st.columns(2)
    with col1:
        student_name = st.text_input("이름 *", placeholder="홍길동")
    with col2:
        student_id = st.text_input("학번 *", placeholder="20241234")

st.divider()

# B. 성과 입력 로직
st.subheader("2. 연구 성과 입력")

def add_item():
    # 기본 템플릿 생성
    st.session_state.research_items.append({
        "type": "논문",  # 기본값
        # 공통 정보
        "class_name": "",
        "prof_name": "",
        # 논문 전용 필드 초기화
        "p_type_code": "국외전문학술지(01)", "p_journal": "", "p_title": "",
        "p_issn": "", "p_doi": "", "p_contrib": 0, "p_first_auth": "",
        "p_co_auth": "", "p_vol": "", "p_sci": "SCI/SSCI/A&HCI(01)",
        "p_page_start": "", "p_page_end": "", "p_impact": 0.0,
        "p_date": datetime.date.today(), "p_abstract": "",
        # 저서/학술대회 전용 필드 초기화
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
        # 삭제 버튼
        if st.button("🗑️ 이 항목 삭제", key=f"del_{i}"):
            remove_item(i)
            st.rerun()

        # 1. 성과 구분 선택
        type_options = ["논문", "저서", "학술대회 발표"]
        selected_type = st.selectbox(
            "성과 구분", 
            type_options, 
            key=f"type_{i}",
            index=type_options.index(item["type"]) if item["type"] in type_options else 0
        )
        st.session_state.research_items[i]["type"] = selected_type

        # -----------------------------------------------------
        # CASE 1: 논문 (학교 포맷 적용)
        # -----------------------------------------------------
        if selected_type == "논문":
            st.markdown("##### 📄 논문 상세 정보 (학교 제출 양식)")
            
            # Row 1: 구분
            c1, c2 = st.columns(2)
            with c1:
                p_type_code = st.selectbox("논문구분 *", ["국외전문학술지(01)", "국내전문학술지(03)"], key=f"p_type_{i}")
            with c2:
                p_sci = st.selectbox("SCI(E)구분 *", ["SCI/SSCI/A&HCI(01)", "비SCI(02)"], help="01: SCI급, 02: 비SCI(Scopus/KCI 등)", key=f"p_sci_{i}")

            # Row 2: 저널명, 논문명
            c1, c2 = st.columns(2)
            with c1:
                p_journal = st.text_input("학술지명(Full Name) *", placeholder="예: Nature", key=f"p_jour_{i}")
            with c2:
                p_title = st.text_input("논문명(Full Name) *", placeholder="학술지에 게재된 제목 그대로", key=f"p_tit_{i}")

            # Row 3: ISSN, DOI
            c1, c2 = st.columns(2)
            with c1:
                p_issn = st.text_input("ISSN *", placeholder="예: 1234-5678 (모르면 0000-0000)", help="하이픈(-) 포함 기재", key=f"p_issn_{i}")
            with c2:
                p_doi = st.text_input("DOI *", placeholder="예: 10.1038/xxx (모르면 0)", help="10.으로 시작", key=f"p_doi_{i}")

            # Row 4: 저자 정보
            c1, c2, c3 = st.columns([2, 1, 2])
            with c1:
                p_first_auth = st.text_input("주저자명(제1저자) *", placeholder="예: Gil-Dong Hong", help="영문 원칙, 저널에 한글이면 한글", key=f"p_fa_{i}")
            with c2:
                p_contrib = st.number_input("기여율(%)", min_value=0, max_value=100, value=item.get("p_contrib", 0), help="모르면 0", key=f"p_con_{i}")
            with c3:
                p_co_auth = st.text_input("공동저자명", placeholder="예: Cheol-Su Kim; Young-Hee Lee", help="2인 이상은 세미콜론(;) 구분", key=f"p_co_{i}")

            # Row 5: 게재 정보
            c1, c2, c3 = st.columns(3)
            with c1:
                p_vol = st.text_input("볼륨번호, 권(호) *", placeholder="예: 12(3)", help="권,호 단위 입력 금지. 모르면 N 입력", key=f"p_vol_{i}")
            with c2:
                p_page_start = st.text_input("시작 페이지 *", placeholder="예: 151 또는 A-10 (모르면 0)", key=f"p_ps_{i}")
            with c3:
                p_page_end = st.text_input("끝 페이지", placeholder="예: 157 (모르면 0)", key=f"p_pe_{i}")

            # Row 6: 날짜 및 IF
            c1, c2 = st.columns(2)
            with c1:
                # 날짜 입력받아 YYYYMMDD로 변환 준비
                p_date_pick = st.date_input("학술지 출판일자 *", value=item.get("p_date", datetime.date.today()), key=f"p_d_{i}")
            with c2:
                p_impact = st.number_input("임팩트팩터(IF)", format="%.5f", step=0.01, value=float(item.get("p_impact", 0.0)), help="최대 소수점 5자리", key=f"p_if_{i}")

            # 초록
            p_abstract = st.text_area("초록 *", placeholder="논문의 초록 내용을 붙여넣으세요.", height=100, key=f"p_abs_{i}")

            # 논문 데이터 업데이트
            st.session_state.research_items[i].update({
                "p_type_code": p_type_code, "p_sci": p_sci, "p_journal": p_journal,
                "p_title": p_title, "p_issn": p_issn, "p_doi": p_doi,
                "p_first_auth": p_first_auth, "p_contrib": int(p_contrib), "p_co_auth": p_co_auth,
                "p_vol": p_vol, "p_page_start": p_page_start, "p_page_end": p_page_end,
                "p_date": p_date_pick, "p_impact": p_impact, "p_abstract": p_abstract
            })

        # -----------------------------------------------------
        # CASE 2 & 3: 저서 / 학술대회 (기존 방식 유지)
        # -----------------------------------------------------
        else:
            st.markdown(f"##### 📘 {selected_type} 상세 정보")
            
            # 역할 선택
            role_options = []
            if selected_type == "저서":
                role_options = ["단독저자", "공동저자(챕터)", "공동저자(전체)", "대표저자"]
            else: 
                role_options = ["발표자", "공동연구자(발표안함)"]
            
            o_role = st.selectbox("참여 역할 *", role_options, key=f"o_r_{i}")

            # 저자 및 기본정보
            c1, c2 = st.columns([3, 1])
            with c1:
                o_authors_all = st.text_input("전체 저자/발표자 명단 *", placeholder="홍길동(본인), 김철수", key=f"o_aa_{i}")
            with c2:
                o_author_count = st.number_input("전체 인원 수", min_value=1, value=item.get("o_author_count", 1), key=f"o_ac_{i}")

            # 라벨링
            if selected_type == "저서":
                lbl_title, lbl_journal, lbl_detail = "저서명 *", "출판사 *", "ISBN / 개정판 정보"
            else:
                lbl_title, lbl_journal, lbl_detail = "발표 제목 *", "학술대회명 *", "개최 장소"

            o_title = st.text_input(lbl_title, key=f"o_t_{i}")
            
            cc1, cc2 = st.columns(2)
            with cc1:
                o_journal = st.text_input(lbl_journal, key=f"o_j_{i}")
            with cc2:
                o_details = st.text_input(lbl_detail, key=f"o_dt_{i}")
            
            o_date_pick = st.date_input("출판/발표 일자 *", value=item.get("o_date", datetime.date.today()), key=f"o_d_{i}")

            # 기타 데이터 업데이트
            st.session_state.research_items[i].update({
                "o_role": o_role, "o_authors_all": o_authors_all, "o_author_count": o_author_count,
                "o_title": o_title, "o_journal": o_journal, "o_details": o_details, "o_date": o_date_pick
            })

        # -----------------------------------------------------
        # 공통: 수업 연계 정보
        # -----------------------------------------------------
        st.markdown("---")
        st.info("💡 **연구성과물과 연계된 교과명 및 담당 교수자 정보를 입력해주세요.**")
        
        col_class, col_prof = st.columns(2)
        with col_class:
            class_name = st.text_input("연계 교과목명", placeholder="예: 디지털인문학", key=f"cl_{i}")
        with col_prof:
            prof_name = st.text_input("담당 교수", placeholder="예: 김철수 교수", key=f"pr_{i}")

        st.session_state.research_items[i].update({
            "class_name": class_name,
            "prof_name": prof_name
        })

st.divider()

# ---------------------------------------------------------
# 3. 제출 및 저장 로직 (유효성 검사 포함)
# ---------------------------------------------------------
if st.button("📤 제출하기", type="primary"):
    # 1. 기본정보 검사
    if not student_name or not student_id:
        st.error("❌ 맨 위의 [이름]과 [학번]을 반드시 입력해주세요.")
    elif len(st.session_state.research_items) == 0:
        st.warning("⚠️ 입력된 성과가 없습니다. '성과 추가하기' 버튼을 눌러주세요.")
    else:
        # 2. 상세 항목 유효성 검사 (필수값 체크)
        validation_error = False
        for idx, item in enumerate(st.session_state.research_items):
            missing_fields = []
            if item["type"] == "논문":
                # 논문 필수값 체크
                if not item["p_journal"]: missing_fields.append("학술지명")
                if not item["p_title"]: missing_fields.append("논문명")
                if not item["p_issn"]: missing_fields.append("ISSN")
                if not item["p_doi"]: missing_fields.append("DOI")
                if not item["p_first_auth"]: missing_fields.append("주저자명")
                if not item["p_vol"]: missing_fields.append("볼륨번호")
                if not item["p_page_start"]: missing_fields.append("시작페이지")
                if not item["p_abstract"]: missing_fields.append("초록")
            else:
                # 저서/학술대회 필수값 체크
                if not item["o_title"]: missing_fields.append("제목")
                if not item["o_journal"]: missing_fields.append("출판사/학술대회명")
                if not item["o_authors_all"]: missing_fields.append("저자/발표자 명단")

            if missing_fields:
                st.error(f"❌ [성과 #{idx+1} - {item['type']}] 필수 항목이 비어있습니다: {', '.join(missing_fields)}")
                validation_error = True

        # 3. 에러가 없을 때만 저장 진행
        if not validation_error:
            try:
                with st.spinner("구글 시트에 저장 중입니다..."):
                    client = get_connection()
                    
                    # *** [중요] 본인의 구글 시트 URL로 교체 필수 ***
                    SHEET_URL = "https://docs.google.com/spreadsheets/d/1nfE8lcFRsUfYkdV-tjpsZfFPWER0YeNR2TaxYLH32JY/edit?usp=sharing"
                    doc = client.open_by_url(SHEET_URL)

                    rows_paper = []
                    rows_book = []
                    rows_conf = []
                    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    for item in st.session_state.research_items:
                        # 공통값
                        common_front = [now_str, student_name, student_id]
                        common_back = [item["class_name"], item["prof_name"], ""] # 비고 포함

                        if item["type"] == "논문":
                            # 논문 전용 코드값 변환 (예: '국외전문학술지(01)' -> '01')
                            t_code = "01" if "01" in item["p_type_code"] else "03"
                            s_code = "01" if "01" in item["p_sci"] else "02"
                            date_yyyymmdd = item["p_date"].strftime("%Y%m%d") # 학교 요구 포맷

                            row = common_front + [
                                t_code, item["p_journal"], item["p_title"], item["p_issn"], item["p_doi"],
                                item["p_contrib"], item["p_first_auth"], item["p_co_auth"], item["p_vol"],
                                s_code, item["p_page_start"], item["p_page_end"], item["p_impact"],
                                date_yyyymmdd, item["p_abstract"]
                            ] + common_back
                            rows_paper.append(row)

                        else: # 저서 또는 학술대회
                            # 기존 포맷 유지 (YYYY-MM-DD)
                            date_std = item["o_date"].strftime("%Y-%m-%d")
                            row = common_front + [
                                item["type"], item["o_role"], item["o_authors_all"], item["o_author_count"],
                                item["o_title"], item["o_journal"], item["o_details"], date_std
                            ] + common_back
                            
                            if item["type"] == "저서":
                                rows_book.append(row)
                            else:
                                rows_conf.append(row)

                    # 시트 저장
                    if rows_paper: doc.worksheet("논문").append_rows(rows_paper)
                    if rows_book: doc.worksheet("저서").append_rows(rows_book)
                    if rows_conf: doc.worksheet("학술대회").append_rows(rows_conf)

                st.success("✅ 모든 데이터가 성공적으로 제출되었습니다!")
                st.session_state.research_items = []
                st.rerun()

            except gspread.WorksheetNotFound:
                st.error("오류: 구글 시트 탭 이름이 '논문', '저서', '학술대회'인지 확인해주세요.")
            except Exception as e:
                st.error(f"저장 중 오류 발생: {e}")