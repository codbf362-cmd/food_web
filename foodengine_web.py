import os
import base64
import mimetypes
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# 이메일 전송용
import smtplib
from email.mime.text import MIMEText

# (선택) 자동완성 컴포넌트
try:
    from streamlit_searchbox import st_searchbox
    HAS_SEARCHBOX = True
except Exception:
    HAS_SEARCHBOX = False

# ============================
# 기본 경로 설정 (웹 / GitHub 용)
# ============================
# 이 파일이 있는 폴더 기준
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 엑셀 데이터는 ./data 폴더 안에 넣기
DATA_DIR = os.path.join(BASE_DIR, "data")
FOOD_FILE = os.path.join(DATA_DIR, "food_insert1.xlsx")
DISEASE_FILE = os.path.join(DATA_DIR, "disease_insert.xlsx")
DISEASE_EXPLAIN_FILE = os.path.join(DATA_DIR, "disease_explanation.xlsx")

# 배경 이미지는 프로젝트 루트에 있는 foodphoto.png 로 사용
BG_IMAGE_FILE = os.path.join(BASE_DIR, "foodphoto.png")

FOOD_NAME_COL = "식품명"
STATE1_COL = "상태 1"
STATE2_COL = "상태2"
CATEGORY_COL = "카테고리"
DISEASE_NAME_COL = "질병명"

# 기본 영양소 컬럼(음식, 질병 공통)
NUTRIENT_COLS = [
    "에너지(kcal)",
    "단백질(g)",
    "지방(g)",
    "탄수화물(g)",
    "식이섬유(g)",
    "당류(g)",
    "나트륨(mg)",
    "칼륨(mg)",
    "인(mg)",
    "포화지방산(g)",
    "콜레스테롤(mg)",
]

# 질병 엑셀에 들어있는 "주의" 컷오프 컬럼 매핑
WARNING_COL_MAP = {
    "단백질(g)": "단백질주의(g)",
    "지방(g)": "지방주의(g)",
    "탄수화물(g)": "탄수화물주의(g)",
    "당류(g)": "당류주의(g)",
    "나트륨(mg)": "나트륨주의(mg)",
    "칼륨(mg)": "칼륨주의(mg)",
    "인(mg)": "인주의(mg)",
    "포화지방산(g)": "포화지방산주의(g)",
    "콜레스테롤(mg)": "콜레스테롤주의(mg)",
}
WARNING_NUMERIC_COLS = list(set(WARNING_COL_MAP.values()))

# ============================
# 이메일 SMTP 설정 (웹용: st.secrets 사용)
# ============================
# Streamlit Cloud / 로컬 둘 다에서:
# .streamlit/secrets.toml 에 [email] 섹션 만들어서 값 넣기
try:
    email_conf = st.secrets["email"]
except Exception:
    # secrets 없으면 빈 설정으로 두고, 전송 시 에러 메시지 반환
    email_conf = {}

SMTP_SERVER = email_conf.get("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(email_conf.get("SMTP_PORT", 587))
SMTP_USER = email_conf.get("SMTP_USER", "")
SMTP_PASSWORD = email_conf.get("SMTP_PASSWORD", "")
CONTACT_RECEIVER = email_conf.get("CONTACT_RECEIVER", SMTP_USER)


def send_contact_email(user_email: str, user_msg: str) -> tuple[bool, str]:
    """
    문의하기 폼에서 입력받은 내용을 실제 이메일로 전송.
    return: (성공여부, 에러메시지)
    """
    # SMTP 설정이 비어 있으면 전송 안 하고 안내만
    if not SMTP_USER or not SMTP_PASSWORD or not CONTACT_RECEIVER:
        return False, "SMTP 설정이 비어 있습니다. secrets.toml의 [email] 값을 확인하세요."

    subject = "[질병별 음식 판정] 문의가 도착했습니다"
    time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    body = (
        f"질병별 음식 판정 시스템에서 새로운 문의가 접수되었습니다.\n\n"
        f"시간: {time_str}\n"
        f"보낸 사람 이메일: {user_email}\n\n"
        f"문의 내용:\n"
        f"{user_msg}\n"
    )

    msg = MIMEText(body, _charset="utf-8")
    msg["Subject"] = subject
    msg["From"] = SMTP_USER
    msg["To"] = CONTACT_RECEIVER

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)

        # 콘솔에도 같이 찍어줌(디버그용)
        print("===== 문의하기 이메일 전송 완료 =====")
        print("To:", CONTACT_RECEIVER)
        print("From:", SMTP_USER)
        print("User Email:", user_email)
        print("Message:")
        print(user_msg)
        print("===================================")

        return True, ""
    except Exception as e:
        # 실패 시 콘솔 로그
        print("===== 문의하기 이메일 전송 실패 =====")
        print("에러:", e)
        print("User Email:", user_email)
        print("Message:")
        print(user_msg)
        print("===================================")
        return False, str(e)


# ============================
# 페이지 설정
# ============================
st.set_page_config(page_title="질병별 음식 판정", layout="wide")


# ============================
# 유틸: 숫자형 변환
# ============================
def coerce_numeric(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = (
                out[c]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.replace(r"[^0-9\-\.\+eE]", "", regex=True)
            )
            out[c] = pd.to_numeric(out[c], errors="coerce")
    return out


# ============================
# 유틸: 이미지 → base64
# ============================
def get_base64_image(image_path: str):
    try:
        with open(image_path, "rb") as img:
            data = img.read()
        encoded = base64.b64encode(data).decode()
        mime = mimetypes.guess_type(image_path)[0] or "image/png"
        return encoded, mime
    except Exception:
        return None, None


# ============================
# 배경 이미지 CSS 적용
# ============================
bg64, bg_mime = get_base64_image(BG_IMAGE_FILE)
if bg64:
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url("data:{bg_mime};base64,{bg64}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )
else:
    st.warning(f"배경 이미지 파일을 불러올 수 없습니다: {BG_IMAGE_FILE}")


# ============================
# 데이터 로드
# ============================
@st.cache_data
def load_data(food_path: str, disease_path: str, explain_path: str):
    foods = pd.read_excel(food_path)
    diseases = pd.read_excel(disease_path)

    # 질병 설명 (없어도 동작)
    try:
        disease_expl = pd.read_excel(explain_path)
    except Exception:
        disease_expl = pd.DataFrame(columns=[DISEASE_NAME_COL, "설명"])

    # 음식 문자열 정리
    for c in [FOOD_NAME_COL, STATE1_COL, STATE2_COL, CATEGORY_COL]:
        if c in foods.columns:
            foods[c] = (
                foods[c]
                .astype("string")
                .str.strip()
                .replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})
            )

    # 질병 컷오프 문자열 정리
    if DISEASE_NAME_COL in diseases.columns:
        diseases[DISEASE_NAME_COL] = (
            diseases[DISEASE_NAME_COL]
            .astype("string")
            .str.strip()
            .replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})
        )

    # 질병 설명 문자열 정리
    if DISEASE_NAME_COL in disease_expl.columns:
        disease_expl[DISEASE_NAME_COL] = (
            disease_expl[DISEASE_NAME_COL]
            .astype("string")
            .str.strip()
            .replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})
        )
    if "설명" in disease_expl.columns:
        disease_expl["설명"] = disease_expl["설명"].astype("string").str.strip()

    # 숫자형 변환
    foods = coerce_numeric(foods, NUTRIENT_COLS)
    diseases = coerce_numeric(diseases, NUTRIENT_COLS + WARNING_NUMERIC_COLS)

    return foods, diseases, disease_expl


# ============================
# 판정 함수 (단일 질병 기준, 주의 컷오프 포함)
# ============================
def evaluate_row(food_row: pd.Series, disease_row: pd.Series):
    violations_fail = []
    violations_warn = []
    checked = []

    for col in NUTRIENT_COLS:
        if col in food_row.index and col in disease_row.index:
            cutoff = disease_row[col]
            val = food_row[col]

            if pd.notna(cutoff) and pd.notna(val):
                checked.append(col)

                warn_col = WARNING_COL_MAP.get(col)
                warn_cut = np.nan
                if warn_col and warn_col in disease_row.index:
                    warn_cut = disease_row[warn_col]

                status = "ok"

                # 경고 컷오프 있는 경우
                if pd.notna(warn_cut):
                    if val > cutoff:
                        status = "fail"
                    elif val >= warn_cut:
                        status = "warn"
                    else:
                        status = "ok"
                else:
                    # 경고 컬럼 없으면 단일 컷오프
                    if val > cutoff:
                        status = "fail"
                    else:
                        status = "ok"

                if status == "fail":
                    violations_fail.append(
                        (col, float(val), float(cutoff), float(val - cutoff))
                    )
                elif status == "warn":
                    violations_warn.append(
                        (col, float(val), float(warn_cut), float(val - warn_cut))
                    )

    has_fail = len(violations_fail) > 0
    has_warning = len(violations_warn) > 0

    if has_fail:
        overall_status = "불합격"
    elif has_warning:
        overall_status = "주의"
    else:
        overall_status = "합격"

    return {
        "checked_cols": checked,
        "violations_fail": violations_fail,
        "violations_warn": violations_warn,
        "has_fail": has_fail,
        "has_warning": has_warning,
        "overall_status": overall_status,
    }


# ============================
# AI 느낌: 영양소 벡터 거리 계산
# ============================
def nutrient_distance(
    original_row: pd.Series,
    candidate_row: pd.Series,
    disease_row: pd.Series,
) -> float:
    """
    영양소 벡터 간 거리 계산.
    - 각 영양소 값 / 질병 컷오프 값 으로 정규화해서 비교
    - 둘 다 값이 있는 영양소만 사용
    - 유클리드 거리 사용
    """
    sq_diffs = []

    for col in NUTRIENT_COLS:
        fo = original_row.get(col, np.nan)
        fc = candidate_row.get(col, np.nan)

        if pd.isna(fo) or pd.isna(fc):
            continue

        cutoff = disease_row.get(col, np.nan)
        if pd.notna(cutoff) and cutoff > 0:
            no = fo / cutoff
            nc = fc / cutoff
        else:
            no = fo
            nc = fc

        diff = float(no - nc)
        sq_diffs.append(diff * diff)

    if not sq_diffs:
        return float("inf")

    return float(np.sqrt(np.mean(sq_diffs)))


# ============================
# 상태 우선순위 (기본/무상태 우선)
# ============================
def state_preference(row: pd.Series) -> tuple[int, int]:
    """
    작은 값일수록 더 우선.
    - 상태1: "" 또는 "기본"이면 0, 나머지는 1
    - 상태2: "" 이면 0, 나머지는 1
    """
    s1 = row.get(STATE1_COL)
    s2 = row.get(STATE2_COL)

    s1_str = "" if pd.isna(s1) else str(s1).strip()
    s2_str = "" if pd.isna(s2) else str(s2).strip()

    s1_flag = 0 if (s1_str == "" or s1_str == "기본") else 1
    s2_flag = 0 if s2_str == "" else 1

    return (s1_flag, s2_flag)


# ============================
# 공통: 후보 집합에서 추천 리스트 만들기
# ============================
def _build_recommendations_from_candidates(
    cand: pd.DataFrame,
    foods: pd.DataFrame,
    disease_row: pd.Series,
    original_row: pd.Series,
    max_rec: int,
) -> pd.DataFrame:
    if cand.empty:
        return pd.DataFrame()

    # 원래 선택한 음식(이름+상태1+상태2) 제거
    def is_same_food(row):
        return (
            str(row.get(FOOD_NAME_COL, "")) == str(original_row.get(FOOD_NAME_COL, ""))
            and str(row.get(STATE1_COL, "")) == str(original_row.get(STATE1_COL, ""))
            and str(row.get(STATE2_COL, "")) == str(original_row.get(STATE2_COL, ""))
        )

    cand = cand[~cand.apply(is_same_food, axis=1)]
    if cand.empty:
        return pd.DataFrame()

    pass_dict: dict[str, dict] = {}
    warn_dict: dict[str, dict] = {}

    for _, row in cand.iterrows():
        res = evaluate_row(row, disease_row)
        status = res["overall_status"]
        if status == "불합격":
            continue

        dist = nutrient_distance(original_row, row, disease_row)
        if not np.isfinite(dist):
            continue

        base_name = str(row.get(FOOD_NAME_COL, ""))
        pref = state_preference(row)

        if status == "합격":
            target_dict = pass_dict
        elif status == "주의":
            target_dict = warn_dict
        else:
            continue

        current = target_dict.get(base_name)
        if current is None:
            target_dict[base_name] = {"dist": dist, "row": row, "pref": pref}
        else:
            if pref < current["pref"] or (pref == current["pref"] and dist < current["dist"]):
                target_dict[base_name] = {"dist": dist, "row": row, "pref": pref}

    if not pass_dict and not warn_dict:
        return pd.DataFrame()

    pass_list = [(v["dist"], v["row"]) for v in pass_dict.values()]
    warn_list = [(v["dist"], v["row"]) for v in warn_dict.values()]

    pass_list.sort(key=lambda x: x[0])
    warn_list.sort(key=lambda x: x[0])

    selected_rows: list[tuple[str, float, pd.Series]] = []

    # 1) 합격 우선
    for dist, row in pass_list:
        if len(selected_rows) >= max_rec:
            break
        selected_rows.append(("합격", dist, row))

    # 2) 부족하면 주의로 채우기
    if len(selected_rows) < max_rec:
        for dist, row in warn_list:
            if len(selected_rows) >= max_rec:
                break
            selected_rows.append(("주의", dist, row))

    if not selected_rows:
        return pd.DataFrame()

    out_rows = []
    for status, dist, row in selected_rows:
        name = str(row.get(FOOD_NAME_COL, ""))
        s1 = row.get(STATE1_COL)
        s2 = row.get(STATE2_COL)
        label = name
        if pd.notna(s1) and str(s1).strip() != "":
            label += f" / {s1}"
        else:
            label += " / 기본"
        if pd.notna(s2) and str(s2).strip() != "":
            label += f" / {s2}"

        out_rows.append(
            {
                "추천 음식": label,
                "카테고리": row.get(CATEGORY_COL),
                "판정": status,
                "유사도거리": round(dist, 4),
            }
        )

    return pd.DataFrame(out_rows)


# ============================
# AI 추천: 카테고리 + fallback(전체)
# ============================
def recommend_alternatives_with_fallback(
    foods: pd.DataFrame,
    disease_row: pd.Series,
    original_row: pd.Series,
    max_rec: int = 4,
) -> pd.DataFrame:
    orig_cat = original_row.get(CATEGORY_COL)

    # 1단계: 같은 카테고리
    if CATEGORY_COL in foods.columns and pd.notna(orig_cat):
        cand_same = foods[foods[CATEGORY_COL] == orig_cat].copy()
        df_same = _build_recommendations_from_candidates(
            cand_same, foods, disease_row, original_row, max_rec
        )
        if not df_same.empty:
            return df_same

    # 2단계: 전체 음식에서 fallback
    cand_all = foods.copy()
    df_all = _build_recommendations_from_candidates(
        cand_all, foods, disease_row, original_row, max_rec
    )
    return df_all


# ============================
# 폴백 자동완성 (searchbox 없을 때)
# ============================
def fallback_live_search(
    label: str, key_text: str, choices: list[str], max_suggest: int = 15
):
    typed = st.text_input(label, key=key_text, placeholder="검색어 입력…")
    typed_norm = (typed or "").strip().lower()
    sel_key = f"{key_text}__selected"
    selected = st.session_state.get(sel_key)

    if typed_norm:
        hits = [c for c in choices if typed_norm in str(c).lower()]
        if hits:
            st.caption("추천:")
            cols = st.columns(3)
            for i, name in enumerate(hits[:max_suggest]):
                if cols[i % 3].button(str(name), key=f"{sel_key}_{i}"):
                    st.session_state[key_text] = name
                    st.session_state[sel_key] = name
                    st.rerun()

    return selected


# ============================
# 히스토리 팝오버
# ============================
def render_history_popover():
    with st.popover("히스토리"):
        hist = st.session_state.get("history", [])
        if not hist:
            st.caption("히스토리가 없습니다.")
            return

        st.caption("최근 10개 기록")
        last10 = hist[-10:][::-1]
        for i, h in enumerate(last10):
            diseases = h.get("diseases")
            if diseases is None:
                d = h.get("disease")
                diseases = [d] if d else []
            disease_label = ", ".join(diseases) if diseases else "(질병 없음)"
            label = f"{h['time']} | {disease_label} | {h['food_label']}"
            if st.button(label, key=f"hist_btn_{i}"):
                st.session_state.final_diseases = diseases
                st.session_state.final_food_row = h["food_row"]
                st.session_state.page = "result"
                st.rerun()


# ============================
# 질병 선택 UI (최대 3개)
# ============================
def disease_input_block(disease_names_all: list[str]) -> list[str]:
    st.subheader("① 질병 입력 (최대 3개)")

    selected_list: list[str | None] = []

    for i in range(1, 4):
        if i == 1 or (len(selected_list) >= i - 1 and selected_list[i - 2]):
            st.markdown(f"**질병 {i}**")
            dcol1, dcol2 = st.columns([5, 2])
            key_base = f"disease_{i}"

            existing_value = st.session_state.get(f"{key_base}_final")

            # 검색
            with dcol1:
                if HAS_SEARCHBOX:
                    disease_selected_search = st_searchbox(
                        lambda p: [
                            d
                            for d in disease_names_all
                            if p.lower() in str(d).lower()
                        ],
                        key=f"{key_base}_search",
                        default=existing_value,
                        default_searchterm=str(existing_value)
                        if existing_value
                        else "",
                        placeholder="Search ...",
                        edit_after_submit="option",
                    )
                else:
                    disease_selected_search = fallback_live_search(
                        "Search ...", f"{key_base}_query", disease_names_all
                    )

                if disease_selected_search:
                    st.session_state[f"{key_base}_final"] = disease_selected_search
                    st.session_state[f"{key_base}_source"] = "search"

            # SELECT (팝오버)
            with dcol2:
                with st.popover("SELECT"):
                    st.write("질병 목록에서 선택")
                    if disease_names_all:
                        options = ["(선택)"] + disease_names_all
                        disease_candidate = st.radio(
                            "질병 목록",
                            options,
                            key=f"{key_base}_pop_radio",
                            index=0,
                        )
                        if disease_candidate != "(선택)":
                            st.session_state[f"{key_base}_final"] = disease_candidate
                            st.session_state[f"{key_base}_source"] = "select"
                    else:
                        st.caption("질병 목록이 없습니다.")

            disease_selected = st.session_state.get(f"{key_base}_final")
            selected_list.append(disease_selected)

            if not disease_selected and i == 1:
                st.warning("최소 1개 이상의 질병을 선택하세요.")
        else:
            selected_list.append(None)

    final_list = [d for d in selected_list if d]

    if final_list:
        st.markdown(
            f"<div style='margin-top:0.5rem; padding:0.5rem; "
            f"background-color:rgba(180,200,255,0.25); border-radius:8px;'>"
            f"<b>선택된 질병 :</b> {', '.join(final_list)}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.caption("선택된 질병이 없습니다.")

    return final_list


# ============================
# 음식 + 상태 선택 UI
# ============================
def food_input_block(
    foods: pd.DataFrame,
    food_names_all: list[str],
    category_all: list[str],
) -> tuple[str | None, dict | None]:
    st.subheader("② 음식 + 상태 선택")

    fcol1, fcol2 = st.columns([6, 1])

    # 검색
    with fcol1:
        existing_food = st.session_state.get("food_selected_final")
        if HAS_SEARCHBOX:
            food_selected_search = st_searchbox(
                lambda p: [f for f in food_names_all if p.lower() in str(f).lower()],
                key="food_search",
                default=existing_food,
                default_searchterm=str(existing_food) if existing_food else "",
                placeholder="음식 검색",
                edit_after_submit="option",
            )
        else:
            food_selected_search = fallback_live_search(
                "음식 검색", "food_query", food_names_all
            )

        if food_selected_search and st.session_state.get("food_source") != "cat":
            st.session_state.food_selected_final = food_selected_search
            st.session_state.food_source = "search"

    # 카테고리 SELECT
    with fcol2:
        with st.popover("SELECT"):
            st.write("카테고리로 고르기")

            if not category_all:
                st.caption("카테고리 데이터 없음")
            else:
                cat_options = ["(선택)"] + category_all
                selected_cat = st.radio(
                    "카테고리", cat_options, key="cat_pop", index=0
                )

                foods_in_cat = []
                if selected_cat != "(선택)":
                    tmp = foods[foods[CATEGORY_COL] == selected_cat][FOOD_NAME_COL]
                    foods_in_cat = (
                        tmp.dropna().unique().tolist() if not tmp.empty else []
                    )

                if selected_cat == "(선택)":
                    st.caption("카테고리를 먼저 선택하세요.")
                elif foods_in_cat:
                    food_options = ["(선택)"] + foods_in_cat
                    food_candidate = st.radio(
                        "음식", food_options, key="food_pop", index=0
                    )
                    if food_candidate != "(선택)":
                        st.session_state.food_selected_final = food_candidate
                        st.session_state.food_source = "cat"
                else:
                    st.caption("해당 카테고리에 음식이 없습니다.")

    food_selected = st.session_state.get("food_selected_final")

    if food_selected:
        st.info(f"현재 선택된 음식: **{food_selected}**")
    else:
        st.warning("검색 또는 SELECT에서 음식을 선택하세요.")

    # 상태 선택
    selected_row = None
    if food_selected:
        subset = foods[foods[FOOD_NAME_COL] == food_selected].copy()

        if subset.empty:
            st.error("해당 음식 데이터가 없습니다.")
        else:
            prev_food = st.session_state.get("prev_food_for_state")
            if prev_food != food_selected:
                for k in ("state1", "state2"):
                    if k in st.session_state:
                        del st.session_state[k]
                st.session_state.prev_food_for_state = food_selected

            # 상태1 후보
            has_blank_state1 = False
            if STATE1_COL in subset.columns:
                s1_series = subset[STATE1_COL].astype("string")
                has_blank_state1 = s1_series.isna().any() or (
                    s1_series.str.strip() == ""
                ).any()
                s1_real = sorted(
                    [
                        x
                        for x in s1_series.dropna().unique().tolist()
                        if str(x).strip() != ""
                    ]
                )
            else:
                s1_real = []

            state1_options = ["(선택)"]
            if has_blank_state1:
                state1_options.append("기본")
            state1_options += s1_real

            if len(state1_options) == 1:
                st.info("상태 없음 → 기본값 사용")
                match = subset
            else:
                chosen_state1 = st.selectbox(
                    "상태", state1_options, index=0, key="state1"
                )
                if chosen_state1 == "(선택)":
                    st.warning("상태를 선택하세요.")
                    match = subset.iloc[0:0]
                else:
                    if chosen_state1 == "기본":
                        match_s1 = subset[
                            subset[STATE1_COL].isna()
                            | (
                                subset[STATE1_COL]
                                .astype("string")
                                .str.strip()
                                == ""
                            )
                        ]
                    else:
                        match_s1 = subset[subset[STATE1_COL] == chosen_state1]

                    # 상태2 후보
                    if STATE2_COL in match_s1.columns:
                        s2_series = match_s1[STATE2_COL].astype("string")
                        s2_vals = sorted(
                            [
                                x
                                for x in s2_series.dropna().unique().tolist()
                                if str(x).strip() != ""
                            ]
                        )
                    else:
                        s2_vals = []

                    if s2_vals:
                        chosen_state2 = st.selectbox(
                            "세부 상태",
                            ["(선택)"] + s2_vals,
                            index=0,
                            key="state2",
                        )
                        if chosen_state2 == "(선택)":
                            match = match_s1.iloc[0:0]
                        else:
                            match = match_s1[match_s1[STATE2_COL] == chosen_state2]
                    else:
                        match = match_s1

            if match.empty:
                st.error("해당 상태 조합 데이터 없음")
            else:
                selected_row = match.iloc[0]

    return food_selected, (None if selected_row is None else selected_row.to_dict())


# ============================
# 문의하기 (st.dialog 사용, 실제 이메일 전송 + 창 자동 닫기)
# ============================
@st.dialog("문의하기")
def open_contact_modal():
    st.markdown("**궁금한 점이나 데이터 추가 요청을 남겨주세요.**")

    user_email = st.text_input("📧 이메일 주소", key="contact_email")
    user_msg = st.text_area("📝 문의 내용", key="contact_message", height=200)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Send", key="contact_send"):
            if not user_email or not user_msg:
                st.warning("이메일과 내용을 모두 입력해주세요.")
            else:
                # 콘솔 출력
                print("===== 문의하기 도착 =====")
                print("보낸 사람:", user_email)
                print("내용:")
                print(user_msg)
                print("=======================")

                # 실제 이메일 전송
                ok, err = send_contact_email(user_email, user_msg)
                if ok:
                    st.session_state.contact_sent = True
                    st.session_state.contact_clear_form = True
                    st.rerun()
                else:
                    st.error(
                        "문의가 접수되었지만 이메일 전송에 실패했습니다.\n"
                        "SMTP 설정 또는 서버 상태를 확인해주세요."
                    )
                    st.caption(f"에러 내용: {err}")

    with col2:
        if st.button("닫기", key="contact_close"):
            st.rerun()


# ============================
# 세션 상태 초기화
# ============================
if "page" not in st.session_state:
    st.session_state.page = "input"

if "final_diseases" not in st.session_state:
    st.session_state.final_diseases = []

if "final_food_row" not in st.session_state:
    st.session_state.final_food_row = None

for i in range(1, 4):
    key_final = f"disease_{i}_final"
    key_source = f"disease_{i}_source"
    if key_final not in st.session_state:
        st.session_state[key_final] = None
    if key_source not in st.session_state:
        st.session_state[key_source] = None

if "food_selected_final" not in st.session_state:
    st.session_state.food_selected_final = None
if "food_source" not in st.session_state:
    st.session_state.food_source = None

if "history" not in st.session_state:
    st.session_state.history = []

if "contact_email" not in st.session_state:
    st.session_state.contact_email = ""
if "contact_message" not in st.session_state:
    st.session_state.contact_message = ""
if "contact_sent" not in st.session_state:
    st.session_state.contact_sent = False
if "contact_clear_form" not in st.session_state:
    st.session_state.contact_clear_form = False

# 폼 비우기 플래그 처리
if st.session_state.contact_clear_form:
    st.session_state.contact_email = ""
    st.session_state.contact_message = ""
    st.session_state.contact_clear_form = False


# ============================
# 엑셀 로드
# ============================
try:
    foods, diseases, disease_expl = load_data(
        FOOD_FILE, DISEASE_FILE, DISEASE_EXPLAIN_FILE
    )
except Exception as e:
    st.error(f"엑셀 불러오기 실패: {e}")
    st.stop()

disease_names_all = (
    diseases[DISEASE_NAME_COL].dropna().unique().tolist()
    if DISEASE_NAME_COL in diseases.columns
    else []
)
food_names_all = (
    foods[FOOD_NAME_COL].dropna().unique().tolist()
    if FOOD_NAME_COL in foods.columns
    else []
)
category_all = (
    sorted(foods[CATEGORY_COL].dropna().unique().tolist())
    if CATEGORY_COL in foods.columns
    else []
)


# ============================
# PAGE 1 — 입력 화면
# ============================
if st.session_state.page == "input":
    h_left, h_right = st.columns([8, 1])
    with h_left:
        st.title("질병별 음식 판정 시스템")
    with h_right:
        render_history_popover()

    col1, col2 = st.columns([1, 2])

    with col1:
        selected_diseases = disease_input_block(disease_names_all)

    with col2:
        food_selected, food_row_dict = food_input_block(
            foods, food_names_all, category_all
        )

    st.markdown("---")

    # OK 버튼
    if st.button("OK", use_container_width=True):
        if not selected_diseases:
            st.error("최소 1개 이상의 질병을 선택하세요.")
        elif not food_selected or food_row_dict is None:
            st.error("음식/상태 선택을 완료하세요.")
        else:
            st.session_state.final_diseases = [str(d) for d in selected_diseases]
            st.session_state.final_food_row = food_row_dict

            time_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            frow = pd.Series(food_row_dict)

            label = str(frow.get(FOOD_NAME_COL, ""))
            s1 = frow.get(STATE1_COL)
            s2 = frow.get(STATE2_COL)

            if pd.notna(s1) and str(s1).strip() != "":
                label += f" / {s1}"
            else:
                label += " / 기본"

            if pd.notna(s2) and str(s2).strip() != "":
                label += f" / {s2}"

            hist = st.session_state.history
            hist.append(
                {
                    "time": time_str,
                    "diseases": st.session_state.final_diseases,
                    "food_label": label,
                    "food_row": food_row_dict,
                }
            )
            st.session_state.history = hist[-100:]

            st.session_state.page = "result"
            st.rerun()

    st.markdown("---")

    # 문의하기 버튼
    if st.button("문의하기", key="contact_button"):
        open_contact_modal()

    # 메일 전송 성공 메시지 (한 번만)
    if st.session_state.contact_sent:
        st.success("메일이 성공적으로 전송되었습니다!")
        st.session_state.contact_sent = False


# ============================
# PAGE 2 — 결과 화면
# ============================
else:
    h_left, h_right = st.columns([8, 1])
    with h_left:
        st.title("판정 결과")
    with h_right:
        render_history_popover()

    dnames = st.session_state.get("final_diseases", [])
    frow_dict = st.session_state.get("final_food_row")

    if not dnames or frow_dict is None:
        st.error("결과를 표시할 데이터가 없습니다. 처음 화면에서 다시 진행해 주세요.")
        if st.button("처음으로"):
            st.session_state.page = "input"
            st.session_state.food_selected_final = None
            st.session_state.food_source = None
            st.rerun()
        st.stop()

    frow = pd.Series(frow_dict)

    # 질병별 판정
    disease_results = []
    for dname in dnames:
        drow_df = diseases[diseases[DISEASE_NAME_COL] == dname]
        if drow_df.empty:
            st.error(f"질병 데이터 없음: {dname}")
            continue
        drow = drow_df.iloc[0]
        res = evaluate_row(frow, drow)
        disease_results.append({"name": dname, "row": drow, "res": res})

    if not disease_results:
        st.error("유효한 질병 데이터가 없어 판정을 수행할 수 없습니다.")
        if st.button("처음으로"):
            st.session_state.page = "input"
            st.rerun()
        st.stop()

    # 전체 판정 (불합격 > 주의 > 합격)
    global_has_fail = any(d["res"]["has_fail"] for d in disease_results)
    global_has_warn = any(d["res"]["has_warning"] for d in disease_results)
    if global_has_fail:
        global_status = "불합격 ❌"
    elif global_has_warn:
        global_status = "주의 ⚠️"
    else:
        global_status = "합격 ✅"

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("", ", ".join(d["name"] for d in disease_results))

    with c2:
        label = str(frow.get(FOOD_NAME_COL, ""))
        s1 = frow.get(STATE1_COL)
        s2 = frow.get(STATE2_COL)

        if pd.notna(s1) and str(s1).strip() != "":
            label += f" / {s1}"
        else:
            label += " / 기본"

        if pd.notna(s2) and str(s2).strip() != "":
            label += f" / {s2}"

        st.metric("음식", label)

    with c3:
        st.metric("최종 판정", global_status)

    st.markdown("---")

    # 불합격 질병 먼저, 그 다음 주의, 마지막 합격
    def sort_key(item):
        r = item["res"]
        if r["has_fail"]:
            return 0
        elif r["has_warning"]:
            return 1
        else:
            return 2

    sorted_results = sorted(disease_results, key=sort_key)

    # 행별 배경색 스타일링 함수
    def style_by_status(row):
        status = row.get("판정")
        if status == "불합격":
            color = "rgba(255, 0, 0, 0.18)"
        elif status == "주의":
            color = "rgba(255, 255, 0, 0.18)"
        elif status == "합격":
            color = "rgba(0, 200, 0, 0.18)"
        else:
            color = ""
        if color:
            return [f"background-color: {color};"] * len(row)
        else:
            return [""] * len(row)

    for item in sorted_results:
        dname = item["name"]
        drow = item["row"]
        res = item["res"]

        if res["has_fail"]:
            status_text = "불합격 ❌"
        elif res["has_warning"]:
            status_text = "주의 ⚠️"
        else:
            status_text = "합격 ✅"

        st.markdown(f"### {dname} — {status_text}")

        # 대체 음식 추천 (불합격일 때만)
        if res["has_fail"]:
            alt_df = recommend_alternatives_with_fallback(
                foods=foods,
                disease_row=drow,
                original_row=frow,
                max_rec=4,
            )

            if alt_df.empty:
                st.info(
                    "추천할 대체 음식을 찾지 못했습니다. "
                    "(같은 카테고리와 전체 음식 모두에서 합격/주의 음식 없음)"
                )
            else:
                # 설명 문구는 빈 캡션만
                st.caption("")
                display_df = alt_df[["추천 음식", "카테고리", "판정"]].copy()
                st.dataframe(display_df, use_container_width=True, hide_index=True)

        # 질병 설명
        disease_explain_text = None
        if (
            isinstance(disease_expl, pd.DataFrame)
            and not disease_expl.empty
            and DISEASE_NAME_COL in disease_expl.columns
            and "설명" in disease_expl.columns
        ):
            row_ex = disease_expl[disease_expl[DISEASE_NAME_COL] == dname]
            if not row_ex.empty:
                disease_explain_text = str(row_ex.iloc[0]["설명"])

        if disease_explain_text:
            with st.expander("질병 설명", expanded=True):
                st.write(disease_explain_text)
        else:
            with st.expander("질병 설명", expanded=True):
                st.info(
                    "해당 질병에 대한 설명이 등록되어 있지 않습니다.\n"
                    "※ disease_explanation.xlsx 에 [질병명, 설명]을 추가하면 여기 표시됩니다."
                )

        # 세부 영양 비교
        with st.expander(
            f"세부 영양 비교 — {dname}",
            expanded=(res["has_fail"] or res["has_warning"]),
        ):
            rows = []
            for col in NUTRIENT_COLS:
                fv = frow.get(col, np.nan)
                cv = drow.get(col, np.nan)

                warn_col = WARNING_COL_MAP.get(col)
                wv = drow.get(warn_col, np.nan) if warn_col in drow.index else np.nan

                status = ""
                excess = np.nan

                if pd.notna(fv) and pd.notna(cv):
                    if pd.notna(wv):
                        if fv > cv:
                            status = "불합격"
                            excess = float(fv - cv)
                        elif fv >= wv:
                            status = "주의"
                        else:
                            status = "합격"
                    else:
                        if fv > cv:
                            status = "불합격"
                            excess = float(fv - cv)
                        else:
                            status = "합격"

                rows.append(
                    {
                        "영양소": col,
                        "음식값": None if pd.isna(fv) else float(fv),
                        "주의시작": None if pd.isna(wv) else float(wv),
                        "컷오프": None if pd.isna(cv) else float(cv),
                        "판정": status if status else None,
                        "초과량": None
                        if (pd.isna(excess) or status != "불합격")
                        else float(excess),
                    }
                )

            df_detail = pd.DataFrame(rows)
            styled = df_detail.style.apply(style_by_status, axis=1)
            st.dataframe(styled, use_container_width=True)

        st.markdown("---")

    if st.button("처음으로", use_container_width=True):
        st.session_state.page = "input"
        st.session_state.food_selected_final = None
        st.session_state.food_source = None
        st.rerun()
