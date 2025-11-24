import os
import re
import base64
import mimetypes
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# 이메일 전송용
import smtplib
from email.mime.text import MIMEText

# OpenAI (AI 대체 음식 추천용)
from openai import OpenAI

# (선택) 자동완성 컴포넌트
try:
    from streamlit_searchbox import st_searchbox
    HAS_SEARCHBOX = True
except Exception:
    HAS_SEARCHBOX = False

# ============================
# Streamlit 페이지 설정
# ============================
st.set_page_config(page_title="질병별 음식 판정", layout="wide")

# ============================
# 툴팁용 CSS / HTML (제목 옆 ?에 사용)
# ============================
tooltip_css = """
<style>
.tooltip-wrapper {
    display: inline-flex;
    align-items: center;
    position: relative;
    margin-left: 8px;
}

.tooltip-icon {
    background-color: #ffffff;
    color: #111111;
    font-weight: 800;
    border-radius: 50%;
    padding: 4px 12px;
    cursor: default;
    border: 1px solid rgba(0,0,0,0.25);
    font-size: 1.15rem;
    line-height: 1.2;
    box-shadow: 0 1px 4px rgba(0,0,0,0.18);
}

.tooltip-box {
    visibility: hidden;
    opacity: 0;
    width: 390px;
    max-width: 92vw;
    background-color: #ffffff;
    color: #333333;
    text-align: left;
    border-radius: 8px;
    padding: 12px 14px;
    position: absolute;
    z-index: 999;
    top: 38px;
    left: -20px;
    font-size: 0.9rem;
    line-height: 1.35rem;
    box-shadow: 0 2px 10px rgba(0,0,0,0.18);
    transition: opacity 0.18s ease-in-out;
}

/* 마우스 올렸을 때에만 보이도록 */
.tooltip-wrapper:hover .tooltip-box {
    visibility: visible;
    opacity: 1;
}
</style>
"""

tooltip_html = """
<span class="tooltip-wrapper">
    <span class="tooltip-icon">?</span>
    <div class="tooltip-box">
        🔍 ‘대표적인 식품명’보다 ‘구체적인 제품명’ 검색 시 정확도가 높을 수 있습니다.<br>
        ⏳ 대량의 데이터를 사용하므로 간헐적으로 로딩이 지연될 수 있습니다.<br>
        ⚖️ 모든 영양소 값은 100g 기준이며, 질병 컷오프 역시 100g 기준으로 판정됩니다.<br>
        📊 데이터 출처: 식품의약품안전처 공식 식품영양성분 DB.<br>
        💬 문의 사항은 페이지 하단의 [문의하기] 기능을 이용해주세요.
    </div>
</span>
"""

st.markdown(tooltip_css, unsafe_allow_html=True)

# ============================
# OpenAI API 키 설정 (AI 추천용)
# ============================
try:
    openai_conf = st.secrets["openai"]
    _api_key = openai_conf.get("OPENAI_API_KEY", "")
    if _api_key:
        client = OpenAI(api_key=_api_key)
        HAS_OPENAI = True
    else:
        client = None
        HAS_OPENAI = False
except Exception:
    client = None
    HAS_OPENAI = False

# ============================
# 기본 경로/상수 설정
# ============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")
FOOD_FILE_BASE = os.path.join(DATA_DIR, "food_insert1.xlsx")    # 대표 음식 DB
FOOD_FILE_DETAIL = os.path.join(DATA_DIR, "food_insert2.xlsx")  # 구체 제품 DB
DISEASE_FILE = os.path.join(DATA_DIR, "disease_insert.xlsx")
DISEASE_EXPLAIN_FILE = os.path.join(DATA_DIR, "disease_explanation.xlsx")

# 배경 이미지
BG_IMAGE_FILE = os.path.join(BASE_DIR, "foodphoto.png")

FOOD_NAME_COL = "식품명"
STATE1_COL = "상태 1"
STATE2_COL = "상태2"
CATEGORY_COL = "카테고리"
DISEASE_NAME_COL = "질병명"

# 기본 영양소 컬럼
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

# 질병 엑셀 "주의" 컷오프 매핑
WARNING_COL_MAP = {
    "단백질(g)": "단백질주의(g)",
    "지방(g)": "지방주의(g)",
    "탄수화물(g)": "탄수화물주의(g)",
    "당류(g)": "당류주의(mg)",
    "나트륨(mg)": "나트륨주의(mg)",
    "칼륨(mg)": "칼륨주의(mg)",
    "인(mg)": "인주의(mg)",
    "포화지방산(g)": "포화지방산주의(g)",
    "콜레스테롤(mg)": "콜레스테롤주의(mg)",
}
WARNING_NUMERIC_COLS = list(set(WARNING_COL_MAP.values()))

# 구체 제품 개수(표시용 대략값)
DETAIL_FOOD_COUNT_APPROX = 148000

# ============================
# 이메일 SMTP 설정
# ============================
try:
    email_conf = st.secrets["email"]
except Exception:
    email_conf = {}

SMTP_SERVER = email_conf.get("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(email_conf.get("SMTP_PORT", 587))
SMTP_USER = email_conf.get("SMTP_USER", "")
SMTP_PASSWORD = email_conf.get("SMTP_PASSWORD", "")
CONTACT_RECEIVER = email_conf.get("CONTACT_RECEIVER", SMTP_USER)


def send_contact_email(user_email: str, user_msg: str) -> tuple[bool, str]:
    """문의하기 폼에서 입력받은 내용을 실제 이메일로 전송."""
    if not SMTP_USER or not SMTP_PASSWORD or not CONTACT_RECEIVER:
        return False, "SMTP 설정이 비어 있습니다. secrets.toml의 [email] 값을 확인하세요."

    subject = "[질병별 음식 판정] 문의가 도착했습니다"
    time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    body = (
        "질병별 음식 판정 시스템에서 새로운 문의가 접수되었습니다.\n\n"
        f"시간: {time_str}\n"
        f"보낸 사람 이메일: {user_email}\n\n"
        "문의 내용:\n"
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

        print("===== 문의하기 이메일 전송 완료 =====")
        print("To:", CONTACT_RECEIVER)
        print("From:", SMTP_USER)
        print("User Email:", user_email)
        print("Message:")
        print(user_msg)
        print("===================================")

        return True, ""
    except Exception as e:
        print("===== 문의하기 이메일 전송 실패 =====")
        print("에러:", e)
        print("User Email:", user_email)
        print("Message:")
        print(user_msg)
        print("===================================")
        return False, str(e)


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
# 데이터 로드 (지연 로딩 구조)
# ============================
def _clean_food_common(foods: pd.DataFrame) -> pd.DataFrame:
    for c in [FOOD_NAME_COL, STATE1_COL, STATE2_COL, CATEGORY_COL]:
        if c in foods.columns:
            foods[c] = (
                foods[c]
                .astype("string")
                .str.strip()
                .replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})
            )
    return foods


@st.cache_data
def load_base_and_disease():
    """대표 음식 + 질병 + 질병 설명만 로드 (초기 진입용)."""
    # 대표 음식 DB
    foods_base = pd.read_excel(FOOD_FILE_BASE)
    foods_base = _clean_food_common(foods_base)
    foods_base = coerce_numeric(foods_base, NUTRIENT_COLS)

    # 질병 컷오프
    diseases = pd.read_excel(DISEASE_FILE)
    if DISEASE_NAME_COL in diseases.columns:
        diseases[DISEASE_NAME_COL] = (
            diseases[DISEASE_NAME_COL]
            .astype("string")
            .str.strip()
            .replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})
        )
    diseases = coerce_numeric(diseases, NUTRIENT_COLS + WARNING_NUMERIC_COLS)

    # 질병 설명
    try:
        disease_expl = pd.read_excel(DISEASE_EXPLAIN_FILE)
    except Exception:
        disease_expl = pd.DataFrame(columns=[DISEASE_NAME_COL, "설명"])

    if DISEASE_NAME_COL in disease_expl.columns:
        disease_expl[DISEASE_NAME_COL] = (
            disease_expl[DISEASE_NAME_COL]
            .astype("string")
            .str.strip()
            .replace({"nan": pd.NA, "None": pd.NA, "": pd.NA})
        )
    if "설명" in disease_expl.columns:
        disease_expl["설명"] = disease_expl["설명"].astype("string").str.strip()

    return foods_base, diseases, disease_expl


@st.cache_data
def load_detail_food():
    """구체적인 제품 DB는 실제로 선택했을 때만 로드."""
    foods_detail = pd.read_excel(FOOD_FILE_DETAIL)
    foods_detail = _clean_food_common(foods_detail)
    foods_detail = coerce_numeric(foods_detail, NUTRIENT_COLS)
    return foods_detail


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

                if pd.notna(warn_cut):
                    if val > cutoff:
                        status = "fail"
                    elif val >= warn_cut:
                        status = "warn"
                    else:
                        status = "ok"
                else:
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
# 벡터 기반 추천용 유틸
# ============================
def nutrient_distance(
    original_row: pd.Series, candidate_row: pd.Series, disease_row: pd.Series
) -> float:
    """질병 컷오프로 정규화한 유클리드 거리"""
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


def state_preference(row: pd.Series) -> tuple[int, int]:
    """상태1/상태2 우선순위 (작을수록 우선)"""
    s1 = row.get(STATE1_COL)
    s2 = row.get(STATE2_COL)

    s1_str = "" if pd.isna(s1) else str(s1).strip()
    s2_str = "" if pd.isna(s2) else str(s2).strip()

    s1_flag = 0 if (s1_str == "" or s1_str == "기본") else 1
    s2_flag = 0 if s2_str == "" else 1

    return (s1_flag, s2_flag)


def _build_vector_recs_from_candidates(
    cand: pd.DataFrame,
    disease_row: pd.Series,
    original_row: pd.Series,
    max_rec: int,
) -> pd.DataFrame:
    if cand.empty:
        return pd.DataFrame()

    # 원래 선택한 음식 제외 (이름+상태 기준)
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
            target = pass_dict
        elif status == "주의":
            target = warn_dict
        else:
            continue

        cur = target.get(base_name)
        if cur is None or pref < cur["pref"] or (
            pref == cur["pref"] and dist < cur["dist"]
        ):
            target[base_name] = {"row": row, "dist": dist, "pref": pref}

    if not pass_dict and not warn_dict:
        return pd.DataFrame()

    pass_list = [(v["dist"], v["row"]) for v in pass_dict.values()]
    warn_list = [(v["dist"], v["row"]) for v in warn_dict.values()]

    pass_list.sort(key=lambda x: x[0])
    warn_list.sort(key=lambda x: x[0])

    selected: list[tuple[str, pd.Series]] = []
    for _, row in pass_list:
        if len(selected) >= max_rec:
            break
        selected.append(("합격", row))

    if len(selected) < max_rec:
        for _, row in warn_list:
            if len(selected) >= max_rec:
                break
            selected.append(("주의", row))

    if not selected:
        return pd.DataFrame()

    out_rows = []
    for status, row in selected:
        name = str(row.get(FOOD_NAME_COL, ""))
        s1 = row.get(STATE1_COL)
        s2 = row.get(STATE2_COL)

        label = name
        if STATE1_COL in row.index and pd.notna(s1) and str(s1).strip() != "":
            label += f" / {s1}"
        else:
            label += " / 기본"

        if STATE2_COL in row.index and pd.notna(s2) and str(s2).strip() != "":
            label += f" / {s2}"

        out_rows.append(
            {
                "추천 음식": label,
                "카테고리": row.get(CATEGORY_COL),
                "판정": status,
            }
        )

    return pd.DataFrame(out_rows)


def recommend_vector_based_alternatives(
    foods: pd.DataFrame,
    disease_row: pd.Series,
    original_row: pd.Series,
    max_rec: int = 4,
) -> pd.DataFrame:
    """
    1순위: 같은 카테고리 + 합격/주의 음식들로 최대 max_rec개 추천
    2순위: 그 외 카테고리 + 합격/주의 음식들로 나머지 개수 채우기
    (기존처럼 불합격은 항상 제외)

    ⚠ Streamlit Cloud 속도 문제 때문에,
      같은 카테고리/다른 카테고리 모두 후보 수를 일정 개수로 제한해서
      너무 많은 행을 돌지 않도록 함.
    """
    orig_cat = original_row.get(CATEGORY_COL)

    # 한 번에 계산할 최대 행 수 (너무 크면 느려짐)
    MAX_SAME = 3000      # 같은 카테고리에서 최대 3000개만
    MAX_OTHERS = 4000    # 다른 카테고리에서 최대 4000개만

    collected_df_list: list[pd.DataFrame] = []

    # ----- 1단계: 같은 카테고리 -----
    if CATEGORY_COL in foods.columns and pd.notna(orig_cat):
        cand_same = foods[foods[CATEGORY_COL] == orig_cat].copy()

        # 너무 많으면 샘플링
        if len(cand_same) > MAX_SAME:
            cand_same = cand_same.sample(MAX_SAME, random_state=0)

        same_df = _build_vector_recs_from_candidates(
            cand_same, disease_row, original_row, max_rec
        )
        if not same_df.empty:
            collected_df_list.append(same_df)

    current_n = sum(len(df) for df in collected_df_list)

    # ----- 2단계: 다른 카테고리 -----
    if current_n < max_rec:
        if CATEGORY_COL in foods.columns and pd.notna(orig_cat):
            others = foods[foods[CATEGORY_COL] != orig_cat].copy()
        else:
            others = foods.copy()

        # 이미 선택된 이름 제외
        already_names: set[str] = set()
        for df in collected_df_list:
            base_names = df["추천 음식"].astype(str).str.split(" /").str[0]
            already_names.update(base_names.tolist())

        if not others.empty and already_names:
            others = others[
                ~others[FOOD_NAME_COL].astype(str).isin(already_names)
            ]

        # 너무 많으면 샘플링
        if len(others) > MAX_OTHERS:
            others = others.sample(MAX_OTHERS, random_state=1)

        remain = max_rec - current_n
        if remain > 0 and not others.empty:
            others_df = _build_vector_recs_from_candidates(
                others, disease_row, original_row, remain
            )
            if not others_df.empty:
                collected_df_list.append(others_df)

    if not collected_df_list:
        return pd.DataFrame()

    result = pd.concat(collected_df_list, ignore_index=True)

    if len(result) > max_rec:
        result = result.iloc[:max_rec].reset_index(drop=True)

    return result


# ============================
# 음식 base type 추론 (면/밥/생선/빵/기타)
# ============================
def infer_base_type(food_name: str, category: str | None) -> str:
    text = (str(food_name) + " " + str(category)).lower()

    if any(k in text for k in ["면", "국수", "파스타", "우동", "소바", "라멘", "라면"]):
        return "면"

    if any(k in text for k in ["밥", "죽", "비빔밥", "덮밥"]):
        return "밥"

    if any(k in text for k in ["생선", "고등어", "갈치", "연어", "오징어", "명태", "물고기"]):
        return "생선"

    if any(k in text for k in ["빵", "토스트", "베이글", "샌드위치"]):
        return "빵"

    return "기타"


# ============================
# AI 대체 음식 추천 (GPT 순수 추론)
# ============================
def get_ai_alternatives(food_row: pd.Series, diseases: list[str]) -> str:
    if not HAS_OPENAI or client is None:
        return "⚠️ OpenAI API 설정(OPENAI_API_KEY)이 되어 있지 않아 AI 대체 음식 추천을 사용할 수 없습니다."

    cur_name = str(food_row.get(FOOD_NAME_COL, ""))
    cur_cat = food_row.get(CATEGORY_COL, None)
    cur_s1 = food_row.get(STATE1_COL, None)
    cur_s2 = food_row.get(STATE2_COL, None)

    food_label = cur_name
    if STATE1_COL in food_row.index and pd.notna(cur_s1) and str(cur_s1).strip():
        food_label += f" / {cur_s1}"
    else:
        food_label += " / 기본"
    if STATE2_COL in food_row.index and pd.notna(cur_s2) and str(cur_s2).strip():
        food_label += f" / {cur_s2}"

    base_type = infer_base_type(cur_name, cur_cat)
    disease_text = ", ".join(diseases) if diseases else "없음"

    prompt = f"""
당신은 한국 음식 전문가이자 영양사입니다.

현재 선택된 음식: {food_label}
음식 유형(base type): {base_type}
사용자의 질병: {disease_text}

역할:
- 사용자가 현재 음식 대신 먹을 수 있는 더 안전한 한국 음식을 4가지 추천한다.
- 반드시 같은 유형(base type)의 음식만 추천한다.
- 영양소(특히 나트륨, 포화지방산, 당류 등)를 고려하여 현재 음식보다 부담이 적을 것 같은 선택을 한다.

형식:
1. 각 줄은 반드시 '음식 이름 - 한 줄 이유' 형식으로 작성한다.
2. 글머리 기호, 번호, 마크다운 표, HTML 태그, 코드블럭은 절대로 사용하지 않는다.
3. 마지막 줄에는 정확히 한 줄로 '이 내용은 의료 진단이 아닌 참고용입니다.' 문장을 넣는다.
"""

    try:
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        return f"⚠️ AI 추천 호출 중 오류가 발생했습니다: {e}"


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
            diseases_h = h.get("diseases")
            if diseases_h is None:
                d = h.get("disease")
                diseases_h = [d] if d else []
            disease_label = ", ".join(diseases_h) if diseases_h else "(질병 없음)"
            label = f"{h['time']} | {disease_label} | {h['food_label']}"
            if st.button(label, key=f"hist_btn_{i}"):
                st.session_state.final_diseases = diseases_h
                st.session_state.final_food_row = h["food_row"]
                st.session_state.page = "result"
                st.rerun()


# ============================
# 질병 선택 UI (최대 3개)
# ============================
def disease_input_block(disease_names_all: list[str]) -> list[str]:
    st.subheader("① 질병 입력 (최대 3개)")

    # 🔧 Search / TextInput 아래 기본 여백 제거해서 안내문을 바로 밑에 붙이기
    st.markdown(
        """
        <style>
            div[data-testid="stSearchInputContainer"] {
                margin-bottom: 0rem !important;
            }
            div[data-testid="stTextInputRoot"] {
                margin-bottom: 0rem !important;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    selected_list: list[str | None] = []

    # modal 타깃 초기화
    if "disease_dialog_target" not in st.session_state:
        st.session_state["disease_dialog_target"] = None

    for i in range(1, 4):
        key_base = f"disease_{i}"
        final_key = f"{key_base}_final"

        # 1번은 항상 보이고, 2·3번은 앞 슬롯이 채워졌을 때만 보이게
        if i == 1 or (len(selected_list) >= i - 1 and selected_list[i - 2]):
            st.markdown(f"**질병 {i}**")

            dcol1, dcol2 = st.columns([5, 1.4])
            existing_value = st.session_state.get(final_key)

            # ------------ 왼쪽: 검색 박스 ------------
            with dcol1:
                if HAS_SEARCHBOX:
                    disease_selected_search = st_searchbox(
                        lambda p: [
                            d for d in disease_names_all
                            if p.lower() in str(d).lower()
                        ],
                        key=f"{key_base}_search",
                        default=existing_value,
                        default_searchterm=str(existing_value) if existing_value else "",
                        placeholder="Search ...",
                        edit_after_submit="option",
                    )
                else:
                    disease_selected_search = fallback_live_search(
                        "Search ...", f"{key_base}_query", disease_names_all
                    )

                # 검색으로 선택된 값이 실제로 변경되었을 때만 최종 값 갱신
                if disease_selected_search and disease_selected_search != existing_value:
                    st.session_state[final_key] = disease_selected_search
                    st.session_state[f"{key_base}_source"] = "search"
                    st.rerun()

                # 🔻 2번째/3번째 질병 안내 문구 (검색 박스 바로 아래에 붙이기)
                if i == 2:
                    st.markdown(
                        """
                        <div style="
                            margin-top:-0.25rem;
                            margin-bottom:0.3rem;
                            color:#cc0000;
                            font-size:0.78rem;
                        ">
                            ※ 2번째 질병은 선택 시 함께 판정되며, 필요 없을 경우 선택하지 않으셔도 됩니다.
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                elif i == 3:
                    st.markdown(
                        """
                        <div style="
                            margin-top:-0.25rem;
                            margin-bottom:0.3rem;
                            color:#cc0000;
                            font-size:0.78rem;
                        ">
                            ※ 3번째 질병은 선택 시 함께 판정되며, 필요 없을 경우 선택하지 않으셔도 됩니다.
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

            # ------------ 오른쪽: SELECT 버튼 → 중앙 팝업 ------------
            with dcol2:
                if st.button("SELECT", key=f"{key_base}_select_btn"):
                    st.session_state["disease_dialog_target"] = i
                    open_disease_select_modal()

            # 현재 확정된 값
            disease_selected = st.session_state.get(final_key)
            selected_list.append(disease_selected)

            # 1번째 질병은 필수 → 비어 있으면 경고
            if not disease_selected and i == 1:
                st.warning("최소 1개 이상의 질병을 선택하세요.")
        else:
            # 아직 이 칸을 보여줄 조건이 안 되면 None으로 채워두기
            selected_list.append(None)

    # None 제거한 확정 리스트
    final_list = [d for d in selected_list if d]

    # 선택된 질병 표시 박스 (폰트 크게)
    if final_list:
        st.markdown(
            f"""
            <div style="
                margin-top:0.7rem;
                padding:0.6rem 0.8rem;
                background-color:rgba(180,200,255,0.30);
                border-radius:8px;
                font-size:1.15rem;
                font-weight:600;
            ">
                <span style="font-weight:800;">선택된 질병 :</span>
                {', '.join(final_list)}
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.caption("선택된 질병이 없습니다.")

    return final_list


# ============================
# 검색 엔진 선택 UI (대표 / 구체)
#  - 대표적인 식품명: 라디오 옵션은 유지, 회색/비활성 처리
#  - 순서: 선택 안 함, 구체적인 제품명, 대표적인 식품명(비활성)
# ============================
def search_engine_block(base_count: int, detailed_count_approx: int) -> str:
    st.subheader("② 검색 엔진 선택")

    opt_none = "선택 안 함"
    opt_base = f"대표적인 식품명으로 검색: 약 {base_count:,}개"
    opt_detail = (
        f"구체적인 제품명으로 검색: 약 {detailed_count_approx / 10000:.1f}만개의 데이터"
    )

    # 보여지는 순서: None → Detail → Base(비활성)
    options = [opt_none, opt_detail, opt_base]

    # 세션 기본값
    if "search_mode" not in st.session_state:
        st.session_state.search_mode = "선택 안 함"

    # 현재 search_mode를 옵션 문자열로 매핑
    if st.session_state.search_mode == "구체적인 제품명으로 검색":
        target = opt_detail
    elif st.session_state.search_mode == "대표적인 식품명으로 검색":
        target = opt_base
    else:
        target = opt_none

    try:
        idx = options.index(target)
    except ValueError:
        idx = 0

    # 이 라디오 블록에만 적용되도록 wrapper + CSS
    st.markdown(
        """
        <style>
        /* 검색 엔진 선택 라디오 안에서만
           3번째 옵션(대표적인 식품명)을 회색 + 클릭 막기 */
        .search-engine-radio div[data-testid="stRadio"] div[role="radiogroup"] > label:nth-child(3) span {
            color: #AAAAAA !important;
        }
        .search-engine-radio div[data-testid="stRadio"] div[role="radiogroup"] > label:nth-child(3) {
            pointer-events: none;      /* 클릭 불가 */
            opacity: 0.6;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="search-engine-radio">', unsafe_allow_html=True)
    mode_full = st.radio(
        " ",
        options,
        horizontal=True,
        index=idx,
        label_visibility="collapsed",
        key="search_mode_radio",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # 내부 로직용 짧은 모드 이름으로 정규화
    if mode_full == opt_none:
        mode = "선택 안 함"
    elif mode_full == opt_base:
        mode = "대표적인 식품명으로 검색"
    else:
        mode = "구체적인 제품명으로 검색"

    # 혹시 CSS가 깨져서 대표적인 식품명이 실제로 선택된 경우 방어
    if mode == "대표적인 식품명으로 검색":
        st.info(
            "대표적인 식품명 검색은 현재 데이터 검토 중이라 사용할 수 없습니다. "
            "[구체적인 제품명]으로 검색을 이용해주세요."
        )
        mode = "선택 안 함"

    st.session_state.search_mode = mode
    return mode


# ============================
# 음식 + 상태 선택 UI (대표 DB용)
# ============================
def food_input_block_base(
    foods: pd.DataFrame,
    food_names_all: list[str],
    category_all: list[str],
) -> tuple[str | None, dict | None]:
    st.subheader("③ 음식 + 상태 선택")

    fcol1, fcol2 = st.columns([6, 1])

    # 음식 검색
    with fcol1:
        existing_food = st.session_state.get("food_selected_final")
        last_source = st.session_state.get("food_source")  # 🔹 마지막 선택 경로

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

        # 🔧 검색 값 → SELECT보다 낮은 우선순위
        #  - 마지막 source가 cat(카테고리 선택) / cat_detail(제품 선택)이라면
        #    이번 rerun에서는 검색으로 덮지 않는다.
        if (
            food_selected_search
            and food_selected_search != existing_food
            and last_source not in ("cat", "cat_detail")
        ):
            st.session_state.food_selected_final = food_selected_search
            st.session_state.food_source = "search"
            st.rerun()

    # 카테고리 SELECT (팝오버 안에서 분류 → 음식)
    with fcol2:
        with st.popover("SELECT"):
            st.write("카테고리로 고르기")

            if not category_all:
                st.caption("카테고리 데이터 없음")
            else:
                stage_key = "base_select_stage"
                cat_key = "base_selected_cat"

                if stage_key not in st.session_state:
                    st.session_state[stage_key] = "cat"
                if cat_key not in st.session_state:
                    st.session_state[cat_key] = None

                stage = st.session_state[stage_key]
                current_cat = st.session_state[cat_key]

                if stage == "cat":
                    cat_options = ["(선택)"] + category_all
                    selected_cat = st.radio(
                        "카테고리",
                        cat_options,
                        key="cat_pop",
                        index=0,
                    )
                    if selected_cat != "(선택)":
                        st.session_state[cat_key] = selected_cat
                        st.session_state[stage_key] = "food"
                        st.rerun()
                else:
                    selected_cat = current_cat
                    st.caption(f"선택된 분류: **{selected_cat}**")

                    if st.button("분류 다시 선택", key="base_cat_reset"):
                        st.session_state[stage_key] = "cat"
                        st.session_state[cat_key] = None
                        st.rerun()

                    tmp = foods[foods[CATEGORY_COL] == selected_cat][FOOD_NAME_COL]
                    foods_in_cat = (
                        tmp.dropna().unique().tolist() if not tmp.empty else []
                    )

                    if not foods_in_cat:
                        st.caption("해당 카테고리에 음식이 없습니다.")
                    else:
                        food_options = ["(선택)"] + foods_in_cat
                        food_candidate = st.radio(
                            "음식",
                            food_options,
                            key="food_pop",
                            index=0,
                        )
                        if food_candidate != "(선택)":
                            st.session_state.food_selected_final = food_candidate
                            st.session_state.food_source = "cat"  # 🔹 카테고리 선택
                            st.success("음식이 선택되었습니다. 창을 닫아주세요.")

    food_selected = st.session_state.get("food_selected_final")

    # 선택된 음식 표시
    if food_selected:
        st.markdown(
            f"<div style='margin-top:0.5rem; padding:0.55rem 0.7rem; "
            f"background-color:rgba(255,255,255,0.85); border-radius:8px; "
            f"border-left:4px solid #4a90e2; font-size:1.05rem; font-weight:600;'>"
            f"<span style='font-weight:700;'>선택된 음식 :</span> {food_selected}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.warning("검색 또는 SELECT에서 음식을 선택하세요.")

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
                            match = match_s1[STATE2_COL] == chosen_state2
                    else:
                        match = match_s1

            if isinstance(match, pd.DataFrame) and match.empty:
                st.error("해당 상태 조합 데이터 없음")
            else:
                if isinstance(match, pd.Series) and match.dtype == bool:
                    match_df = subset[match]
                    if match_df.empty:
                        st.error("해당 상태 조합 데이터 없음")
                    else:
                        selected_row = match_df.iloc[0]
                else:
                    selected_row = match.iloc[0]

    return food_selected, (None if selected_row is None else selected_row.to_dict())


# ============================
# 제품 선택 UI (구체 제품 DB용, 상태 없음)
# ============================
def food_input_block_detail(
    foods: pd.DataFrame,
) -> tuple[str | None, dict | None]:
    st.subheader("③ 제품 선택")

    # 카테고리 목록 (있으면)
    if CATEGORY_COL in foods.columns:
        category_all = (
            sorted(foods[CATEGORY_COL].dropna().unique().tolist())
            if not foods.empty
            else []
        )
    else:
        category_all = []

    fcol1, fcol2 = st.columns([6, 1])

    # 1) 제품 검색 (자동완성: DataFrame 필터)
    with fcol1:
        existing_food = st.session_state.get("food_selected_final")

        def detail_searcher(query: str):
            if not query:
                return []
            df = foods[
                foods[FOOD_NAME_COL]
                .astype("string")
                .str.contains(query, case=False, na=False)
            ]
            return (
                df[FOOD_NAME_COL]
                .dropna()
                .drop_duplicates()
                .head(30)
                .tolist()
            )

        if HAS_SEARCHBOX:
            # 🔹 검색 박스의 key를 고정해서, SELECT로 고른 값도 같이 표시되도록
            food_selected_search = st_searchbox(
                detail_searcher,
                key="food_search_detail",
                default=existing_food,
                default_searchterm=str(existing_food) if existing_food else "",
                placeholder="제품명 검색",
                edit_after_submit="option",
            )
        else:
            typed = st.text_input("제품명 검색", key="food_query_detail")
            food_selected_search = None
            if typed:
                df = foods[
                    foods[FOOD_NAME_COL]
                    .astype("string")
                    .str.contains(typed, case=False, na=False)
                ]
                if not df.empty:
                    candidates = (
                        df[FOOD_NAME_COL]
                        .dropna()
                        .drop_duplicates()
                        .head(30)
                        .tolist()
                    )
                    st.caption("추천:")
                    for i, name in enumerate(candidates):
                        if st.button(name, key=f"detail_suggest_{i}"):
                            food_selected_search = name
                            break

        # 🔧 검색으로 선택된 값이 "기존 값과 실제로 다를 때만" 최종 선택 갱신
        if food_selected_search and food_selected_search != existing_food:
            st.session_state.food_selected_final = food_selected_search
            st.session_state.food_source = "search_detail"
            #st.session_state["food_search_detail"] = food_selected_search
            st.rerun()

    # 2) 카테고리 / 제품 SELECT → 중앙 모달
    with fcol2:
        if st.button("SELECT", key="detail_select_btn"):
            st.session_state["_detail_foods_df"] = foods
            open_detail_select_modal()

    # 3) 선택된 제품 표시
    food_selected = st.session_state.get("food_selected_final")

    if food_selected:
        st.markdown(
            f"<div style='margin-top:0.5rem; padding=0.55rem 0.7rem; "
            f"background-color:rgba(255,255,255,0.9); border-radius:8px; "
            f"border-left:4px solid #4a90e2; font-size:1.05rem; font-weight:600;'>"
            f"<span style='font-weight:700;'>선택된 음식 :</span> {food_selected}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.warning("검색 또는 SELECT에서 제품을 선택하세요.")

    # 4) 최종 row 선택 (동명이인 처리)
    selected_row = None
    if food_selected:
        subset = foods[foods[FOOD_NAME_COL] == food_selected].copy()

        if subset.empty:
            st.error("해당 제품 데이터가 없습니다.")
        else:
            subset_reset = subset.reset_index(drop=True)
            if len(subset_reset) == 1:
                selected_row = subset_reset.iloc[0]
            else:
                options = []
                for i, (_, r) in enumerate(subset_reset.iterrows()):
                    cat = r.get(CATEGORY_COL)
                    label = f"{i + 1}. {r.get(FOOD_NAME_COL, '')}"
                    if pd.notna(cat):
                        label += f" ({cat})"
                    options.append(label)

                chosen = st.radio(
                    "같은 제품명이 여러 개 있습니다. 선택하세요.",
                    options,
                    key="detail_row_choice",
                )
                idx = options.index(chosen)
                selected_row = subset_reset.iloc[idx]

    return food_selected, (None if selected_row is None else selected_row.to_dict())


# ============================
# 문의하기 (st.dialog 사용)
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
                print("===== 문의하기 도착 =====")
                print("보낸 사람:", user_email)
                print("내용:")
                print(user_msg)
                print("=======================")

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


@st.dialog("질병 선택")
def open_disease_select_modal():
    i = st.session_state.get("disease_dialog_target", 1)
    disease_names_all = st.session_state.get("_disease_names_all", [])

    key_base = f"disease_{i}"
    final_key = f"{key_base}_final"

    existing_value = st.session_state.get(final_key)

    if disease_names_all:
        options = ["(선택)"] + disease_names_all

        if existing_value and existing_value in disease_names_all:
            default_idx = disease_names_all.index(existing_value) + 1
        else:
            default_idx = 0

        choice = st.radio(
            "질병 선택",
            options,
            index=default_idx,
            key=f"disease_modal_radio_{i}",
        )

        if choice != "(선택)" and choice != existing_value:
            # 최종 질병 선택 값 저장
            st.session_state[final_key] = choice
            st.session_state[f"{key_base}_source"] = "select"

            # 🔹 disease searchbox 내부 상태 리셋
            search_key = f"{key_base}_search"   # st_searchbox 키
            query_key = f"{key_base}_query"     # fallback_live_search 키 (혹시 모를 때 대비)

            if search_key in st.session_state:
                del st.session_state[search_key]
            if query_key in st.session_state:
                del st.session_state[query_key]

            st.session_state["disease_dialog_target"] = None
            st.rerun()
    else:
        st.caption("질병 목록이 없습니다.")

    if st.button("닫기", key="disease_modal_close"):
        st.session_state["disease_dialog_target"] = None
        st.rerun()


# ============================
# 사이트 정보 (st.dialog 사용)
# ============================
@st.dialog("사이트 정보")
def open_site_info_modal():
    st.markdown(
        """
        **Creator** : Song Chae Yul  
        **GIT_ID** : codbf362-cmd  
        **Access Date** : 2025/11/19  
        **Made for Capstone**
        """
    )
    st.markdown("---")
    if st.button("닫기", key="siteinfo_close"):
        st.rerun()


# ============================
# 벡터 기반 대체 음식 추천 모달
# ============================
@st.dialog("벡터 기반 대체 음식 추천")
def open_vector_modal():
    mode = st.session_state.get("search_mode")

    if mode == "구체적인 제품명으로 검색":
        foods_df = st.session_state.get("_detail_foods_df")
        if foods_df is None:
            try:
                foods_df = load_detail_food()
                st.session_state["_detail_foods_df"] = foods_df
            except Exception:
                foods_df = None
    else:
        foods_df = st.session_state.get("_base_foods_df")

    if foods_df is None or foods_df.empty:
        st.info("⚠ 대체 음식 추천을 위한 음식 데이터가 없습니다.")
        return

    vec_df = recommend_vector_based_alternatives(
        foods=foods_df,
        disease_row=drow,
        original_row=frow,
        max_rec=4,
    )

    if vec_df.empty:
        st.info("⚠ 벡터 기반으로 추천할 수 있는 음식이 없습니다.")
    else:
        st.dataframe(vec_df, width="stretch", hide_index=True)

    # 굳이 버튼 없이 X로만 닫게 하고 싶으면 아래는 아예 지우거나 주석 처리
    # if st.button("닫기", key="vector_close"):
    #     st.rerun()




@st.dialog("제품 선택")
def open_detail_select_modal():
    foods = st.session_state.get("_detail_foods_df")

    chosen_food = None

    if foods is None or foods.empty:
        st.info("제품 데이터를 불러올 수 없습니다.")
        if st.button("닫기", key="detail_modal_close_empty"):
            st.rerun()
        return

    if CATEGORY_COL in foods.columns:
        category_all = (
            sorted(foods[CATEGORY_COL].dropna().unique().tolist())
            if not foods.empty
            else []
        )
    else:
        category_all = []

    selected_cat = st.selectbox(
        "카테고리",
        ["(선택)"] + category_all,
        key="detail_modal_cat",
    )

    foods_in_cat = []
    if selected_cat != "(선택)":
        tmp = foods[foods[CATEGORY_COL] == selected_cat][FOOD_NAME_COL]
        foods_in_cat = (
            tmp.dropna().drop_duplicates().tolist()
            if not tmp.empty
            else []
        )

    if selected_cat == "(선택)" or not foods_in_cat:
        st.caption("카테고리를 선택하면 해당 제품 목록이 표시됩니다.")
    else:
        food_options = ["(선택)"] + foods_in_cat
        chosen_food = st.radio(
            "제품 선택",
            food_options,
            key="detail_modal_food",
        )
        if chosen_food == "(선택)":
            chosen_food = None

    if chosen_food:
        # 최종 선택 값 저장
        st.session_state.food_selected_final = chosen_food
        st.session_state.food_source = "cat_detail"

        # 🔹 searchbox 내부 상태를 초기화해서,
        #    다음 rerun 때 default=chosen_food 이 화면에 그대로 뜨게 함
        if "food_search_detail" in st.session_state:
            del st.session_state["food_search_detail"]
        if "food_query_detail" in st.session_state:
            del st.session_state["food_query_detail"]

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

if "final_search_mode" not in st.session_state:
    st.session_state.final_search_mode = "선택 안 함"

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

if "show_vector_modal" not in st.session_state:
    st.session_state.show_vector_modal = False

if "search_mode" not in st.session_state:
    st.session_state.search_mode = "선택 안 함"

if "_detail_foods_df" not in st.session_state:
    st.session_state["_detail_foods_df"] = None

# 문의 폼 초기화 플래그
if st.session_state.contact_clear_form:
    st.session_state.contact_email = ""
    st.session_state.contact_message = ""
    st.session_state.contact_clear_form = False


# ============================
# 엑셀 로드 (초기엔 대표+질병만)
# ============================
try:
    foods_base, diseases, disease_expl = load_base_and_disease()
except Exception as e:
    st.error(f"엑셀 불러오기 실패: {e}")
    st.stop()

st.session_state["_base_foods_df"] = foods_base

disease_names_all = (
    diseases[DISEASE_NAME_COL].dropna().unique().tolist()
    if DISEASE_NAME_COL in diseases.columns
    else []
)
st.session_state["_disease_names_all"] = disease_names_all

food_names_base = (
    foods_base[FOOD_NAME_COL].dropna().unique().tolist()
    if FOOD_NAME_COL in foods_base.columns
    else []
)
category_base = (
    sorted(foods_base[CATEGORY_COL].dropna().unique().tolist())
    if CATEGORY_COL in foods_base.columns
    else []
)

base_food_count = len(foods_base)


# ============================
# PAGE 1 — 입력 화면
# ============================
if st.session_state.page == "input":
    h_left, h_right = st.columns([8, 1])

    with h_left:
        st.markdown(
            f"""
<div style="display:flex; align-items:center; gap:8px;">
    <span style="font-size:2.4rem; font-weight:800;">
        질병별 음식 판정 시스템
    </span>
    {tooltip_html}
</div>
            """,
            unsafe_allow_html=True,
        )

    with h_right:
        render_history_popover()

    col1, col2 = st.columns([1, 2])

    with col1:
        selected_diseases = disease_input_block(disease_names_all)

    with col2:
        current_mode = search_engine_block(base_food_count, DETAIL_FOOD_COUNT_APPROX)

        if current_mode == "선택 안 함":
            food_selected, food_row_dict = None, None
        elif current_mode == "대표적인 식품명으로 검색":
            food_selected, food_row_dict = food_input_block_base(
                foods_base, food_names_base, category_base
            )
        else:  # 구체적인 제품명으로 검색
            foods_detail = load_detail_food()
            st.session_state["_detail_foods_df"] = foods_detail
            food_selected, food_row_dict = food_input_block_detail(foods_detail)

    st.markdown("---")

    # OK 버튼
    if st.button("OK", use_container_width=True):
        if st.session_state.search_mode == "선택 안 함":
            st.error("검색 엔진을 먼저 선택하세요.")
        elif not selected_diseases:
            st.error("최소 1개 이상의 질병을 선택하세요.")
        elif not food_selected or food_row_dict is None:
            st.error("음식/제품 선택을 완료하세요.")
        else:
            st.session_state.final_diseases = [str(d) for d in selected_diseases]
            st.session_state.final_food_row = food_row_dict
            st.session_state.final_search_mode = st.session_state.search_mode

            time_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            frow = pd.Series(food_row_dict)

            label = str(frow.get(FOOD_NAME_COL, ""))
            s1 = frow.get(STATE1_COL)
            s2 = frow.get(STATE2_COL)

            if STATE1_COL in frow.index and pd.notna(s1) and str(s1).strip() != "":
                label += f" / {s1}"
            else:
                label += " / 기본"

            if STATE2_COL in frow.index and pd.notna(s2) and str(s2).strip() != "":
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

    left_block, _ = st.columns([0.14, 0.86])
    with left_block:
        b1, b2 = st.columns([1, 1])
        with b1:
            if st.button("문의하기", key="contact_button"):
                open_contact_modal()
        with b2:
            if st.button("사이트 정보", key="siteinfo_button"):
                open_site_info_modal()

# ============================
# PAGE 2 — 결과 화면
# ============================
else:
    st.markdown(
        """
        <style>
        div[data-testid="stTextInputRoot"],
        div[data-testid="stSearchInputContainer"] {
            display: none !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

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
            st.session_state.show_vector_modal = False
            st.rerun()
        st.stop()

    frow = pd.Series(frow_dict)

    food_name_for_icon = str(frow.get(FOOD_NAME_COL, ""))
    food_cat_for_icon = frow.get(CATEGORY_COL, None)
    base_type_for_icon = infer_base_type(food_name_for_icon, food_cat_for_icon)
    icon_map = {
        "면": "🍜",
        "밥": "🍚",
        "생선": "🐟",
        "빵": "🍞",
        "기타": "🥗",
    }
    ai_food_icon = icon_map.get(base_type_for_icon, "🍽️")

    with st.spinner("LOADING..."):
        disease_results = []
        for dname in dnames:
            drow_df = diseases[diseases[DISEASE_NAME_COL] == dname]
            if drow_df.empty:
                continue
            drow = drow_df.iloc[0]
            res = evaluate_row(frow, drow)
            disease_results.append({"name": dname, "row": drow, "res": res})

        if not disease_results:
            st.error("유효한 질병 데이터가 없어 판정을 수행할 수 없습니다.")
            st.stop()

        global_has_fail = any(d["res"]["has_fail"] for d in disease_results)
        global_has_warn = any(d["res"]["has_warning"] for d in disease_results)
        if global_has_fail:
            global_status = "불합격 ❌"
        elif global_has_warn:
            global_status = "주의 ⚠️"
        else:
            global_status = "합격 ✅"

        def sort_key(item):
            r = item["res"]
            if r["has_fail"]:
                return 0
            elif r["has_warning"]:
                return 1
            else:
                return 2

        sorted_results = sorted(disease_results, key=sort_key)

        ai_target_diseases = [
            d["name"]
            for d in disease_results
            if d["res"]["has_fail"] or d["res"]["has_warning"]
        ]

        ai_text = None
        if ai_target_diseases and HAS_OPENAI and client is not None:
            ai_text = get_ai_alternatives(frow, ai_target_diseases)

    st.session_state["_current_disease_results"] = disease_results
    st.session_state["_current_food_row"] = frow_dict

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("질병", ", ".join(d["name"] for d in disease_results))

    with c2:
        label = str(frow.get(FOOD_NAME_COL, ""))
        s1 = frow.get(STATE1_COL)
        s2 = frow.get(STATE2_COL)

        if STATE1_COL in frow.index and pd.notna(s1) and str(s1).strip() != "":
            label += f" / {s1}"
        else:
            label += " / 기본"

        if STATE2_COL in frow.index and pd.notna(s2) and str(s2).strip() != "":
            label += f" / {s2}"

        st.metric("음식", label)

    with c3:
        st.metric("최종 판정", global_status)

    st.markdown("---")

    # ===== AI 대체 음식 추천 =====
    if ai_target_diseases:
        if not ai_text:
            st.info(
                "AI 대체 음식 추천을 사용하려면 .streamlit/secrets.toml "
                "에 [openai] OPENAI_API_KEY 값을 설정하세요."
            )
        else:
            if str(ai_text).lstrip().startswith("⚠️"):
                st.write(ai_text)
            else:
                lines = [ln.strip() for ln in str(ai_text).splitlines() if ln.strip()]
                disclaimer = ""
                rec_lines = lines

                if len(lines) >= 1 and "참고용" in lines[-1]:
                    disclaimer = lines[-1]
                    rec_lines = lines[:-1]

                rows = []
                for line in rec_lines:
                    if " - " in line:
                        name, reason = line.split(" - ", 1)
                    elif "-" in line:
                        name, reason = line.split("-", 1)
                    else:
                        name, reason = line, ""
                    rows.append(
                        {
                            "추천 음식": f"{ai_food_icon} {name.strip()}",
                            "설명": reason.strip(),
                        }
                    )

                df_ai = pd.DataFrame(rows)

                ai_css = """
                <style>
                .ai-card {
                    padding: 0;
                    margin: 0;
                    background-color: transparent;
                    border: none;
                    box-shadow: none;
                    width: 100%;
                }
                .ai-card-title {
                    font-size: 1.7rem;
                    font-weight: 800;
                    margin: 0;
                    text-align: left;
                }
                .ai-card-subtitle {
                    text-align: left;
                    color:#555;
                    margin: 0.25rem 0 0.0rem 0;
                    font-size:1.0rem;
                }
                .ai-alt-table table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-left: auto;
                    margin-right: auto;
                    text-align: center;
                }
                .ai-alt-table th, .ai-alt-table td {
                    text-align: center;
                    font-size: 1.15rem;
                    padding: 0.65rem 1.0rem;
                    border-bottom: 1px solid rgba(0,0,0,0.05);
                }
                .ai-alt-table th {
                    font-weight: 700;
                }
                </style>
                """
                st.markdown(ai_css, unsafe_allow_html=True)

                html_table = df_ai.to_html(
                    index=False,
                    escape=False,
                    border=0,
                )

                header_left, header_right = st.columns([8.7, 1.3])
                with header_left:
                    st.markdown(
                        """
                        <div>
                            <h3 class="ai-card-title">🤖 AI 대체 음식 추천</h3>
                            <p class="ai-card-subtitle">
                                현재 선택한 음식을 대신해 먹을 수 있는 대체 음식입니다.
                            </p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with header_right:
                    st.markdown(
                        "<div style='text-align:right; margin-top:0.6rem; width:100%;'>",
                        unsafe_allow_html=True,
                    )
                    if st.button("다른 알고리즘으로     추천받기", key="vector_button"):
                        st.session_state.show_vector_modal = True
                    st.markdown("</div>", unsafe_allow_html=True)

                st.markdown('<div class="ai-card">', unsafe_allow_html=True)
                st.markdown(
                    f'<div class="ai-alt-table">{html_table}</div>',
                    unsafe_allow_html=True,
                )
                st.markdown("</div>", unsafe_allow_html=True)

                if disclaimer:
                    disclaimer_clean = re.sub(r"<.*?>", "", disclaimer).strip()
                else:
                    disclaimer_clean = (
                        "이 내용은 의료 진단이 아닌 참고용 정보이며, "
                        "실제 식단·치료는 반드시 의료 전문가와 상의해야 합니다."
                    )

                st.markdown(
                    f'<p style="font-size:0.85rem; color:#777; '
                    f'margin-top:0.1rem; margin-bottom:1.2rem; text-align:center;">'
                    f'{disclaimer_clean}'
                    f"</p>",
                    unsafe_allow_html=True,
                )

    if st.session_state.show_vector_modal:
        open_vector_modal()

    # ===== 질병별 상세 표 =====
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

        with st.expander(
            f"세부 영양 비교 — {dname}",
            expanded=True,
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

            st.dataframe(styled, width="stretch")

        st.markdown("---")

    if st.button("처음으로", use_container_width=True):
        st.session_state.page = "input"
        st.session_state.food_selected_final = None
        st.session_state.food_source = None
        st.session_state.show_vector_modal = False
        st.rerun()