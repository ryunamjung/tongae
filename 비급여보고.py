# main.py
from __future__ import annotations

import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

# ============================================================
# ✅ 버전 표시 (지금 실행 중인 파일 확인용)
# ============================================================
APP_VERSION = "2025-12-30 v3 (계산 절대 비필수 / ALIASES 적용)"

# ============================================================
# ✅ 너가 준 ALIASES 그대로 사용
# ============================================================
ALIASES: Dict[str, List[str]] = {
    "차트번호": ["차트번호", "차트", "chartno", "chart_no", "chart"],
    "오더코드": ["오더코드", "오더 코드", "처방코드", "처방 코드", "ordercode", "order_code"],
    "청구코드": ["청구코드", "청구 코드", "edi코드", "edi 코드", "claimcode", "claim_code"],
    "오더금액": ["오더금액", "오더 금액", "금액", "청구금액", "amount", "orderamt", "order_amt"],
    "단가": ["단가", "수가", "수가단가", "price", "unitprice", "unit_price"],
    "계산": ["계산", "계산용량", "계산수량", "수량", "횟수", "산정횟수", "산정수량", "qty", "quantity"],
    "일수": ["일수", "days", "day", "기간", "투약일수", "재원일수"],
    "오더명칭": ["오더명칭", "오더 명칭", "처방명", "처방명칭", "명칭", "항목명", "ordername", "order_name"],
}

# ============================================================
# ✅ 필수 컬럼: '계산'은 절대 포함하지 않는다
# ============================================================
REQUIRED_COLS = ["차트번호", "오더코드", "청구코드", "오더금액", "단가", "일수", "오더명칭"]
# 위 소계표는 7컬럼 보여야 하므로 계산 포함
DISPLAY_COLS = ["오더코드", "청구코드", "오더금액", "단가", "계산", "일수", "오더명칭"]


# =========================
# 유틸
# =========================
def _norm(s: str) -> str:
    """컬럼 비교용 정규화: 공백/개행 제거 + 소문자"""
    s = str(s).replace("\n", "").replace("\r", "")
    s = re.sub(r"\s+", "", s)
    return s.lower()


def _clean_sheet_name(name: str, used: set[str]) -> str:
    """엑셀 시트명 제약(31자/금지문자/중복) 처리"""
    name = (name or "").strip()
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    if not name:
        name = "Sheet"
    base = name[:31]
    if base not in used:
        used.add(base)
        return base
    i = 2
    while True:
        suf = f"_{i}"
        cand = f"{base[:31-len(suf)]}{suf}"
        if cand not in used:
            used.add(cand)
            return cand
        i += 1


def _to_num(s: pd.Series) -> pd.Series:
    """문자/쉼표/공백 섞여도 숫자로 안전 변환"""
    s = s.astype(str).str.strip()
    s = s.replace({"nan": "", "None": ""})
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace(r"[^0-9\.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def _find_target_sheet(xls: pd.ExcelFile) -> str:
    """
    차트번호 후보 컬럼이 있는 시트를 우선 선택.
    없으면 첫 시트.
    """
    candidates_norm = [_norm(x) for x in ALIASES.get("차트번호", ["차트번호"])]
    for sh in xls.sheet_names:
        try:
            head = pd.read_excel(xls, sheet_name=sh, nrows=5)
            cols_norm = [_norm(c) for c in head.columns]
            if any(cn in cols_norm for cn in candidates_norm):
                return sh
        except Exception:
            continue
    return xls.sheet_names[0]


def _load_original(uploaded) -> Tuple[pd.DataFrame, str]:
    """원본 DF를 그대로 읽어오기"""
    data = uploaded.getvalue()
    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    sh = _find_target_sheet(xls)
    df = pd.read_excel(xls, sheet_name=sh, engine="openpyxl")
    return df, sh


def _canonical_view(df: pd.DataFrame) -> pd.DataFrame:
    """
    원본은 그대로 두고, 필터/요약을 위한 '표준 컬럼 뷰'만 만든다.
    - ALIASES 기반으로 rename 시도
    """
    norm_to_raw = {_norm(c): c for c in df.columns}  # 정규화 -> 원본 컬럼명

    rename_map: Dict[str, str] = {}  # 원본컬럼명 -> 표준컬럼명
    for std, cand_list in ALIASES.items():
        for cand in cand_list:
            k = _norm(cand)
            if k in norm_to_raw:
                rename_map[norm_to_raw[k]] = std
                break

    return df.rename(columns=rename_map).copy()


def _make_filtered(df_original: pd.DataFrame) -> pd.DataFrame:
    """
    1) 표준뷰 생성(aliases rename)
    2) REQUIRED_COLS(계산 제외)만 검사
    3) 계산 없으면 빈 컬럼 생성
    4) 차트번호=='소계' 필터
    5) 소계표는 DISPLAY_COLS(7컬럼)만 추출
    """
    dfw = _canonical_view(df_original)

    # ✅ 절대 '계산'은 검사하지 않는다
    missing = [c for c in REQUIRED_COLS if c not in dfw.columns]
    if missing:
        raise ValueError(
            f"필수 컬럼 누락: {missing}\n"
            f"현재 원본 컬럼: {list(df_original.columns)}"
        )

    if "계산" not in dfw.columns:
        dfw["계산"] = ""  # 없으면 빈 칸

    chart = dfw["차트번호"].astype(str).str.strip()
    sub = dfw.loc[chart.eq("소계"), DISPLAY_COLS].copy()

    # 숫자 변환
    sub["오더금액"] = _to_num(sub["오더금액"])
    sub["단가"] = _to_num(sub["단가"])

    return sub


def _build_excel(
    per_orig: Dict[str, pd.DataFrame],
    per_sub: Dict[str, pd.DataFrame],
    summary_df: pd.DataFrame
) -> bytes:
    """
    저장 규칙:
    - 각 파일 시트:
      (1) 위: 소계(7컬럼) 표 (헤더 포함)
      (2) 그 아래 2줄 띄우고
      (3) 원본 전체를 행/열 하나도 빠짐없이 그대로(헤더 포함)
    - 요약 시트:
      파일별 소계(7컬럼) 오더금액 합계
    """
    out = io.BytesIO()
    used: set[str] = set()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # 요약 시트
        sum_name = _clean_sheet_name("요약", used)
        summary_df.to_excel(writer, sheet_name=sum_name, index=False)

        # 파일별 시트
        for label in per_orig.keys():
            sh = _clean_sheet_name(label, used)

            df_f = per_sub[label]
            df_o = per_orig[label]

            # 1) 위쪽: 소계표
            df_f.to_excel(writer, sheet_name=sh, index=False, startrow=0)

            # 2줄 아래: 원본 전체
            startrow = (1 + len(df_f)) + 2
            df_o.to_excel(writer, sheet_name=sh, index=False, startrow=startrow)

    out.seek(0)
    return out.getvalue()


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="소계(7컬럼)+원본전체+요약", layout="wide")
st.write(f"✅ RUNNING: {APP_VERSION}")  # 실행 파일 확인용

st.title("엑셀 업로드 → 소계(7컬럼) 위 + 2줄 아래 원본 전체 + 요약")

st.markdown(
    """
- 위쪽: `차트번호 == "소계"` 행만 → 7컬럼(오더코드/청구코드/오더금액/단가/계산/일수/오더명칭)
- 아래쪽: 2줄 띄우고 **원본 전체를 하나도 빠짐없이 그대로**
- 요약 시트: 파일(시트)별 **소계 범위 오더금액 합계**
- `계산` 컬럼이 파일에 없으면 자동으로 **빈 컬럼 생성** (절대 에러 안 남)
"""
)

files = st.file_uploader("엑셀 파일 여러 개 업로드(.xlsx)", type=["xlsx"], accept_multiple_files=True)
if not files:
    st.stop()

if st.button("처리 & 결과 생성", type="primary"):
    per_orig: Dict[str, pd.DataFrame] = {}
    per_sub: Dict[str, pd.DataFrame] = {}
    summary_rows: List[dict] = []
    errors: List[str] = []

    for f in files:
        try:
            label = re.sub(r"\.xlsx$", "", f.name, flags=re.IGNORECASE).strip() or f.name

            df_o, used_sheet = _load_original(f)
            df_f = _make_filtered(df_o)

            per_orig[label] = df_o
            per_sub[label] = df_f

            summary_rows.append({
                "시트(파일명)": label,
                "원본시트": used_sheet,
                "소계 행수": int(len(df_f)),
                "오더금액 합계": float(df_f["오더금액"].sum()),
            })
        except Exception as e:
            errors.append(f"[{f.name}] {e}")

    if errors:
        st.error("오류가 발생했습니다. 아래 확인:")
        for msg in errors:
            st.write(f"- {msg}")
        st.stop()

    summary_df = (
        pd.DataFrame(summary_rows)
        .sort_values("오더금액 합계", ascending=False)
        .reset_index(drop=True)
    )

    st.subheader("요약 시트 미리보기 (소계 기준 오더금액 합계)")
    st.dataframe(summary_df, use_container_width=True)

    st.subheader("파일별 미리보기")
    tabs = st.tabs(list(per_orig.keys()))
    for name, tab in zip(per_orig.keys(), tabs):
        with tab:
            st.write(f"### {name}")
            st.write("**(위) 소계(7컬럼)**")
            st.dataframe(per_sub[name], use_container_width=True)
            st.write("**(아래) 원본 전체(하나도 빠짐없이)**")
            st.dataframe(per_orig[name], use_container_width=True)

    excel_bytes = _build_excel(per_orig, per_sub, summary_df)

    st.download_button(
        label="결과 엑셀 다운로드",
        data=excel_bytes,
        file_name="소계필터_상단표_원본전체_요약포함.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.success("완료! 다운로드 버튼으로 결과 엑셀을 받으세요.")
