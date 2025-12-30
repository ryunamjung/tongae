# main.py
from __future__ import annotations

import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


# =========================
# 설정
# =========================
# "계산"은 파일에 없을 수 있으므로 필수에서 제외
REQUIRED_COLS = ["차트번호", "오더코드", "청구코드", "오더금액", "단가", "일수", "오더명칭"]
# 소계 표는 7컬럼(계산 포함)로 보여줘야 하므로 표시 컬럼엔 포함
FILTER_COLS = ["오더코드", "청구코드", "오더금액", "단가", "계산", "일수", "오더명칭"]


# =========================
# 유틸
# =========================
def _clean_sheet_name(name: str, used: set[str]) -> str:
    """
    엑셀 시트명 제약:
    - 길이 31자 제한
    - \\ / * ? : [ ] 금지
    - 중복이면 _2, _3 ... 붙임
    """
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
        suffix = f"_{i}"
        trimmed = base[: 31 - len(suffix)]
        cand = f"{trimmed}{suffix}"
        if cand not in used:
            used.add(cand)
            return cand
        i += 1


def _to_number_series(s: pd.Series) -> pd.Series:
    """문자/쉼표/공백 섞여도 숫자로 안전 변환"""
    s = s.astype(str).str.strip()
    s = s.replace({"nan": "", "None": ""})
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace(r"[^0-9\.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def _find_target_sheet(xls: pd.ExcelFile) -> str:
    """차트번호 컬럼이 있는 시트를 우선 선택, 없으면 첫 시트"""
    for sh in xls.sheet_names:
        try:
            df_head = pd.read_excel(xls, sheet_name=sh, nrows=5)
            if "차트번호" in df_head.columns:
                return sh
        except Exception:
            continue
    return xls.sheet_names[0]


def load_original_df(uploaded) -> Tuple[pd.DataFrame, str]:
    """원본 DF를 그대로 읽어오기"""
    data = uploaded.getvalue()
    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    target_sheet = _find_target_sheet(xls)
    df = pd.read_excel(xls, sheet_name=target_sheet, engine="openpyxl")
    return df, target_sheet


def make_filtered_df(df_original: pd.DataFrame) -> pd.DataFrame:
    """
    1) 필수 컬럼 검사(계산 제외)
    2) 계산 컬럼 없으면 빈 컬럼 생성
    3) 차트번호 == 소계 필터
    4) 7컬럼만 추출
    """
    missing = [c for c in REQUIRED_COLS if c not in df_original.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing}")

    # 계산 컬럼은 없을 수 있으니 없으면 생성(빈값)
    if "계산" not in df_original.columns:
        df_original = df_original.copy()
        df_original["계산"] = ""

    chart = df_original["차트번호"].astype(str).str.strip()
    sub = df_original.loc[chart.eq("소계"), FILTER_COLS].copy()

    # 오더금액/단가 숫자 변환
    sub["오더금액"] = _to_number_series(sub["오더금액"])
    sub["단가"] = _to_number_series(sub["단가"])

    return sub


def build_output_excel(
    per_file_original: Dict[str, pd.DataFrame],
    per_file_filtered: Dict[str, pd.DataFrame],
    summary_df: pd.DataFrame
) -> bytes:
    """
    저장 규칙(요구사항 그대로):
    - 각 파일 시트:
      (1) 위: 소계(7컬럼) 표 (헤더 포함)
      (2) 그 아래 2줄 띄우고
      (3) 원본 전체를 행/열 하나도 빠짐없이 그대로 (헤더 포함)
    - 요약 시트:
      파일(시트)별 소계(7컬럼)에서 오더금액 합계만
    """
    out = io.BytesIO()
    used_sheet_names: set[str] = set()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # 요약 시트
        sum_sheet = _clean_sheet_name("요약", used_sheet_names)
        summary_df.to_excel(writer, sheet_name=sum_sheet, index=False)

        # 파일별 시트
        for file_label in per_file_original.keys():
            sh = _clean_sheet_name(file_label, used_sheet_names)
            df_f = per_file_filtered[file_label]
            df_o = per_file_original[file_label]

            # (1) 위쪽: 소계(7컬럼) 표 (A1부터)
            df_f.to_excel(writer, sheet_name=sh, index=False, startrow=0)

            # (2) + (3) : 2줄 아래 원본 전체
            # 소계 표가 차지하는 행 수 = 헤더 1 + 데이터 len
            filtered_rows = 1 + len(df_f)
            startrow_original = filtered_rows + 2  # 정확히 2줄 띄움
            df_o.to_excel(writer, sheet_name=sh, index=False, startrow=startrow_original)

    out.seek(0)
    return out.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="소계(7컬럼) + 2줄 아래 원본 전체 + 요약", layout="wide")

st.title("엑셀 업로드 → 소계(7컬럼) 위에 표시 + 2줄 아래 원본 전체 그대로 + 요약")

st.markdown(
    """
### 저장 결과(각 파일 시트)
- **위쪽**: `차트번호 == "소계"` 행만 필터 →  
  `오더코드, 청구코드, 오더금액, 단가, 계산, 일수, 오더명칭` **7컬럼만 표로 표시**
- 그 아래 **2줄 띄우고**
- **원본 데이터를 행/열 하나도 빠짐없이 그대로 출력**

### 요약 시트
- 파일(시트)별로, 위 소계(7컬럼) 범위에서 **오더금액 합계**만 표 생성

> ⚠️ `계산` 컬럼이 파일에 없으면, 소계표에서는 **빈 컬럼으로 자동 생성**합니다.
"""
)

files = st.file_uploader("엑셀 파일 여러 개 업로드(.xlsx)", type=["xlsx"], accept_multiple_files=True)

if not files:
    st.info("엑셀 파일을 업로드하세요.")
    st.stop()

if st.button("처리 & 결과 생성", type="primary"):
    per_file_original: Dict[str, pd.DataFrame] = {}
    per_file_filtered: Dict[str, pd.DataFrame] = {}
    summary_rows: List[dict] = []
    errors: List[str] = []

    for f in files:
        try:
            label = re.sub(r"\.xlsx$", "", f.name, flags=re.IGNORECASE).strip() or f.name

            # 원본 로드(그대로)
            df_o, used_sheet = load_original_df(f)

            # 소계(7컬럼) 표 생성 (계산 없으면 빈컬럼으로 자동 생성)
            df_f = make_filtered_df(df_o)

            per_file_original[label] = df_o
            per_file_filtered[label] = df_f

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

    st.subheader("요약 시트 미리보기 (소계 7컬럼 기준 오더금액 합계)")
    st.dataframe(summary_df, use_container_width=True)

    st.subheader("파일별 미리보기")
    tabs = st.tabs(list(per_file_original.keys()))
    for name, tab in zip(per_file_original.keys(), tabs):
        with tab:
            st.write(f"### {name}")
            st.write("**(위) 소계(7컬럼)**")
            st.dataframe(per_file_filtered[name], use_container_width=True)
            st.write("**(아래) 원본 전체(하나도 빠짐없이)**")
            st.dataframe(per_file_original[name], use_container_width=True)

    excel_bytes = build_output_excel(per_file_original, per_file_filtered, summary_df)

    st.download_button(
        label="결과 엑셀 다운로드",
        data=excel_bytes,
        file_name="소계필터_상단표_원본전체_요약포함.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.success("완료! 다운로드 버튼으로 결과 엑셀을 받으세요.")
