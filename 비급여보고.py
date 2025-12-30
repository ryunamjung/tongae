# main.py
from __future__ import annotations

import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


# ============================================================
# ✅ 필수 컬럼(계산은 절대 필수 아님!)
# ============================================================
REQUIRED_COLS = ["차트번호", "오더코드", "청구코드", "오더금액", "단가", "일수", "오더명칭"]

# ✅ 소계 표(위쪽 표)는 7컬럼으로 보여줘야 하므로 계산 포함
DISPLAY_COLS = ["오더코드", "청구코드", "오더금액", "단가", "계산", "일수", "오더명칭"]

# ============================================================
# ✅ 컬럼 별칭(엑셀마다 컬럼명이 달라도 자동 매칭)
#   - '계산'은 파일에 없을 수 있으니 대체 후보를 넉넉히 둠
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


def _norm(s: str) -> str:
    s = str(s).replace("\n", "").replace("\r", "")
    s = re.sub(r"\s+", "", s)
    return s.lower()


def _clean_sheet_name(name: str, used: set[str]) -> str:
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
    s = s.astype(str).str.strip()
    s = s.replace({"nan": "", "None": ""})
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace(r"[^0-9\.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0)


def _find_target_sheet(xls: pd.ExcelFile) -> str:
    for sh in xls.sheet_names:
        try:
            head = pd.read_excel(xls, sheet_name=sh, nrows=5)
            cols = [_norm(c) for c in head.columns]
            # 차트번호 후보가 하나라도 있으면 그 시트 사용
            if any(_norm(a) in cols for a in ALIASES["차트번호"]):
                return sh
        except Exception:
            continue
    return xls.sheet_names[0]


def _load_original(uploaded) -> Tuple[pd.DataFrame, str]:
    data = uploaded.getvalue()
    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    sh = _find_target_sheet(xls)
    df = pd.read_excel(xls, sheet_name=sh, engine="openpyxl")
    return df, sh


def _canonical_view(df: pd.DataFrame) -> pd.DataFrame:
    # 정규화 컬럼명 -> 원본 컬럼명
    norm_to_raw = {_norm(c): c for c in df.columns}
    rename_raw_to_std: Dict[str, str] = {}

    for std, candidates in ALIASES.items():
        for cand in candidates:
            k = _norm(cand)
            if k in norm_to_raw:
                rename_raw_to_std[norm_to_raw[k]] = std
                break

    return df.rename(columns=rename_raw_to_std).copy()


def _make_filtered(df_original: pd.DataFrame) -> pd.DataFrame:
    dfw = _canonical_view(df_original)

    # ✅ 여기서 '계산'은 절대 검사하지 않음!
    missing = [c for c in REQUIRED_COLS if c not in dfw.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing}\n현재 컬럼: {list(df_original.columns)}")

    # 계산 컬럼은 없으면 빈칸 생성
    if "계산" not in dfw.columns:
        dfw["계산"] = ""

    chart = dfw["차트번호"].astype(str).str.strip()
    sub = dfw.loc[chart.eq("소계"), DISPLAY_COLS].copy()

    sub["오더금액"] = _to_num(sub["오더금액"])
    sub["단가"] = _to_num(sub["단가"])
    return sub


def _build_excel(per_orig: Dict[str, pd.DataFrame],
                per_sub: Dict[str, pd.DataFrame],
                summary_df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    used: set[str] = set()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        _clean_sheet_name("요약", used)
        summary_df.to_excel(writer, sheet_name="요약", index=False)

        for label in per_orig.keys():
            sh = _clean_sheet_name(label, used)
            df_f = per_sub[label]
            df_o = per_orig[label]

            # 위: 소계(7컬럼)
            df_f.to_excel(writer, sheet_name=sh, index=False, startrow=0)

            # 2줄 아래: 원본 전체(하나도 빠짐없이)
            startrow = (1 + len(df_f)) + 2
            df_o.to_excel(writer, sheet_name=sh, index=False, startrow=startrow)

    out.seek(0)
    return out.getvalue()


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="소계(7컬럼)+원본전체+요약", layout="wide")
st.title("엑셀 업로드 → 소계(7컬럼) 위 + 2줄 아래 원본 전체 + 요약")

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

    summary_df = pd.DataFrame(summary_rows).sort_values("오더금액 합계", ascending=False).reset_index(drop=True)
    st.dataframe(summary_df, use_container_width=True)

    excel_bytes = _build_excel(per_orig, per_sub, summary_df)
    st.download_button(
        "결과 엑셀 다운로드",
        data=excel_bytes,
        file_name="소계필터_상단표_원본전체_요약포함.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.success("완료!")
