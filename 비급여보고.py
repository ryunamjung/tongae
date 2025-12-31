# main.py
from __future__ import annotations

import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

APP_VERSION = "비급여보고통계 남양주백병원 / vCalcSubtotalOnly"

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

REQUIRED_COLS = ["차트번호", "오더코드", "청구코드", "오더금액", "단가", "일수", "오더명칭"]
DISPLAY_COLS = ["오더코드", "청구코드", "오더금액", "단가", "계산", "일수", "오더명칭"]


def _norm(s: str) -> str:
    s = str(s).replace("\n", "").replace("\r", "")
    s = re.sub(r"\s+", "", s)
    return s.lower()


def _to_num(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip()
    s = s.replace({"nan": "", "None": ""})
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace(r"[^0-9\.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0)


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


def _find_target_sheet(xls: pd.ExcelFile) -> str:
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
    data = uploaded.getvalue()
    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    sh = _find_target_sheet(xls)
    df = pd.read_excel(xls, sheet_name=sh, engine="openpyxl")
    return df, sh


def _canonical_view(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    norm_to_raw = {_norm(c): c for c in df.columns}
    rename_map_raw_to_std: Dict[str, str] = {}
    picked_std_to_raw: Dict[str, str] = {}

    for std, cand_list in ALIASES.items():
        for cand in cand_list:
            k = _norm(cand)
            if k in norm_to_raw:
                raw_col = norm_to_raw[k]
                rename_map_raw_to_std[raw_col] = std
                picked_std_to_raw[std] = raw_col
                break

    dfw = df.rename(columns=rename_map_raw_to_std).copy()
    return dfw, picked_std_to_raw


def _make_filtered(df_original: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    dfw, picked = _canonical_view(df_original)

    safe_required = [c for c in REQUIRED_COLS if c != "계산"]
    missing = [c for c in safe_required if c not in dfw.columns]
    if missing:
        raise ValueError(
            f"필수 컬럼 누락: {missing}\n"
            f"현재 원본 컬럼(일부): {list(df_original.columns)[:30]} ... (총 {len(df_original.columns)}개)"
        )

    if "계산" not in dfw.columns:
        dfw["계산"] = ""
        picked["계산"] = "(없음→빈칸생성)"

    # ✅ 소계 판정 강화 ("소 계"도 소계로)
    chart_raw = dfw["차트번호"].astype(str).fillna("").str.strip()
    chart_key = chart_raw.str.replace(r"\s+", "", regex=True)
    is_subtotal = chart_key.eq("소계")

    sub = dfw.loc[is_subtotal, DISPLAY_COLS].copy()

    # ✅ 오더금액 합계처럼, 소계표(df_f) 안에서 계산도 숫자화해 둠
    sub["오더금액"] = _to_num(sub["오더금액"])
    sub["단가"] = _to_num(sub["단가"])
    sub["계산"] = _to_num(sub["계산"])   # ✅ 핵심

    return sub, picked


def _build_excel(
    per_orig: Dict[str, pd.DataFrame],
    per_sub: Dict[str, pd.DataFrame],
    summary_df: pd.DataFrame
) -> bytes:
    out = io.BytesIO()
    used: set[str] = set()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        sum_name = _clean_sheet_name("요약", used)
        summary_df.to_excel(writer, sheet_name=sum_name, index=False)

        for label in per_orig.keys():
            sh = _clean_sheet_name(label, used)
            df_f = per_sub[label]
            df_o = per_orig[label]

            df_f.to_excel(writer, sheet_name=sh, index=False, startrow=0)
            startrow = (1 + len(df_f)) + 2
            df_o.to_excel(writer, sheet_name=sh, index=False, startrow=startrow)

    out.seek(0)
    return out.getvalue()


# =========================
# UI
# =========================
st.set_page_config(page_title="비급여관련 통계도우미", layout="wide")
st.write(f"✅ RUNNING: {APP_VERSION}")

st.title("코드별 엑셀파일 xlsx로 재저장후 여러파일 한번에 업로드 → 요약, 각각 파일별 sheet 구성됨")

files = st.file_uploader("엑셀 파일 여러 개 업로드(.xlsx)", type=["xlsx"], accept_multiple_files=True)
if not files:
    st.stop()

show_debug = st.checkbox("디버그(컬럼 매핑 확인) 보기", value=True)

if st.button("처리 & 결과 생성", type="primary"):
    per_orig: Dict[str, pd.DataFrame] = {}
    per_sub: Dict[str, pd.DataFrame] = {}
    summary_rows: List[dict] = []
    debug_rows: List[dict] = []
    errors: List[str] = []

    for f in files:
        try:
            label = re.sub(r"\.xlsx$", "", f.name, flags=re.IGNORECASE).strip() or f.name

            df_o, used_sheet = _load_original(f)
            df_f, picked = _make_filtered(df_o)

            per_orig[label] = df_o
            per_sub[label] = df_f

            # ✅ 오더금액합계와 동일 방식: 소계표(df_f)에서 합산
            order_sum = float(df_f["오더금액"].sum()) if "오더금액" in df_f.columns else 0.0
            calc_sum = float(df_f["계산"].sum()) if "계산" in df_f.columns else 0.0

            summary_rows.append({
                "시트(파일명)": label,
                "원본시트": used_sheet,
                "소계 행수": int(len(df_f)),
                "오더금액 합계": order_sum,
                "계산 합계": calc_sum,  # ✅ 요청: 소계(df_f)에서 계산용량(=계산) 합계
            })

            debug_rows.append({
                "파일": f.name,
                "원본시트": used_sheet,
                "차트번호(매핑)": picked.get("차트번호", ""),
                "오더코드(매핑)": picked.get("오더코드", ""),
                "청구코드(매핑)": picked.get("청구코드", ""),
                "오더금액(매핑)": picked.get("오더금액", ""),
                "단가(매핑)": picked.get("단가", ""),
                "계산(매핑)": picked.get("계산", ""),
                "일수(매핑)": picked.get("일수", ""),
                "오더명칭(매핑)": picked.get("오더명칭", ""),
                "소계표 계산 상위5": ", ".join(map(str, df_f["계산"].head(5).tolist())) if "계산" in df_f.columns else "",
                "소계표 계산 dtype": str(df_f["계산"].dtype) if "계산" in df_f.columns else "(없음)",
            })

        except Exception as e:
            errors.append(f"[{f.name}] {e}")

    if errors:
        st.error("오류가 발생했습니다. 아래 확인:")
        for msg in errors:
            st.write(f"- {msg}")
        st.stop()

    summary_df = pd.DataFrame(summary_rows).sort_values("오더금액 합계", ascending=False).reset_index(drop=True)
    st.subheader("요약 시트 미리보기 (소계 기준 오더금액 합계 + 계산 합계)")
    st.dataframe(summary_df, use_container_width=True)

    if show_debug:
        st.subheader("디버그: 표준 컬럼 매핑 결과/값 확인")
        st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)

    excel_bytes = _build_excel(per_orig, per_sub, summary_df)

    st.download_button(
        label="결과 엑셀 다운로드",
        data=excel_bytes,
        file_name="소계필터_상단표_원본전체_요약포함.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.success("완료! 다운로드 버튼으로 결과 엑셀을 받으세요.")
