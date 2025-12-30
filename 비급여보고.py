# main.py
from __future__ import annotations

import io
import re
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st

APP_VERSION = "2025-12-30 v4 (계산 강제 비필수 + 계산용량 자동매핑 + 디버그표시)"

# ✅ 너가 준 ALIASES 그대로
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

# ✅ 필수 컬럼(계산은 원래부터 제외)
REQUIRED_COLS = ["차트번호", "오더코드", "청구코드", "오더금액", "단가", "일수", "오더명칭"]

# ✅ 위쪽 소계표는 7컬럼(계산 포함)으로 보여줘야 함
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
    """
    alias 기반 rename을 적용한 작업용 df + (표준컬럼 -> 실제 사용된 원본컬럼명) 매핑 리턴
    """
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
    """
    - 계산은 어떤 경우에도 필수검사에서 제외(강제 안전장치)
    - 계산은 계산용량/횟수/수량 등으로 매핑되면 그걸 사용
    - 매핑이 안되면 빈 컬럼 생성
    """
    dfw, picked = _canonical_view(df_original)

    # ✅ 강제 안전장치: 혹시 누가 REQUIRED_COLS에 계산을 넣어도 무조건 제거
    safe_required = [c for c in REQUIRED_COLS if c != "계산"]

    missing = [c for c in safe_required if c not in dfw.columns]
    if missing:
        raise ValueError(
            f"필수 컬럼 누락: {missing}\n"
            f"현재 원본 컬럼(일부): {list(df_original.columns)[:30]} ... (총 {len(df_original.columns)}개)"
        )

    # ✅ 계산 컬럼 확보: 없으면 빈칸 생성
    if "계산" not in dfw.columns:
        dfw["계산"] = ""
        picked["계산"] = "(없음→빈칸생성)"

    chart = dfw["차트번호"].astype(str).str.strip()
    sub = dfw.loc[chart.eq("소계"), DISPLAY_COLS].copy()

    sub["오더금액"] = _to_num(sub["오더금액"])
    sub["단가"] = _to_num(sub["단가"])

    return sub, picked


def _build_excel(
    per_orig: Dict[str, pd.DataFrame],
    per_sub: Dict[str, pd.DataFrame],
    summary_df: pd.DataFrame
) -> bytes:
    out = io.BytesIO()
    used: set[str] = set()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # 요약 시트
        sum_name = _clean_sheet_name("요약", used)
        summary_df.to_excel(writer, sheet_name=sum_name, index=False)

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


# =========================
# UI
# =========================
st.set_page_config(page_title="소계(7컬럼)+원본전체+요약", layout="wide")
st.write(f"✅ RUNNING: {APP_VERSION}")

st.title("엑셀 업로드 → 소계(7컬럼) 위 + 2줄 아래 원본 전체 + 요약")

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

            summary_rows.append({
                "시트(파일명)": label,
                "원본시트": used_sheet,
                "소계 행수": int(len(df_f)),
                "오더금액 합계": float(df_f["오더금액"].sum()),
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
            })

        except Exception as e:
            errors.append(f"[{f.name}] {e}")

    if errors:
        st.error("오류가 발생했습니다. 아래 확인:")
        for msg in errors:
            st.write(f"- {msg}")
        st.stop()

    summary_df = pd.DataFrame(summary_rows).sort_values("오더금액 합계", ascending=False).reset_index(drop=True)
    st.subheader("요약 시트 미리보기 (소계 기준 오더금액 합계)")
    st.dataframe(summary_df, use_container_width=True)

    if show_debug:
        st.subheader("디버그: 표준 컬럼 매핑 결과(어떤 원본 컬럼을 사용했는지)")
        st.dataframe(pd.DataFrame(debug_rows), use_container_width=True)

    excel_bytes = _build_excel(per_orig, per_sub, summary_df)

    st.download_button(
        label="결과 엑셀 다운로드",
        data=excel_bytes,
        file_name="소계필터_상단표_원본전체_요약포함.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.success("완료! 다운로드 버튼으로 결과 엑셀을 받으세요.")
