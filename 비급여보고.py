import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import numbers  # 날짜 포맷용

def process_excel(file):
    # 1. 엑셀 파일 읽기
    df = pd.read_excel(file)

    # 2. 필요한 컬럼만 추출
    needed_cols = ['병동','보험','차트번호','환자성명','입원일시','처방의사',
                   '청구코드','오더코드','단가','처방용량','횟수','계산용량',
                   '오더명칭','오더일자','계산유형']
    df = df[needed_cols]

    # 3. '계산용량' 3 이상인 행 제외
    df = df[df['계산용량'] < 3]

    # 4. '오더금액' 컬럼 만들기 ('단가' * '계산용량')
    df['오더금액'] = df['단가'] * df['계산용량']

    # 5. 요약 데이터 생성 ('오더코드'별 총합)
    summary = df.groupby('오더코드').agg({
        '청구코드': 'first',
        '오더금액': 'sum',
        '단가': 'first',
        '계산용량': 'sum',
        '오더명칭': 'first'
    }).reset_index()

    # 6. 컬럼 순서 정리
    summary = summary[['오더코드','청구코드','오더금액','단가','계산용량','오더명칭']]

    # 7. 엑셀로 저장 (A8부터 df, B1부터 summary)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet = 'Sheet1'
        
        # 원본 데이터 쓰기 (A8부터)
        df.to_excel(writer, sheet_name=sheet, startrow=7, index=False)
        
        # 요약 데이터 쓰기 (B1부터)
        summary.to_excel(writer, sheet_name=sheet, startrow=0, startcol=1, index=False)

        # 날짜 포맷 적용
        workbook = writer.book
        worksheet = writer.sheets[sheet]

        for idx, col_name in enumerate(df.columns):
            if col_name in ['입원일시', '오더일자']:
                col_letter = chr(65 + idx)  # A, B, C...
                for row in range(8, 8 + len(df)):  # 시작은 A8 → 엑셀은 1-index
                    cell = worksheet[f'{col_letter}{row}']
                    cell.number_format = numbers.FORMAT_DATE_DATETIME

    # 엑셀 파일 내용을 Bytes로 변환
    processed_data = output.getvalue()
    return processed_data

# Streamlit UI
st.title('Excel 데이터 처리 및 저장')

uploaded_file = st.file_uploader('엑셀 파일 업로드', type=['xlsx', 'xls'])
if uploaded_file:
    result = process_excel(uploaded_file)
    st.success('처리 완료!')

    st.download_button(
        label='처리된 엑셀 파일 다운로드',
        data=result,
        file_name='processed_data.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )



