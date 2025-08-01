import streamlit as st
import pandas as pd
from io import BytesIO

def process_excel(file):
    df = pd.read_excel(file)
    needed_cols = ['병동','보험','차트번호','환자성명','입원일시','처방의사',
                   '청구코드','오더코드','단가','처방용량','횟수','계산용량',
                   '오더명칭','오더일자','계산유형']
    df = df[needed_cols]
    df = df[df['계산용량'] < 3]
    df['오더금액'] = df['단가'] * df['계산용량']

    summary = df.groupby('오더코드').agg({
        '청구코드': 'first',
        '오더금액': 'sum',
        '단가': 'first',
        '계산용량': 'sum',
        '오더명칭': 'first'
    }).reset_index()
    summary = summary[['오더코드','청구코드','오더금액','단가','계산용량','오더명칭']]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet = 'Sheet1'
        df = df[needed_cols]
        df.to_excel(writer, sheet_name=sheet, startrow=7, index=False)
        summary.to_excel(writer, sheet_name=sheet, startrow=0, startcol=1, index=False)

    processed_data = output.getvalue()

    return df, summary, processed_data

st.title('Excel 데이터 처리 및 저장')

uploaded_file = st.file_uploader('엑셀 파일 업로드', type=['xlsx', 'xls'])
if uploaded_file:
    df, summary, result = process_excel(uploaded_file)

    st.success('처리 완료!')

    st.subheader('원본 데이터 (필터 및 계산 적용 후)')
    st.dataframe(df)

    st.subheader('요약 데이터')
    st.dataframe(summary)


    st.download_button(
        label='처리된 엑셀 파일 다운로드',
        data=result,
        file_name='processed_data.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )







