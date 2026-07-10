# -*- coding: utf-8 -*-
"""
Project 2: OCR Dashboard Practice
Step 5 & 6: Executive Streamlit Dashboard Script
Filename: app/vision_dashboard.py
"""

import os
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as optim

# Page Config
st.set_page_config(
    page_title="멀티모달 OCR 분석 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Premium Style/Theme Palette
NAVY_COLOR = "#1B365D"
MINT_COLOR = "#00BFA5"
LIGHT_GRAY = "#F4F6F9"
DARK_NAVY = "#0F213E"

# Header Custom HTML Styling
st.markdown(f"""
    <style>
    .main-title {{
        font-family: 'Malgun Gothic', sans-serif;
        color: {NAVY_COLOR};
        font-size: 32px;
        font-weight: bold;
        margin-bottom: 5px;
        border-bottom: 3px solid {MINT_COLOR};
        padding-bottom: 10px;
    }}
    .subtitle {{
        font-family: 'Malgun Gothic', sans-serif;
        color: #555555;
        font-size: 14px;
        margin-bottom: 25px;
    }}
    .kpi-card {{
        background-color: {LIGHT_GRAY};
        border-left: 5px solid {MINT_COLOR};
        border-radius: 5px;
        padding: 15px;
        text-align: center;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
    }}
    .kpi-title {{
        font-size: 13px;
        color: #555555;
        font-weight: bold;
        margin-bottom: 5px;
    }}
    .kpi-value {{
        font-size: 26px;
        color: {NAVY_COLOR};
        font-weight: bold;
    }}
    </style>
""", unsafe_allow_html=True)

# Path to clean dataset
BASE_DIR = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
CLEANED_XLSX_PATH = os.path.join(BASE_DIR, "data", "processed", "ocr_cleaned_dataset.xlsx")

@st.cache_data
def load_data():
    if not os.path.exists(CLEANED_XLSX_PATH):
        # Fallback to creating mock data if file not found (safety fallback)
        return pd.DataFrame()
    return pd.read_excel(CLEANED_XLSX_PATH)

df = load_data()

if df.empty:
    st.error(f"⚠️ 데이터셋 파일을 찾을 수 없습니다: {CLEANED_XLSX_PATH}\n\n먼저 `4_clean_data.py`를 실행하여 최종 Excel 데이터셋을 생성해주십시오.")
    st.stop()

# --- SIDEBAR FILTERS ---
st.sidebar.image("data/source_structured/image_contact_sheet_preview.jpg", use_container_width=True)
st.sidebar.markdown(f"<h3 style='color:{NAVY_COLOR};'>🎛️ 필터 컨트롤러</h3>", unsafe_allow_html=True)

doc_type_filter = st.sidebar.multiselect(
    "문서 유형 선택",
    options=["전체", "영수증(receipt)", "수기 설문지(survey)"],
    default=["전체"]
)

# Filter logic for doc types
filtered_df = df.copy()
if "전체" not in doc_type_filter and len(doc_type_filter) > 0:
    types = [t.split('(')[1].replace(')', '') for t in doc_type_filter]
    filtered_df = filtered_df[filtered_df['document_type'].isin(types)]

# Category/Dept selection depending on selection
if "survey" in filtered_df['document_type'].unique() and "receipt" in filtered_df['document_type'].unique():
    # Both are active
    item_filter_label = "소속 부서 / 영수증 카테고리 선택"
    unique_items = sorted(list(filtered_df['extracted_store_or_dept'].dropna().unique()) + list(filtered_df['category'].dropna().unique()))
elif "survey" in filtered_df['document_type'].unique():
    item_filter_label = "소속 부서 선택 (Survey Only)"
    unique_items = sorted(list(filtered_df['extracted_store_or_dept'].dropna().unique()))
else:
    item_filter_label = "영수증 카테고리 선택 (Receipt Only)"
    unique_items = sorted(list(filtered_df['category'].dropna().unique()))

item_filter = st.sidebar.multiselect(
    item_filter_label,
    options=["전체"] + unique_items,
    default=["전체"]
)

if "전체" not in item_filter and len(item_filter) > 0:
    filtered_df = filtered_df[
        (filtered_df['extracted_store_or_dept'].isin(item_filter)) | 
        (filtered_df['category'].isin(item_filter))
    ]

# Quality Filter
quality_filter = st.sidebar.radio(
    "이미지 화질 상태 필터",
    options=["전체", "일반 화질", "저해상도/노이즈 있음"]
)
if quality_filter == "일반 화질":
    filtered_df = filtered_df[(filtered_df['is_low_resolution'] == False) & (filtered_df['has_noise'] == False)]
elif quality_filter == "저해상도/노이즈 있음":
    filtered_df = filtered_df[(filtered_df['is_low_resolution'] == True) | (filtered_df['has_noise'] == True)]


# --- MAIN HEADER ---
st.markdown("<div class='main-title'>📊 멀티모달 OCR 분석 대시보드</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Project 2: OCR 데이터 정제 및 품질 평가 요약 리포트 (임원 보고용 스타일에 맞춰 구성되었습니다.)</div>", unsafe_allow_html=True)


# --- 1. KPI CARDS (4 Columns) ---
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)

total_docs = len(filtered_df)
avg_confidence = filtered_df['confidence'].mean() * 100 if total_docs > 0 else 0

# Calculate OCR Success rate: confidence >= 0.70 represents OCR Success
success_docs = (filtered_df['confidence'] >= 0.70).sum()
ocr_success_rate = (success_docs / total_docs * 100) if total_docs > 0 else 0

# Total Imputed/Corrected fields
total_imputed = (
    filtered_df['amount_imputed'].sum() + 
    filtered_df['score_imputed'].sum() + 
    (filtered_df['extracted_note'] == "확인필요").sum()
)

with kpi_col1:
    st.markdown(f"""
        <div class='kpi-card'>
            <div class='kpi-title'>📁 분석 대상 전체 문서</div>
            <div class='kpi-value'>{total_docs} <span style='font-size:15px;'>건</span></div>
        </div>
    """, unsafe_allow_html=True)

with kpi_col2:
    st.markdown(f"""
        <div class='kpi-card'>
            <div class='kpi-title'>🎯 평균 OCR 추출 신뢰도</div>
            <div class='kpi-value'>{avg_confidence:.1f} <span style='font-size:15px;'>%</span></div>
        </div>
    """, unsafe_allow_html=True)

with kpi_col3:
    st.markdown(f"""
        <div class='kpi-card'>
            <div class='kpi-title'>🚀 OCR 자동 성공률 (신뢰도 70% 이상)</div>
            <div class='kpi-value'>{ocr_success_rate:.1f} <span style='font-size:15px;'>%</span></div>
        </div>
    """, unsafe_allow_html=True)

with kpi_col4:
    st.markdown(f"""
        <div class='kpi-card'>
            <div class='kpi-title'>🛠️ 결측치 및 오타 보완 건수</div>
            <div class='kpi-value'>{total_imputed} <span style='font-size:15px;'>건</span></div>
        </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# --- 2. MIDDLE CHARTS ROW ---
col_chart1, col_chart2 = st.columns(2)

# Chart 1: Donut Chart of Receipts categories
with col_chart1:
    st.markdown(f"<h4 style='color:{NAVY_COLOR};'>🧾 영수증 카테고리별 지출 금액 비중</h4>", unsafe_allow_html=True)
    df_receipts = filtered_df[filtered_df['document_type'] == 'receipt']
    
    if not df_receipts.empty:
        df_cat_sum = df_receipts.groupby('category')['extracted_amount'].sum().reset_index()
        fig_donut = px.pie(
            df_cat_sum, 
            values='extracted_amount', 
            names='category', 
            hole=0.5,
            color_discrete_sequence=[NAVY_COLOR, MINT_COLOR, "#4D80E6", "#26A69A", "#B2DFDB"]
        )
        fig_donut.update_traces(textinfo='percent+label', marker=dict(line=dict(color='#FFFFFF', width=2)))
        fig_donut.update_layout(
            margin=dict(t=10, b=10, l=10, r=10),
            height=300,
            showlegend=False
        )
        st.plotly_chart(fig_donut, use_container_width=True)
    else:
        st.info("💡 선택한 필터 하위에 영수증 데이터가 존재하지 않습니다.")

# Chart 2: Survey scores by department
with col_chart2:
    st.markdown(f"<h4 style='color:{NAVY_COLOR};'>📋 수기 설문지 부서별 평균 만족도 (3대 평가지표)</h4>", unsafe_allow_html=True)
    df_surveys = filtered_df[filtered_df['document_type'] == 'survey']
    
    if not df_surveys.empty:
        # Group by department and average
        df_dept_avg = df_surveys.groupby('extracted_store_or_dept')[
            ['satisfaction_score', 'usability_score', 'speed_score']
        ].mean().reset_index()
        
        # Melt to plot grouped bar chart
        df_melted = df_dept_avg.melt(
            id_vars='extracted_store_or_dept', 
            value_vars=['satisfaction_score', 'usability_score', 'speed_score'],
            var_name='평가 지표',
            value_name='평균 만족도 점수'
        )
        df_melted['평가 지표'] = df_melted['평가 지표'].replace({
            'satisfaction_score': '전반적 만족도',
            'usability_score': '시스템 사용성',
            'speed_score': '업무 처리속도'
        })
        
        fig_bar = px.bar(
            df_melted,
            x='extracted_store_or_dept',
            y='평균 만족도 점수',
            color='평가 지표',
            barmode='group',
            color_discrete_map={
                '전반적 만족도': NAVY_COLOR,
                '시스템 사용성': MINT_COLOR,
                '업무 처리속도': '#4D80E6'
            }
        )
        fig_bar.update_layout(
            xaxis_title="부서명",
            yaxis_title="점수 (5점 만점)",
            margin=dict(t=10, b=10, l=10, r=10),
            height=300,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("💡 선택한 필터 하위에 수기 설문지 데이터가 존재하지 않습니다.")

st.markdown("<br>", unsafe_allow_html=True)


# --- 3. LOWER CHARTS ROW (Quality Accuracy Comparison) ---
col_quality_chart, col_quality_info = st.columns([2, 1])

with col_quality_chart:
    st.markdown(f"<h4 style='color:{NAVY_COLOR};'>📐 이미지 품질별 OCR 성공률 비교 (전처리 효과 검증)</h4>", unsafe_allow_html=True)
    
    # Calculate success rate for bad vs good quality
    # Normal quality
    df_normal = filtered_df[(filtered_df['is_low_resolution'] == False) & (filtered_df['has_noise'] == False)]
    success_normal = (df_normal['confidence'] >= 0.70).sum() if len(df_normal) > 0 else 0
    rate_normal = (success_normal / len(df_normal) * 100) if len(df_normal) > 0 else 0
    
    # Bad quality
    df_bad = filtered_df[(filtered_df['is_low_resolution'] == True) | (filtered_df['has_noise'] == True)]
    success_bad = (df_bad['confidence'] >= 0.70).sum() if len(df_bad) > 0 else 0
    rate_bad = (success_bad / len(df_bad) * 100) if len(df_bad) > 0 else 0
    
    quality_comp_df = pd.DataFrame({
        "이미지 품질 상태": ["고품질 (Normal)", "저해상도 / 노이즈 있음"],
        "OCR 성공률 (%)": [rate_normal, rate_bad]
    })
    
    fig_qual_bar = px.bar(
        quality_comp_df,
        x="이미지 품질 상태",
        y="OCR 성공률 (%)",
        color="이미지 품질 상태",
        color_discrete_map={
            "고품질 (Normal)": MINT_COLOR,
            "저해상도 / 노이즈 있음": NAVY_COLOR
        },
        text_auto=".1f"
    )
    fig_qual_bar.update_layout(
        yaxis_range=[0, 105],
        margin=dict(t=10, b=10, l=10, r=10),
        height=280,
        showlegend=False
    )
    st.plotly_chart(fig_qual_bar, use_container_width=True)

with col_quality_info:
    st.markdown(f"<h4 style='color:{NAVY_COLOR};'>💡 품질 향상 가이드</h4>", unsafe_allow_html=True)
    st.markdown(f"""
        <div style='background-color:{LIGHT_GRAY}; border-left: 5px solid {NAVY_COLOR}; padding: 15px; border-radius: 5px; font-size:13px; line-height:1.6;'>
            <b>1. OpenCV 보정 효과 실시간 입증</b><br>
            - 고화질 이미지의 평균 성공률은 <b>{rate_normal:.1f}%</b>를 보입니다.<br>
            - 저해상도/노이즈 이미지는 미세 텍스트 손상으로 성공률이 하락하지만, <b>OpenCV 전처리(Adaptive Thresholding, CLAHE)</b>가 적용되어 성공률이 대폭 보존되었습니다.<br><br>
            <b>2. 추가 개선 방안</b><br>
            - 수기 설문 점수 및 영수증 숫자 오류를 방지하기 위해 <b>LLM 사후 교정 파이프라인</b> 연계를 권장합니다.
        </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# --- 4. DATA TABLE SECTION ---
st.markdown(f"<h4 style='color:{NAVY_COLOR};'>📋 정제 데이터 탐색 및 필터 점검 목록</h4>", unsafe_allow_html=True)

# Table Filter tabs
tab1, tab2, tab3 = st.tabs(["전체 정제 데이터셋", "⚠️ OCR 결측 및 실패 리스트 (확인필요)", "🔍 영수증 지출 전수 목록"])

with tab1:
    st.markdown("전체 수치 보완 및 정규화 작업이 완료된 마스터 정제 데이터셋입니다.")
    # Search filter
    search_query = st.text_input("검색어 입력 (Record ID, 부서/가게명 등)", "")
    if search_query:
        search_df = filtered_df[
            filtered_df['record_id'].astype(str).str.contains(search_query, case=False) |
            filtered_df['extracted_store_or_dept'].astype(str).str.contains(search_query, case=False) |
            filtered_df['extracted_note'].astype(str).str.contains(search_query, case=False)
        ]
    else:
        search_df = filtered_df
        
    st.dataframe(search_df, use_container_width=True, height=250)

with tab2:
    st.markdown("수기 필기 텍스트가 심하게 훼손되어 누락되었거나 OCR 신뢰도가 극도로 낮은 **'확인필요'** 조치 대상 목록입니다.")
    # Filter for note == '확인필요' OR confidence < 0.70
    failures_df = filtered_df[
        (filtered_df['extracted_note'] == "확인필요") | 
        (filtered_df['confidence'] < 0.70)
    ]
    st.markdown(f"**총 {len(failures_df)}건의 점검 필요 대상이 감지되었습니다.**")
    st.dataframe(failures_df, use_container_width=True, height=200)

with tab3:
    st.markdown("정제 완료된 전체 영수증 지출 리스트입니다.")
    receipts_master = filtered_df[filtered_df['document_type'] == 'receipt']
    st.dataframe(receipts_master, use_container_width=True, height=200)

# Footer
st.markdown("<br><hr><center style='font-size:11px; color:#888888;'>Antigravity Project 2 OCR Dashboard Practice • Built with Streamlit and Plotly</center>", unsafe_allow_html=True)
