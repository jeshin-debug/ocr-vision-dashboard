# -*- coding: utf-8 -*-
"""
Project 2: OCR Dashboard Practice
Step 4: Missing Value Imputation and Data Cleaning Script
Filename: 4_clean_data.py
"""

import os
import sys
import json
import pandas as pd
import numpy as np

# Ensure openpyxl is installed
try:
    import openpyxl
except ImportError:
    import subprocess
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except:
        pass

# Set console output encoding to UTF-8 for Korean support immediately
if sys.platform.startswith('win'):
    import subprocess
    try:
        subprocess.run('chcp 65001', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except:
        pass
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except:
        pass

# Paths
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
GT_CSV_PATH = os.path.join(BASE_DIR, "data", "source_structured", "ground_truth_multimodal_240.csv")
OCR_CSV_PATH = os.path.join(BASE_DIR, "data", "ocr", "ocr_extracted_raw.csv")
PROCESSED_DIR = os.path.join(BASE_DIR, "data", "processed")
OUTPUT_XLSX_PATH = os.path.join(PROCESSED_DIR, "ocr_cleaned_dataset.xlsx")

def safe_parse_json(val):
    if pd.isna(val) or not val:
        return {}
    try:
        return json.loads(val)
    except:
        return {}

def main():
    print("=" * 60)
    print(" Project 2: OCR 결측치 보간 및 데이터 정제 정규화")
    print("=" * 60)
    
    if not os.path.exists(GT_CSV_PATH) or not os.path.exists(OCR_CSV_PATH):
        print("[오류] 입력 파일이 부족합니다. 이전 단계를 먼저 수행하십시오.")
        sys.exit(1)
        
    df_gt = pd.read_csv(GT_CSV_PATH, encoding='utf-8')
    df_ocr = pd.read_csv(OCR_CSV_PATH, encoding='utf-8')
    
    # Copy raw OCR as a base for clean dataset
    df_clean = df_ocr.copy()
    
    # 1. Parse survey scores JSON into individual columns in df_clean
    df_clean['satisfaction_score'] = np.nan
    df_clean['usability_score'] = np.nan
    df_clean['speed_score'] = np.nan
    
    for idx, row in df_clean.iterrows():
        if row['document_type'] == 'survey':
            scores = safe_parse_json(row['extracted_scores'])
            df_clean.at[idx, 'satisfaction_score'] = scores.get('satisfaction_score')
            df_clean.at[idx, 'usability_score'] = scores.get('usability_score')
            df_clean.at[idx, 'speed_score'] = scores.get('speed_score')
            
    # Keep track of imputation statistics
    imputed_dates = 0
    imputed_amounts = 0
    imputed_scores = 0
    imputed_notes = 0
    
    df_clean['amount_imputed'] = False
    df_clean['score_imputed'] = False
    
    # Pre-calculate department average scores from the OCR extracted values to use for survey imputation
    dept_scores_avg = df_clean[df_clean['document_type']=='survey'].groupby('extracted_store_or_dept')[
        ['satisfaction_score', 'usability_score', 'speed_score']
    ].transform('mean')
    
    # Fallback overall averages in case an entire department is missing
    overall_avg_sat = df_clean[df_clean['document_type']=='survey']['satisfaction_score'].mean()
    overall_avg_usa = df_clean[df_clean['document_type']=='survey']['usability_score'].mean()
    overall_avg_spd = df_clean[df_clean['document_type']=='survey']['speed_score'].mean()
    
    # Apply cleaning and imputation policies
    for idx, row in df_clean.iterrows():
        rec_id = row['record_id']
        doc_type = row['document_type']
        
        # Get ground truth matching row
        gt_row = df_gt[df_gt['record_id'] == rec_id].iloc[0]
        
        # Policy 1: Date missing -> 보완
        if pd.isna(row['extracted_date']):
            df_clean.at[idx, 'extracted_date'] = gt_row['doc_date']
            imputed_dates += 1
            
        # Policy 2: Amount missing (receipts only) -> 보완 및 amount_imputed = True
        if doc_type == 'receipt':
            if pd.isna(row['extracted_amount']):
                df_clean.at[idx, 'extracted_amount'] = gt_row['total_amount']
                df_clean.at[idx, 'amount_imputed'] = True
                imputed_amounts += 1
                
        # Policy 3: Survey scores missing -> 부서 평균 보간
        if doc_type == 'survey':
            dept = row['extracted_store_or_dept'] # For survey, this column contains department name
            
            # Satisfaction Imputation
            if pd.isna(row['satisfaction_score']):
                # Find other surveys in same department to calculate average
                dept_surveys = df_clean[(df_clean['document_type']=='survey') & (df_clean['extracted_store_or_dept']==dept)]
                avg_val = dept_surveys['satisfaction_score'].mean()
                fill_val = round(avg_val) if not pd.isna(avg_val) else round(overall_avg_sat)
                df_clean.at[idx, 'satisfaction_score'] = fill_val
                df_clean.at[idx, 'score_imputed'] = True
                imputed_scores += 1
                
            # Usability Imputation
            if pd.isna(row['usability_score']):
                dept_surveys = df_clean[(df_clean['document_type']=='survey') & (df_clean['extracted_store_or_dept']==dept)]
                avg_val = dept_surveys['usability_score'].mean()
                fill_val = round(avg_val) if not pd.isna(avg_val) else round(overall_avg_usa)
                df_clean.at[idx, 'usability_score'] = fill_val
                df_clean.at[idx, 'score_imputed'] = True
                imputed_scores += 1
                
            # Speed Imputation
            if pd.isna(row['speed_score']):
                dept_surveys = df_clean[(df_clean['document_type']=='survey') & (df_clean['extracted_store_or_dept']==dept)]
                avg_val = dept_surveys['speed_score'].mean()
                fill_val = round(avg_val) if not pd.isna(avg_val) else round(overall_avg_spd)
                df_clean.at[idx, 'speed_score'] = fill_val
                df_clean.at[idx, 'score_imputed'] = True
                imputed_scores += 1
                
        # Policy 4: Memo missing (surveys only) -> '확인필요'로 대체
        if doc_type == 'survey':
            if pd.isna(row['extracted_note']):
                df_clean.at[idx, 'extracted_note'] = "확인필요"
                imputed_notes += 1
                
    # Add other ground truth metadata columns helpful for dashboards
    df_clean['is_low_resolution'] = df_gt['is_low_resolution']
    df_clean['has_noise'] = df_gt['has_noise']
    
    # Merge category from ground truth (receipt categories like '식비', '도서' etc)
    df_clean['category'] = df_gt['category']
    
    # Save to Excel
    os.makedirs(PROCESSED_DIR, exist_ok=True)
    df_clean.to_excel(OUTPUT_XLSX_PATH, index=False)
    
    print("-" * 50)
    print("🎉 결측치 보간 및 데이터 정제 작업이 성공적으로 완료되었습니다!")
    print(f"-> 결과 저장 파일: {OUTPUT_XLSX_PATH}")
    print("--------------------------------------------------")
    print(" 🛠️ 정제 및 보간 처리 통계 (Imputation Summary)")
    print("--------------------------------------------------")
    print(f"1. 누락된 날짜 보완 건수 (Date Filled): {imputed_dates}건")
    print(f"2. 누락된 영수증 금액 정답 보완 건수 (Amount Filled): {imputed_amounts}건")
    print(f"3. 누락된 설문 만족도 점수 부서평균 보간 건수 (Scores Imputed): {imputed_scores}건")
    print(f"4. 누락된 설문 수기메모 '확인필요' 보정 건수 (Notes Filled): {imputed_notes}건")
    print("==================================================")

if __name__ == "__main__":
    main()
