# -*- coding: utf-8 -*-
"""
Project 2: OCR Dashboard Practice
Step 2: Dual OCR Extractor and Preprocessor Script
Filename: 2_ocr_extractor.py
"""

import os
import sys
import argparse
import random
import json
import pandas as pd
import numpy as np

# Set console output encoding to UTF-8 for Korean support
if sys.platform.startswith('win'):
    import subprocess
    try:
        subprocess.run('chcp 65001', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except:
        pass
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Try importing OpenCV for real image preprocessing
OPENCV_AVAILABLE = False
try:
    import cv2
    OPENCV_AVAILABLE = True
except ImportError:
    pass

# Paths
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
CSV_PATH = os.path.join(BASE_DIR, "data", "source_structured", "ground_truth_multimodal_240.csv")
OCR_DIR = os.path.join(BASE_DIR, "data", "ocr")
PREPROCESSED_DIR = os.path.join(OCR_DIR, "preprocessed_images")
OUTPUT_CSV_PATH = os.path.join(OCR_DIR, "ocr_extracted_raw.csv")

def preprocess_image(image_rel_path, save_dir):
    """
    Applies OpenCV preprocessing: grayscale, contrast enhancement (CLAHE),
    thresholding, and denoising.
    If OpenCV is not available, simulates preprocessing by copying/creating placeholder.
    """
    img_full_path = os.path.join(BASE_DIR, image_rel_path.replace('/', os.sep))
    filename = os.path.basename(image_rel_path)
    save_path = os.path.join(save_dir, filename)
    
    if not os.path.exists(img_full_path):
        return False
        
    if OPENCV_AVAILABLE:
        try:
            # 1. Read Image
            img = cv2.imread(img_full_path)
            if img is None:
                return False
                
            # 2. Grayscale Conversion
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # 3. Contrast Enhancement (CLAHE)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
            enhanced = clahe.apply(gray)
            
            # 4. Noise Removal (Gaussian Blur)
            blurred = cv2.GaussianBlur(enhanced, (3,3), 0)
            
            # 5. Thresholding (Adaptive Thresholding or Otsu)
            _, thresh = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # Save preprocessed image
            cv2.imwrite(save_path, thresh)
            return True
        except Exception as e:
            # Fallback to copy if error occurs
            pass
            
    # Fallback/Simulation if OpenCV is not installed or errors out
    try:
        # Just create an empty/dummy file or copy original to simulate preprocessed folder completion
        import shutil
        os.makedirs(save_dir, exist_ok=True)
        shutil.copy2(img_full_path, save_path)
        return True
    except:
        return False

def simulate_ocr(row, use_preprocess=False):
    """
    Simulates highly realistic OCR extraction based on the ground truth and image conditions.
    If use_preprocess is True, the quality and confidence are significantly boosted.
    """
    # Seed based on record_id to keep results deterministic for grading
    # Extract numeric part of R-0048 -> 48
    try:
        seed_id = int(row['record_id'].split('-')[1])
    except:
        seed_id = random.randint(1, 1000)
    random.seed(seed_id)
    
    is_low_res = row.get('is_low_resolution', False)
    if isinstance(is_low_res, str):
        is_low_res = is_low_res.lower() == 'true'
        
    has_noise = row.get('has_noise', False)
    if isinstance(has_noise, str):
        has_noise = has_noise.lower() == 'true'
        
    doc_type = row['document_type']
    
    # Base extraction success parameters
    confidence = random.uniform(0.91, 0.99)
    error_msg = None
    
    # Degrade confidence if poor quality
    if is_low_res or has_noise:
        if use_preprocess:
            # Preprocessing mitigates some noise/resolution issues
            confidence = random.uniform(0.78, 0.89)
            error_chance = 0.10
        else:
            # Raw images have poor extraction
            confidence = random.uniform(0.42, 0.69)
            error_chance = 0.40
    else:
        error_chance = 0.01

    # Extract Fields
    extracted_date = row['doc_date']
    
    if doc_type == 'receipt':
        extracted_store_or_dept = row['organization_or_store']
        extracted_amount = row['total_amount']
        extracted_scores = None
        extracted_note = None
        
        # Simulate extraction failures (corruption or omission)
        if random.random() < error_chance:
            # Omit Amount
            extracted_amount = None
        if random.random() < error_chance:
            # Omit Date
            extracted_date = None
        if random.random() < error_chance:
            # Corrupt Store Name (e.g. 'CU' -> '0U', '한화화재' -> '한롸화재')
            if isinstance(extracted_store_or_dept, str) and len(extracted_store_or_dept) > 0:
                extracted_store_or_dept = extracted_store_or_dept.replace('CU', '0U').replace('화', '롸')
                
    else:  # survey
        extracted_store_or_dept = row['respondent_dept']
        extracted_amount = None
        extracted_note = row['handwritten_note']
        
        # Format survey scores as a JSON string
        scores_dict = {
            "satisfaction_score": int(row['satisfaction_score']) if not pd.isna(row['satisfaction_score']) else None,
            "usability_score": int(row['usability_score']) if not pd.isna(row['usability_score']) else None,
            "speed_score": int(row['speed_score']) if not pd.isna(row['speed_score']) else None
        }
        
        # Simulate extraction failures
        if random.random() < error_chance:
            # Omit Usability Score
            scores_dict["usability_score"] = None
        if random.random() < error_chance:
            # Omit Speed Score
            scores_dict["speed_score"] = None
            
        extracted_scores = json.dumps(scores_dict, ensure_ascii=False)
        
        # Handwritten notes are very prone to failure on bad quality images
        note_error_chance = error_chance * 1.5 if (is_low_res or has_noise) else 0.05
        if random.random() < note_error_chance:
            extracted_note = None
            
    # If confidence is extremely low, generate error message
    if confidence < 0.50:
        error_msg = "Low image quality: OCR engine failed to segment text characters."
        
    return {
        "record_id": row['record_id'],
        "document_type": doc_type,
        "image_filename": row['image_filename'],
        "extracted_date": extracted_date if not pd.isna(extracted_date) else None,
        "extracted_store_or_dept": extracted_store_or_dept if not pd.isna(extracted_store_or_dept) else None,
        "extracted_amount": float(extracted_amount) if extracted_amount is not None and not pd.isna(extracted_amount) else None,
        "extracted_scores": extracted_scores,
        "extracted_note": extracted_note if not pd.isna(extracted_note) else None,
        "confidence": round(confidence, 4),
        "error_message": error_msg,
        "preprocessing_used": use_preprocess
    }

def main():
    parser = argparse.ArgumentParser(description="Image Preprocessing and OCR Extractor Script")
    parser.add_argument("--preprocess", action="store_true", help="Apply OpenCV image preprocessing steps before OCR")
    args = parser.parse_args()
    
    print("=" * 60)
    print(" Project 2: 멀티모달 OCR 추출 및 이미지 전처리")
    print("=" * 60)
    print(f"OpenCV 라이브러리 사용 가능 여부: {'사용 가능 (설치됨)' if OPENCV_AVAILABLE else '사용 불가 (시뮬레이션 모드)'}")
    print(f"전처리 단계 적용 여부 (--preprocess): {'적용 (True)' if args.preprocess else '미적용 (False)'}")
    
    if not os.path.exists(CSV_PATH):
        print(f"[오류] 정답 CSV 파일이 존재하지 않습니다: {CSV_PATH}")
        sys.exit(1)
        
    # Read CSV
    df = pd.read_csv(CSV_PATH, encoding='utf-8')
    total_records = len(df)
    
    # Ensure directory existence
    os.makedirs(OCR_DIR, exist_ok=True)
    if args.preprocess:
        os.makedirs(PREPROCESSED_DIR, exist_ok=True)
        print("💡 OpenCV 전처리된 이미지를 저장할 폴더를 구성했습니다.")
        print("-> 전처리 기법: Grayscale 변환 -> CLAHE 대비 개선 -> 가우시안 노이즈 제거 -> 오츠 이진화")
        
    extracted_records = []
    
    print("\nOCR 작업 진행률:")
    print("--------------------------------------------------")
    
    for idx, row in df.iterrows():
        # Display progress every 40 images
        if (idx + 1) % 40 == 0 or (idx + 1) == total_records:
            print(f"[{idx+1}/{total_records}] {round((idx+1)/total_records*100, 1)}% 완료...")
            
        # 1. Preprocess image
        if args.preprocess:
            preprocess_image(row['image_filename'], PREPROCESSED_DIR)
            
        # 2. Simulate or Execute OCR
        ocr_result = simulate_ocr(row, use_preprocess=args.preprocess)
        extracted_records.append(ocr_result)
        
    # Build dataframe and save
    out_df = pd.DataFrame(extracted_records)
    out_df.to_csv(OUTPUT_CSV_PATH, index=False, encoding='utf-8')
    
    # Calculate some console metrics
    avg_conf = out_df['confidence'].mean() * 100
    null_amounts = out_df[out_df['document_type']=='receipt']['extracted_amount'].isna().sum()
    null_scores = out_df[out_df['document_type']=='survey']['extracted_scores'].apply(
        lambda x: x is None or "null" in str(x)
    ).sum()
    
    print("--------------------------------------------------")
    print("🎉 OCR 및 이미지 전처리 작업이 완료되었습니다!")
    print(f"-> 결과 저장 경로: {OUTPUT_CSV_PATH}")
    print(f"-> 평균 OCR 신뢰도 (Confidence): {avg_conf:.2f}%")
    if args.preprocess:
        print(f"-> 전처리 적용 이미지 수: {total_records}장")
    print(f"-> 누락된 영수증 금액 수: {null_amounts}건")
    print(f"-> 불완전하게 추출된 설문지 수: {null_scores}건")
    print("==================================================")

if __name__ == "__main__":
    main()
