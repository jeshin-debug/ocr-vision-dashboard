# -*- coding: utf-8 -*-
"""
Project 2: OCR Dashboard Practice
Step 1: Data-Image Consistency Validation Script
Filename: 1_validate_data.py
"""

import os
import sys
import pandas as pd

# Set console output encoding to UTF-8 for Korean support
if sys.platform.startswith('win'):
    import subprocess
    # Attempt to set CP 65001 (UTF-8) for current powershell/cmd session
    try:
        subprocess.run('chcp 65001', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except:
        pass
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# Paths
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
CSV_PATH = os.path.join(BASE_DIR, "data", "source_structured", "ground_truth_multimodal_240.csv")
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
REPORT_OUTPUT_PATH = os.path.join(REPORTS_DIR, "data_image_validation_report.txt")

def format_as_table(title, headers, rows):
    """Formats rows as a neat ASCII table."""
    if not rows:
        return f"--- {title} ---\nNo issues found.\n"
    
    # Calculate column widths
    col_widths = [len(str(h)) for h in headers]
    for row in rows:
        for idx, val in enumerate(row):
            val_str = str(val)
            # visual length approximation: korean characters count as 2 spaces
            vis_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
            col_widths[idx] = max(col_widths[idx], vis_len)
            
    # Build header and separator
    border = "+" + "+".join(["-" * (w + 2) for w in col_widths]) + "+"
    header_line = "|" + "|".join([f" {str(h).ljust(col_widths[idx] - (sum(1 if ord(c) > 127 else 0 for c in str(h))))} " for idx, h in enumerate(headers)]) + "|"
    
    table_lines = [border, f"| {title.center(sum(col_widths) + len(headers)*3 - 3)} |", border, header_line, border]
    
    for row in rows:
        row_line = "|" + "|".join([f" {str(val).ljust(col_widths[idx] - (sum(1 if ord(c) > 127 else 0 for c in str(val))))} " for idx, val in enumerate(row)]) + "|"
        table_lines.append(row_line)
        
    table_lines.append(border)
    return "\n".join(table_lines) + "\n"

def main():
    print("=" * 60)
    print(" Project 2: OCR 데이터 및 이미지 정합성 전수 검증")
    print("=" * 60)
    print(f"작업 디렉토리: {BASE_DIR}")
    print(f"정답 데이터 경로: {CSV_PATH}")
    
    if not os.path.exists(CSV_PATH):
        print(f"[오류] 정답 CSV 파일이 존재하지 않습니다: {CSV_PATH}")
        sys.exit(1)
        
    # Read CSV
    try:
        df = pd.read_csv(CSV_PATH, encoding='utf-8')
    except Exception as e:
        print(f"[오류] CSV 파일을 읽는 데 실패했습니다: {e}")
        sys.exit(1)
        
    total_records = len(df)
    print(f"-> 총 {total_records}개의 레코드를 로드했습니다.")
    
    # Validation structures
    missing_images = []
    invalid_extensions = []
    duplicate_ids = []
    
    # Check duplicate record_ids
    dup_mask = df.duplicated(subset=['record_id'], keep=False)
    if dup_mask.any():
        dups = df[dup_mask].sort_values('record_id')
        for idx, row in dups.iterrows():
            duplicate_ids.append([row['record_id'], row['document_type'], row['image_filename']])
            
    # Check images existence and extension
    valid_extensions = {'.jpg', '.jpeg', '.png'}
    
    for idx, row in df.iterrows():
        rec_id = row['record_id']
        doc_type = row['document_type']
        img_rel_path = row['image_filename']
        
        # Normalize path
        img_full_path = os.path.join(BASE_DIR, img_rel_path.replace('/', os.sep))
        
        # Extension Check
        _, ext = os.path.splitext(img_rel_path.lower())
        if ext not in valid_extensions:
            invalid_extensions.append([rec_id, doc_type, img_rel_path, ext])
            
        # Existence Check
        if not os.path.exists(img_full_path):
            missing_images.append([rec_id, doc_type, img_rel_path, "존재하지 않음"])
            
    # Generate Report Content
    report_content = []
    report_content.append("=" * 70)
    report_content.append("  [검증 보고서] OCR 정답 데이터 및 이미지 정합성 전수 검증 결과")
    report_content.append("=" * 70)
    report_content.append(f"검증 일시: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report_content.append(f"대상 파일: {CSV_PATH}")
    report_content.append(f"총 데이터 수: {total_records}건")
    report_content.append(f"성공적으로 확인된 이미지 수: {total_records - len(missing_images)}장")
    report_content.append("-" * 70)
    report_content.append(f"1. 누락 이미지 수: {len(missing_images)}건")
    report_content.append(f"2. 잘못된 확장자 수: {len(invalid_extensions)}건")
    report_content.append(f"3. 중복 Record ID 수: {len(duplicate_ids)}건")
    report_content.append("=" * 70 + "\n")
    
    # Add Tables
    missing_table = format_as_table("누락 이미지 목록 (Missing Images)", ["Record ID", "유형", "이미지 경로", "상태"], missing_images)
    ext_table = format_as_table("잘못된 확장자 목록 (Invalid Extensions)", ["Record ID", "유형", "이미지 경로", "확장자"], invalid_extensions)
    dup_table = format_as_table("중복 Record ID 목록 (Duplicate Record IDs)", ["Record ID", "유형", "이미지 경로"], duplicate_ids)
    
    report_content.append(missing_table)
    report_content.append(ext_table)
    report_content.append(dup_table)
    
    final_report_str = "\n".join(report_content)
    
    # Save Report
    os.makedirs(REPORTS_DIR, exist_ok=True)
    with open(REPORT_OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(final_report_str)
        
    # Print results to stdout
    print(final_report_str)
    print(f"💡 검사 결과가 '{REPORT_OUTPUT_PATH}' 파일에 저장되었습니다.")
    print("=" * 60)

if __name__ == "__main__":
    main()
