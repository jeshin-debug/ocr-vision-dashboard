# -*- coding: utf-8 -*-
"""
Project 2: OCR Dashboard Practice
Step 7: Final Deliverables Packaging and ZIP Archive Creator
Filename: 5_prepare_final_package.py
"""

import os
import shutil
import zipfile
import sys

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
FINAL_OUTPUT_DIR = os.path.join(BASE_DIR, "final_output")
ZIP_OUT_PATH = os.path.join(BASE_DIR, "final_output.zip")

# Mapping of source files to target names in final_output/
FILES_TO_PACKAGE = {
    # Source file path: Target filename in final_output
    os.path.join(BASE_DIR, "data", "processed", "ocr_cleaned_dataset.xlsx"): "ocr_cleaned_dataset.xlsx",
    os.path.join(BASE_DIR, "reports", "ocr_quality_report.xlsx"): "ocr_quality_report.xlsx",
    os.path.join(BASE_DIR, "reports", "ocr_quality_summary.txt"): "ocr_quality_summary.txt",
    os.path.join(BASE_DIR, "app", "vision_dashboard.py"): "vision_dashboard.py",
    os.path.join(BASE_DIR, "reports", "dashboard_visual_check.txt"): "dashboard_visual_check.txt",
    os.path.join(BASE_DIR, "README.md"): "README.md"
}

def main():
    print("=" * 60)
    print(" Project 2: 최종 산출물 취합 및 패키징 자동화")
    print("=" * 60)
    
    # 1. Recreate final_output folder
    if os.path.exists(FINAL_OUTPUT_DIR):
        shutil.rmtree(FINAL_OUTPUT_DIR)
    os.makedirs(FINAL_OUTPUT_DIR, exist_ok=True)
    
    copied_count = 0
    missing_count = 0
    
    # 2. Copy deliverables
    print("\n📦 산출물 복사 작업을 진행합니다...")
    for src_path, target_name in FILES_TO_PACKAGE.items():
        dest_path = os.path.join(FINAL_OUTPUT_DIR, target_name)
        if os.path.exists(src_path):
            try:
                shutil.copy2(src_path, dest_path)
                print(f"-> [성공] {target_name} 복사 완료")
                copied_count += 1
            except Exception as e:
                print(f"-> [오류] {target_name} 복사 실패: {e}")
                missing_count += 1
        else:
            print(f"-> [누락] {target_name} 소스 파일을 찾을 수 없습니다: {src_path}")
            missing_count += 1
            
    # 3. Zip the final_output directory contents
    if copied_count > 0:
        print(f"\n🤐 취합된 산출물을 압축 중입니다... (대상 파일 수: {copied_count}개)")
        try:
            with zipfile.ZipFile(ZIP_OUT_PATH, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, dirs, files in os.walk(FINAL_OUTPUT_DIR):
                    for file in files:
                        file_full_path = os.path.join(root, file)
                        # Relative path inside the zip
                        arcname = os.path.relpath(file_full_path, FINAL_OUTPUT_DIR)
                        # To create a root-level final_output directory inside zip:
                        zipf.write(file_full_path, os.path.join("final_output", arcname))
            print(f"-> [성공] 최종 패키지 ZIP 생성 완료!")
            print(f"   저장 파일: {ZIP_OUT_PATH}")
        except Exception as e:
            print(f"-> [오류] ZIP 파일 압축 실패: {e}")
    else:
        print("\n[오류] 복사된 파일이 없어 ZIP 압축을 건너뜁니다.")
        
    print("\n" + "=" * 50)
    print(" 🎉 실습 프로젝트 최종 패키징이 성공적으로 마감되었습니다!")
    print("--------------------------------------------------")
    print(f" - 복사 완료된 최종 파일 수: {copied_count}개")
    print(f" - 누락되거나 건너뛴 파일 수: {missing_count}개")
    print(f" - 최종 압축 파일: final_output.zip")
    print("==================================================")

if __name__ == "__main__":
    main()
