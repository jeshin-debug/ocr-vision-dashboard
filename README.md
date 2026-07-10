# 📊 Project 2: 멀티모달 OCR 분석 대시보드 실습 프로젝트

본 실습 프로젝트는 제공된 원본 정량/정형 데이터와 이미지 데이터를 활용하여 **데이터 정합성 검증, OpenCV 기반 이미지 전처리, 멀티모달 OCR 추출 시뮬레이션, 데이터 품질 측정(QA), 결측치 보간(Imputation),** 그리고 최종 **임원 보고용 고급 Streamlit 대시보드 시각화**까지 데이터 파이프라인의 전 과정을 순차적으로 학습할 수 있는 실습 패키지입니다.

---

## 📂 프로젝트 주요 구성 요소 및 산출물

```
새 폴더/
├─ data/
│  ├─ source_structured/               # 제공된 원본 정답 데이터 및 가이드 이미지
│  ├─ input_images/                    # 영수증(120장) 및 설문지(120장) 원본 이미지
│  ├─ ocr/                              # OCR 원시 데이터 추출 저장소
│  │  └─ preprocessed_images/          # OpenCV 전처리 이진화 완료 이미지
│  └─ processed/                        # 최종 결측치 보간 완료 마스터 데이터셋
├─ app/                                 # Streamlit 대시보드 소스 코드
├─ reports/                             # 단계별 검증 및 품질 보고서
├─ final_output/                        # 최종 완료본 취합 및 ZIP 패키지
├─ 1_validate_data.py                   # [1단계] 데이터-이미지 정합성 전수 검사
├─ 2_ocr_extractor.py                   # [2단계] OpenCV 전처리 및 OCR 추출기 (시뮬레이션 겸용)
├─ 3_analyze_quality.py                 # [3단계] OCR 정확도 평가 및 보고서 생성 (openpyxl 스타일링)
├─ 4_clean_data.py                      # [4단계] 도메인 규칙 기반 결측치 보간 및 정제
├─ 5_prepare_final_package.py           # [5단계] 최종 산출물 취합 및 자동 압축 패키징
└─ README.md                            # 본 실습 가이드 문서
```

---

## 🛠️ 실습 환경 설정 및 필수 패키지 설치

실습에 필요한 라이브러리를 설치합니다. 본 프로젝트는 라이브러리가 미설치되어 있어도 예외 처리 및 시뮬레이터 모드로 원활하게 진행할 수 있도록 방어 코딩이 완료되어 있습니다.

```bash
# 필수 데이터 분석 및 시각화 패키지 전수 설치
pip install pandas numpy openpyxl streamlit plotly opencv-python
```

---

## 🏃‍♂️ 데이터 파이프라인 순차 실행 가이드

수강생은 코드를 직접 수정할 필요 없이, 아래 순서대로 실행 명령어만 터미널에 입력하여 전 파이프라인을 실습할 수 있습니다.

### 1단계: 원본 데이터 및 이미지 정합성 전수 검증
제공된 CSV의 이미지 경로와 파일 시스템 상의 이미지 파일이 누락 없이 일치하는지 전수 대조합니다.
```bash
python 1_validate_data.py
```
- **출력 결과물**: `reports/data_image_validation_report.txt` 생성 및 콘솔 테이블 출력.

### 2단계: OpenCV 전처리 적용 및 멀티모달 OCR 추출
저해상도 및 어둡고 오염된 이미지를 인식하기 위해 OpenCV(이진화, 노이즈 필터, 대비 향상)를 적용하여 raw OCR 추출 결과를 도출합니다.
```bash
python 2_ocr_extractor.py --preprocess
```
- **주요 플래그**: `--preprocess` (전처리 활성화)
- **출력 결과물**: 
  - `data/ocr/ocr_extracted_raw.csv` (11개 컬럼 원시 데이터 생성)
  - `data/ocr/preprocessed_images/` (전처리 적용 흑백 이미지 240장 저장)

### 3단계: OCR 추출 품질 측정 및 평가 리포트 생성
추출된 원시 OCR 데이터와 정답(Ground Truth) 데이터를 일대일 비교하여 수치 오차 및 분류 정확도를 산출합니다.
```bash
python 3_analyze_quality.py
```
- **출력 결과물**: 
  - `reports/ocr_quality_summary.txt` (요약 통계 텍스트 보고서)
  - `reports/ocr_quality_report.xlsx` (Classic Navy / Mint 테마 적용 보고서 - openpyxl 스타일링)

### 4단계: 도메인 비즈니스 규칙 기반 결측치 정제 및 보간
품질 저하 이미지로 인해 OCR 단계에서 발생한 누락 금액(정답 매칭 보완), 날짜, 누락 만족도(부서 평균 보간), 메모 누락('확인필요' 마크) 등의 결측치 정책을 적용하여 최종 골드 데이터셋을 만듭니다.
```bash
python 4_clean_data.py
```
- **출력 결과물**: `data/processed/ocr_cleaned_dataset.xlsx` (최종 정제 마스터 엑셀셋)

### 5단계: Streamlit 임원 보고용 대시보드 가동
정제 완료된 최종 마스터 엑셀셋을 연계하여 Horizontal KPI 카드, 도넛 차트, 멀티 바 차트, 테이블 탭 컨트롤러가 탑재된 대시보드를 가동합니다.
```bash
streamlit run app/vision_dashboard.py
```
- **서버 접속 주소**: 웹 브라우저를 열고 `http://localhost:8501` 에 접속하여 모니터링합니다.

### 6단계: 최종 산출물 정리 및 압축 자동화
모든 실습 산출물을 한 곳에 수집하고 배포용 ZIP 압축 파일로 패키징합니다.
```bash
python 5_prepare_final_package.py
```
- **출력 결과물**: `final_output/` 폴더 내 최종본 취합 및 `final_output.zip` 압축 파일 자동 완성!

---
💡 **Antigravity AI Assistant**에 의해 수강생 중심의 완성형 실습 패키지로 안전하게 설계 및 빌드되었습니다.
