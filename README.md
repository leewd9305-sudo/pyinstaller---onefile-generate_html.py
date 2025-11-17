# 🧰 HUSK Guide Generator  
### 단색/별색/일반 제작 가이드 자동 생성 프로그램  
*자동화 기획 · 개발 : 이운도 (Lee Woondo)*  

---

## 📌 Overview  
**HUSK Guide Generator**는 제작가이드용 엑셀 파일을 기반으로  
단색 / 별색 / 일반 **이미지 가이드 TXT**를 자동 생성하고,  
병합 텍스트 + 3종 로그 + 자동 압축까지 수행하는 **사내 자동화 프로그램**입니다.

검수팀의 반복 업무를 자동화하여  
작업 시간을 대폭 감소시키는 것을 목표로 제작되었습니다.  
(현재 실제 업무에서 사용 중)

---

## ✨ 주요 기능

### 🟦 1. 단색 / 별색 / 일반 TXT 자동 생성  
- 엑셀의 시트명에 따라 자동 분류  
  - `파일명 리스트(단색)`  
  - `파일명 리스트(별색)`  
  - 그 외 → 일반 TXT 처리  
- 이미지 파일명 기반으로 HTML 템플릿 생성  
- 제품명 기반 파일명 자동 생성

---

### 🟩 2. 단색 기준 병합 TXT 생성  
- 단색 TXT를 기준으로 **단색/별색 양쪽 화면을 하나로 병합한 TXT 생성**
- 병합 레이아웃 자동 구성  
- 공정 안내 툴팁 이미지 자동 포함

---

### 🟧 3. 3종 로그 자동 생성 (mono / spot / normal)  
- 각 TXT 생성 내용이 모두 로그로 기록  
- 시트명 분류 기준으로 자동 분리  
- 로그가 필요한 검수팀/디자인팀 업무에 활용 가능

---

### 📦 4. 자동 ZIP 압축  
모든 결과물(TXT + 로그)을 하나의 ZIP으로 자동 패킹하여  
검수팀이 `Downloads` 폴더에서 바로 확인할 수 있도록 구성.

---

## 🖥️ 사용 방법

### 1) exe 실행  
프로그램 실행 시 자동으로 파일 선택 UI가 팝업됩니다.

### 2) "제작가이드 엑셀 파일" 선택  
엑셀 파일 구조는 다음과 같아야 합니다:  
- 헤더 2줄  
- 1열: 순번  
- 2열: 제품명  
- 3열~: 이미지 파일명  

### 3) 자동 생성 프로세스 진행  
- TXT 생성  
- 로그 생성  
- 병합 TXT 생성  
- ZIP 자동 저장  

### 4) 결과 ZIP 위치  
C:/Users/사용자명/Downloads/husk_guide_output_날짜.zip

yaml
Copy code

---

## 📁 프로젝트 구조

📦 HUSK Guide Generator
├─ data/ # XLSX 테스트 파일 (선택)
├─ output/ # 실행 시 자동 생성되는 파일
├─ build/ # PyInstaller 빌드 정보
├─ dist/ # exe 빌드 폴더
├─ generate_html.py # 메인 스크립트
├─ generate_html.spec # PyInstaller 설정
└─ README.md

yaml
Copy code

---

## 🚀 개발 스택

- Python 3.11  
- Pandas  
- Tkinter  
- PyInstaller  
- Unicode Normalization (NFKC)  
- 정규식 기반 HTML 템플릿 파서  

---

## 🧑‍💻 개발/기획 담당자

**이운도 (Lee Woondo)**  
- UX/UI 디자이너 → 실제 업무 자동화를 위한 간단한 개발까지 확장  
- 자동화 기획, UI 흐름, 파일 구조 설계, 코드 작성 전부 직접 수행  
- PyInstaller 패키징 및 배포 시스템 구축  

---

## 🏷️ 버전 정보

### **v1.0.0**
- 단색/별색/일반 TXT 자동 생성  
- 단색 기준 병합 TXT 생성  
- 3종 로그 자동 생성  
- ZIP 자동 압축  
- exe 버전 최초 배포  

---

## 🔗 다운로드

**Latest Release:**  
👉 https://github.com/leewd9305-sudo/pyinstaller---onefile-generate_html.py/releases/latest  

---

## 📜 라이선스  
Internal-use only (사내용 사용자 자동화 프로그램)
