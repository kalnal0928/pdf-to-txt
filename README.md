# PDF to TXT 변환기

PDF 파일을 텍스트 파일로 변환하는 Python 스크립트입니다. GUI와 명령행 인터페이스를 모두 지원합니다.

## 설치

```bash
pip install -r requirements.txt
```

## 사용법

### 1. GUI 버전 (권장)

```bash
# GUI 실행
python pdf_to_txt_gui.py

# 또는 메인 스크립트에서
python pdf_to_txt.py --gui

# Windows에서 배치 파일 사용
run.bat
```

**GUI 기능:**
- 단일 파일 또는 여러 파일 선택
- 폴더 내 모든 PDF 파일 일괄 선택
- 출력 폴더 지정
- 추출 방법 선택 (pdfplumber/PyPDF2)
- 실시간 진행 상황 표시
- 사용자 친화적인 인터페이스

### 2. 명령행 버전

**단일 파일 변환:**
```bash
# 기본 사용 (출력 파일명 자동 생성)
python pdf_to_txt.py document.pdf

# 출력 파일명 지정
python pdf_to_txt.py document.pdf output.txt

# 추출 방법 지정
python pdf_to_txt.py document.pdf output.txt pdfplumber
```

**일괄 변환 (폴더 내 모든 PDF 파일):**
```bash
# 같은 폴더에 TXT 파일 생성
python pdf_to_txt.py --batch ./pdfs

# 다른 폴더에 TXT 파일 생성
python pdf_to_txt.py --batch ./pdfs ./output

# 추출 방법 지정
python pdf_to_txt.py --batch ./pdfs ./output pypdf2
```

## 추출 방법

1. **pdfplumber** (기본값, 권장)
   - 더 정확한 텍스트 추출
   - 표와 복잡한 레이아웃 처리 우수

2. **pypdf2**
   - 빠른 텍스트 추출
   - 단순한 문서에 적합

## 주의사항

- 이미지 기반 PDF (스캔된 문서)는 텍스트 추출이 어려울 수 있습니다
- 보호된 PDF 파일은 추출이 제한될 수 있습니다
- 복잡한 레이아웃의 경우 텍스트 순서가 원본과 다를 수 있습니다

## 예시

```python
from pdf_to_txt import pdf_to_txt, batch_convert

# 단일 파일 변환
pdf_to_txt("document.pdf", "output.txt")

# 폴더 일괄 변환
batch_convert("./pdfs", "./texts")
```
