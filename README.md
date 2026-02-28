# any2pdf

원본 폴더 내 문서들을 PDF로 일괄 변환하는 프로젝트입니다.

지원 형식:
- `hwp`, `hwpx`, `doc`, `docx`, `md`

구성:
- `any2pdf/any2pdf.py`: 폴더 단위 배치 변환기
- `md2docx/`: Markdown → DOCX 변환 스크립트
- `hwp2docx/`: HWP/HWPX 변환 스크립트(폴백 포함)

## 요구 사항

- Python 3.10+
- LibreOffice (필수)
- Python 패키지

```bash
uv pip install -r any2pdf/requirements.txt
uv pip install -r md2docx/requirements.txt
uv pip install -r hwp2docx/requirements.txt
```

## 빠른 실행

```bash
# 1) 가상환경(옵션)
python -m venv .venv
source .venv/bin/activate

# 2) 의존성 설치
uv pip install -r any2pdf/requirements.txt
uv pip install -r md2docx/requirements.txt
uv pip install -r hwp2docx/requirements.txt

# 3) 변환 실행
python any2pdf/any2pdf.py "입력_폴더" -o "출력_폴더" --recursive
```

출력 파일명 규칙:
- `<원본이름>_<확장자>.pdf`
  - 예) `문서.docx` -> `문서_docx.pdf`, `문서.hwp` -> `문서_hwp.pdf`

## 환경별 LibreOffice 준비

- macOS
  - `brew install --cask libreoffice`
- Windows
  - `winget install --id TheDocumentFoundation.LibreOffice -e --accept-package-agreements --accept-source-agreements`

상세한 설치/진단은 `any2pdf/README.md`를 참고하세요.
