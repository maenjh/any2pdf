# hwp2docx

## 역할

한글 문서 `(.hwp, .hwpx)`를 Word `(.docx)`로 변환합니다.

`any2pdf`는 주로 PDF 변환을 담당하지만, LibreOffice 변환이 어려운 `.hwp`는 필요 시 본 모듈의 fallback 경로를 사용합니다.

## 사용법

```bash
cd /Users/moon/Desktop/any2pdf
source .venv/bin/activate
python hwp2docx/hwp_to_docx.py /path/to/file.hwp
```

폴더 변환도 가능합니다.

```bash
python hwp2docx/hwp_to_docx.py /path/to/folder
```

### 설치

```bash
uv pip install -r hwp2docx/requirements.txt
```

### 변환 방식

1. LibreOffice CLI(`soffice`/`libreoffice`)로 우선 변환
2. LibreOffice 실패 시(및 `hwp` 파일에서만) `hwp5txt + python-docx` 텍스트 기반 변환 시도

### 참고

- `README_hwp_to_docx.md`에는 상세 동작 흐름과 예시가 더 포함되어 있습니다.
