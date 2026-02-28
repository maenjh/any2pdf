# hwp2docx

## .hwp/.hwpx → .docx 변환기

`hwp_to_docx.py`는 `.hwp` 또는 `.hwpx` 파일을 `.docx`로 변환합니다.
`any2pdf` 파이프라인에서는 이 모듈을 통해 HWP/HWPX 입력을 먼저 DOCX로 만들고, 이후 PDF로 변환합니다.

## 사용법

```bash
cd /Users/moon/Desktop/any2pdf
source .venv/bin/activate

# 단일 파일
python hwp2docx/hwp_to_docx.py /path/to/file.hwp
python hwp2docx/hwp_to_docx.py /path/to/file.hwpx

# 폴더 입력
python hwp2docx/hwp_to_docx.py /path/to/folder_containing_hwp_and_hwpx

# 출력 디렉토리 지정 + 덮어쓰기
python hwp2docx/hwp_to_docx.py /path/to/file1.hwp /path/to/file2.hwpx -o /path/to/output_dir --overwrite
```

## 동작 방식

1. 먼저 LibreOffice CLI(`soffice` 또는 `libreoffice`)를 시도합니다.
2. `.hwp` 파일은 LibreOffice 실패 시, `hwp5txt + python-docx` 텍스트 폴백을 시도합니다.

> `.hwpx`는 폴백 텍스트 변환을 지원하지 않습니다.

## 권장 설치

- LibreOffice 설치 (필수)
- 텍스트 폴백 사용 시:
  - `hwp5txt` (예: `pyhwp` 계열 패키지)
  - `python-docx`

## 의존성

```bash
uv pip install -r hwp2docx/requirements.txt
```

## 참고

- LibreOffice가 설치되어 있으면 안정적이며, 하드웨어/형식 이슈가 적습니다.
