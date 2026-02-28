# .hwp / .hwpx → .docx 변환기

`hwp_to_docx.py`는 `.hwp` 또는 `.hwpx` 파일을 `.docx`로 변환합니다.

## 사용법

```bash
python3 /Users/moon/Desktop/hwp_to_docx.py /path/to/file.hwp
python3 /Users/moon/Desktop/hwp_to_docx.py /path/to/dir_containing_hwp_and_hwpx
python3 /Users/moon/Desktop/hwp_to_docx.py /path/to/file1.hwp /path/to/file2.hwpx -o /path/to/output_dir --overwrite
```

## 동작 방식

1. 우선 LibreOffice CLI(`soffice` 또는 `libreoffice`)를 이용해 변환을 시도합니다.
2. LibreOffice 실패 + `.hwp` 파일인 경우:
   - `hwp5txt` + `python-docx` 조합으로 텍스트 기반 폴백 변환을 시도합니다.

## 권장 설치

- `LibreOffice` 설치 (필수 권장)
- 폴백을 원할 때:
  - `hwp5txt` (선택)
  - `python-docx` (선택): `pip install python-docx`

## 참고

- `.hwpx`는 형식상 문서 구조가 달라서 폴백 텍스트 변환을 지원하지 않으며, LibreOffice가 설치되어 있어야 안정적으로 변환됩니다.
