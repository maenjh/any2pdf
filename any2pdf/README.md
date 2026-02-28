# any2pdf (폴더 단위 PDF 변환기)

## 역할

한 폴더 안의 `hwp`, `hwpx`, `doc`, `docx`, `md` 파일을 PDF로 일괄 변환하는 CLI 스크립트입니다.

## 사용법

```bash
cd /path/to/any2pdf
source .venv/bin/activate
python any2pdf/any2pdf.py /입력/폴더 -o /출력/폴더 --recursive
```

### 주요 옵션

- `-o, --out`: 출력 폴더(기본값은 각 입력 파일의 부모 폴더)
- `--recursive`: 하위 폴더까지 재귀 탐색
- `--overwrite`: 기존 PDF 덮어쓰기
- `--timeout`: 변환 타임아웃(초, 기본 180)
- `--no-fallback`: `.hwp` 텍스트 폴백(hwp5txt) 비활성화

### 출력 파일명 규칙

- `<원본이름>_<확장자>.pdf`
  - 예) `docx` 파일 → `문서_docx.pdf`, `hwp` 파일 → `문서_hwp.pdf`

### 선행 조건

- Python 패키지: `uv pip install -r any2pdf/requirements.txt`
- LibreOffice CLI(`soffice` 또는 `libreoffice`) 설치
- `soffice --version` 또는 `libreoffice --version` 실행 확인

### 문제 해결

실패 시 로그의 `[FAIL]` 메시지를 기준으로 파일별 원인을 확인하세요.
- `soffice` 미설치: 환경 점검 필요
- `hwp` 변환 실패: 폴백 옵션과 `hwp5txt` 의존성 상태 점검
