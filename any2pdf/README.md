# any2pdf

`any2pdf`는 한 폴더 안의 `hwp`, `hwpx`, `doc`, `docx`, `md` 파일을 PDF로 일괄 변환하는 스크립트입니다.

## 사용법

```bash
cd /Users/moon/Desktop/text_files
source .venv/bin/activate
python any2pdf/any2pdf.py /경로/폴더명 -o /경로/출력폴더 --recursive
```

- 출력 파일명은 충돌 방지를 위해 `원본파일명_확장자.pdf` 형식으로 저장됩니다.
  - 예) `문서.docx` -> `문서_docx.pdf`, `문서.hwp` -> `문서_hwp.pdf`

- `--recursive`: 하위 폴더까지 모두 탐색
- `--overwrite`: 기존 PDF를 덮어씀
- `--timeout`: 실행 타임아웃(초, 기본 180)
- `--no-fallback`: `hwp` 파일에서 `hwp5txt` 폴백 비활성화

## 의존성

- Python 패키지:

```bash
uv pip install -r any2pdf/requirements.txt
```

- LibreOffice CLI(`soffice` 또는 `libreoffice`)가 **필수 의존성**입니다. (없으면 변환이 실패합니다)

### macOS

```bash
brew install --cask libreoffice
```

### Windows

1) 자동 설치(권장)

```powershell
winget install --id TheDocumentFoundation.LibreOffice -e --accept-package-agreements --accept-source-agreements
```

2) 수동 설치

- 공식 설치 페이지:

```text
https://www.libreoffice.org/download
```

- 설치 후 PATH 등록 확인:

```powershell
$env:Path -split ';' | Select-String 'LibreOffice.*program'
```

- 또는 기본 경로(`C:\Program Files\LibreOffice\program`)가 없으면 수동으로 PATH에 추가하세요.
- 확인:

```powershell
soffice --version
```

## 실행 전/후 점검 체크리스트

### 실행 전

- 입력 폴더에 `.hwp/.hwpx/.doc/.docx/.md` 파일이 있는지 확인
- 가상환경 활성화 여부 확인
- `uv pip install -r any2pdf/requirements.txt` 실행 여부 확인
- `soffice --version` 또는 `libreoffice --version` 동작 여부 확인

### 실행 후

- 출력 폴더에 `<파일명>_<원본확장자>.pdf` 생성 확인
- 실패 로그의 `[FAIL]` 항목이 있으면 파일별 원인 메시지 점검
- `hwp` 처리 실패 시 필요 시 `--overwrite`로 재시도 또는 `--no-fallback` 정책 점검

### 실행 전/후 즉시 점검 (복붙용)

#### macOS

```bash
cd /Users/moon/Desktop/text_files
source .venv/bin/activate
uv pip install -r any2pdf/requirements.txt
soffice --version || libreoffice --version || { echo "LibreOffice 미설치"; exit 1; }
python any2pdf/any2pdf.py "입력_폴더_경로" -o "출력_폴더_경로" --recursive
ls -1 "출력_폴더_경로" | head
```

#### Windows (PowerShell)

```powershell
cd C:\Users\<user>\Desktop\text_files
.venv\Scripts\Activate.ps1
uv pip install -r any2pdf\requirements.txt
soffice --version
if ($LASTEXITCODE -ne 0) { libreoffice --version }
python any2pdf\any2pdf.py "입력_폴더_경로" -o "출력_폴더_경로" --recursive
Get-ChildItem "출력_폴더_경로" -File | Select-Object -ExpandProperty Name | Select-Object -First 20
```
