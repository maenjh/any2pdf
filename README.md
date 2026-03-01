# any2pdf

`any2pdf`는 문서 변환 파이프라인 저장소입니다.

- `any2pdf/`: 입력 폴더의 `hwp`, `hwpx`, `doc`, `docx`, `md` 파일을 한 번에 PDF로 변환
- `md2docx/`: Markdown(`.md`)을 Word(`.docx`)로 변환
- `hwp2docx/`: 한글 문서(`.hwp`, `.hwpx`)를 Word(`.docx`)로 변환(필요 시 텍스트 폴백)

## 빠른 시작

```bash
# 1) 저장소 진입
cd /path/to/any2pdf

# 2) 가상환경(선택)
python -m venv .venv
source .venv/bin/activate

# 3) uv 환경 준비 및 의존성 동기화
uv venv
source .venv/bin/activate
uv sync

# 4) 폴더 단위 PDF 변환 실행
uv run any2pdf-cli "입력_폴더" -o "출력_폴더" --recursive
```

## GUI 실행

```bash
cd /path/to/any2pdf
source .venv/bin/activate
uv run any2pdf-gui
```

화면이 안 뜨는 환경(서버/SSH)이라면 아래처럼 실행해 보세요.

```bash
sudo apt-get install xvfb   # 또는 OS에 맞는 패키지 관리자 사용
xvfb-run -a uv run any2pdf-gui
```

GUI에서 제공되는 기능

- `any2pdf` 탭: 파일 또는 폴더의 `.hwp`, `.hwpx`, `.doc`, `.docx`, `.md`를 재귀 탐색 가능하게 PDF 변환
- `hwp2docx` 탭: 파일 또는 폴더의 `.hwp`, `.hwpx`를 DOCX로 변환
- `md2docx` 탭: 파일 또는 폴더의 `.md`를 DOCX로 변환

출력 파일명은 충돌을 피하기 위해 다음 형식으로 생성됩니다.
- `<원본이름>_<원본확장자>.pdf`
  - 예) `report.docx` → `report_docx.pdf`, `report.hwp` → `report_hwp.pdf`

## 폴더별 가이드

- `any2pdf/README.md`: 배치 PDF 변환기 사용법 및 환경(LibreOffice) 점검 절차
- `md2docx/README.md`: `md -> docx` 변환 사용법
- `hwp2docx/README.md`: `hwp/hwpx -> docx` 변환 사용법

## 환경 요구사항

- Python 3.10 이상
- LibreOffice CLI(`soffice` 또는 `libreoffice`) 설치 (any2pdf 실행에 필수)
- 패키지: `uv`, `python-docx`, `markdown`, `beautifulsoup4`, `pyhwp`, `six`

## 참고

- `any2pdf`의 실행 로그를 보면 입력 파일별 성공/실패를 바로 확인할 수 있습니다.
- `hwp2docx`는 기본적으로 LibreOffice 변환을 우선 시도하고, `.hwp`만 텍스트 기반 폴백 경로(`hwp5txt`)를 시도합니다.
