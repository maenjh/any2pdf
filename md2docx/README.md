# md2docx

## 역할

Markdown(`.md`) 파일을 Word(`.docx`) 형식으로 변환합니다.

`any2pdf`에서는 `md` 입력을 PDF로 바꾸기 전에 이 모듈을 통해 중간 `docx`를 생성합니다.

## 사용법

```bash
cd /Users/moon/Desktop/any2pdf
source .venv/bin/activate
python md2docx/md_to_docx.py input.md [output.docx]
```

### 옵션

- `input`: 필수, 변환할 `.md` 파일
- `output`: 선택, 저장할 `.docx` 경로(미지정 시 동일 경로에 동일 이름 저장)
- `--font`: 본문 한글 폰트(기본: `Malgun Gothic`)
- `--code-font`: 코드용 폰트(기본: `Consolas`)

### 설치

```bash
uv pip install -r md2docx/requirements.txt
```

### 유의사항

- 입력 파일은 UTF-8 인코딩을 권장합니다.
