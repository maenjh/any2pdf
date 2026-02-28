#!/usr/bin/env python3
"""
hwp/hwpx -> docx 변환 유틸리티

- 기본 변환: LibreOffice CLI(`soffice` 또는 `libreoffice`) 사용
- 보조 폴백: `hwp5txt` + `python-docx` (텍스트 기반) 사용
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path
import shutil
import tempfile
from dataclasses import dataclass
from typing import Iterable, List, Optional, Tuple


SUPPORTED_EXTS = {".hwp", ".hwpx"}


@dataclass
class ConvertResult:
    src: Path
    success: bool
    message: str


def find_soffice() -> Optional[str]:
    for cmd in ("soffice", "libreoffice"):
        path = shutil.which(cmd)
        if path:
            return path
    return None


def convert_with_soffice(soffice: str, src: Path, out_dir: Path, timeout: int) -> Path:
    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--norestore",
        "--convert-to",
        "docx",
        "--outdir",
        str(out_dir),
        str(src),
    ]
    completed = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=timeout,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            f"LibreOffice 변환 실패 ({completed.returncode}).\n"
            f"stdout: {completed.stdout.strip()}\nstderr: {completed.stderr.strip()}"
        )

    target = out_dir / f"{src.stem}.docx"
    if not target.exists():
        # 드물게 경로가 충돌하거나 확장자가 바뀌는 경우를 대비한 탐색
        candidates = list(out_dir.glob(f"{src.stem}*.docx"))
        if not candidates:
            raise RuntimeError("LibreOffice 실행은 성공했지만 결과 파일을 찾지 못했습니다.")
        target = candidates[0]
    return target


def convert_with_hwp5txt(src: Path, out_dir: Path, timeout: int) -> Optional[Path]:
    hwp5txt = shutil.which("hwp5txt")
    if not hwp5txt:
        return None

    try:
        from docx import Document  # type: ignore
    except Exception:
        return None

    with tempfile.NamedTemporaryFile(suffix=".txt", delete=True) as tmp:
        tmp_path = Path(tmp.name)
        completed = subprocess.run(
            [hwp5txt, str(src)],
            stdout=tmp,
            stderr=subprocess.PIPE,
            text=True,
            timeout=timeout,
        )
        if completed.returncode != 0:
            raise RuntimeError(
                "hwp5txt 변환 실패.\n"
                f"stderr: {completed.stderr.strip()}"
            )

        doc = Document()
        text = tmp_path.read_text(encoding="utf-8", errors="ignore")
        lines = text.splitlines() or [""]
        for line in lines:
            doc.add_paragraph(line)

        out = out_dir / f"{src.stem}.docx"
        doc.save(str(out))
        return out


def ensure_out_dir(path: Optional[Path], default: Path) -> Path:
    out = path if path else default
    out.mkdir(parents=True, exist_ok=True)
    return out


def convert_one(src: Path, out_dir: Path, overwrite: bool, timeout: int, no_fallback: bool) -> ConvertResult:
    src = src.expanduser().resolve()
    if not src.exists():
        return ConvertResult(src=src, success=False, message="파일이 존재하지 않습니다.")
    if not src.is_file():
        return ConvertResult(src=src, success=False, message="입력은 파일이어야 합니다.")
    if src.suffix.lower() not in SUPPORTED_EXTS:
        return ConvertResult(src=src, success=False, message="지원되지 않는 확장자입니다. (.hwp, .hwpx만 지원)")

    out_dir.mkdir(parents=True, exist_ok=True)
    expected = out_dir / f"{src.stem}.docx"
    if expected.exists() and not overwrite:
        return ConvertResult(src=src, success=False, message=f"이미 출력 파일이 존재합니다: {expected}")

    soffice = find_soffice()
    if soffice:
        try:
            out = convert_with_soffice(soffice, src, out_dir, timeout)
            return ConvertResult(src=src, success=True, message=f"변환 완료: {out}")
        except Exception as e:
            if no_fallback:
                return ConvertResult(src=src, success=False, message=str(e))
            # .hwp의 경우에만 텍스트 기반 폴백 시도

    if not no_fallback and src.suffix.lower() == ".hwp":
        out = convert_with_hwp5txt(src, out_dir, timeout)
        if out:
            return ConvertResult(src=src, success=True, message=f"폴백(hwp5txt)로 변환 완료: {out}")
        return ConvertResult(src=src, success=False, message="폴백 변환 불가: hwp5txt 또는 python-docx 미설치")

    return ConvertResult(
        src=src,
        success=False,
        message="변환할 수 없습니다. LibreOffice가 필요합니다. (soffice 또는 libreoffice 설치 필요)",
    )


def collect_inputs(raw_inputs: Iterable[str]) -> List[Path]:
    files: List[Path] = []
    for item in raw_inputs:
        p = Path(item).expanduser()
        if p.is_dir():
            for ext in SUPPORTED_EXTS:
                files.extend(sorted(p.rglob(f"*{ext}")))
            continue
        files.append(p)
    return files


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=".hwp, .hwpx 파일을 docx로 변환합니다.")
    p.add_argument("inputs", nargs="+", help="변환할 파일 또는 폴더")
    p.add_argument("-o", "--out", type=Path, default=None, help="출력 폴더(기본: 입력 파일과 같은 폴더)")
    p.add_argument("--overwrite", action="store_true", help="기존 docx 덮어쓰기")
    p.add_argument("--timeout", type=int, default=180, help="LibreOffice 실행 타임아웃(초)")
    p.add_argument("--no-fallback", action="store_true", help="텍스트 폴백(hwp5txt) 비활성화")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    inputs = collect_inputs(args.inputs)
    if not inputs:
        print("변환할 대상(.hwp/.hwpx)이 없습니다.")
        sys.exit(1)

    overall_ok = True
    for src in inputs:
        if args.out:
            out_dir = ensure_out_dir(args.out, src.parent)
        else:
            out_dir = src.parent

        result = convert_one(src, out_dir, args.overwrite, args.timeout, args.no_fallback)
        prefix = "[OK]" if result.success else "[FAIL]"
        print(f"{prefix} {src} -> {result.message}")
        overall_ok = overall_ok and result.success

    if not overall_ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
