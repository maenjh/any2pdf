#!/usr/bin/env python3
"""Convert multiple office-like documents in one folder to PDF.

Supported inputs: .hwp, .hwpx, .doc, .docx, .md

Flow:
- Use LibreOffice CLI (soffice/libreoffice) whenever possible
- For .md: convert to .docx first via existing md_to_docx script, then to PDF
- For .hwp: try LibreOffice first, then fallback to hwp5txt -> python-docx text conversion
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional


SUPPORTED_EXTS = {".hwp", ".hwpx", ".doc", ".docx", ".md"}


@dataclass
class ConvertResult:
    src: Path
    success: bool
    output: Optional[Path]
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
        "pdf",
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

    target = out_dir / f"{src.stem}.pdf"
    if not target.exists():
        candidates = list(out_dir.glob(f"{src.stem}*.pdf"))
        if not candidates:
            raise RuntimeError("LibreOffice 실행은 성공했지만 PDF 결과 파일을 찾지 못했습니다.")
        candidates.sort()
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

    out = out_dir / f"{src.stem}.docx"
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

        text = tmp_path.read_text(encoding="utf-8", errors="ignore")
        doc = Document()
        for line in (text.splitlines() or [""]):
            doc.add_paragraph(line)
        doc.save(str(out))
    return out


def convert_md_to_docx(src: Path, out_dir: Path, timeout: int) -> Path:
    script = Path(__file__).resolve().parent.parent / "md2docx" / "md_to_docx.py"
    out = out_dir / f"{src.stem}.docx"
    completed = subprocess.run(
        [sys.executable, str(script), str(src), str(out)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=timeout,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            "md2docx 변환 실패.\n"
            f"stdout: {completed.stdout.strip()}\nstderr: {completed.stderr.strip()}"
        )
    return out


def ensure_out_dir(path: Optional[Path], default: Path) -> Path:
    out = path if path else default
    out.mkdir(parents=True, exist_ok=True)
    return out


def pdf_output_path(src: Path, out_dir: Path) -> Path:
    ext = src.suffix.lower().lstrip(".") or "file"
    return out_dir / f"{src.stem}_{ext}.pdf"


def move_pdf(source_pdf: Path, target_pdf: Path, overwrite: bool) -> Path:
    if source_pdf == target_pdf:
        return target_pdf
    if target_pdf.exists():
        if not overwrite:
            raise RuntimeError(f"이미 출력 파일이 존재합니다: {target_pdf}")
        target_pdf.unlink()
    source_pdf.rename(target_pdf)
    return target_pdf


def convert_one(src: Path, out_dir: Path, overwrite: bool, timeout: int, no_fallback: bool) -> ConvertResult:
    src = src.expanduser().resolve()
    if not src.exists() or not src.is_file():
        return ConvertResult(src, False, None, "입력 파일이 존재하지 않습니다.")

    ext = src.suffix.lower()
    if ext not in SUPPORTED_EXTS:
        return ConvertResult(src, False, None, "지원되지 않는 확장자입니다.")

    out_dir.mkdir(parents=True, exist_ok=True)
    expected_pdf = pdf_output_path(src, out_dir)
    if expected_pdf.exists() and not overwrite:
        return ConvertResult(src, False, expected_pdf, f"이미 출력 파일이 존재합니다: {expected_pdf}")

    soffice = find_soffice()
    if not soffice:
        return ConvertResult(src, False, None, "LibreOffice가 필요합니다. (soffice 또는 libreoffice 설치 필요)")

    try:
        if ext == ".md":
            with tempfile.TemporaryDirectory() as work:
                work_dir = Path(work)
                docx = convert_md_to_docx(src, work_dir, timeout)
                pdf = move_pdf(convert_with_soffice(soffice, docx, out_dir, timeout), expected_pdf, overwrite)
                return ConvertResult(src, True, pdf, f"변환 완료: {pdf}")

        if ext in {".hwp", ".hwpx"}:
            try:
                pdf = move_pdf(convert_with_soffice(soffice, src, out_dir, timeout), expected_pdf, overwrite)
                return ConvertResult(src, True, pdf, f"변환 완료: {pdf}")
            except Exception as e:
                if no_fallback:
                    return ConvertResult(src, False, None, str(e))
                try:
                    with tempfile.TemporaryDirectory() as work:
                        work_dir = Path(work)
                        fallback_docx = convert_with_hwp5txt(src, work_dir, timeout)
                        if not fallback_docx:
                            return ConvertResult(src, False, None, "폴백 변환 불가: hwp5txt 또는 python-docx 미설치")
                        pdf = move_pdf(convert_with_soffice(soffice, fallback_docx, out_dir, timeout), expected_pdf, overwrite)
                        return ConvertResult(
                            src,
                            True,
                            pdf,
                            f"폴백(hwp5txt)으로 변환 완료: {pdf}",
                        )
                except Exception as e2:
                    return ConvertResult(src, False, None, f"폴백(hwp5txt) 변환 실패: {e2}")

        pdf = move_pdf(convert_with_soffice(soffice, src, out_dir, timeout), expected_pdf, overwrite)
        return ConvertResult(src, True, pdf, f"변환 완료: {pdf}")
    except Exception as e:
        return ConvertResult(src, False, None, str(e))


def collect_inputs(path: Path, recursive: bool) -> List[Path]:
    p = path.expanduser().resolve()
    if p.is_file():
        return [p] if p.suffix.lower() in SUPPORTED_EXTS else []
    if not p.is_dir():
        return []

    files: List[Path] = []
    for child in p.iterdir():
        if child.is_file() and child.suffix.lower() in SUPPORTED_EXTS:
            files.append(child)
        elif child.is_dir() and recursive:
            files.extend(sorted(collect_inputs(child, recursive=True)))
    return sorted(files)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="폴더 내 문서를 PDF로 일괄 변환")
    p.add_argument("input", type=Path, help="변환할 폴더 경로")
    p.add_argument("-o", "--out", type=Path, default=None, help="출력 폴더(기본: 입력 파일과 같은 폴더)")
    p.add_argument("--timeout", type=int, default=180, help="변환 타임아웃(초)")
    p.add_argument("--overwrite", action="store_true", help="기존 PDF 덮어쓰기")
    p.add_argument("--recursive", action="store_true", help="하위 폴더까지 모두 탐색")
    p.add_argument("--no-fallback", action="store_true", help="hwp 텍스트 폴백(hwp5txt) 비활성화")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    inputs = collect_inputs(args.input, args.recursive)
    if not inputs:
        print("변환 대상(.hwp/.hwpx/.doc/.docx/.md)이 없습니다.")
        sys.exit(1)

    overall_ok = True
    for src in inputs:
        out_dir = ensure_out_dir(args.out, src.parent)
        result = convert_one(src, out_dir, args.overwrite, args.timeout, args.no_fallback)
        status = "[OK]" if result.success else "[FAIL]"
        print(f"{status} {src} -> {result.message}")
        overall_ok = overall_ok and result.success

    if not overall_ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
