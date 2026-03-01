#!/usr/bin/env python3
"""GUI launcher for the converters in this repository."""

from __future__ import annotations

import importlib
import os
import shutil
import queue
import subprocess
import threading
import tkinter as tk
import sys
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter import ttk


class ConverterGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("any2pdf 통합 변환기")
        self.geometry("900x780")
        self.minsize(820, 640)

        self.log_queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self.running_threads: list[threading.Thread] = []
        self.running = False

        self._build_ui()
        self._drain_log_queue()

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=10)
        container.pack(fill="both", expand=True)

        notebook = ttk.Notebook(container)
        notebook.pack(fill="both", expand=True)

        self._build_pdf_tab(notebook)
        self._build_hwp_tab(notebook)
        self._build_md_tab(notebook)

        log_frame = ttk.LabelFrame(container, text="실행 로그")
        log_frame.pack(fill="both", expand=True, pady=(10, 0))
        self.log_text = tk.Text(log_frame, height=16, wrap="word", state="disabled")
        self.log_text.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.pack(side="right", fill="y", pady=8)
        self.log_text.configure(yscrollcommand=log_scroll.set)

        bottom = ttk.Frame(container)
        bottom.pack(fill="x", pady=(6, 0))
        self.status_label = ttk.Label(bottom, text="대기 중")
        self.status_label.pack(side="left")
        self.clear_btn = ttk.Button(bottom, text="로그 지우기", command=self._clear_log)
        self.clear_btn.pack(side="right")

    def _build_pdf_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=12)
        notebook.add(frame, text="any2pdf (PDF)")

        self.any_input_var = tk.StringVar()
        self.any_output_var = tk.StringVar()
        self.any_recursive_var = tk.BooleanVar(value=True)
        self.any_overwrite_var = tk.BooleanVar(value=False)
        self.any_no_fallback_var = tk.BooleanVar(value=False)
        self.any_timeout_var = tk.StringVar(value="180")

        row = self._build_tab_header(
            frame,
            "입력 경로",
            self.any_input_var,
            self._run_any2pdf,
            filetypes=[("지원 파일", "*.hwp *.hwpx *.doc *.docx *.md"), ("모든 파일", "*")],
            allow_file=True,
        )
        row = self._build_path_row(frame, "출력 폴더", self.any_output_var, start_row=row)
        row = self._build_option_row(
            frame,
            [
                ("하위 폴더까지", self.any_recursive_var),
                ("기존 파일 덮어쓰기", self.any_overwrite_var),
                ("hwp 텍스트 폴백 비활성", self.any_no_fallback_var),
            ],
            row,
        )
        row = self._build_entry_row(frame, "타임아웃(초)", self.any_timeout_var, row)

        note = ttk.Label(
            frame,
            text="지원 형식: .hwp, .hwpx, .doc, .docx, .md",
            foreground="gray",
        )
        note.grid(row=row, column=0, columnspan=3, sticky="w", pady=(12, 0))

    def _build_hwp_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=12)
        notebook.add(frame, text="hwp2docx (DOCX)")

        self.hwp_input_var = tk.StringVar()
        self.hwp_output_var = tk.StringVar()
        self.hwp_recursive_var = tk.BooleanVar(value=True)
        self.hwp_overwrite_var = tk.BooleanVar(value=False)
        self.hwp_no_fallback_var = tk.BooleanVar(value=False)
        self.hwp_timeout_var = tk.StringVar(value="180")

        row = self._build_tab_header(
            frame,
            "입력 경로",
            self.hwp_input_var,
            self._run_hwp2docx,
            filetypes=[("HWP/HWPX", "*.hwp *.hwpx"), ("모든 파일", "*")],
            allow_file=True,
        )
        row = self._build_path_row(frame, "출력 폴더", self.hwp_output_var, start_row=row)
        row = self._build_option_row(
            frame,
            [
                ("하위 폴더까지", self.hwp_recursive_var),
                ("기존 파일 덮어쓰기", self.hwp_overwrite_var),
                ("hwp 텍스트 폴백 비활성", self.hwp_no_fallback_var),
            ],
            row,
        )
        row = self._build_entry_row(frame, "타임아웃(초)", self.hwp_timeout_var, row)
        note = ttk.Label(
            frame,
            text="지원 형식: .hwp, .hwpx",
            foreground="gray",
        )
        note.grid(row=row, column=0, columnspan=3, sticky="w", pady=(12, 0))

    def _build_md_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=12)
        notebook.add(frame, text="md2docx")

        self.md_input_var = tk.StringVar()
        self.md_output_var = tk.StringVar()
        self.md_recursive_var = tk.BooleanVar(value=True)
        self.md_overwrite_var = tk.BooleanVar(value=False)
        self.md_korean_font_var = tk.StringVar(value="Malgun Gothic")
        self.md_code_font_var = tk.StringVar(value="Consolas")

        row = self._build_tab_header(
            frame,
            "입력 경로",
            self.md_input_var,
            self._run_md2docx,
            filetypes=[("Markdown", "*.md"), ("모든 파일", "*")],
            allow_file=True,
        )
        row = self._build_path_row(frame, "출력 폴더", self.md_output_var, start_row=row)
        row = self._build_option_row(
            frame,
            [
                ("하위 폴더까지", self.md_recursive_var),
                ("기존 파일 덮어쓰기", self.md_overwrite_var),
            ],
            row,
        )
        row = self._build_entry_row(frame, "한글 본문 폰트", self.md_korean_font_var, row)
        row = self._build_entry_row(frame, "코드 폰트", self.md_code_font_var, row)

        note = ttk.Label(
            frame,
            text="지원 형식: .md",
            foreground="gray",
        )
        note.grid(row=row, column=0, columnspan=3, sticky="w", pady=(12, 0))

    def _build_tab_header(
        self,
        parent: ttk.Frame,
        label: str,
        path_var: tk.StringVar,
        action,
        filetypes: list[tuple[str, str]] | None = None,
        allow_file: bool = False,
        start_row: int = 0,
    ) -> int:
        ttk.Label(parent, text=label).grid(row=start_row, column=0, sticky="w")
        path_row = ttk.Frame(parent)
        path_row.grid(row=start_row + 1, column=0, columnspan=3, sticky="ew", pady=(4, 10))
        path_row.columnconfigure(0, weight=1)
        entry = ttk.Entry(path_row, textvariable=path_var)
        entry.grid(row=0, column=0, sticky="ew")
        btn_frame = ttk.Frame(path_row)
        btn_frame.grid(row=0, column=1, padx=(6, 0))
        ttk.Button(btn_frame, text="폴더 선택", command=lambda: self._pick_folder(path_var)).pack(side="left")
        if allow_file:
            ttk.Button(
                btn_frame,
                text="파일 선택",
                command=lambda: self._pick_input_file(path_var, filetypes or [("모든 파일", "*")]),
            ).pack(side="left", padx=(6, 0))
        start_btn = ttk.Button(parent, text="실행", command=action)
        start_btn.grid(row=start_row + 1, column=2, padx=(6, 0))
        self._register_start_button(start_btn)
        return start_row + 2

    def _build_path_row(self, parent: ttk.Frame, label: str, var: tk.StringVar, start_row: int) -> int:
        row = start_row
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(0, 2))
        row_frame = ttk.Frame(parent)
        row_frame.grid(row=row + 1, column=0, columnspan=3, sticky="ew")
        row_frame.columnconfigure(0, weight=1)
        entry = ttk.Entry(row_frame, textvariable=var)
        entry.grid(row=0, column=0, sticky="ew")
        ttk.Button(row_frame, text="선택", command=lambda: self._pick_folder(var)).grid(row=0, column=1, padx=(6, 0))
        return row + 2

    def _build_option_row(self, parent: ttk.Frame, options: list[tuple[str, tk.BooleanVar]], row: int) -> int:
        option_frame = ttk.Frame(parent)
        option_frame.grid(row=row, column=0, columnspan=3, sticky="w", pady=(8, 0))
        for text, var in options:
            ttk.Checkbutton(option_frame, text=text, variable=var).pack(side="left", padx=(0, 14))
        return row + 1

    def _build_entry_row(self, parent: ttk.Frame, label: str, var: tk.StringVar, row: int) -> int:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(8, 2))
        entry = ttk.Entry(parent, textvariable=var, width=25)
        entry.grid(row=row, column=1, sticky="w")
        return row + 1

    def _register_start_button(self, btn: ttk.Button) -> None:
        if not hasattr(self, "_start_buttons"):
            self._start_buttons = []
        self._start_buttons.append(btn)

    def _set_running_state(self, running: bool) -> None:
        self.running = running
        for btn in self._start_buttons:
            btn.config(state=("disabled" if running else "normal"))
        self.status_label.config(text=("작업 중..." if running else "대기 중"))

    def _pick_folder(self, target: tk.StringVar) -> None:
        path = filedialog.askdirectory(title="폴더 선택")
        if path:
            target.set(path)

    def _pick_input_file(self, target: tk.StringVar, filetypes: list[tuple[str, str]]) -> None:
        path = filedialog.askopenfilename(title="파일 선택", filetypes=filetypes)
        if path:
            target.set(path)

    def _log(self, message: str) -> None:
        self.log_queue.put(("log", message))

    def _log_done(self, success: int, fail: int, label: str) -> None:
        self.log_queue.put(("done", success, fail, label))

    def _clear_log(self) -> None:
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _append_log(self, message: str) -> None:
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _drain_log_queue(self) -> None:
        while True:
            try:
                item = self.log_queue.get_nowait()
            except queue.Empty:
                break

            kind = item[0]
            if kind == "log":
                self._append_log(item[1])  # type: ignore[index]
            elif kind == "done":
                success, fail, label = item[1], item[2], item[3]
                self._append_log(f"{label} 종료 - 성공: {success}, 실패: {fail}")
                self._set_running_state(False)

        self.after(120, self._drain_log_queue)

    def _collect_files(self, root: Path, exts: set[str], recursive: bool) -> list[Path]:
        if root.is_file():
            return [root] if root.suffix.lower() in exts else []
        if not root.exists() or not root.is_dir():
            return []

        if recursive:
            files: list[Path] = []
            for ext in sorted(exts):
                files.extend(root.rglob(f"*{ext}"))
            return sorted(set(files))

        return sorted([p for p in root.iterdir() if p.is_file() and p.suffix.lower() in exts])

    def _safe_int(self, raw: str) -> int:
        try:
            value = int(raw.strip())
            if value < 1:
                raise ValueError
            return value
        except Exception as e:
            raise ValueError("타임아웃은 1 이상의 정수여야 합니다.") from e

    def _run_any2pdf(self) -> None:
        if self.running:
            return
        self._start_conversion(
            target="any2pdf",
            input_var=self.any_input_var,
            output_var=self.any_output_var,
            recursive=self.any_recursive_var.get(),
            overwrite=self.any_overwrite_var.get(),
            timeout=self.any_timeout_var.get(),
            extra=self.any_no_fallback_var.get(),
            worker=self._any2pdf_worker,
        )

    def _run_hwp2docx(self) -> None:
        if self.running:
            return
        self._start_conversion(
            target="hwp2docx",
            input_var=self.hwp_input_var,
            output_var=self.hwp_output_var,
            recursive=self.hwp_recursive_var.get(),
            overwrite=self.hwp_overwrite_var.get(),
            timeout=self.hwp_timeout_var.get(),
            extra=self.hwp_no_fallback_var.get(),
            worker=self._hwp2docx_worker,
        )

    def _run_md2docx(self) -> None:
        if self.running:
            return
        self._start_conversion(
            target="md2docx",
            input_var=self.md_input_var,
            output_var=self.md_output_var,
            recursive=self.md_recursive_var.get(),
            overwrite=self.md_overwrite_var.get(),
            timeout="180",
            extra=None,
            worker=self._md2docx_worker,
        )

    def _start_conversion(self, target: str, input_var: tk.StringVar, output_var: tk.StringVar,
                         recursive: bool, overwrite: bool, timeout: str, extra: bool | None,
                         worker) -> None:
        path_text = input_var.get().strip()
        if not path_text:
            messagebox.showwarning("경고", "입력 폴더를 먼저 선택하세요.")
            return

        in_path = Path(path_text).expanduser()
        if not in_path.exists():
            messagebox.showerror("오류", "입력 경로를 찾을 수 없습니다.")
            return

        output = output_var.get().strip()
        out_path = Path(output).expanduser() if output else None
        if out_path:
            out_path.mkdir(parents=True, exist_ok=True)

        try:
            timeout_int = self._safe_int(timeout)
        except Exception as e:
            messagebox.showerror("오류", str(e))
            return

        self._set_running_state(True)
        self._log(f"[{target}] 시작 - 입력: {in_path}")
        thread = threading.Thread(
            target=worker,
            args=(in_path, out_path, recursive, overwrite, timeout_int, extra),
            daemon=True,
        )
        self.running_threads.append(thread)
        thread.start()

    def _any2pdf_worker(self, input_path: Path, out_path: Path | None, recursive: bool,
                        overwrite: bool, timeout: int, no_fallback: bool) -> None:
        try:
            mod = importlib.import_module("any2pdf.any2pdf")
            files = mod.collect_inputs(input_path, recursive)
            if not files:
                self._log("[any2pdf] 지원하는 파일이 없습니다.")
                self._log_done(0, 0, "any2pdf")
                return

            success = fail = 0
            for src in files:
                target_dir = out_path if out_path else src.parent
                result = mod.convert_one(src, target_dir, overwrite, timeout, no_fallback)
                if result.success:
                    success += 1
                    self._log(f"[OK] {src} -> {result.output}")
                else:
                    fail += 1
                    self._log(f"[FAIL] {src} -> {result.message}")
            self._log_done(success, fail, "any2pdf")
        except Exception as e:
            self._log(f"[any2pdf] 실행 오류: {e}")
            self._log_done(0, 1, "any2pdf")

    def _hwp2docx_worker(self, input_path: Path, out_path: Path | None, recursive: bool,
                         overwrite: bool, timeout: int, no_fallback: bool) -> None:
        try:
            mod = importlib.import_module("hwp2docx.hwp_to_docx")
            files = self._collect_files(input_path, {".hwp", ".hwpx"}, recursive)
            if not files:
                self._log("[hwp2docx] 지원하는 파일이 없습니다.")
                self._log_done(0, 0, "hwp2docx")
                return

            success = fail = 0
            for src in files:
                target_dir = out_path if out_path else src.parent
                result = mod.convert_one(src, target_dir, overwrite, timeout, no_fallback)
                if result.success:
                    success += 1
                    self._log(f"[OK] {src} -> {result.message}")
                else:
                    fail += 1
                    self._log(f"[FAIL] {src} -> {result.message}")
            self._log_done(success, fail, "hwp2docx")
        except Exception as e:
            self._log(f"[hwp2docx] 실행 오류: {e}")
            self._log_done(0, 1, "hwp2docx")

    def _md2docx_worker(self, input_path: Path, out_path: Path | None, recursive: bool,
                        overwrite: bool, timeout: int, no_fallback: bool | None) -> None:
        _ = timeout
        _ = no_fallback
        try:
            mod = importlib.import_module("md2docx.md_to_docx")
            files = self._collect_files(input_path, {".md"}, recursive)
            if not files:
                self._log("[md2docx] 지원하는 파일이 없습니다.")
                self._log_done(0, 0, "md2docx")
                return

            korean_font = self.md_korean_font_var.get().strip() or "Malgun Gothic"
            code_font = self.md_code_font_var.get().strip() or "Consolas"

            success = fail = 0
            for src in files:
                target_dir = out_path if out_path else src.parent
                target_dir.mkdir(parents=True, exist_ok=True)
                output_path = target_dir / f"{src.stem}.docx"
                if output_path.exists() and not overwrite:
                    fail += 1
                    self._log(f"[FAIL] {src} -> 이미 출력 파일이 존재합니다: {output_path}")
                    continue

                try:
                    mod.convert_md_to_docx(src, output_path, korean_font, code_font)
                    success += 1
                    self._log(f"[OK] {src} -> {output_path}")
                except Exception as e:
                    fail += 1
                    self._log(f"[FAIL] {src} -> {e}")
            self._log_done(success, fail, "md2docx")
        except Exception as e:
            self._log(f"[md2docx] 실행 오류: {e}")
            self._log_done(0, 1, "md2docx")

def _run_via_xvfb(argv: list[str]) -> int:
    if not sys.platform.startswith("linux"):
        return 1

    if os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"):
        return 1

    if os.environ.get("ANY2PDF_XVFB_RETRY") == "1":
        return 1

    xvfb = shutil.which("xvfb-run")
    if not xvfb:
        return 1

    print("GUI 표시 환경이 없어 xvfb-run으로 재시도합니다.")
    script_dir = Path(__file__).resolve().parent
    env = os.environ.copy()
    env["PYTHONPATH"] = str(script_dir)
    env["ANY2PDF_XVFB_RETRY"] = "1"
    return subprocess.run(
        [xvfb, "-a", sys.executable, "-m", "gui"] + argv[1:],
        cwd=str(script_dir),
        env=env,
    ).returncode


def _has_display() -> bool:
    if os.name == "nt":
        return True
    if sys.platform == "darwin":
        return True
    return bool(os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"))


def _is_display_error(exc: tk.TclError) -> bool:
    msg = str(exc).lower()
    return "display" in msg or "cannot connect" in msg or "xdg runtime dir" in msg


def main() -> None:
    if not _has_display():
        code = _run_via_xvfb(list(sys.argv))
        if code != 0:
            print("GUI를 시작할 수 없습니다.")
            print("현재 환경에 디스플레이가 없습니다. Xvfb가 없거나 실행 실패입니다.")
            print("로컬 데스크톱(또는 X/Wayland 포워딩)에서 실행하거나 xvfb-run을 설치해 주세요.")
            print("예: sudo apt-get install xvfb")
        sys.exit(code if code != 0 else 0)

    try:
        app = ConverterGUI()
        app.mainloop()
    except tk.TclError as e:
        if _is_display_error(e):
            code = _run_via_xvfb(list(sys.argv))
            if code == 0:
                return
            print("GUI 표시 연결이 실패해 xvfb-run 재시도도 실패했습니다.")
            print("현재 환경에 디스플레이가 없습니다. Xvfb가 없거나 실행 실패입니다.")
            print("로컬 데스크톱(또는 X/Wayland 포워딩)에서 실행하거나 xvfb-run을 설치해 주세요.")
            print("예: sudo apt-get install xvfb")
        else:
            print("GUI를 시작할 수 없습니다.")
            print(f"오류: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
