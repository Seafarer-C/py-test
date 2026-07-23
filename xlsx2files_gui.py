#!/usr/bin/env python3
"""xlsx2files 的桌面图形界面。"""

from __future__ import annotations

import argparse
import os
import queue
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from PIL import Image, ImageTk

import xlsx2files


def default_browser_profile() -> Path:
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
        return base / "LingTu" / "ExcelOrderDownloader" / "browser-profile"
    return Path.home() / ".xlsx2files-browser-profile"


def resource_path(relative: str) -> Path:
    base = Path(sys._MEIPASS) if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS") else Path(__file__).parent
    return base / relative


class QueueWriter:
    def __init__(self, messages: queue.Queue):
        self.messages = messages
        self.buffer = ""

    def write(self, text: str) -> int:
        self.buffer += text.replace("\r", "\n")
        while "\n" in self.buffer:
            line, self.buffer = self.buffer.split("\n", 1)
            if line.strip():
                self.messages.put(("log", line))
        return len(text)

    def flush(self) -> None:
        if self.buffer.strip():
            self.messages.put(("log", self.buffer))
        self.buffer = ""


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel 订单素材下载工具")
        self.geometry("900x650")
        self.minsize(760, 560)
        logo_png = resource_path("assets/lingtu-logo.png")
        logo_ico = resource_path("assets/lingtu-logo.ico")
        self.logo_image = ImageTk.PhotoImage(Image.open(logo_png).resize((44, 44), Image.Resampling.LANCZOS))
        self.iconphoto(True, self.logo_image)
        if sys.platform == "win32" and logo_ico.exists():
            self.iconbitmap(default=str(logo_ico))
        self.messages: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None

        self.xlsx_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str(Path.cwd() / "output"))
        self.limit_enabled = tk.BooleanVar(value=False)
        self.limit_var = tk.StringVar(value="3")
        self.browser_var = tk.StringVar(value="visible")
        self.status_var = tk.StringVar(value="请选择 Excel 文件")
        self.progress_text = tk.StringVar(value="0 / 0")
        self._build()
        self.after(100, self._poll_messages)

    def _build(self) -> None:
        outer = ttk.Frame(self, padding=16)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(8, weight=1)

        brand = ttk.Frame(outer)
        brand.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        brand.columnconfigure(0, weight=1)
        ttk.Label(brand, text="Excel 订单素材下载工具", font=("TkDefaultFont", 20, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(brand, text="Power by 灵图（ipoddy.cn）", foreground="#666666").grid(
            row=0, column=1, sticky="e", padx=(20, 0)
        )

        ttk.Label(outer, text="Excel 文件").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(outer, textvariable=self.xlsx_var).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(outer, text="选择", command=self._choose_xlsx).grid(row=1, column=2)

        ttk.Label(outer, text="输出目录").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(outer, textvariable=self.output_var).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(outer, text="选择", command=self._choose_output).grid(row=2, column=2)

        options = ttk.Frame(outer)
        options.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        ttk.Checkbutton(options, text="仅处理前", variable=self.limit_enabled).pack(side="left")
        ttk.Entry(options, textvariable=self.limit_var, width=6).pack(side="left", padx=5)
        ttk.Label(options, text="条数据").pack(side="left")
        ttk.Label(options, text="浏览器回退").pack(side="left", padx=(28, 6))
        fallback = ttk.Combobox(
            options,
            textvariable=self.browser_var,
            values=("visible", "headless", "off"),
            state="readonly",
            width=12,
        )
        fallback.pack(side="left")
        ttk.Label(options, text="（可见登录 / 无头 / 关闭）").pack(side="left", padx=6)

        buttons = ttk.Frame(outer)
        buttons.grid(row=4, column=0, columnspan=3, sticky="ew", pady=10)
        self.start_button = ttk.Button(buttons, text="开始处理", command=self._start)
        self.start_button.pack(side="left")
        self.stop_button = ttk.Button(buttons, text="停止", command=self._stop, state="disabled")
        self.stop_button.pack(side="left", padx=8)
        ttk.Button(buttons, text="清空日志", command=lambda: self.log.delete("1.0", "end")).pack(side="right")

        status = ttk.Frame(outer)
        status.grid(row=5, column=0, columnspan=3, sticky="ew")
        status.columnconfigure(0, weight=1)
        ttk.Label(status, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        ttk.Label(status, textvariable=self.progress_text).grid(row=0, column=1, sticky="e")

        self.progress = ttk.Progressbar(outer, mode="determinate", maximum=100)
        self.progress.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(5, 10))

        ttk.Label(outer, text="运行日志（包含当前订单、下载文件、解压结果和失败原因）").grid(
            row=7, column=0, columnspan=3, sticky="w"
        )
        self.log = scrolledtext.ScrolledText(outer, height=20, wrap="word", state="normal")
        self.log.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=(5, 0))

    def _choose_xlsx(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")])
        if path:
            self.xlsx_var.set(path)
            if self.output_var.get().endswith("output"):
                self.output_var.set(str(Path(path).with_name(Path(path).stem + "_输出")))

    def _choose_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_var.set(path)

    def _start(self) -> None:
        xlsx = Path(self.xlsx_var.get().strip())
        output = Path(self.output_var.get().strip())
        if not xlsx.is_file():
            messagebox.showerror("无法开始", "请选择有效的 xlsx 文件。")
            return
        try:
            limit = int(self.limit_var.get()) if self.limit_enabled.get() else None
            if limit is not None and limit <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("无法开始", "测试条数必须是正整数。")
            return

        self.cancel_event.clear()
        self.progress["value"] = 0
        self.progress_text.set("0 / 0")
        self.status_var.set("正在读取 Excel…")
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        args = argparse.Namespace(
            xlsx=xlsx,
            output=output,
            sheet=None,
            limit=limit,
            skip_downloads=False,
            browser_fallback=self.browser_var.get(),
            browser_profile=default_browser_profile(),
            cancel_event=self.cancel_event,
        )
        self.worker = threading.Thread(target=self._run, args=(args,), daemon=True)
        self.worker.start()

    def _on_progress(self, event: xlsx2files.ProgressEvent) -> None:
        self.messages.put(("progress", event))

    def _run(self, args: argparse.Namespace) -> None:
        writer = QueueWriter(self.messages)
        old_stdout, old_stderr = sys.stdout, sys.stderr
        try:
            sys.stdout = sys.stderr = writer
            code = xlsx2files.process(args, progress_callback=self._on_progress)
            writer.flush()
            self.messages.put(("done", code))
        except Exception:
            writer.write(traceback.format_exc())
            writer.flush()
            self.messages.put(("done", 1))
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr

    def _stop(self) -> None:
        self.cancel_event.set()
        self.status_var.set("已请求停止，将在当前下载结束后停止…")
        self.stop_button.configure(state="disabled")

    def _poll_messages(self) -> None:
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                if kind == "log":
                    line = str(payload)
                    self.log.insert("end", line + "\n")
                    self.log.see("end")
                elif kind == "progress":
                    event = payload
                    if event.order_total:
                        self.progress["value"] = event.order_index * 100 / event.order_total
                        self.progress_text.set(f"{event.order_index} / {event.order_total}")
                    status = event.message or event.order_no
                    if event.item_total:
                        status = f"{event.order_no}（{event.item_index}/{event.item_total}） {event.message}".strip()
                    self.status_var.set(status)
                elif kind == "done":
                    code = int(payload)
                    self.start_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    if code == 0:
                        self.progress["value"] = 100
                        self.status_var.set("全部处理完成")
                        messagebox.showinfo("处理完成", "全部数据处理完成，没有失败项。")
                    elif code == 130:
                        self.status_var.set("任务已停止")
                    else:
                        self.status_var.set("处理完成，但存在失败项，请查看日志")
                        messagebox.showwarning(
                            "存在失败项",
                            "部分内容未成功处理，请查看运行日志及带“_下载失败”后缀的订单目录。",
                        )
        except queue.Empty:
            pass
        self.after(100, self._poll_messages)


if __name__ == "__main__":
    App().mainloop()
