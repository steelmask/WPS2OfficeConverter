"""Convert WPS proprietary files to Office-compatible formats on Windows."""

from __future__ import annotations

import shutil
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


FORMAT_GROUPS = {
    "spreadsheet": {
        "label": "WPS 表格", "sources": (".et", ".ett"),
        "targets": {".xlsx": "Excel 工作簿 (.xlsx)", ".xls": "Excel 97-2003 (.xls)"},
        "program_id": "et.Application", "com_formats": {".xlsx": 51, ".xls": 56},
        "lo_formats": {".xlsx": "Calc MS Excel 2007 XML", ".xls": "MS Excel 97"},
    },
    "document": {
        "label": "WPS 文字", "sources": (".wps", ".wpt"),
        "targets": {".docx": "Word 文档 (.docx)", ".doc": "Word 97-2003 (.doc)"},
        "program_id": "wps.Application", "com_formats": {".docx": 16, ".doc": 0},
        "lo_formats": {".docx": "Office Open XML Text", ".doc": "MS Word 97"},
    },
    "presentation": {
        "label": "WPS 演示", "sources": (".dps", ".dpt"),
        "targets": {".pptx": "PowerPoint 演示文稿 (.pptx)", ".ppt": "PowerPoint 97-2003 (.ppt)"},
        "program_id": "wpp.Application", "com_formats": {".pptx": 24, ".ppt": 1},
        "lo_formats": {".pptx": "Impress MS PowerPoint 2007 XML", ".ppt": "MS PowerPoint 97"},
    },
    "fixed_layout": {
        "label": "OFD 版式文档", "sources": (".ofd",),
        "targets": {".pdf": "PDF 文档 (.pdf)"},
        "program_id": "wps.Application", "com_formats": {".pdf": 17},
        "lo_formats": {},
    },
}


def file_group(source: Path) -> str:
    extension = source.suffix.lower()
    for name, group in FORMAT_GROUPS.items():
        if extension in group["sources"]:
            return name
    supported = ", ".join(ext for group in FORMAT_GROUPS.values() for ext in group["sources"])
    raise ConversionError(f"不支持的文件类型：{extension or '无扩展名'}。支持：{supported}")


class ConversionError(RuntimeError):
    pass


def convert_with_wps(source: Path, destination: Path) -> None:
    """Use WPS's corresponding COM application, which preserves formatting best."""
    group_name = file_group(source)
    group = FORMAT_GROUPS[group_name]
    try:
        import win32com.client  # type: ignore[import-not-found]
        import pythoncom  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ConversionError("未安装 pywin32，无法使用 WPS 自动化。") from exc

    # This function runs in a worker thread. COM must be initialized per thread.
    pythoncom.CoInitialize()
    app = None
    book = None
    try:
        app = win32com.client.DispatchEx(group["program_id"])
        app.Visible = False
        app.DisplayAlerts = False
        if group_name == "spreadsheet":
            book = app.Workbooks.Open(str(source.resolve()))
        elif group_name in ("document", "fixed_layout"):
            book = app.Documents.Open(str(source.resolve()))
        else:
            book = app.Presentations.Open(str(source.resolve()), ReadOnly=True, Untitled=False, WithWindow=False)
        book.SaveAs(str(destination.resolve()), FileFormat=group["com_formats"][destination.suffix.lower()])
    except Exception as exc:  # COM errors vary by installed WPS version
        raise ConversionError(f"WPS 转换失败：{exc}") from exc
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def find_soffice() -> str | None:
    """Find LibreOffice without requiring it to be on PATH."""
    candidates = [
        shutil.which("soffice"),
        shutil.which("soffice.exe"),
        r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ]
    return next((item for item in candidates if item and Path(item).exists()), None)


def convert_with_libreoffice(source: Path, destination: Path) -> None:
    soffice = find_soffice()
    if not soffice:
        raise ConversionError("找不到 WPS 或 LibreOffice。请安装 WPS 表格，或安装 LibreOffice 并加入 PATH。")

    group = FORMAT_GROUPS[file_group(source)]
    if destination.suffix.lower() not in group["lo_formats"]:
        raise ConversionError("OFD 转 PDF 需要安装 WPS Office；LibreOffice 不支持可靠导入 OFD。")
    # LibreOffice writes to an output directory rather than an exact filename.
    output_dir = destination.parent
    filter_name = group["lo_formats"][destination.suffix.lower()]
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", f"{destination.suffix[1:]}:{filter_name}",
         "--outdir", str(output_dir), str(source)],
        capture_output=True,
        text=True,
        timeout=120,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    generated = output_dir / f"{source.stem}{destination.suffix}"
    if result.returncode != 0 or not generated.exists():
        detail = (result.stderr or result.stdout).strip()
        raise ConversionError(f"LibreOffice 转换失败。{detail}")
    if generated.resolve() != destination.resolve():
        if destination.exists():
            destination.unlink()
        generated.replace(destination)


def convert(source: Path, destination: Path) -> str:
    try:
        convert_with_wps(source, destination)
        return "已使用 WPS 完成转换"
    except ConversionError as wps_error:
        try:
            convert_with_libreoffice(source, destination)
            return "已使用 LibreOffice 完成转换"
        except ConversionError as libre_error:
            raise ConversionError(f"{wps_error}\n\n备用方案也失败：{libre_error}") from libre_error


class ConverterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("WPS2OfficeConverter")
        self.geometry("620x255")
        self.resizable(False, False)
        self.source = tk.StringVar()
        self.format = tk.StringVar()
        self.format_label = tk.StringVar()
        self.kind = tk.StringVar(value="尚未选择文件")
        self.status = tk.StringVar(value="请选择 WPS 专有格式文件")
        self._build()

    def _build(self) -> None:
        root = ttk.Frame(self, padding=24)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)

        ttk.Label(root, text="源文件：").grid(row=0, column=0, sticky="w", pady=(0, 16))
        ttk.Entry(root, textvariable=self.source, state="readonly").grid(row=0, column=1, sticky="ew", pady=(0, 16))
        ttk.Button(root, text="选择文件", command=self.pick_source).grid(row=0, column=2, padx=(12, 0), pady=(0, 16))

        ttk.Label(root, text="文件类型：").grid(row=1, column=0, sticky="w")
        ttk.Label(root, textvariable=self.kind).grid(row=1, column=1, columnspan=2, sticky="w")
        ttk.Label(root, text="导出格式：").grid(row=2, column=0, sticky="w", pady=(16, 0))
        self.format_box = ttk.Combobox(root, textvariable=self.format_label, state="readonly", width=35)
        self.format_box.grid(row=2, column=1, columnspan=2, sticky="w", pady=(16, 0))

        self.convert_button = ttk.Button(root, text="转换并保存", command=self.start_conversion)
        self.convert_button.grid(row=3, column=1, sticky="w", pady=(26, 12))
        ttk.Label(root, textvariable=self.status, foreground="#444").grid(row=4, column=0, columnspan=3, sticky="w")

    def pick_source(self) -> None:
        patterns = " ".join(ext for group in FORMAT_GROUPS.values() for ext in group["sources"])
        filename = filedialog.askopenfilename(title="选择 WPS 文件", filetypes=[("WPS 专有格式", patterns), ("所有文件", "*.*")])
        if filename:
            source = Path(filename)
            try:
                group = FORMAT_GROUPS[file_group(source)]
            except ConversionError as exc:
                messagebox.showerror("不支持的文件", str(exc))
                return
            self.source.set(filename)
            self.kind.set(group["label"])
            targets = group["targets"]
            self.format_box["values"] = list(targets.values())
            extension, label = next(iter(targets.items()))
            self.format_label.set(label)
            self.format.set(extension)
            self.format_box.bind("<<ComboboxSelected>>", lambda _event: self._sync_extension(group))
            self.status.set("已选择文件，点击“转换并保存”继续")

    def _sync_extension(self, group: dict) -> None:
        label = self.format_label.get()
        for extension, target_label in group["targets"].items():
            if target_label == label:
                self.format.set(extension)
                return

    def start_conversion(self) -> None:
        source_text = self.source.get()
        if not source_text:
            messagebox.showwarning("尚未选择文件", "请先选择一个 .et 文件。")
            return
        source = Path(source_text)
        if not source.is_file():
            messagebox.showerror("文件不存在", "所选文件已不存在或无法访问。")
            return
        group = FORMAT_GROUPS[file_group(source)]
        extension = self.format.get()
        if not extension.startswith("."):
            self._sync_extension(group)
            extension = self.format.get()
        destination_text = filedialog.asksaveasfilename(
            title="保存 Office 文件", defaultextension=extension,
            initialfile=f"{source.stem}{extension}",
            filetypes=[(label, f"*{suffix}") for suffix, label in group["targets"].items()],
        )
        if not destination_text:
            return
        destination = Path(destination_text).with_suffix(extension)
        self.convert_button.config(state="disabled")
        self.status.set("正在转换，请稍候…")
        threading.Thread(target=self._convert_worker, args=(source, destination), daemon=True).start()

    def _convert_worker(self, source: Path, destination: Path) -> None:
        try:
            result = convert(source, destination)
            self.after(0, lambda: self._complete(True, f"{result}：{destination}"))
        except Exception as exc:
            # Exception variables are cleared when an except block ends. Bind the
            # message now, otherwise this deferred callback raises NameError and
            # the user never sees the real conversion error.
            detail = str(exc)
            self.after(0, lambda message=detail: self._complete(False, message))

    def _complete(self, success: bool, detail: str) -> None:
        self.convert_button.config(state="normal")
        self.status.set(detail)
        if success:
            messagebox.showinfo("转换完成", detail)
        else:
            messagebox.showerror("无法转换", detail)


if __name__ == "__main__":
    if sys.platform != "win32":
        raise SystemExit("此工具目前仅支持 Windows。")
    ConverterApp().mainloop()
