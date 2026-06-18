import datetime
import logging
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import app_bien_muc_60s as core
from app_bien_muc_60s import A090_PLACEHOLDER


APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(APP_DIR, "app_bien_muc_60s_trua.log")
APP_TITLE = "Tự động biên mục 60s trưa"

TRUA_CONFIG = {
    "label": "60s trưa",
    "list_prefix": "BT60STRUA_",
    "output_prefix": "Map_BT60STRUA",
    "a911": "Nguyễn Thị Quyên",
}

ENDING_CREW_LABELS = (
    "Chịu trách nhiệm nội dung",
    "Biên tập",
    "Biên dịch",
    "Dẫn chương trình",
    "Đạo diễn",
    "Kỹ thuật",
    "Thư ký biên tập",
    "Đồ họa",
    "Tư vấn trang phục",
    "Trang phục",
    "Trang điểm",
    "Website",
    "Fanpage",
    "Kênh Youtube",
)


logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
    force=True,
)


def extract_ending_crew_rows(input_dir: str, log) -> list[str]:
    ending_path = os.path.join(input_dir, "ENDING.rtf")
    if not os.path.exists(ending_path):
        raise RuntimeError("Không tìm thấy file ENDING.rtf để lấy thông tin $a500.")

    log("Đọc ekip từ ENDING.rtf")
    text = core.rtf_to_text(core.read_rtf_raw(ending_path)).replace("\r", "")
    values = {}
    for raw_line in text.split("\n"):
        line = core.clean_line(raw_line)
        if ":" not in line:
            continue
        label, value = line.split(":", 1)
        label = core.clean_line(label)
        if label in ENDING_CREW_LABELS:
            values[label] = core.clean_line(value)

    return [f"{label}: {values.get(label, '')}" for label in ENDING_CREW_LABELS]


def build_map_trua(input_dir: str, output_dir: str, a090: str, a911: str, log) -> str:
    original_extract_crew_data = core.extract_crew_data
    original_build_crew_rows = core.build_crew_rows

    try:
        core.extract_crew_data = extract_ending_crew_rows
        core.build_crew_rows = lambda crew_rows: crew_rows
        return core.build_map(input_dir, output_dir, a090, a911, TRUA_CONFIG, log)
    finally:
        core.extract_crew_data = original_extract_crew_data
        core.build_crew_rows = original_build_crew_rows


class BienMuc60sTruaApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("860x610")

        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.a090 = tk.StringVar()
        self.a911 = tk.StringVar(value=TRUA_CONFIG["a911"])
        self.a090_entry = None

        self._build_ui()

    def _build_ui(self):
        frm_title = ttk.LabelFrame(self.root, text="Bản tin", padding=10)
        frm_title.pack(fill="x", padx=10, pady=(10, 0))
        ttk.Label(frm_title, text="60s trưa").pack(anchor="w")

        frm_paths = ttk.LabelFrame(self.root, text="Đường dẫn", padding=10)
        frm_paths.pack(fill="x", padx=10, pady=10)

        ttk.Label(frm_paths, text="Thư mục input").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frm_paths, textvariable=self.input_dir, width=76).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(frm_paths, text="Chọn...", command=self.choose_input).grid(row=0, column=2)

        ttk.Label(frm_paths, text="Thư mục output").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frm_paths, textvariable=self.output_dir, width=76).grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Button(frm_paths, text="Chọn...", command=self.choose_output).grid(row=1, column=2)
        ttk.Label(
            frm_paths,
            text="(Mặc định tạo folder 'output' trong input nếu để trống)",
            font=("Arial", 8, "italic"),
        ).grid(row=2, column=1, sticky="w", padx=6, pady=(0, 4))
        frm_paths.columnconfigure(1, weight=1)

        frm_note = ttk.LabelFrame(self.root, text="Lưu ý định dạng đầu vào để bóc tách đúng", padding=(10, 6))
        frm_note.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Label(
            frm_note,
            justify="left",
            font=("Arial", 9),
            text=(
                f"• File LIST (Excel): Bắt đầu bằng '{TRUA_CONFIG['list_prefix']}' (.xlsx).\n"
                "  Cột A: tên file bắt đầu bằng '60s', 'gat60s ', '60st' hoặc 'live -'.\n"
                "  Không bắt nếu Cột A bắt đầu bằng '60s W '.\n"
                "  Bỏ qua nếu Cột A chứa: 'coming up', 'nhung nguoi thuc hien', ' end'.\n"
                "  Cột C: bắt buộc có ID dạng số; mã bắt đầu bằng chữ sẽ không bắt.\n"
                "• File RTF tin tức: tên file nên khớp với Cột A trong LIST để app tìm đúng kịch bản."
            ),
        ).pack(anchor="w")

        frm_meta = ttk.LabelFrame(self.root, text="Thông tin Map", padding=10)
        frm_meta.pack(fill="x", padx=10, pady=(0, 10))

        ttk.Label(frm_meta, text="Mã bản tin $a090").grid(row=0, column=0, sticky="w", pady=4)
        self.a090_entry = tk.Entry(frm_meta, textvariable=self.a090, width=24)
        self.a090_entry.grid(row=0, column=1, sticky="w", padx=6)
        self._install_placeholder(self.a090_entry, A090_PLACEHOLDER)

        ttk.Label(frm_meta, text="Người biên mục $a911").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frm_meta, textvariable=self.a911, width=40).grid(row=1, column=1, sticky="w", padx=6)

        frm_actions = ttk.Frame(self.root, padding=(10, 0))
        frm_actions.pack(fill="x")
        self.btn_start = ttk.Button(frm_actions, text="Bắt đầu biên mục", command=self.start)
        self.btn_start.pack(side="right")

        self.log_box = scrolledtext.ScrolledText(self.root, state="disabled", wrap="word", font=("Consolas", 10))
        self.log_box.pack(fill="both", expand=True, padx=10, pady=10)

    def _install_placeholder(self, entry: tk.Entry, text: str):
        entry.insert(0, text)
        entry.configure(foreground="grey")

        def clear_placeholder(_event=None):
            if entry.get() == text and str(entry.cget("foreground")) == "grey":
                entry.delete(0, tk.END)
                entry.configure(foreground="black")

        def restore_placeholder(_event=None):
            if not entry.get().strip():
                entry.insert(0, text)
                entry.configure(foreground="grey")

        entry.bind("<FocusIn>", clear_placeholder)
        entry.bind("<FocusOut>", restore_placeholder)

    def choose_input(self):
        path = filedialog.askdirectory(title="Chọn thư mục input")
        if path:
            self.input_dir.set(path)

    def choose_output(self):
        path = filedialog.askdirectory(title="Chọn thư mục output")
        if path:
            self.output_dir.set(path)

    def log(self, message: str):
        logging.info(message)
        self.root.after(0, self._append_log, message)

    def _append_log(self, message: str):
        self.log_box.configure(state="normal")
        self.log_box.insert(tk.END, f"[{datetime.datetime.now():%H:%M:%S}] {message}\n")
        self.log_box.see(tk.END)
        self.log_box.configure(state="disabled")

    def show_success_dialog(self, out_path: str):
        top = tk.Toplevel(self.root)
        top.title("Thành công")
        top.resizable(False, False)
        top.transient(self.root)
        top.grab_set()

        frm = ttk.Frame(top, padding=16)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Đã tạo file Map thành công:").pack(anchor="w")
        ttk.Label(frm, text=out_path, wraplength=560).pack(anchor="w", pady=(6, 14))

        buttons = ttk.Frame(frm)
        buttons.pack(anchor="e")

        def open_output_folder():
            folder = os.path.dirname(out_path)
            try:
                os.startfile(folder)
            except Exception as exc:
                messagebox.showerror("Lỗi", f"Không mở được thư mục output:\n{exc}", parent=top)

        ttk.Button(buttons, text="Open output folder", command=open_output_folder).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="OK", command=top.destroy).pack(side="left")

        top.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - top.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - top.winfo_height()) // 2
        top.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def start(self):
        input_dir = self.input_dir.get().strip()
        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showerror("Lỗi", "Vui lòng chọn thư mục input hợp lệ.")
            return

        output_dir = self.output_dir.get().strip() or os.path.join(input_dir, "output")
        a090 = self.a090.get().strip()
        if self.a090_entry and a090 == A090_PLACEHOLDER and str(self.a090_entry.cget("foreground")) == "grey":
            a090 = ""
        if not a090:
            messagebox.showerror("Lỗi", "Vui lòng nhập mã $a090.")
            return

        self.btn_start.configure(state="disabled")
        self.log(f"Bắt đầu biên mục {TRUA_CONFIG['label']}...")

        def worker():
            try:
                out_path = build_map_trua(input_dir, output_dir, a090, self.a911.get().strip(), self.log)
                self.log(f"Hoàn tất: {out_path}")
                self.root.after(0, lambda: self.show_success_dialog(out_path))
            except Exception as exc:
                logging.exception("Lỗi khi biên mục %s", TRUA_CONFIG["label"])
                self.log(f"LỖI: {exc}")
                self.root.after(0, lambda: messagebox.showerror("Lỗi", str(exc)))
            finally:
                self.root.after(0, lambda: self.btn_start.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = BienMuc60sTruaApp(root)
    root.mainloop()
