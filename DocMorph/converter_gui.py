import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from FilesChange import DocumentConverter
from ttkthemes import ThemedTk
import glob
import threading
import queue

class DocumentConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 文档格式转换器")

        os.makedirs("./input", exist_ok=True)
        os.makedirs("./output", exist_ok=True)

        self.converter = DocumentConverter()

        self.input_paths = []
        self.input_dir = tk.StringVar(value="./input")
        self.output_dir = tk.StringVar(value="./output")
        self.from_format = tk.StringVar()
        self.to_format = tk.StringVar()

        self.failed_files = []
        self._task_queue = queue.Queue()

        self.create_widgets()

    def create_widgets(self):
        s = ttk.Style()
        s.configure('TButton', font=('Arial', 10))
        s.configure('TLabel', font=('Arial', 10))
        s.configure('TEntry', font=('Arial', 10))

        file_frame = ttk.LabelFrame(self.root, text="选择文件")
        file_frame.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        ttk.Button(file_frame, text="📁 选择多个文件", command=self.select_input_files).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(file_frame, text="📂 选择文件夹", command=self.select_input_folder).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(file_frame, text="输出目录:").grid(row=1, column=0, sticky="e", padx=5)
        ttk.Entry(file_frame, textvariable=self.output_dir, width=40).grid(row=1, column=1)
        ttk.Button(file_frame, text="浏览", command=self.select_output_folder).grid(row=1, column=2)

        setting_frame = ttk.LabelFrame(self.root, text="转换设置")
        setting_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        ttk.Label(setting_frame, text="源格式:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(setting_frame, textvariable=self.from_format, width=15).grid(row=0, column=1, sticky="w", padx=5)

        ttk.Label(setting_frame, text="目标格式:").grid(row=1, column=0, sticky="e", padx=5)
        self.format_combo = ttk.Combobox(setting_frame, textvariable=self.to_format, width=15, state="readonly")
        self.format_combo.grid(row=1, column=1, sticky="w", padx=5)

        btn_frame = ttk.Frame(self.root)
        btn_frame.grid(row=2, column=0, columnspan=3, pady=10)

        ttk.Button(btn_frame, text="🚀 执行转换", command=self.convert).grid(row=0, column=0, padx=10)
        ttk.Button(btn_frame, text="⚡ 一键转换 input 文件夹", command=self.quick_convert_input_folder).grid(row=0, column=1, padx=10)

        self.progress = ttk.Progressbar(self.root, length=500, mode='determinate')
        self.progress.grid(row=3, column=0, columnspan=3, pady=5)

    def select_input_files(self):
        paths = filedialog.askopenfilenames(initialdir="./input", title="选择多个源文件")
        if paths:
            self.input_paths = list(paths)
            ext = os.path.splitext(paths[0])[1][1:].lower()
            self.from_format.set(ext)
            formats = self.converter.get_supported_conversions(ext)
            self.format_combo["values"] = formats
            if formats:
                self.to_format.set(formats[0])

    def select_input_folder(self):
        path = filedialog.askdirectory(initialdir="./input", title="选择源文件夹")
        if path:
            self.input_paths = []
            self.input_dir.set(path)
            messagebox.showinfo("提示", f"已选择目录：{path}\n将在批量模式下处理该目录")

    def select_output_folder(self):
        path = filedialog.askdirectory(initialdir="./output")
        if path:
            self.output_dir.set(path)

    def convert(self):
        from_fmt = self.from_format.get()
        to_fmt = self.to_format.get()
        out_dir = self.output_dir.get()

        if not from_fmt or not to_fmt:
            messagebox.showerror("错误", "请填写源格式和目标格式")
            return

        try:
            os.makedirs(out_dir, exist_ok=True)
            failed_files = []
            total_files = len(self.input_paths)
            self.progress["value"] = 0
            self.progress["maximum"] = total_files

            success = 0
            for idx, file_path in enumerate(self.input_paths):
                fname = os.path.splitext(os.path.basename(file_path))[0]
                out_path = os.path.join(out_dir, f"{fname}.{to_fmt}")
                try:
                    if self.converter.convert(file_path, out_path, from_fmt, to_fmt):
                        success += 1
                    else:
                        failed_files.append(file_path)
                except Exception:
                    failed_files.append(file_path)

                self.progress["value"] = idx + 1
                self.root.update_idletasks()

            msg = f"✅ 成功转换 {success} 个文件"
            if failed_files:
                msg += f"\n❌ 失败 {len(failed_files)} 个:\n" + "\n".join(os.path.basename(f) for f in failed_files)
            messagebox.showinfo("转换完成", msg)

        except Exception as e:
            messagebox.showerror("转换失败", str(e))

    def quick_convert_input_folder(self):
        to_fmt = self.to_format.get()
        supported_types = self.converter.supported_conversions.keys()

        if not to_fmt:
            messagebox.showerror("错误", "请指定目标格式")
            return

        all_files = []
        for ext in supported_types:
            all_files += glob.glob(f"./input/**/*.{ext}", recursive=True)

        if not all_files:
            messagebox.showinfo("提示", "input 文件夹中没有可转换的文件")
            return

        self.progress["value"] = 0
        self.progress["maximum"] = len(all_files)
        self.failed_files = []

        def convert_task():
            for idx, file_path in enumerate(all_files):
                from_fmt = os.path.splitext(file_path)[1][1:].lower()
                if to_fmt not in self.converter.get_supported_conversions(from_fmt):
                    self.failed_files.append(file_path)
                else:
                    try:
                        filename = os.path.splitext(os.path.basename(file_path))[0]
                        out_path = os.path.join("./output", f"{filename}.{to_fmt}")
                        self.converter.convert(file_path, out_path, from_fmt, to_fmt)
                    except Exception:
                        self.failed_files.append(file_path)
                self._task_queue.put(idx + 1)

            self._task_queue.put("done")

        threading.Thread(target=convert_task, daemon=True).start()
        self.root.after(100, self._process_queue)

    def _process_queue(self):
        try:
            while True:
                value = self._task_queue.get_nowait()
                if value == "done":
                    msg = f"✅ 已尝试转换 {self.progress['maximum']} 个文件"
                    if self.failed_files:
                        msg += f"\n❌ 失败 {len(self.failed_files)} 个:\n" + \
                               "\n".join(os.path.basename(f) for f in self.failed_files)
                    messagebox.showinfo("一键转换完成", msg)
                    return
                else:
                    self.progress["value"] = value
        except queue.Empty:
            self.root.after(100, self._process_queue)


if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    root.geometry("600x300")
    app = DocumentConverterGUI(root)
    root.mainloop()
