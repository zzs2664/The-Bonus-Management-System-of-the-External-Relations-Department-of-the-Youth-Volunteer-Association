"""青志协加分数据库 — 图形界面"""

import sys
import os
import threading
import io
import traceback

sys.stdout.reconfigure(encoding='utf-8')

import customtkinter as ctk
from tkinter import Toplevel, Text, Scrollbar, END, WORD, filedialog, messagebox

import import_scores
import generate_list

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(PROJECT_DIR, '文件模板')
OUTPUT_DIR = os.path.join(PROJECT_DIR, '输出文件夹')


def show_output_window(title_key, text):
    """在 Toplevel 文本窗口中展示运行输出，窗口关闭后自动刷新主界面。"""
    titles = {'import': '导入分数 - 输出', 'generate': '生成名单 - 输出'}

    def _show():
        win = Toplevel()
        win.title(titles.get(title_key, '输出'))
        win.geometry("700x500")
        win.transient()

        frame = ctk.CTkFrame(win)
        frame.pack(fill="both", expand=True, padx=8, pady=8)

        txt = Text(
            frame, wrap=WORD, font=("Consolas", 10),
            borderwidth=0, relief="flat",
        )
        scroll = Scrollbar(frame, command=txt.yview)
        txt.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        txt.pack(fill="both", expand=True)
        txt.insert(END, text or "(无输出)")
        txt.configure(state="disabled")

        # 关闭窗口时恢复主窗口状态
        def _on_close():
            try:
                # 尝试通知主窗口刷新状态
                if hasattr(win, 'master') and win.master:
                    for attr_name in dir(win.master):
                        if attr_name == 'status_var':
                            var = getattr(win.master, 'status_var')
                            if hasattr(var, 'set'):
                                var.set("就绪")
                    if hasattr(win.master, 'deiconify'):
                        win.master.deiconify()
            except Exception:
                pass
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", _on_close)

    # 必须在主线程创建 Tk 组件
    if threading.current_thread() is threading.main_thread():
        _show()
    else:
        # 使用 after_idle 调度到主线程
        import tkinter as tk
        try:
            root = tk._default_root
        except AttributeError:
            root = None
        if root:
            root.after(0, _show)


def run_in_console(script_key, *args):
    """在子线程中运行导入/生成脚本，输出捕获到弹窗。"""
    module_map = {
        'import': import_scores,
        'generate': generate_list,
    }

    def _run():
        old_argv = sys.argv[:]
        sys.argv = [script_key] + list(args)

        buf = io.StringIO()
        old_stdout = sys.stdout
        sys.stdout = buf

        try:
            module_map[script_key].main()
        except SystemExit:
            pass
        except Exception:
            traceback.print_exc()
        finally:
            sys.stdout = old_stdout
            sys.argv = old_argv

        show_output_window(script_key, buf.getvalue())

    threading.Thread(target=_run, daemon=True).start()


# ── 导入分数弹窗 ──

class ImportDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("导入分数")
        self.geometry("500x190")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        # 居中于父窗口
        self.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        w, h = 500, 190
        self.geometry(f"{w}x{h}+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

        ctk.CTkLabel(
            self, text="导入分数（Word → Excel）",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(anchor="w", padx=20, pady=(16, 10))

        r1 = ctk.CTkFrame(self, fg_color="transparent")
        r1.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(r1, text="Word 文件:").pack(side="left")
        self.word_var = ctk.StringVar()
        ctk.CTkEntry(r1, textvariable=self.word_var, width=280).pack(side="left", padx=(8, 6))
        ctk.CTkButton(
            r1, text="浏览...", width=60,
            command=self._browse_word,
        ).pack(side="left")

        r2 = ctk.CTkFrame(self, fg_color="transparent")
        r2.pack(fill="x", padx=20)
        ctk.CTkLabel(r2, text="Excel 目标:").pack(side="left")
        self.excel_var = ctk.StringVar()
        updated = os.path.join(OUTPUT_DIR, '年级花名册数据表_updated.xlsx')
        if os.path.exists(updated):
            self.excel_var.set(updated)
        ctk.CTkEntry(r2, textvariable=self.excel_var, width=280).pack(side="left", padx=(8, 6))
        ctk.CTkButton(
            r2, text="浏览...", width=60,
            command=self._browse_excel,
        ).pack(side="left")

        ctk.CTkButton(
            self, text="导入", width=80,
            command=self._run,
        ).pack(pady=(12, 0))

        self._status = ctk.CTkLabel(self, text="", text_color="gray")
        self._status.pack(pady=(4, 0))

    def _browse_word(self):
        path = filedialog.askopenfilename(
            title="选择 Word 文件",
            initialdir=TEMPLATE_DIR,
            filetypes=[("Word 文件", "*.docx")],
        )
        if path:
            self.word_var.set(path)

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 目标文件",
            initialdir=OUTPUT_DIR,
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.excel_var.set(path)

    def _run(self):
        path = self.word_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择 Word 文件。", parent=self)
            return
        if not os.path.exists(path):
            messagebox.showerror("错误", f"文件不存在:\n{path}", parent=self)
            return

        excel_path = self.excel_var.get().strip()
        if not excel_path:
            messagebox.showwarning("提示", "请先选择 Excel 目标文件。", parent=self)
            return

        self._status.configure(text="导入已启动（查看弹出的 cmd 窗口）")
        run_in_console('import', path, '--excel', excel_path)


# ── 下拉菜单 ──

class DropdownMenu(ctk.CTkToplevel):
    def __init__(self, parent, items, x, y):
        super().__init__(parent)
        self.overrideredirect(True)
        self.transient(parent)

        frame = ctk.CTkFrame(self, border_width=1)
        frame.pack(fill="both", expand=True)

        for label, cmd in items:
            btn = ctk.CTkButton(
                frame, text=label, anchor="w",
                fg_color="transparent", text_color=("gray10", "gray90"),
                hover_color=("gray70", "gray30"),
                command=lambda c=cmd: self._select(c),
            )
            btn.pack(fill="x", padx=2, pady=(2, 0))

        self.update_idletasks()
        w, h = 150, len(items) * 36 + 6
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.focus_set()
        self.bind("<FocusOut>", lambda e: self.destroy())

    def _select(self, cmd):
        self.destroy()
        cmd()


# ── 生成名单测试弹窗 ──

class GenerateTestDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("生成名单-测试")
        self.geometry("580x180")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        w, h = 580, 180
        self.geometry(f"{w}x{h}+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

        ctk.CTkLabel(
            self, text="⚠ 导入的表格文件必须为年级花名册数据表的过往版本，建议开发者使用",
            font=ctk.CTkFont(size=12),
            text_color="#E67E22",
        ).pack(anchor="w", padx=20, pady=(16, 10))

        r1 = ctk.CTkFrame(self, fg_color="transparent")
        r1.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(r1, text="Excel 文件:").pack(side="left")
        self.excel_var = ctk.StringVar()
        ctk.CTkEntry(r1, textvariable=self.excel_var, width=360).pack(side="left", padx=(10, 6))
        ctk.CTkButton(
            r1, text="浏览...", width=60,
            command=self._browse,
        ).pack(side="left")

        r2 = ctk.CTkFrame(self, fg_color="transparent")
        r2.pack(fill="x", padx=20)
        ctk.CTkLabel(r2, text="活动日期:").pack(side="left")
        self.date_var = ctk.StringVar()
        ctk.CTkEntry(r2, textvariable=self.date_var, width=110).pack(side="left", padx=(10, 0))

        ctk.CTkButton(
            self, text="生成名单", width=100,
            command=self._run,
        ).pack(pady=(14, 0))

        self._status = ctk.CTkLabel(self, text="", text_color="gray")
        self._status.pack(pady=(4, 0))

    def _browse(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            initialdir=OUTPUT_DIR,
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.excel_var.set(path)

    def _run(self):
        path = self.excel_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择 Excel 文件。", parent=self)
            return
        if not os.path.exists(path):
            messagebox.showerror("错误", f"文件不存在:\n{path}", parent=self)
            return

        date_str = self.date_var.get().strip()
        if not date_str:
            messagebox.showwarning("提示", "请输入活动日期。", parent=self)
            return

        self._status.configure(text="生成已启动（查看弹出的 cmd 窗口）")
        run_in_console('generate', date_str, '--excel', path)


# ── 主窗口 ──

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("青志协加分数据库")
        self.geometry("560x260")
        self.resizable(False, False)

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        # ── 菜单栏 ──
        menubar = ctk.CTkFrame(self, height=36, corner_radius=0)
        menubar.pack(fill="x", side="top")

        self._tools_btn = ctk.CTkButton(
            menubar, text="工具 ▾", width=80, height=30,
            command=self._toggle_tools,
        )
        self._tools_btn.pack(side="left", padx=(10, 2), pady=3)

        ctk.CTkButton(
            menubar, text="打开输出文件夹", width=130, height=30,
            fg_color="transparent", border_width=1,
            command=self._open_output_dir,
        ).pack(side="right", padx=(2, 10), pady=3)

        ctk.CTkButton(
            menubar, text="最近文件", width=90, height=30,
            fg_color="transparent", border_width=1,
            command=self._open_recent,
        ).pack(side="right", padx=2, pady=3)

        # ── 生成名单区域 ──
        gen_frame = ctk.CTkFrame(self)
        gen_frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            gen_frame, text="生成名单（Excel → Word）",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(anchor="w", padx=20, pady=(16, 10))

        r1 = ctk.CTkFrame(gen_frame, fg_color="transparent")
        r1.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(r1, text="Excel 文件:").pack(side="left")
        self.excel_path_var = ctk.StringVar()
        ctk.CTkEntry(r1, textvariable=self.excel_path_var, width=300).pack(side="left", padx=(10, 6))
        ctk.CTkButton(
            r1, text="浏览...", width=65,
            command=self._browse_excel,
        ).pack(side="left")

        r2 = ctk.CTkFrame(gen_frame, fg_color="transparent")
        r2.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkLabel(r2, text="活动日期:").pack(side="left")
        self.date_var = ctk.StringVar()
        ctk.CTkEntry(r2, textvariable=self.date_var, width=110).pack(side="left", padx=(10, 0))

        ctk.CTkButton(
            gen_frame, text="生成名单", width=120, height=34,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._on_generate,
        ).pack(pady=(16, 0))

        # 状态栏
        self.status_var = ctk.StringVar(value="就绪")
        ctk.CTkLabel(
            self, textvariable=self.status_var,
            font=ctk.CTkFont(size=11), text_color="gray",
        ).pack(side="bottom", anchor="w", padx=20, pady=(0, 10))

        self._set_defaults()

    # ── 默认值 ──

    def _set_defaults(self):
        updated = os.path.join(OUTPUT_DIR, '年级花名册数据表_updated.xlsx')
        if os.path.exists(updated):
            self.excel_path_var.set(updated)

    # ── 工具菜单 ──

    def _toggle_tools(self):
        x = self.winfo_x() + 10
        y = self.winfo_y() + 42
        DropdownMenu(self, [
            ("导入分数", self._open_import),
            ("生成名单-测试", self._open_generate_test),
        ], x, y)

    def _open_import(self):
        ImportDialog(self)

    def _open_generate_test(self):
        GenerateTestDialog(self)

    # ── 文件浏览 ──

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            initialdir=OUTPUT_DIR,
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.excel_path_var.set(path)

    # ── 生成名单 ──

    def _on_generate(self):
        path = self.excel_path_var.get().strip()
        if not path:
            path = os.path.join(OUTPUT_DIR, '年级花名册数据表_updated.xlsx')
            if os.path.exists(path):
                self.excel_path_var.set(path)
            else:
                messagebox.showwarning("提示", "请先选择 Excel 文件或运行一次导入分数。")
                return
        if not os.path.exists(path):
            messagebox.showerror("错误", f"文件不存在:\n{path}")
            return

        date_str = self.date_var.get().strip()
        if not date_str:
            messagebox.showwarning("提示", "请输入活动日期。")
            return

        self.status_var.set("正在生成名单...")
        run_in_console('generate', date_str, '--excel', path)
        self.status_var.set("生成已启动（查看弹出的 cmd 窗口）")

    # ── 快捷操作 ──

    def _open_output_dir(self):
        if os.path.isdir(OUTPUT_DIR):
            os.startfile(OUTPUT_DIR)

    def _open_recent(self):
        docx_files = []
        if not os.path.isdir(OUTPUT_DIR):
            messagebox.showinfo("提示", "输出文件夹尚不存在。")
            return
        for f in os.listdir(OUTPUT_DIR):
            if f.endswith('.docx'):
                full = os.path.join(OUTPUT_DIR, f)
                docx_files.append((os.path.getmtime(full), full))
        if docx_files:
            docx_files.sort(reverse=True)
            os.startfile(docx_files[0][1])
            self.status_var.set(f"已打开: {os.path.basename(docx_files[0][1])}")
        else:
            messagebox.showinfo("提示", "输出文件夹中暂无 Word 文件。")


def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
