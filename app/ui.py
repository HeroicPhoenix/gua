# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog  # 新增：用于选择数据库文件/Excel
from datetime import datetime

from workers import AutoClickWorker, MonitorClickWorker
from io_parse import insert_blank_cols, COL_ORDER
import sys, os

def resource_path(filename: str) -> str:
    """兼容 PyInstaller --onefile 的资源路径解析"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)


class LiuyaoGUI:
    """
    仅负责 UI：输入参数、日志展示、启动/停止对应 Worker。
    ——默认参数“写死”在 UI 内部——
    """
    DEFAULT_TITLE = "六爻正式"
    DEFAULT_BUTTON_TEXT = "电脑起卦"
    DEFAULT_EXCEL_PATH = "./gua_auto_results.xlsx"
    DEFAULT_INTERVAL_SEC = 5
    DEFAULT_BACKEND = "win32"
    DEFAULT_WAIT_TIMEOUT = 5.0
    DEFAULT_WAIT_POLL = 0.15

    def __init__(self):
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("LiuyaoReader")
        except Exception:
            pass
        self.root = tk.Tk()
        self.root.title("自动读取工具")
        self.root.geometry("720x640")  # 高度略增以容纳新增区块

        frm = tk.Frame(self.root)
        frm.pack(pady=10, padx=10, fill="x")

        def add_entry(label, var, default):
            row = tk.Frame(frm)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label, width=15, anchor="e").pack(side="left")
            entry = tk.Entry(row, textvariable=var)
            entry.pack(side="left", fill="x", expand=True, padx=6)
            var.set(default)
            return row

        # ======= 原有参数 =======
        self.title_var = tk.StringVar()
        add_entry("窗口标题：", self.title_var, self.DEFAULT_TITLE)

        self.button_var = tk.StringVar()
        add_entry("按钮文字：", self.button_var, self.DEFAULT_BUTTON_TEXT)

        # Excel 路径 + 选择按钮（新增）
        self.excel_var = tk.StringVar(value=self.DEFAULT_EXCEL_PATH)
        row_exl = tk.Frame(frm)
        row_exl.pack(fill="x", pady=3)
        tk.Label(row_exl, text="Excel 路径：", width=15, anchor="e").pack(side="left")
        tk.Entry(row_exl, textvariable=self.excel_var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row_exl, text="选择…", width=8, command=self._choose_excel).pack(side="left", padx=6)

        # === 新增：数据库文件路径（可选择） ===
        self.db_var = tk.StringVar()
        row_db = tk.Frame(frm)
        row_db.pack(fill="x", pady=3)
        tk.Label(row_db, text="数据库文件：", width=15, anchor="e").pack(side="left")
        tk.Entry(row_db, textvariable=self.db_var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(row_db, text="选择…", width=8, command=self._choose_db).pack(side="left", padx=6)
        self.db_var.set("./gua.gdbx")  # 默认值，可自行修改

        self.mode_var = tk.StringVar(value="manual")
        tk.Radiobutton(frm, text="监测点击模式（点击或回车触发）",
                       variable=self.mode_var, value="manual",
                       command=self._toggle_interval).pack(anchor="w", padx=20)
        tk.Radiobutton(frm, text="自动点击模式",
                       variable=self.mode_var, value="auto",
                       command=self._toggle_interval).pack(anchor="w", padx=20)

        self.interval_var = tk.StringVar(value=str(self.DEFAULT_INTERVAL_SEC))
        self.interval_row = add_entry("间隔秒数：", self.interval_var, str(self.DEFAULT_INTERVAL_SEC))

        # ======= 新增：批量插入空列配置区 =======
        sep = tk.Label(self.root, text="—— 批量在 Excel 中插入空列 ——", fg="#555")
        sep.pack(pady=(8, 4))

        cfg = tk.Frame(self.root)
        cfg.pack(padx=10, fill="x")

        # X：插入位置（第 X 列之前，1 起）
        self.x_var = tk.StringVar(value="1")
        row_x = tk.Frame(cfg)
        row_x.pack(fill="x", pady=3)
        tk.Label(row_x, text="插入位置 X（1 起）：", width=18, anchor="e").pack(side="left")
        tk.Entry(row_x, textvariable=self.x_var, width=10).pack(side="left", padx=6)

        # N：插入列数
        self.n_var = tk.StringVar(value="0")
        row_n = tk.Frame(cfg)
        row_n.pack(fill="x", pady=3)
        tk.Label(row_n, text="插入列数 N：", width=18, anchor="e").pack(side="left")
        ent_n = tk.Entry(row_n, textvariable=self.n_var, width=10)
        ent_n.pack(side="left", padx=6)

        # 列名输入区（动态）
        self.names_frame = tk.LabelFrame(cfg, text="列名（依次插入在第 X 列之前）")
        self.names_frame.pack(fill="x", pady=6)
        self.name_vars = []  # 存储 N 个 StringVar

        # 监听 N 变化，动态重建列名输入框
        self.n_var.trace_add("write", lambda *_: self._rebuild_name_inputs())

        # 执行按钮放在一排
        btn_row = tk.Frame(cfg)
        btn_row.pack(fill="x", pady=4)
        tk.Button(btn_row, text="批量增加列", width=15, command=self.add_blank_cols).pack(side="left", padx=6)

        # ======= 原有开始/停止与日志 =======
        btnfrm = tk.Frame(self.root)
        btnfrm.pack(pady=10)
        tk.Button(btnfrm, text="▶ 开始读取", width=15, command=self.start).pack(side="left", padx=10)
        tk.Button(btnfrm, text="■ 停止读取", width=15, command=self.stop).pack(side="left", padx=10)
        tk.Button(btnfrm, text="打印设置…", width=12, command=self._open_print_config).pack(side="left", padx=10)

        # 文本 + 垂直滚动条（支持拖动）
        log_frame = tk.Frame(self.root)
        log_frame.pack(padx=10, pady=10, fill="both", expand=True)
        ybar = tk.Scrollbar(log_frame)
        ybar.pack(side="right", fill="y")
        self.log_text = tk.Text(log_frame, height=16, wrap="word", yscrollcommand=ybar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        ybar.config(command=self.log_text.yview)

        self.thread = None
        self._toggle_interval()
        self._rebuild_name_inputs()  # 根据初始 N=0 初始化一次

        # ======= 新增：运行状态提示/图标 =======
        self._overlay = None          # 运行中大提示窗口
        self._icon_state = "idle"     # 'idle' / 'running'
        self._set_app_icon("idle")
        self.print_fields = ["卦象名字"]

    # ===== UI 工具 =====
    def _choose_excel(self):
        path = filedialog.asksaveasfilename(
            title="选择或新建 Excel 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.excel_var.set(path)
            self.log(f"已选择 Excel 文件：{path}")

    def _choose_db(self):
        path = filedialog.askopenfilename(
            title="选择数据库文件",
            filetypes=[("数据库/伪装文件", "*.gdbx *.db *.sqlite *.data *.*"), ("所有文件", "*.*")]
        )
        if path:
            self.db_var.set(path)
            self.log(f"已选择数据库文件：{path}")

    def _open_print_config(self):
        """多选配置：日志打印哪些字段。"""
        top = tk.Toplevel(self.root)
        top.title("选择要打印的字段")
        top.grab_set()  # 模态
        top.geometry("420x520")

        # 说明
        tk.Label(top, text="勾选要打印到日志的字段（按列名）：").pack(anchor="w", padx=10, pady=(10, 4))

        # 滚动区
        frame = tk.Frame(top)
        frame.pack(fill="both", expand=True, padx=10, pady=6)

        canvas = tk.Canvas(frame)
        ybar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        wrap = tk.Frame(canvas)

        wrap.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=wrap, anchor="nw")
        canvas.configure(yscrollcommand=ybar.set)

        canvas.pack(side="left", fill="both", expand=True)
        ybar.pack(side="right", fill="y")

        # 让滚轮也能滚动（同时保留拖动滚动条）
        def _on_mousewheel_win(event):
            # Windows / macOS (tk 8.6+) 通常有 event.delta，单位 120
            delta = int(-1 * (event.delta / 120)) if event.delta else -1
            canvas.yview_scroll(delta, "units")

        def _on_mousewheel_linux_up(event):
            canvas.yview_scroll(-3, "units")

        def _on_mousewheel_linux_down(event):
            canvas.yview_scroll(3, "units")

        # 绑定到顶层/画布/内部容器，确保鼠标在复选框上滚轮也生效
        for w in (top, canvas, wrap):
            w.bind("<MouseWheel>", _on_mousewheel_win, add="+")  # Windows / macOS
            w.bind("<Button-4>", _on_mousewheel_linux_up, add="+")  # Linux 上滚
            w.bind("<Button-5>", _on_mousewheel_linux_down, add="+")  # Linux 下滚

        # 复选框变量
        vars_map = {}
        # 按 COL_ORDER 原顺序显示字段
        ordered = [c for c in COL_ORDER if isinstance(c, str) and c.strip()]

        for name in ordered:
            v = tk.BooleanVar(value=(name in self.print_fields))
            cb = tk.Checkbutton(wrap, text=name, variable=v, anchor="w")
            cb.pack(fill="x", pady=2, anchor="w")
            vars_map[name] = v

        # 全选/全不选/默认按钮
        btnbar = tk.Frame(top)
        btnbar.pack(fill="x", padx=10, pady=6)

        def _set_all(flag: bool):
            for v in vars_map.values():
                v.set(flag)

        def _set_default():
            # 默认只显示“卦象名字”
            for k, v in vars_map.items():
                v.set(k == "卦象名字")

        tk.Button(btnbar, text="全选", width=8, command=lambda: _set_all(True)).pack(side="left", padx=4)
        tk.Button(btnbar, text="全不选", width=8, command=lambda: _set_all(False)).pack(side="left", padx=4)
        tk.Button(btnbar, text="恢复默认", width=10, command=_set_default).pack(side="left", padx=4)

        # 底部：确定/取消
        bot = tk.Frame(top)
        bot.pack(fill="x", padx=10, pady=(6, 10))

        def _ok():
            selected = [k for k, v in vars_map.items() if v.get()]
            # 允许为空：表示不打印任何字段
            self.print_fields = selected
            if selected:
                self.log(f"已更新打印字段：{', '.join(self.print_fields)}")
            else:
                self.log("已更新打印字段：<无>（将仅显示“记录完成”提示）")
            top.destroy()

        tk.Button(bot, text="取消", width=10, command=top.destroy).pack(side="right")
        tk.Button(bot, text="确定", width=10, command=_ok).pack(side="right", padx=6)


    def _toggle_interval(self):
        if self.mode_var.get() == "auto":
            self.interval_row.pack(fill="x", pady=3)
        else:
            self.interval_row.pack_forget()

    def log(self, text):
        self.log_text.insert(tk.END, f"{datetime.now():%H:%M:%S}  {text}\n")
        self.log_text.see(tk.END)

    def alert_error(self, msg: str):
        messagebox.showerror("错误", msg)

    def _rebuild_name_inputs(self):
        """根据 N 的值动态重建列名输入框。"""
        # 清空旧控件
        for child in self.names_frame.winfo_children():
            child.destroy()
        self.name_vars.clear()

        # 解析 N
        try:
            n = int(self.n_var.get().strip())
            if n < 0:
                n = 0
        except Exception:
            n = 0

        # 生成 n 个输入框
        for i in range(n):
            v = tk.StringVar()
            row = tk.Frame(self.names_frame)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=f"列名 {i+1}：", width=10, anchor="e").pack(side="left")
            tk.Entry(row, textvariable=v).pack(side="left", fill="x", expand=True, padx=6)
            self.name_vars.append(v)

        if n == 0:
            tk.Label(self.names_frame, text="将根据 N 自动生成列名输入框").pack(anchor="w", padx=6, pady=2)

    # ===== 运行状态提示/图标 =====
    def _set_app_icon(self, mode: str):
        """根据模式切换窗口/任务栏图标：mode ∈ {'idle', 'running'}"""
        try:
            ico_name = "app_running.ico" if mode == "running" else "app_idle.ico"
            ico_path = resource_path(ico_name)  # ← 关键：兼容 PyInstaller
            if os.path.exists(ico_path):
                # 窗口与任务栏（大部分场景）都会用到这个 icon
                self.root.iconbitmap(ico_path)
            self._icon_state = mode
        except Exception:
            pass

    def _show_overlay(self, text="正在运行…（按 Esc 关闭提示）"):
        """显示一个置顶、半透明、无边框的大提示层。"""
        if self._overlay and str(self._overlay.winfo_exists()) == "1":
            for w in self._overlay.winfo_children():
                if isinstance(w, tk.Label):
                    w.config(text=text)
            self._overlay.lift()
            self._overlay.attributes("-topmost", True)
            return

        ol = tk.Toplevel(self.root)
        ol.withdraw()
        ol.overrideredirect(True)           # 无边框
        ol.attributes("-topmost", True)     # 置顶
        try:
            ol.attributes("-alpha", 0.88)   # 半透明
        except Exception:
            pass

        frame = tk.Frame(ol, bg="#111111")
        frame.pack(fill="both", expand=True)

        lbl = tk.Label(
            frame,
            text=text,
            fg="white",
            bg="#111111",
            font=("Microsoft YaHei UI", 28, "bold"),
            padx=40, pady=30,
            justify="center"
        )
        lbl.pack()

        # 呼吸动画（轻微透明度变化）
        def _pulse(op=88, step=-3):
            try:
                new_op = max(70, min(95, op + step))
                if new_op in (70, 95):
                    step = -step
                ol.attributes("-alpha", new_op / 100)
                ol.after(60, _pulse, new_op, step)
            except Exception:
                pass

        try:
            _pulse()
        except Exception:
            pass

        # Esc 关闭提示层（不影响任务）
        def _on_key(evt):
            if evt.keysym.lower() == "escape":
                try:
                    ol.destroy()
                except Exception:
                    pass
        ol.bind("<Key>", _on_key)

        # 点击任意位置隐藏
        # def _on_click(_):
        #     try:
        #         ol.destroy()
        #     except Exception:
        #         pass
        # ol.bind("<Button-1>", _on_click)

        # 居屏显示
        ol.update_idletasks()
        sw = ol.winfo_screenwidth()
        sh = ol.winfo_screenheight()
        w, h = int(sw * 0.6), 160
        x = (sw - w) // 2
        y = int(sh * 0.12)
        ol.geometry(f"{w}x{h}+{x}+{y}")
        ol.deiconify()
        self._overlay = ol

    def _hide_overlay(self):
        if self._overlay and str(self._overlay.winfo_exists()) == "1":
            try:
                self._overlay.destroy()
            except Exception:
                pass
        self._overlay = None

    # ===== 任务控制 =====
    def start(self):
        if self.thread and self.thread.is_alive():
            messagebox.showinfo("提示", "已经在运行中。")
            return

        # 自动模式校验间隔
        if self.mode_var.get() == "auto":
            try:
                interval = int(self.interval_var.get())
                if interval < 3:
                    messagebox.showwarning("警告", "间隔秒数不能小于3。")
                    return
            except ValueError:
                messagebox.showerror("错误", "间隔秒数必须为整数。")
                return

        # 创建对应 worker（与原实现一致:contentReference[oaicite:2]{index=2}）
        if self.mode_var.get() == "auto":
            self.thread = AutoClickWorker(
                gui=self,
                backend=self.DEFAULT_BACKEND,
                wait_timeout=self.DEFAULT_WAIT_TIMEOUT,
                wait_poll=self.DEFAULT_WAIT_POLL,
                interval_sec=int(self.interval_var.get())
            )
        else:
            self.thread = MonitorClickWorker(
                gui=self,
                backend=self.DEFAULT_BACKEND,
                wait_timeout=self.DEFAULT_WAIT_TIMEOUT,
                wait_poll=self.DEFAULT_WAIT_POLL
            )

        self.thread.start()
        self.log(f"开始运行（模式：{'自动点击' if self.mode_var.get() == 'auto' else '监测点击'}）")

        # —— 新增：运行中图标与大提示层 ——
        self._set_app_icon("running")
        self._show_overlay("正在自动读取…（按 Esc 关闭提示）")

    def stop(self):
        if self.thread and self.thread.is_alive():
            self.thread.stop()
            self.log("已请求停止...")
        else:
            self.log("当前无运行任务。")

        # —— 新增：恢复图标与关闭提示层 ——
        self._set_app_icon("idle")
        self._hide_overlay()

    def add_blank_cols(self):
        """点击【批量增加列】后触发：按 UI 配置在 Excel 中插入空列。"""
        path = self.excel_var.get().strip()
        try:
            x = int(self.x_var.get().strip())
        except Exception:
            self.alert_error("X 必须是正整数（1 起）。")
            return

        # 收集列名
        names = [v.get().strip() for v in self.name_vars]
        if any(n == "" for n in names):
            self.alert_error("列名不能为空，请补全。")
            return

        try:
            insert_blank_cols(path, x, names)
            self.log(f"已在 {path} 的第 {x} 列前插入 {len(names)} 列：{names}")
            messagebox.showinfo("成功", "已完成插入。若 Excel 正在打开，请先关闭文件后再执行。")
        except Exception as e:
            self.alert_error(f"插入失败：{e}")

    def run(self):
        self.root.mainloop()
