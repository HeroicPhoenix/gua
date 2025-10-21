# -*- coding: utf-8 -*-
"""
excel_to_sqlite_gui.py

功能：
- 从 Excel 导入数据（全部按字符串读取，不产生 1.0 / 2.0）；
- 选择主列与参数列（可调整顺序）；
- 每次导入前清空表；
- 主列不允许重复，如重复则仅保留最后一条；
- 写入 SQLite 数据库（伪装后缀 .gdbx）；
- 表名 gui_para(name TEXT, param_order INTEGER, param_value TEXT)。
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import sqlite3
import os
import time

DEFAULT_DB_EXT = ".gdbx"  # 可改为伪装后缀


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel导入工具")
        self.root.geometry("860x700")

        # --- Excel 文件选择 ---
        f1 = tk.Frame(self.root)
        f1.pack(fill="x", padx=10, pady=(10, 4))
        tk.Label(f1, text="选择 Excel 文件：").pack(side="left")
        self.excel_var = tk.StringVar()
        tk.Entry(f1, textvariable=self.excel_var, width=60).pack(side="left", padx=6)
        tk.Button(f1, text="浏览", command=self.load_excel).pack(side="left")

        # --- 数据库文件选择 ---
        f2 = tk.Frame(self.root)
        f2.pack(fill="x", padx=10, pady=(0, 6))
        tk.Label(f2, text="选择/新建 数据库文件：").pack(side="left")
        self.db_var = tk.StringVar(value="gua" + DEFAULT_DB_EXT)
        tk.Entry(f2, textvariable=self.db_var, width=60).pack(side="left", padx=6)
        tk.Button(f2, text="选择/新建", command=self.select_db).pack(side="left")

        # --- 列选择区 ---
        frm = tk.Frame(self.root)
        frm.pack(fill="both", expand=True, padx=10, pady=8)

        # 主列
        tk.Label(frm, text="主列（唯一标识）").grid(row=0, column=0, sticky="w")
        self.main_list = tk.Listbox(frm, exportselection=False, height=12)
        self.main_list.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)

        # 参数列
        tk.Label(frm, text="参数列（可多选）").grid(row=0, column=1, sticky="w")
        self.param_list = tk.Listbox(frm, selectmode="multiple", height=12)
        self.param_list.grid(row=1, column=1, sticky="nsew", padx=6, pady=4)

        # 已选择参数顺序
        tk.Label(frm, text="已选参数（顺序可上下调整）").grid(row=0, column=2, sticky="w")
        self.chosen_list = tk.Listbox(frm, exportselection=False, height=12)
        self.chosen_list.grid(row=1, column=2, sticky="nsew", padx=6, pady=4)

        # 控制按钮列
        btns = tk.Frame(frm)
        btns.grid(row=1, column=3, sticky="ns", padx=6)
        tk.Button(btns, text="添加 →", command=self.add_selected).pack(pady=6)
        tk.Button(btns, text="← 移除", command=self.remove_selected).pack(pady=6)
        tk.Button(btns, text="↑ 上移", command=self.move_up).pack(pady=6)
        tk.Button(btns, text="↓ 下移", command=self.move_down).pack(pady=6)

        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(2, weight=1)

        # 导入按钮
        tk.Button(self.root, text="导入到数据库", bg="#0078D7", fg="white",
                  font=("Arial", 12), command=self.import_to_db).pack(pady=10)

        # 日志区（带滚动条）
        log_frame = tk.Frame(self.root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        ybar = tk.Scrollbar(log_frame)
        ybar.pack(side="right", fill="y")
        self.log_text = tk.Text(log_frame, height=10, wrap="none", yscrollcommand=ybar.set)
        self.log_text.pack(fill="both", expand=True)
        ybar.config(command=self.log_text.yview)

        self.df = None

    # ---------- UI 辅助 ----------
    def log(self, txt: str):
        t = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"{t}  {txt}\n")
        self.log_text.see(tk.END)

    def load_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if not p:
            return
        try:
            # 全部按字符串读取，缺失值先转成空串，后续统一判断
            df = pd.read_excel(p, dtype=str).fillna("")
            self.df = df
            self.excel_var.set(p)
            # 更新列列表
            self.main_list.delete(0, tk.END)
            self.param_list.delete(0, tk.END)
            self.chosen_list.delete(0, tk.END)
            for col in df.columns:
                self.main_list.insert(tk.END, col)
                self.param_list.insert(tk.END, col)
            self.log(f"已加载 Excel：{os.path.basename(p)}，{len(df)} 行，{len(df.columns)} 列（全部按字符串读取）。")
        except Exception as e:
            messagebox.showerror("错误", f"加载 Excel 失败：{e}")

    def select_db(self):
        p = filedialog.asksaveasfilename(defaultextension=DEFAULT_DB_EXT,
                                         filetypes=[("数据库文件", f"*{DEFAULT_DB_EXT}"), ("所有文件", "*.*")])
        if p:
            self.db_var.set(p)
            self.log(f"已选择数据库文件：{p}")

    def add_selected(self):
        sel = self.param_list.curselection()
        for i in sel:
            val = self.param_list.get(i)
            if val not in self.chosen_list.get(0, tk.END):
                self.chosen_list.insert(tk.END, val)

    def remove_selected(self):
        sel = list(self.chosen_list.curselection())
        for i in reversed(sel):
            self.chosen_list.delete(i)

    def move_up(self):
        sel = self.chosen_list.curselection()
        if not sel:
            return
        i = sel[0]
        if i == 0:
            return
        val = self.chosen_list.get(i)
        self.chosen_list.delete(i)
        self.chosen_list.insert(i - 1, val)
        self.chosen_list.selection_set(i - 1)

    def move_down(self):
        sel = self.chosen_list.curselection()
        if not sel:
            return
        i = sel[0]
        if i >= self.chosen_list.size() - 1:
            return
        val = self.chosen_list.get(i)
        self.chosen_list.delete(i)
        self.chosen_list.insert(i + 1, val)
        self.chosen_list.selection_set(i + 1)

    # ---------- 导入逻辑 ----------
    def import_to_db(self):
        if self.df is None:
            messagebox.showwarning("提示", "请先加载 Excel 文件。")
            return

        main_sel = self.main_list.curselection()
        if not main_sel:
            messagebox.showwarning("提示", "请选择主列。")
            return
        main_col = self.main_list.get(main_sel[0])

        param_cols = list(self.chosen_list.get(0, tk.END))
        if not param_cols:
            messagebox.showwarning("提示", "请选择至少一个参数列。")
            return

        db_path = self.db_var.get().strip()
        if not db_path:
            messagebox.showwarning("提示", "请选择数据库文件路径。")
            return

        try:
            os.makedirs(os.path.dirname(os.path.abspath(db_path)) or ".", exist_ok=True)
            conn = sqlite3.connect(db_path)
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS gui_para (
                    name TEXT,
                    param_order INTEGER,
                    param_value TEXT
                )
            """)
            # 全量导入前清空
            cur.execute("DELETE FROM gui_para")
            conn.commit()

            rows_written = 0
            null_count = 0

            # 小工具：把单元格值转换为将要入库的值（空 -> None）
            def cell_to_db_value(v):
                # df 里已是字符串或空串，这里统一裁剪
                if v is None:
                    return None
                s = str(v).strip()
                if s == "":
                    return None  # -> SQLite 的 NULL
                return s

            for _, row in self.df.iterrows():
                name_val = cell_to_db_value(row.get(main_col, ""))
                if name_val in (None, ""):
                    # 主键标识为空，无法建唯一占位，跳过整行
                    continue

                # 先删除旧记录再插入（保证主列唯一、只保留最后一次）
                cur.execute("DELETE FROM gui_para WHERE name=?", (name_val,))

                for idx, pcol in enumerate(param_cols, start=1):
                    db_val = cell_to_db_value(row.get(pcol, ""))
                    if db_val is None:
                        null_count += 1
                    # 关键：即使为 None（NULL），也插入一条记录实现“占位”
                    cur.execute(
                        "INSERT INTO gui_para (name, param_order, param_value) VALUES (?, ?, ?)",
                        (name_val, idx, db_val)
                    )
                    rows_written += 1

            conn.commit()
            conn.close()
            self.log(f"导入完成：共写入 {rows_written} 条记录（包含 NULL 占位 {null_count} 条）到 {db_path}")
            messagebox.showinfo(
                "完成",
                f"导入完成：共写入 {rows_written} 条（展开后）。\n其中空值占位（NULL）{null_count} 条。\n文件：{db_path}"
            )
        except Exception as e:
            messagebox.showerror("错误", f"导入失败：{e}")
            self.log(f"导入失败：{e}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
