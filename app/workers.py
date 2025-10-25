# -*- coding: utf-8 -*-
import threading
import time
from datetime import datetime
import os
import sqlite3

import win32gui
import win32con
import win32api

from io_parse import build_excel_row, save_row_to_excel
from winops import connect_main, find_controls, wait_text_change


class BaseWorker(threading.Thread):
    _RECORD_COOLDOWN_SEC = 0.6
    READS_PER_CLICK = 3        # 每次触发最多尝试读取次数
    READ_DELAY_SEC = 0.2       # 触发（点击/回车）后延迟再读，避免半成品文本
    SETTLE_POLLS = 6           # 稳定轮询次数上限
    SETTLE_GAP_SEC = 0.08      # 稳定轮询间隔

    def __init__(self, gui, backend: str = "win32", wait_timeout: float = 5.0, wait_poll: float = 0.15):
        super().__init__(daemon=True)
        self.gui = gui
        self.backend = backend
        self.wait_timeout = wait_timeout
        self.wait_poll = wait_poll

        self.stop_flag = False
        self.last_text = ""
        self.cooldown_until = 0
        self.main = None
        self.btn = None
        self.result_edit = None
        self.intro_static = None  # 右上角简介 Static
        self._last_shown_gua = None

        # 数据库状态
        self._db_ok = False
        self._db_path = ""

    def stop(self):
        self.stop_flag = True

    def _prepare(self):
        # 连接窗口与控件
        main, matched_title = connect_main(self.gui.title_var.get(), backend=self.backend)
        self.gui.log(f"匹配到窗口：{matched_title}")
        btn, result_edit, intro_static = find_controls(main, self.gui.button_var.get())

        self.main = main
        self.btn = btn
        self.result_edit = result_edit
        self.intro_static = intro_static
        try:
            self.last_text = result_edit.window_text()
        except Exception:
            self.last_text = ""

        # ==== 数据库连通性检查（启动时一次性提示） ====
        self._db_ok = False
        self._db_path = ""
        try:
            if hasattr(self.gui, "db_var"):
                self._db_path = (self.gui.db_var.get() or "").strip()
            if not self._db_path:
                self.gui.log("未配置数据库文件（可在“数据库文件”栏选择），将不追加参数列。")
            elif not os.path.exists(self._db_path):
                self.gui.log(f"数据库文件不存在：{self._db_path}，将不追加参数列。")
            else:
                try:
                    conn = sqlite3.connect(self._db_path)
                    cur = conn.cursor()
                    cur.execute("SELECT COUNT(1) FROM sqlite_master WHERE type='table' AND name='gui_para'")
                    has_table = (cur.fetchone() or [0])[0] == 1
                    if not has_table:
                        self.gui.log("数据库连接成功，但未找到表 gui_para，将不追加参数列。")
                    else:
                        cur.execute("SELECT COUNT(*) FROM gui_para")
                        cnt = (cur.fetchone() or [0])[0]
                        self.gui.log(f"数据库连接成功：{os.path.basename(self._db_path)}（gui_para 共 {cnt} 条）")
                        self._db_ok = True
                    conn.close()
                except Exception as db_e:
                    self.gui.log(f"数据库连接失败：{db_e}（将不追加参数列）")
        except Exception as e:
            self.gui.log(f"数据库检查异常：{e}（将不追加参数列）")
        # ==== 检查结束 ====

    # ---- 稳定性辅助 ----
    @staticmethod
    def _looks_complete(t: str) -> bool:
        """判断文本是否“看起来完整”：≥5行非空，且包含关键字段至少3个。"""
        if not t:
            return False
        lines = [ln for ln in t.replace("\r\n", "\n").split("\n") if ln.strip()]
        if len(lines) < 5:
            return False
        keys = ("公历", "农历", "干支", "旬空")
        hit = sum(1 for k in keys if k in t)
        return hit >= 3

    def _read_stable_text(self) -> str:
        """
        先等待 READ_DELAY_SEC，再多次轮询取文本，直到：
        - 看起来“完整”；且
        - 两次读取长度不下降（基本稳定）
        """
        # 触发后延迟，避免拿到半成品
        if self.READ_DELAY_SEC > 0:
            time.sleep(self.READ_DELAY_SEC)

        best = ""
        prev_len = -1
        for _ in range(self.SETTLE_POLLS):
            try:
                cur = self.result_edit.window_text()
            except Exception:
                cur = ""
            if cur and len(cur) >= prev_len:
                best = cur
                prev_len = len(cur)
            if self._looks_complete(cur):
                return cur
            time.sleep(self.SETTLE_GAP_SEC)
        # 兜底：返回最佳一次（可能不完整，调用方会再判断/跳过）
        return best

    # ---- 录入一次 ----
    def _record_once(self):
        now = time.time()
        if now < self.cooldown_until:
            return
        self.cooldown_until = now + self._RECORD_COOLDOWN_SEC

        handled = False  # 本次触发是否已成功写入

        for _ in range(int(getattr(self, "READS_PER_CLICK", 3))):
            # 仅等待“非空文本”，随后走稳定读取
            _ = wait_text_change(
                self.result_edit,
                "",
                timeout=min(self.wait_timeout, 0.6),
                poll=self.wait_poll
            )

            # 稳定读取（防止第一次半成品）
            new_text = self._read_stable_text()
            if not new_text or not new_text.strip():
                time.sleep(getattr(self, "SETTLE_GAP_SEC", 0.08))
                continue
            if not self._looks_complete(new_text):
                # 文本还不完整，不抛错不提示，静默再试
                time.sleep(getattr(self, "SETTLE_GAP_SEC", 0.08))
                continue

            if not handled:
                # 抓右上角简介
                intro_text = ""
                try:
                    if self.intro_static:
                        intro_text = self.intro_static.window_text()
                except Exception:
                    intro_text = ""

                # 解析
                try:
                    row = build_excel_row(new_text, intro_text, write_dt=datetime.now())
                except Exception:
                    # 出现半成品解析错误时不要打扰用户；静默跳过本次
                    time.sleep(getattr(self, "SETTLE_GAP_SEC", 0.08))
                    continue

                # === 查询数据库，拿到“参数1..参数N” ===
                # === 查询数据库，拿到“参数1..参数N” ===
                extra_params = []
                if self._db_ok and self._db_path:
                    try:
                        name_exact = row.get("卦象名字", "") or ""
                        # 备用：用“本卦简称之变卦简称”
                        name_fallback = ""
                        if row.get("本卦简称") or row.get("变卦简称"):
                            name_fallback = f"{row.get('本卦简称', '')}之{row.get('变卦简称', '')}"

                        def _normalize(s: str) -> str:
                            import re
                            s = (s or "").strip()
                            s = re.sub(r"\s+", "", s)
                            s = s.replace("；", ";").replace("，", ",").replace("：", ":")
                            return s

                        conn = sqlite3.connect(self._db_path)
                        cur = conn.cursor()

                        # 1) 精确匹配（去空白）
                        cur.execute(
                            "SELECT param_value FROM gui_para WHERE REPLACE(REPLACE(name,' ',''),'　','') = ? ORDER BY param_order ASC",
                            (_normalize(name_exact),)
                        )
                        fetched = cur.fetchall()
                        rows = []
                        for r in fetched:
                            val = r[0]
                            if val is None:
                                rows.append("")  # ← 将数据库中的 NULL 转为空单元格
                            else:
                                val = str(val).strip()
                                rows.append(val)

                        # 2) 备用匹配（当精确匹配无结果）
                        if not rows and name_fallback.strip():
                            cur.execute(
                                "SELECT param_value FROM gui_para WHERE REPLACE(REPLACE(name,' ',''),'　','') = ? ORDER BY param_order ASC",
                                (_normalize(name_fallback),)
                            )
                            fetched = cur.fetchall()
                            rows = []
                            for r in fetched:
                                val = r[0]
                                if val is None:
                                    rows.append("")
                                else:
                                    val = str(val).strip()
                                    rows.append(val)

                        conn.close()
                        extra_params = rows
                        # self.gui.log(f"参数命中 {len(extra_params)} 项（{name_exact or name_fallback}）")
                    except Exception as e:
                        self.gui.log(f"查询数据库失败：{e}（已忽略，不影响记录）")

                # === 写入 Excel（把参数列带上）===
                try:
                    result = save_row_to_excel(row, self.gui.excel_var.get(), extra_params=extra_params)
                except TypeError:
                    # 兼容旧版 save_row_to_excel(没有 extra_params)
                    result = save_row_to_excel(row, self.gui.excel_var.get())

                row_snapshot = getattr(result, "snapshot", None)
                written = bool(result)

                if written:
                    # 组装打印文本：按界面配置的字段顺序打印
                    fields = getattr(self.gui, "print_fields", None)
                    if fields is None:
                        fields = ["卦象名字"]
                    parts = []
                    for f in fields:
                        if row_snapshot is not None and f in row_snapshot:
                            val = row_snapshot.get(f, "")
                        else:
                            val = row.get(f, "")
                        s = "" if val is None else str(val).strip()
                        parts.append(f"{f}：{s}")

                    if parts:
                        line = " | ".join(parts)
                        msg = f"记录完成 ✅ {line}"
                    else:
                        msg = "记录完成 ✅"

                    # 防刷屏：若包含“卦象名字”，用其与上次比较；否则用整行比较
                    key = row.get("卦象名字", None)
                    comp = key if (key is not None and str(key).strip() != "") else msg

                    if comp != self._last_shown_gua:
                        self.gui.log(msg)
                        self._last_shown_gua = comp

                self.last_text = new_text
                handled = True

            # 已经写入过了，不重复插入/提示
            break


class AutoClickWorker(BaseWorker):
    def __init__(self, gui, backend="win32", wait_timeout=5.0, wait_poll=0.15, interval_sec=5):
        super().__init__(gui, backend, wait_timeout, wait_poll)
        self.interval_sec = interval_sec

    def run(self):
        try:
            self._prepare()
        except Exception as e:
            self.gui.alert_error(f"无法连接窗口: {e}")
            return

        self.gui.log("进入自动点击模式…")

        while not self.stop_flag:
            start = time.time()
            try:
                self.btn.wait("enabled", timeout=5)
                self.btn.click_input()
            except Exception as e:
                self.gui.log(f"点击失败：{e}")

            self._record_once()

            elapsed = time.time() - start
            sleep_left = max(0, self.interval_sec - elapsed)
            for _ in range(int(sleep_left * 10)):
                if self.stop_flag:
                    break
                time.sleep(0.1)


class MonitorClickWorker(BaseWorker):
    def run(self):
        try:
            self._prepare()
        except Exception as e:
            self.gui.alert_error(f"无法连接窗口: {e}")
            return

        main_handle = self.main.handle
        self.gui.log("进入监测点击模式（检测按钮焦点/回车键）…")

        while not self.stop_flag:
            try:
                fg = win32gui.GetForegroundWindow()
                if fg == main_handle:
                    focus = self.main.get_focus()
                    # 鼠标点在“电脑起卦”按钮上（按钮获得焦点）
                    if focus and self.btn and focus.handle == self.btn.handle:
                        self._record_once()

                    # 监听回车键（静默）
                    if win32api.GetAsyncKeyState(win32con.VK_RETURN) & 0x8000:
                        self._record_once()
            except Exception:
                # 任何 UI 小抖动都吞掉，避免刷异常
                pass
            time.sleep(0.05)
