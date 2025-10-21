# -*- coding: utf-8 -*-
import re
import time
from pywinauto import Desktop, Application

def connect_main(title_pattern: str, backend: str = "win32"):
    desk = Desktop(backend=backend)
    pattern = re.escape(title_pattern)
    for w in desk.windows():
        if re.search(pattern, w.window_text(), re.IGNORECASE):
            matched_title = w.window_text()
            app = Application(backend=backend).connect(handle=w.handle)
            main = app.window(handle=w.handle)
            main.wait("visible", timeout=5)
            return main, matched_title
    raise RuntimeError(f"未找到匹配窗口（标题包含：{title_pattern}）")

def find_controls(main, button_text: str):
    """
    返回 (btn, result_edit, intro_static)
    - btn：触发“电脑起卦”按钮
    - result_edit：中部大 Edit（包含“公历/农历/干支/旬空 …”）
    - intro_static：右上角 Static（包含“观之…；月卦身…；世身在…；…\\n神煞：…”），找不到时为 None
    """
    btn = main.child_window(title=button_text, class_name="Button")
    if not btn.exists():
        btn = main.child_window(title=button_text)

    result_edit = None
    intro_static = None

    # 找 Edit
    cand_edit = []
    for c in main.descendants():
        try:
            if c.friendly_class_name() == "Edit":
                t = c.window_text().strip()
                if not t or "点击盘中元素显示相应提示" in t:
                    continue
                # 优先以“公历”开头且文本更长的
                cand_edit.append((len(t), t.startswith("公历"), c))
        except Exception:
            pass
    if cand_edit:
        cand_edit.sort(key=lambda x: (not x[1], -x[0]))
        result_edit = cand_edit[0][2]
    if not result_edit:
        raise RuntimeError("未找到结果 Edit 控件。")

    # 找“简介 Static”（包含神煞/分号；更稳妥的方法是按长度和是否含“神煞”）
    best_len = -1
    for c in main.descendants():
        try:
            if c.friendly_class_name() == "Static":
                t = c.window_text().strip()
                if not t:
                    continue
                if ("神煞" in t) or ("；" in t and "\n" in t):
                    if len(t) > best_len:
                        best_len = len(t)
                        intro_static = c
        except Exception:
            pass

    return btn, result_edit, intro_static

def wait_text_change(edit, old_text: str, timeout: float = 5.0, poll: float = 0.15) -> str:
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            t = edit.window_text()
        except Exception:
            t = ""
        if t and t != old_text:
            return t
        time.sleep(poll)
    try:
        return edit.window_text()
    except Exception:
        return ""
