# probe_win32.py
from pywinauto import Desktop, Application

TITLE = "六爻正式PC版.*"  # 你的完整标题
BACKEND = "win32"

desk = Desktop(backend=BACKEND)
win = desk.window(title_re=TITLE)      # 若标题偶有变化可改成 title_re=".*六爻正式PC版3.1b.*"
win.wait("visible", timeout=5)

# 连接 exe 以便进一步操作
app = Application(backend=BACKEND).connect(handle=win.handle)
main = app.window(handle=win.handle)

print("主窗口：", main.window_text(), main.class_name(), hex(main.handle))
# 建议先看前2~3层，避免太长
main.print_control_identifiers(depth=3)
