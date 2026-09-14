# -*- coding: utf-8 -*-
"""跨平台适配层：把 macOS / Windows 的差异全部收在这里。

约定：主流程里不允许出现 `if mac`，一律调用本模块提供的能力。
这样将来支持第三个平台、或修正某个平台的差异时，只改这一个文件。

注意：本模块对 pyautogui / pyperclip 采用「函数内延迟导入」。
原因是主程序在 import pyautogui 之前，会先安装一个 numpy 拦截器来规避
某些机器上装坏的 numpy/cv2；若本模块在模块级导入 pyautogui，会抢在拦截器
生效之前触发导入，导致那条保护失效。
"""
import os
import sys

IS_MAC = sys.platform == "darwin"
IS_WIN = sys.platform == "win32"

# macOS 的修饰键是 Command，Windows 是 Ctrl。
# 粘贴 / 全选必须用对应键，否则中文小程序名根本粘不进微信搜索框。
MOD = "command" if IS_MAC else "ctrl"


def open_file(path):
    """用系统默认程序打开文件（macOS: open / Windows: startfile / Linux: xdg-open）。"""
    if IS_MAC:
        import subprocess
        subprocess.run(["open", path])
    elif IS_WIN:
        os.startfile(path)
    else:
        import subprocess
        subprocess.run(["xdg-open", path])


def select_all():
    """全选当前焦点控件里的文本，用于清空上一次的残留。"""
    import pyautogui
    pyautogui.hotkey(MOD, "a")


def paste():
    """把剪贴板内容粘贴到当前焦点控件。"""
    import pyautogui
    pyautogui.hotkey(MOD, "v")


def input_text(text):
    """把文本送进当前焦点控件（先全选清空，再粘贴）。

    中文必须走剪贴板：pyautogui.typewrite() 只能发 ASCII，打不出「得乐运动」这种中文名。
    这是整个自动化方案成立的前提，不要改成 typewrite。
    """
    import pyperclip
    pyperclip.copy(text)
    select_all()
    paste()


def permission_hint():
    """运行前需要用户手动处理的环境说明（各平台不同）。"""
    if IS_MAC:
        return ("macOS 需要授权：系统设置 → 隐私与安全性 → 辅助功能 与 屏幕录制，"
                "把「终端 / Terminal」勾上。否则脚本既动不了鼠标键盘，也截不了图。")
    if IS_WIN:
        return "Windows 无需额外授权，但请保持微信窗口位置固定，运行期间不要移动或最小化。"
    return "请保持微信窗口位置固定，运行期间不要移动或最小化。"


def modifier_name():
    """给界面文案用的修饰键名（如「Cmd+V」/「Ctrl+V」）。"""
    return "Cmd" if IS_MAC else "Ctrl"
