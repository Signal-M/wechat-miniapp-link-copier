# -*- coding: utf-8 -*-
"""
微信小程序「复制链接」批量自动化 —— 图形界面版 (GUI v1)
========================================================
相比命令符版：全程窗口按钮操作，标定坐标用"倒计时自动捕获"（不用碰键盘、不会丢焦点）。
只做一件事：打开小程序 → 点「···」→ 点「复制链接」→ 读剪贴板 → 解析 AppID/Path → 存 Excel。

依赖（Windows 上装一次）：
    pip install pyautogui openpyxl pyperclip
（tkinter 是 Python 自带，无需另装）

用法：
    把本文件放到桌面，双击或用命令：python wechat_scraper_gui.py
界面按钮：
    [标定坐标]  按提示逐点捕获（SKIP_SEARCH 模式只需 2 个点）
    [开始运行]  批量跑（SKIP_SEARCH 模式会逐个弹窗让你手动开小程序）
    [检查剪贴板] 看当前剪贴板内容，验证"复制链接"格式
    [打开输入表] 打开 miniapps_input.xlsx（A 列填小程序名）
"""

import os, sys, time, json, re, threading, urllib.parse, traceback, types
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

BASE = os.path.dirname(os.path.abspath(__file__))
# 跨平台差异统一走 platform_adapt / locator 两个模块，本文件不再出现 `if mac`。

# ---------- 崩溃自报告：避免双击运行时一闪而过看不到错误 ----------
def _report_crash(title, detail=None):
    """把崩溃信息写到脚本同目录的 .error.log，并尽量弹窗提示。"""
    tb = detail if isinstance(detail, str) else traceback.format_exc()
    try:
        with open(os.path.join(BASE, "wechat_scraper_gui.error.log"), "w", encoding="utf-8") as f:
            f.write(tb)
    except Exception:
        pass
    try:
        r = tk.Tk(); r.withdraw()
        msg = (tb[-1500:] if len(tb) > 1500 else tb) + "\n\n（详情已写入 wechat_scraper_gui.error.log）"
        messagebox.showerror(title, msg)
    except Exception:
        pass

def _show_import_error(e):
    _report_crash("缺少依赖，无法启动",
                  "启动失败：缺少第三方依赖 pyautogui / openpyxl / pyperclip。\n"
                  "请在命令行运行：\n    pip install pyautogui openpyxl pyperclip\n\n具体错误：%s" % e)

# ---------- 兼容坏依赖（Python 3.14 等 numpy/cv2 装坏的环境）----------
# 本脚本只用鼠标/键盘控制，完全不需要截图功能。但 import pyautogui 会经 pyscreeze 间接
# import cv2 -> numpy；而部分环境的 numpy/cv2 装坏（import 阶段就直接 AttributeError 崩溃）。
# 由于从不用截图，这里在 import pyautogui 之前做两件事：
#   1) 把 cv2 顶成一个"占位模块桩"——让 pyscreeze 的 `import cv2` 成功且不触发坏 numpy；
#   2) 把 numpy 的 import 直接"拦截并抛 ModuleNotFoundError(属于 ImportError)"——这样
#      pyscreeze 和 openpyxl 都会走各自"无 numpy 降级分支"：
#        - pyscreeze：置 _useOpenCV=False，关掉 OpenCV 截图（本脚本用不到）；
#        - openpyxl：用 (int, float, decimal.Decimal) 作为 NUMERIC_TYPES。
# 关键点：必须"拦截成 ImportError"，而不是给 numpy 一个桩模块——因为 openpyxl 会
# `from numpy import ...` 组装 NUMERIC_TYPES 并用于 isinstance()，任何桩都会污染该元组，
# 导致写 xlsx 时 safe_string 的 isnan 崩溃（已实测）。本机鼠标/键盘能力不依赖 cv2/numpy。
import importlib, importlib.abc, importlib.util

class _Cv2Stub(types.ModuleType):
    __version__ = '4.0.0'
    def __getattr__(self, name):
        return 'stub'

class _NumpyBlockLoader(importlib.abc.Loader):
    def create_module(self, spec):
        return None
    def exec_module(self, module):
        raise ModuleNotFoundError("numpy 已被拦截（坏依赖规避）：%s" % module.__name__)

class _NumpyBlocker(importlib.abc.MetaPathFinder):
    """让 `import numpy` / `from numpy import ...` 直接抛 ModuleNotFoundError。"""
    def find_spec(self, fullname, path, target=None):
        if fullname == 'numpy' or fullname.startswith('numpy.'):
            return importlib.util.spec_from_loader(
                fullname, _NumpyBlockLoader(), is_package=(fullname == 'numpy'))
        return None

if 'cv2' not in sys.modules:
    sys.modules['cv2'] = _Cv2Stub('cv2')
sys.meta_path.insert(0, _NumpyBlocker())

# 第三方依赖（缺 pyautogui/pyperclip/openpyxl 会在启动时报错，并提示 pip install）
try:
    import pyautogui
    import pyperclip
    from openpyxl import load_workbook, Workbook
except ImportError as e:
    _show_import_error(e)
    sys.exit(1)

pyautogui.FAILSAFE = True  # 鼠标移到屏幕左上角可紧急中止

import platform_adapt as pa   # 双端适配层：修饰键 / 打开文件 / 权限提示
import locator                # 定位抽象层：坐标实现，预留模板匹配接口

# 输入 / 输出统一放在 data/ 子目录：演示时只需要盯这一个文件夹
DATA_DIR = os.path.join(BASE, "data")
INPUT_XLSX = os.path.join(DATA_DIR, "输入.xlsx")
OUTPUT_XLSX = os.path.join(DATA_DIR, "结果.xlsx")
# 标定点定义与坐标存取已移到 locator 模块，此处仅做别名便于界面代码引用
POINT_DEFS = locator.POINT_DEFS


def ensure_data_ready():
    """确保 data/ 目录与输入表存在。

    演示友好：第一次打开不会因为缺文件就报错，而是直接拿到一张带表头的空输入表，
    填好名称就能开跑。异常一律静默忽略，不阻塞界面启动。
    """
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        if not os.path.exists(INPUT_XLSX):
            wb = Workbook()
            ws = wb.active
            ws.append(["小程序名称"])
            wb.save(INPUT_XLSX)
    except Exception:
        pass


ensure_data_ready()
WAIT_OPEN = 5.0       # 打开小程序后默认等待(秒)，界面可调；加载慢的调到 8~10
WAIT_SEARCH_RESULT = 3.0  # 搜完回车后、点第一个结果前的等待(秒)，界面可调；结果加载慢调到 4~5
WAIT_MENU = 0.8
WAIT_CLOSE = 0.8  # 关闭小程序窗口后等待
COPY_WAIT = 1.0        # 点完「复制链接」后等剪贴板的基础时间
RETRY_BACKOFF = 1.5    # 每次重试额外增加的等待(秒)，应对加载慢
MAX_COPY_TRIES = 3     # 复制链接最大尝试次数（链接异常/未找到时自动重试）
CAPTURE_SECONDS = 5  # 标定倒计时秒数


def is_valid_link(text):
    """判断剪贴板内容是否像一条真实的小程序分享链接（而非错误提示/空）。"""
    if not text:
        return False
    t = text.strip()
    if "SCRAPER_SENTINEL" in t:
        return False
    # 真实小程序链接特征：含 appid=wx… / 以 #小程序:// 开头 / 是 http(s) 链接
    if re.search(r'appid=wx[0-9a-zA-Z]+', t):
        return True
    if t.startswith("#小程序://"):
        return True
    if re.match(r'https?://', t):
        return True
    return False


# ---------- 数据与解析 ----------
def read_names():
    if not os.path.exists(INPUT_XLSX):
        return None
    wb = load_workbook(INPUT_XLSX)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        return []
    header = rows[0]
    # 表头识别：首行含 场地/名称/小程序/链接 等字样则视为表头，跳过
    HINTS = ('场地', '名称', '小程序', '链接', 'appid', 'path', '状态', 'id')
    looks_like_header = any(
        c and any(h in str(c).lower() for h in HINTS) for c in header[:2]
    )
    # 名称列判定：对照表格式里「网球场名称/小程序名称」在 B 列(index 1)，
    # 普通单列表单在 A 列(index 0)。自动选对列，避免拿 场地ID 当搜索词。
    name_col = 1 if (len(header) > 1 and header[1]
                     and any(k in str(header[1]) for k in ('名称', '小程序'))) else 0
    data_rows = rows[1:] if looks_like_header else rows
    names = []
    for row in data_rows:
        if name_col < len(row):
            v = row[name_col]
            if v and str(v).strip():
                names.append(str(v).strip())
    return names


def is_success_status(status):
    """只有这些状态算"已完成"，失败/异常的不算，下次运行会留作待重试。

    "链接已存但无AppID" 是旧版遗留状态（现已不再产生，改为直接记 OK），
    保留在列表里是为了兼容已有的结果文件，避免这些条目被重复重跑。
    """
    return status in ("OK", "链接已存但无AppID")


def load_done():
    done = {}
    if os.path.exists(OUTPUT_XLSX):
        wb = load_workbook(OUTPUT_XLSX)
        ws = wb.active
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and r[0] and r[4] is not None:
                # 仅把"成功"的记为已完成；失败/异常仍留作待重试
                if is_success_status(str(r[4]).strip()):
                    done[str(r[0]).strip()] = True
    return done


def load_failed():
    """返回上次输出表中"失败/异常"的名称集合（用于重试提示）。"""
    failed = set()
    if os.path.exists(OUTPUT_XLSX):
        wb = load_workbook(OUTPUT_XLSX)
        ws = wb.active
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and r[0] and r[4] is not None:
                if not is_success_status(str(r[4]).strip()):
                    failed.add(str(r[0]).strip())
    return failed


def remove_rows_by_name(name):
    """重写输出表，删除指定名称的所有旧行（避免失败重试时残留重复行）。表头保留。
    若输出 Excel 被占用则跳过（交由 write_row 在重试窗口内统一处理）。"""
    if not os.path.exists(OUTPUT_XLSX):
        return
    try:
        wb = load_workbook(OUTPUT_XLSX)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
    except PermissionError:
        return  # 文件被占用，暂时跳过去重；write_row 会在重试窗口内覆盖
    if not rows:
        return
    header, body = rows[0], rows[1:]
    kept = [header] + [r for r in body if not (r and r[0] and str(r[0]).strip() == name)]
    if len(kept) == len(rows):
        return  # 没有该名称的行，无需重写
    try:
        wb2 = Workbook()
        ws2 = wb2.active
        for r in kept:
            ws2.append(r)
        wb2.save(OUTPUT_XLSX)
    except PermissionError:
        return


def write_row(name, link, appid, path, status):
    """写一行结果。若输出 Excel 被占用（如用户在 Excel 里打开），会自动重试，
    避免 PermissionError 把整轮采集打断。"""
    # 写新结果前，先清掉该名称的旧行（尤其是上次失败的残留），避免重复
    remove_rows_by_name(name)
    last_err = None
    for _ in range(30):  # 最多等 ~15 秒让文件释放
        try:
            if os.path.exists(OUTPUT_XLSX):
                wb = load_workbook(OUTPUT_XLSX)
                ws = wb.active
            else:
                wb = Workbook()
                ws = wb.active
                ws.append(["名称", "复制链接", "AppID", "Path", "状态"])
            ws.append([name, link or "", appid or "", path or "", status])
            wb.save(OUTPUT_XLSX)
            return
        except PermissionError as e:
            last_err = e
            time.sleep(0.5)
    raise last_err


def parse_link(text):
    appid = None
    m = re.search(r'appid=(wx[0-9a-zA-Z]+)', text)
    if m:
        appid = m.group(1)
    path = None
    m2 = re.search(r'url=([^&\s]+)', text)
    if m2:
        try:
            path = urllib.parse.unquote(m2.group(1))
        except Exception:
            path = m2.group(1)
    return appid, path


# ---------- 界面 ----------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("小程序复制链接采集器 (GUI)")
        self.root.geometry("640x520")
        self.skip_search = tk.BooleanVar(value=True)
        self.locator = locator.create_locator()
        self.continue_event = threading.Event()
        self.running = False

        # 顶部：模式开关
        top = ttk.Frame(root)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Checkbutton(top, text="手动打开小程序（只自动复制链接，推荐）",
                        variable=self.skip_search).pack(side="left")
        ttk.Label(top, text="打开后等待(秒):").pack(side="left", padx=(12, 2))
        self.wait_open_var = tk.StringVar(value=str(WAIT_OPEN))
        self.wait_open_entry = ttk.Entry(top, textvariable=self.wait_open_var, width=5)
        self.wait_open_entry.pack(side="left")
        ttk.Label(top, text="搜→点等(秒):").pack(side="left", padx=(8, 2))
        self.wait_search_var = tk.StringVar(value=str(WAIT_SEARCH_RESULT))
        ttk.Entry(top, textvariable=self.wait_search_var, width=5).pack(side="left")
        ttk.Button(top, text="打开输入表", command=self.open_input).pack(side="right")
        ttk.Button(top, text="打开数据目录", command=self.open_data_dir).pack(side="right")

        # 按钮区
        bf = ttk.Frame(root)
        bf.pack(fill="x", padx=10, pady=4)
        ttk.Button(bf, text="① 标定坐标", command=self.start_teach).pack(side="left", padx=4)
        ttk.Button(bf, text="② 开始运行", command=self.start_run).pack(side="left", padx=4)
        ttk.Button(bf, text="③ 清空结果重跑", command=self.reset_and_run).pack(side="left", padx=4)
        ttk.Button(bf, text="检查剪贴板", command=self.check_clip).pack(side="left", padx=4)

        # 标定说明 + 进度
        self.instr = ttk.Label(root, text="先点「① 标定坐标」，按提示把鼠标移到对应位置后点「捕获」。",
                               wraplength=600, justify="left")
        self.instr.pack(fill="x", padx=10, pady=4)
        self.capture_btn = ttk.Button(root, text="捕获当前鼠标坐标", state="disabled",
                                      command=self.on_capture)
        self.capture_btn.pack(pady=2)
        self.pt_label = ttk.Label(root, text=self._points_text(), foreground="#555")
        self.pt_label.pack(fill="x", padx=10)

        # 日志
        ttk.Label(root, text="运行日志：").pack(anchor="w", padx=10)
        self.log = scrolledtext.ScrolledText(root, height=14, state="normal")
        self.log.pack(fill="both", expand=True, padx=10, pady=4)

        self.teach_idx = 0
        self.log_msg("就绪。数据目录：" + DATA_DIR)
        self.log_msg("结果将保存到：" + OUTPUT_XLSX)
        self.log_msg(pa.permission_hint())
        if self.locator.calibrated_file_exists():
            self.log_msg("已检测到历史标定坐标（位于用户主目录），无需重新打点，可直接点「② 开始运行」。")
        else:
            self.log_msg("提示：首次使用请先点「① 标定坐标」。标定后坐标会保存在用户主目录，之后更新脚本也不丢失。")

    def _points_text(self):
        pts = self.locator.all()
        if not pts:
            return "已标定坐标：无"
        return "已标定坐标：" + "  ".join(f"{k}={v}" for k, v in pts.items())

    def log_msg(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)

    def open_input(self):
        if not os.path.exists(INPUT_XLSX):
            self.log_msg("输入表不存在：" + INPUT_XLSX)
            return
        pa.open_file(INPUT_XLSX)

    def open_data_dir(self):
        """打开数据目录（Finder / 资源管理器），演示时一键找到输入与结果。"""
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
        except Exception:
            pass
        pa.open_file(DATA_DIR)
        self.log_msg("数据目录：" + DATA_DIR)

    def check_clip(self):
        self.log_msg("剪贴板内容：" + repr(pyperclip.paste()))

    # ---- 标定 ----
    def start_teach(self):
        self.locator.clear()
        self.teach_idx = 0
        self.log_msg("=== 开始标定 ===")
        self.next_teach_point()

    def next_teach_point(self):
        if self.teach_idx >= len(POINT_DEFS):
            self.locator.save()
            self.pt_label.config(text=self._points_text())
            self.capture_btn.config(state="disabled")
            self.instr.config(text="标定完成！现在点「② 开始运行」。")
            self.log_msg("坐标已保存。")
            return
        key, desc = POINT_DEFS[self.teach_idx]
        if (self.skip_search.get() and key in ("focus", "search_box", "search_result", "first_tab")) or (pa.IS_MAC and key == "first_tab"):
            self.log_msg(f"(跳过 {key}，手动模式不需要)")
            self.teach_idx += 1
            self.next_teach_point()
            return
        self.instr.config(text=f"第 {self.teach_idx+1} 点 · {key}：\n{desc}\n\n就位后点下面「捕获当前鼠标坐标」。")
        self.capture_btn.config(state="normal")

    def on_capture(self):
        key, _ = POINT_DEFS[self.teach_idx]
        self.capture_btn.config(state="disabled")
        self.root.iconify()
        tl = tk.Toplevel()
        tl.overrideredirect(True)
        tl.attributes("-topmost", True)
        sw = self.root.winfo_screenwidth()
        tl.geometry(f"260x90+{sw-280}+20")
        lbl = ttk.Label(tl, text="", font=("Microsoft YaHei", 12), background="#222", foreground="#fff")
        lbl.pack(expand=True, fill="both")
        for i in range(CAPTURE_SECONDS, 0, -1):
            lbl.config(text=f"{i} 秒后自动捕获\n把鼠标移到目标位置")
            tl.update()
            time.sleep(1)
        x, y = pyautogui.position()
        tl.destroy()
        self.root.deiconify()
        self.locator.set(key, (x, y))
        self.pt_label.config(text=self._points_text())
        self.log_msg(f"  已记录 {key} = ({x}, {y})")
        self.teach_idx += 1
        self.next_teach_point()

    # ---- 运行 ----
    def start_run(self):
        if self.running:
            return
        if not self.locator.has("dots") or not self.locator.has("copy_link"):
            messagebox.showerror("缺少坐标", "请先点「① 标定坐标」记录 dots / copy_link 两个点。")
            return
        if not self.skip_search.get():
            # 全自动模式必需的坐标点。first_tab 是 Windows 微信专属（macOS 标定时会自动跳过），
            # 故不计入必需项，避免 Mac 用户被一个无意义的校验挡在门外。
            # 这里必须显式校验，否则缺哪个点只会在运行时抛 KeyError，用户看不懂。
            required = ["focus", "search_box", "search_result"]
            missing = [k for k in required if not self.locator.has(k)]
            if missing:
                messagebox.showerror("缺少坐标",
                    "全自动模式还需要标定这些点：" + "、".join(missing) + "\n"
                    "请点「① 标定坐标」补上。\n"
                    "（first_tab 是 Windows 微信专属，macOS 不需要标定）")
                return
            if not self.locator.has("first_tab"):
                # Mac 微信通常没有 Windows 那种「搜索结果 tab」，且每次循环都会「全选+粘贴」清空搜索框，
                # 因此 first_tab 复位在 Mac 上不是必需；仅在 Windows 下若遇干扰才需要标定。
                self.log_msg("⚠️ 未标定 first_tab：Mac 微信通常无此 tab，将靠每次「全选+粘贴」清空搜索框；"
                             "若发现上一项残留干扰，可在 Windows 下补标定 first_tab。")
            if not self.locator.has("close_app"):
                messagebox.showwarning("建议补充点位",
                    "全自动模式建议标定 close_app（关闭小程序按钮），否则会累积很多小程序窗口。\n"
                    "点确定将按 Esc 关菜单（小程序窗口不关闭），仍可能有残留窗口。")
        names = read_names()
        if names is None:
            messagebox.showerror("缺少输入表", "找不到 " + INPUT_XLSX)
            return
        if not names:
            messagebox.showinfo("输入表是空的",
                "请先在输入表的 A 列填入小程序名称（一行一个）：\n" + INPUT_XLSX)
            return
        self.running = True
        threading.Thread(target=self.run_worker, daemon=True).start()

    def reset_and_run(self):
        """清空已有结果表后全量重跑（不区分成功/失败，所有输入名都重抓）。"""
        if self.running:
            return
        if not messagebox.askyesno("确认清空", "这会清空 " + OUTPUT_XLSX + " 里所有已采集结果并从头重跑，确定？"):
            return
        # 只清空数据行、保留文件本身。
        # 旧实现是 os.remove 直接删文件：一旦这一轮没跑出任何结果（中途中止 / 出错），
        # 历史结果就永久丢失且无法恢复。
        if os.path.exists(OUTPUT_XLSX):
            try:
                wb = load_workbook(OUTPUT_XLSX)
                ws = wb.active
                rows_to_del = ws.max_row - 1
                if rows_to_del > 0:
                    ws.delete_rows(2, rows_to_del)
                wb.save(OUTPUT_XLSX)
            except Exception as e:
                messagebox.showerror("清空失败", str(e))
                return
        self.log_msg("已清空结果（保留文件），准备全量重跑。")
        self.start_run()

    def run_worker(self):
        names = read_names()
        done = load_done()
        failed = load_failed()
        pending = [n for n in names if n not in done]
        self.gui_log(f"输入 {len(names)} 个，已成功 {len(done)} 个，待处理 {len(pending)} 个"
                     + (f"（其中上次失败待重试 {len(failed & set(pending))} 个）" if failed else ""))
        if not pending:
            self.gui_log("全部已成功，无需处理。如需强制重跑请点「③ 清空结果重跑」。")
            self.running = False
            return
        if not self.skip_search.get():
            self.gui_log("全自动模式：请保持微信窗口固定，开始自动搜索。")
        for i, name in enumerate(pending, 1):
            try:
                # 读取可调的"打开后等待"与"搜→点等待"时间（界面输入框，非法则回退默认）
                try:
                    wait_open = float(self.wait_open_var.get())
                    if wait_open < 0:
                        wait_open = WAIT_OPEN
                except Exception:
                    wait_open = WAIT_OPEN
                try:
                    wait_search = float(self.wait_search_var.get())
                    if wait_search < 0:
                        wait_search = WAIT_SEARCH_RESULT
                except Exception:
                    wait_search = WAIT_SEARCH_RESULT
                if self.skip_search.get():
                    self.gui_log(f"\n--- [{i}/{len(pending)}] {name} ---")
                    self.ask_open(name)
                    self.continue_event.wait()
                    self.continue_event.clear()
                else:
                    self.gui_log(f"\n--- [{i}/{len(pending)}] {name} ---")
                    pyautogui.click(*self.locator.point("focus")); time.sleep(0.4)
                    # 【可选，仅 Windows 微信】每次搜索前先「回到第一个 tab」复位搜索界面：
                    # 上一项若搜到公众号/文章会另开 tab，先点回第一个 tab 丢弃这些干扰。
                    # macOS 微信没有这个 tab，标定阶段会自动跳过它 —— 所以这里必须判存在再点，
                    # 否则取不到会抛 KeyError，导致整批全部失败。
                    if self.locator.has("first_tab"):
                        pyautogui.click(*self.locator.point("first_tab")); time.sleep(0.6)
                    # 再点搜索框聚焦（不要用 Ctrl+F，微信里 Ctrl+F 会触发聊天内搜索）
                    pyautogui.click(*self.locator.point("search_box")); time.sleep(0.4)
                    # 中文名必须走剪贴板粘贴（typewrite 只能发 ASCII），修饰键由适配层统一处理
                    pa.input_text(name); time.sleep(0.5)
                    pyautogui.press("enter"); time.sleep(wait_search)
                    # 直接点第一个搜索结果（搜索结果页「小程序」标签位置不固定/可能不存在，故不再做标签过滤；
                    # 若结果实际是公众号/文章，复制链接后会被 looks_like_article 兜底标为疑似非小程序）
                    pyautogui.click(*self.locator.point("search_result")); time.sleep(wait_open)
                link, appid, path = self.do_copy_link()
                # 兜底校验：即使前面点了小程序标签，仍可能拿到公众号/文章链接
                # 文章链接形如 https://mp.weixin.qq.com/...  —— 说明搜到/点的不是小程序
                looks_like_article = bool(re.search(r'mp\.weixin\.qq\.com', link or ""))
                if appid:
                    status = "OK"
                    self.gui_log(f"  ✅ 已获取到链接，链接预览 {link} 即可")
                    self.gui_log(f"     AppID={appid}  Path={path or ''}")
                elif looks_like_article:
                    # 复制到了公众号/文章链接 → 标成"疑似非小程序"，不计入成功，下次自动复跑
                    status = "疑似非小程序链接(公众号/文章)"
                    self.gui_log(f"  ⚠️ {status} —— 搜索结果可能非小程序，请核对名称后重试。")
                elif is_valid_link(link):
                    # 拿到有效链接即算成功，不再提示「无 AppID」。
                    # 原因：小程序面板「复制链接」给出的链接经常是 #小程序://名称/路径 这种格式，
                    # 它本身就不含 appid 参数，这一环节本来就取不到 AppID，
                    # 显示「无 AppID」只会让人误以为是失败。
                    status = "OK"
                    self.gui_log(f"  ✅ 已获取到链接，链接预览 {link} 即可")
                else:
                    status = "失败(多次重试仍无有效链接)"
                    self.gui_log(f"  ❌ {status}  最后内容：{(link or '')[:60]}")
                write_row(name, link, appid, path, status)
                if not self.skip_search.get():
                    if looks_like_article:
                        # 公众号/文章场景：没有真正打开小程序窗口，改回「第一个 tab」关掉那些文章 tab，
                        # 复位搜索界面，忽略后边打开的 tab，直接进入下一项
                        self.gui_log("  ↩ 回到第一个 tab 复位搜索界面（忽略公众号/文章 tab）")
                        if self.locator.has("first_tab"):
                            pyautogui.click(*self.locator.point("first_tab")); time.sleep(0.6)
                        else:
                            pyautogui.press("escape"); time.sleep(0.4)
                    elif self.locator.has("close_app"):
                        # 优先点「关闭小程序」按钮，彻底关掉当前小程序窗口（避免累积）
                        pyautogui.click(*self.locator.point("close_app")); time.sleep(WAIT_CLOSE)
                        # 若「···」菜单还开着，补一次 Esc 收掉菜单
                        pyautogui.press("escape"); time.sleep(0.4)
                    else:
                        # 未标定关闭按钮时退回旧行为：连按 Esc 关菜单（小程序窗口可能残留）
                        pyautogui.press("escape"); time.sleep(0.5)
                        pyautogui.press("escape"); time.sleep(0.5)
            except pyautogui.FailSafeException:
                self.gui_log("触发 FailSafe（鼠标移到左上角），已中止。")
                break
            except Exception as e:
                hint = ""
                if isinstance(e, PermissionError) or "Permission denied" in str(e):
                    hint = "  👉 请先关闭已打开的 miniapps_links.xlsx（Excel 占用会导致写不进），关闭后点「② 开始运行」会自动补抓跳过的项。"
                self.gui_log(f"  ❌ 出错（已跳过，继续下一个）：{e}{hint}")
                try:
                    write_row(name, "", "", "", "异常:" + str(e)[:30])
                except Exception as we:
                    self.gui_log(f"  ⚠️ 结果写入失败（请关闭 Excel 后点开始运行补抓）：{we}")
                continue
        self.gui_log("\n=== 完成！结果在 " + OUTPUT_XLSX)
        self.running = False

    def ask_open(self, name):
        def on_ok():
            self.cont_win.destroy()
            self.continue_event.set()
        self.root.after(0, lambda: self._show_ask(name, on_ok))

    def _show_ask(self, name, on_ok):
        self.cont_win = tk.Toplevel(self.root)
        self.cont_win.title("请手动打开小程序")
        self.cont_win.attributes("-topmost", True)
        ttk.Label(self.cont_win, text=f"请手动打开小程序：\n「{name}」\n并使其置前，然后点继续。",
                  wraplength=260, justify="center", padding=12).pack()
        ttk.Button(self.cont_win, text="已打开，继续", command=on_ok).pack(pady=10)

    def do_copy_link(self):
        """打开小程序后点「···」→「复制链接」，复制到真实链接即返回。
        慢加载场景下第一次可能拿到错误提示/空，这里做渐进式重试：
        每失败一次多等 RETRY_BACKOFF 秒（让小程序继续加载），最多 MAX_COPY_TRIES 次。"""
        sentinel = "___SCRAPER_SENTINEL_%d___" % int(time.time())
        last_link = ""
        for attempt in range(1, MAX_COPY_TRIES + 1):
            pyperclip.copy(sentinel)
            time.sleep(0.3)
            pyautogui.click(*self.locator.point("dots")); time.sleep(WAIT_MENU)
            pyautogui.click(*self.locator.point("copy_link"))
            # 第1次用基础等待，重试时叠加 backoff（应对加载慢）
            time.sleep(COPY_WAIT + (attempt - 1) * RETRY_BACKOFF)
            link = pyperclip.paste()
            if is_valid_link(link):
                appid, path = parse_link(link)
                if attempt > 1:
                    self.gui_log(f"  ↻ 第{attempt}次重试成功")
                return link, appid, path
            last_link = "" if link == sentinel else link
            if attempt < MAX_COPY_TRIES:
                self.gui_log(f"  ⚠️ 第{attempt}次未取到有效链接（可能还在加载），"
                             f"多等{RETRY_BACKOFF}s后重试…")
        # 用尽重试仍失败：把最后一次拿到的（错误提示/空）原样返回，状态标记失败
        appid, path = (None, None) if (not last_link or "SCRAPER_SENTINEL" in last_link) else parse_link(last_link)
        return last_link, appid, path

    def gui_log(self, msg):
        self.root.after(0, lambda: self.log_msg(msg))


def main():
    try:
        root = tk.Tk()
        App(root)
        root.mainloop()
    except Exception as e:
        _report_crash("脚本崩溃", e)


if __name__ == "__main__":
    main()
