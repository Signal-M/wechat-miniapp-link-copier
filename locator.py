# -*- coding: utf-8 -*-
"""定位抽象层：主流程只依赖 locator.point(name)，不关心坐标是怎么来的。

当前实现
--------
CoordinateLocator：读用户标定好的屏幕坐标（JSON 文件）。

未来可替换
----------
TemplateLocator：用图像模板匹配去找「···」按钮等控件。
届时主流程一行都不用改，只需把工厂函数 create_locator() 的返回换掉即可 ——
这就是「预留 Vision 接口」的真实含义，不是空头承诺。

为什么必须有这一层
------------------
Windows 微信是 Qt 自绘，UI Automation 读不到控件名；Mac 微信虽是原生 App，
但小程序面板是 webview，同样读不到控件。所以「按坐标点」是双端唯一通用解。
代价：微信改版 / 换分辨率 / 换机器都会让坐标失效。
把定位隔离成一层，是为了将来能无痛换成更稳的定位方式，也便于在此集中提示用户重新标定。
"""
import os
import json

# 坐标存到用户主目录（与脚本位置无关）：更新或移动脚本都不会丢失，无需每次重新打点
POINTS_FILE = os.path.join(os.path.expanduser("~"), "wechat_scraper_points.json")

# 标定点的定义：(key, 展示给用户的说明)
# 说明文字保持双端通用，不写死 Windows 的 Ctrl+F 之类平台专有操作。
POINT_DEFS = [
    ("focus",
     "把鼠标移到微信主窗口【标题栏空白】处（仅用于聚焦窗口，别点到按钮）【全自动模式才需要】"),
    ("search_box",
     "把鼠标移到微信顶部的【搜索输入框】上【全自动模式才需要】"),
    ("search_result",
     "搜一个小程序名并回车后，把鼠标移到出现的【第一个搜索结果】上"
     "（建议先人工确认是小程序，后续也有链接校验兜底）【全自动模式才需要】"),
    ("first_tab",
     "搜索界面【最左侧第一个结果 tab】——每次搜索前先点它复位，"
     "忽略上一项打开的公众号/文章 tab（仅 Windows 全自动模式需要，macOS 会自动跳过）"),
    ("dots",
     "打开该小程序后，把鼠标移到小程序面板【右上角「···」按钮】上"),
    ("copy_link",
     "点开「···」菜单后，把鼠标移到【「复制链接」菜单项】上"),
    ("close_app",
     "小程序窗口【右上角「✕ 关闭」按钮】（在「···」左边）——"
     "复制完链接后点它关掉当前小程序，避免窗口越开越多"),
]


class Locator(object):
    """定位器接口。任何实现都要提供这些方法，主流程只依赖这个接口。"""

    def point(self, name):
        """返回目标位置的 (x, y)。未标定时应抛 KeyError。"""
        raise NotImplementedError

    def has(self, name):
        """该点是否已标定。"""
        raise NotImplementedError

    def set(self, name, xy):
        raise NotImplementedError

    def clear(self):
        raise NotImplementedError

    def save(self):
        raise NotImplementedError

    def all(self):
        """返回 {name: (x, y)} 的副本。"""
        raise NotImplementedError

    def calibrated_file_exists(self):
        """标定文件是否已存在（用于首启提示）。"""
        raise NotImplementedError


class CoordinateLocator(Locator):
    """基于标定坐标的定位器（当前实现）。"""

    def __init__(self, path=POINTS_FILE):
        self.path = path
        self.points = {}
        self.load()

    def load(self):
        if os.path.exists(self.path):
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    self.points = json.load(f)
            except Exception:
                self.points = {}
        else:
            self.points = {}

    def save(self):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.points, f, ensure_ascii=False, indent=2)

    def has(self, name):
        return name in self.points

    def point(self, name):
        return self.points[name]

    def set(self, name, xy):
        self.points[name] = list(xy)

    def clear(self):
        self.points = {}

    def all(self):
        return dict(self.points)

    def calibrated_file_exists(self):
        return os.path.exists(self.path)


def create_locator(path=POINTS_FILE):
    """工厂函数：将来要换成模板匹配，只改这里一行。"""
    return CoordinateLocator(path)
