import itertools
import sys
import os
import json
import threading
import time
import uuid
from copy import deepcopy
from datetime import datetime, date, timedelta

import openpyxl, random, itertools
import pyautogui

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QPushButton, QGroupBox, QLabel, QLineEdit,
    QComboBox, QTimeEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QFormLayout, QSpinBox,
    QListWidgetItem, QCheckBox, QMenu, QFrame, QStyle,
    QSplitter, QSizePolicy, QDialog, QDialogButtonBox,
    QGridLayout, QFileDialog, QMessageBox, QDoubleSpinBox,
    QTextEdit, QPlainTextEdit, QSystemTrayIcon, QScrollArea, QInputDialog, QDateEdit, QDateTimeEdit, QWidgetAction
)
from PySide6.QtCore import Qt, QTime, QSize, QSettings, Signal, QObject, QTimer, QPointF, QRectF, QDate, QDateTime, \
    QPoint, QRect, QThread, Slot
from PySide6.QtGui import QIcon, QAction, QFont, QPalette, QColor, QLinearGradient, QTextCursor, QKeySequence, QPixmap, \
    QBrush, QPainterPath, QPainter, QPen, QMouseEvent

import os
from pathlib import Path

from pynput import keyboard


# 工具函数
def resource_path(relative_path: str) -> str:
    """打包 / 开发环境下通用的资源路径解析"""
    try:
        base_path = sys._MEIPASS           # PyInstaller 运行时
    except AttributeError:
        base_path = os.path.abspath(".")   # 开发环境
    return os.path.join(base_path, relative_path)

# hotkey_listener.py


class HotkeyListener(QThread):
    # 自定义信号：当按下 Esc 时触发
    hotkey_activated = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.stop_event = threading.Event()  # 控制监听循环退出

    def run(self):
        """QThread 的主执行函数"""
        def on_press(key):
            if self.stop_event.is_set():
                return False  # 停止监听器
            try:
                if key == keyboard.Key.esc:
                    self.hotkey_activated.emit()  # 安全发射信号到主线程
            except Exception as e:
                print(f"热键监听错误: {e}")

        # 启动 pynput 键盘监听（阻塞）
        with keyboard.Listener(on_press=on_press) as listener:
            listener.join()

    def stop(self):
        """安全停止监听线程"""
        self.stop_event.set()  # 触发退出
        self.quit()           # 请求线程退出
        self.wait()           # 等待线程结束

class RegionCaptureOverlay(QWidget):
    finished = Signal(QRect)   # 自定义信号，返回选区

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(Qt.FramelessWindowHint
                            | Qt.WindowStaysOnTopHint
                            | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setCursor(Qt.CrossCursor)
        self.setFixedSize(QApplication.primaryScreen().size())

        self.start_pos = QPoint()
        self.end_pos   = QPoint()

    # ---------- 事件 ----------
    def mousePressEvent(self, event: QMouseEvent):
        self.start_pos = event.globalPosition().toPoint()
        self.end_pos   = self.start_pos

    def mouseMoveEvent(self, event: QMouseEvent):
        self.end_pos = event.globalPosition().toPoint()
        self.update()

    def mouseReleaseEvent(self, event: QMouseEvent):
        rect = QRect(self.start_pos, self.end_pos).normalized()
        self.finished.emit(rect)
        # self.close()
        # print("🖱️ 鼠标释放，发送区域信号并关闭覆盖层")

    # ---------- 绘制 ----------
    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        # 半透明背景
        p.setBrush(QColor(0, 0, 0, 100))
        p.setPen(Qt.NoPen)
        p.drawRect(self.rect())

        # 红色选框
        if not self.start_pos.isNull():
            p.setPen(QPen(Qt.red, 2))
            p.setBrush(Qt.NoBrush)
            p.drawRect(QRect(self.start_pos, self.end_pos).normalized())
class StepTableHelper:
    """负责把步骤对象渲染成表格行的工具类，可放到主窗口里复用"""
    FIXED_ROW_HEIGHT = 32          # 统一行高（像素）
    ICON_SIZE = 20          # 左侧图标宽/高
    IMG_HEIGHT = 32         # 如果是图片，缩略图高度

    @staticmethod
    def desc_of(step: dict) -> str:
        """给每种步骤生成一句精炼描述，用于第1列"""
        t = step["type"]
        p = step["params"]
        # 使用当前时间，只显示时分秒
        time_str = datetime.now().strftime("%H:%M:%S")
        if t == "鼠标点击":
            return f"点击 · {os.path.basename(p.get('image_path', ''))} · {time_str}"
        if t == "文本输入":
            txt = p.get("text", "")
            if txt:
                return f"键盘 · {txt[:10]}{'…' if len(txt) > 10 else ''} · {time_str}"
            mode = p.get("mode", "顺序")
            file = os.path.basename(p.get("excel_path", ""))
            return f"键盘 · {mode}·{file} · {time_str}"
        if t == "等待":
            return f"等待 · {p.get('seconds', 0)}s · {time_str}"
        if t == "截图":
            return f"截图 · {os.path.basename(p.get('save_path', ''))} · {time_str}"
        if t == "鼠标滚轮":
            return f"滚轮 · {p.get('direction', '向下')}{p.get('clicks', 3)}格 · {time_str}"
        if t == "拖拽":
            return f"拖拽 · ({p.get('start_x', 0)},{p.get('start_y', 0)})→({p.get('end_x', 0)},{p.get('end_y', 0)}) · {time_str}"
        return t

    @staticmethod
    def widget_of(step: dict) -> QWidget:
        """
        返回一个可直接塞进 QTableWidget 的 QWidget，
        内部 QLabel 负责显示图标/文字/图片 + 时间
        """
        t = step["type"]
        p = step["params"]
        time_str = p.get("step_time",datetime.now().strftime("%H:%M:%S"))

        # 主容器
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(4, 2, 4, 2)

        # 左侧图标或图片
        icon_label = QLabel()
        icon_label.setFixedSize(StepTableHelper.ICON_SIZE, StepTableHelper.ICON_SIZE)
        icon_label.setScaledContents(True)

        # 中间文字/图片
        content_label = QLabel()
        content_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        content_label.setStyleSheet("""color:#ffffff;
background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
border-radius:6px;
padding:2px 6px;
font-weight:bold;""")

        # 右侧时间
        time_label = QLabel(time_str)
        time_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        font = QFont()
        font.setPointSize(8)
        time_label.setFont(font)
        time_label.setStyleSheet("""color:#ffffff;
background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #0078d4,stop:1 #00bcf2);
border-radius:6px;
padding:2px 6px;
font-weight:bold;""")

        # 根据类型生成内容
        if t == "鼠标点击":
            img_path = p.get("image_path", "")
            if os.path.isfile(img_path):
                pm = QPixmap(img_path).scaledToHeight(StepTableHelper.IMG_HEIGHT, Qt.SmoothTransformation)
                content_label.setPixmap(pm)
            else:
                content_label.setText(os.path.basename(img_path))
            icon_label.setText("👆")

        elif t == "文本输入":
            txt = p.get("text", "")
            if txt:
                txt = txt[:10] + "…" if len(txt) > 10 else txt
                content_label.setText(txt)
            else:
                mode = p.get("mode", "顺序")
                file = os.path.basename(p.get("excel_path", ""))
                content_label.setText(f"{mode}·{file}")
            icon_label.setText("⌨")

        elif t == "等待":
            content_label.setText(f"{p.get('seconds', 0)}s")
            icon_label.setText("⏱")

        elif t == "截图":
            save_path = p.get("save_path", "")
            if os.path.isfile(save_path):
                pm = QPixmap(save_path).scaledToHeight(StepTableHelper.IMG_HEIGHT, Qt.SmoothTransformation)
                content_label.setPixmap(pm)
            else:
                content_label.setText(os.path.basename(save_path))
            icon_label.setText("📸")

        elif t == "鼠标滚轮":
            dire = p.get("direction", "向下")
            clicks = p.get("clicks", 3)
            content_label.setText(f"{dire}{clicks}格")
            icon_label.setText("⚙")
        elif t == "键盘热键":
            hotkey = p.get("hotkey", "ctrl+c").upper()
            delay = p.get("delay_ms", 100)
            content_label.setText(f"{hotkey}")
            time_label.setText(f"{delay} ms")
            icon_label.setText("⌨")
        elif t == "拖拽":
            sx, sy = p.get("start_x", 0), p.get("start_y", 0)
            ex, ey = p.get("end_x", 0), p.get("end_y", 0)
            content_label.setText(f"({sx},{sy})→({ex},{ey})")
            icon_label.setText("✋")

        else:
            content_label.setText(t)
            icon_label.setText("?")

        # 加入布局
        layout.addWidget(icon_label)
        layout.addWidget(content_label, 1)   # 伸缩
        layout.addWidget(time_label)

        return container


    @staticmethod
    def thumb_widget(img_path: str, row_height: int) -> QWidget:
        """返回一个已设置好缩略图的 QLabel，高度=row_height，宽度自适应"""
        label = QLabel()
        label.setScaledContents(True)
        label.setAlignment(Qt.AlignCenter)

        # 读图并缩放到行高
        pixmap = QPixmap(img_path)
        if not pixmap.isNull():
            pixmap = pixmap.scaledToHeight(row_height, Qt.SmoothTransformation)
        label.setPixmap(pixmap)

        # 用 QWidget 包一层，方便后续扩展
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(2, 2, 2, 2)
        lay.addWidget(label)
        return w

class ATIcon:
    """
    为「Auto Tool」桌面自动办公软件生成一枚
    64×64 带毛玻璃效果、渐变背景的「AT」图标。
    """
    SIZE = 64
    _cache = {}          # 缓存，避免重复渲染

    @classmethod
    def pixmap(cls, size=SIZE) -> QPixmap:
        """返回渲染好的 QPixmap，可自由缩放"""
        if size in cls._cache:
            return cls._cache[size]

        px = QPixmap(size, size)
        px.fill(Qt.transparent)

        p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)

        # 1. 圆角矩形背景 -------------------------------------------------
        rect = QRectF(0, 0, size, size)
        radius = size * 0.18
        path = QPainterPath()
        path.addRoundedRect(rect, radius, radius)

        # 2. 渐变填充 ------------------------------------------------------
        g = QLinearGradient(QPointF(0, 0), QPointF(size, size))
        g.setColorAt(0.0, QColor("#6A11CB"))   # 紫
        g.setColorAt(1.0, QColor("#2575FC"))   # 蓝
        p.fillPath(path, QBrush(g))

        # 3. 毛玻璃：一层极低不透明度白色蒙版 -------------------------------
        blur_layer = QPainterPath()
        blur_layer.addRoundedRect(rect, radius, radius)
        p.fillPath(blur_layer, QColor(255, 255, 255, 35))

        # 4. 字母 “AT” ----------------------------------------------------
        font = QFont("Segoe UI", size * 0.32, QFont.Bold)
        p.setFont(font)
        p.setPen(QPen(Qt.white))
        p.drawText(rect, Qt.AlignCenter, "AT")

        p.end()
        cls._cache[size] = px
        return px

    @classmethod
    def icon(cls, size=SIZE) -> QIcon:
        """直接拿到 QIcon，可设给窗口、托盘、按钮等"""
        return QIcon(cls.pixmap(size))

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("AboutDialog")
        self.setWindowTitle("关于")
        self.setModal(True)
        self.resize(480, 520)

        # 根布局
        root = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        root.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)
        lay = QVBoxLayout(content)
        lay.setAlignment(Qt.AlignTop)

        # 1. 标题
        title = QLabel("自动化任务管理器")
        title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignCenter)
        lay.addWidget(title)
        # 2. 版本 + 作者 + 头像
        author_layout = QHBoxLayout()
        author_layout.setSpacing(12)

        # 头像
        self.avatar = QLabel()
        self.avatar.setFixedSize(64, 64)
        self.avatar.setObjectName("avatarLabel")
        self.load_avatar()

        # 作者信息
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.addRow("版　本：", QLabel("1.0.0"))
        form.addRow("作　者：", QLabel("B_arbarian from UESTEC"))
        author_layout.addWidget(self.avatar)
        author_layout.addLayout(form)
        author_layout.addStretch()
        lay.addLayout(author_layout)
        # 3. 联系方式（带超链接）
        link_lbl = QLabel(
            'B站主页：<a href="https://space.bilibili.com/521967044">'
            '<span style="color:#409EFF;">点击访问</span></a><br>'
            '邮　　箱：<a href="mailto:264214429@qq.com">'
            '<span style="color:#409EFF;">264214429@qq.com</span></a>'
        )
        link_lbl.setObjectName("linkLabel")
        link_lbl.setOpenExternalLinks(True)
        link_lbl.setTextInteractionFlags(Qt.TextBrowserInteraction)
        lay.addWidget(link_lbl, alignment=Qt.AlignCenter)

        # 4. 简介
        intro = QTextEdit()
        intro.setObjectName("introText")
        intro.setMaximumHeight(180)
        intro.setPlainText(
            "自动化任务管理器是一款强大的桌面自动化工具，可以帮助您自动化执行重复的计算机操作，提高工作效率。\n\n"
            "主要功能：\n"
            "• 基于图像识别的鼠标操作\n"
            "• 文本输入自动化\n"
            "• 定时任务执行\n"
            "• 详细执行日志记录\n\n"
            "感谢使用本软件！如有任何问题或建议，请通过上述联系方式与我们联系。"
        )
        lay.addWidget(intro)
        # 5. 打赏二维码
        qr_lay = QHBoxLayout()
        qr_lay.setSpacing(16)
        qr_lay.addStretch()

        self.wx_qr = QLabel()
        self.wx_qr.setObjectName("qrLabel")
        self.wx_qr.setFixedSize(160, 160)
        self.load_qr(self.wx_qr, "img/donate.png", "微信赞赏")

        self.zfb_qr = QLabel()
        self.zfb_qr.setObjectName("qrLabel")
        self.zfb_qr.setFixedSize(160, 160)
        self.load_qr(self.zfb_qr, "img/zhifubao.jpg", "支付宝打赏")

        qr_lay.addWidget(self.wx_qr)
        qr_lay.addWidget(self.zfb_qr)
        qr_lay.addStretch()
        lay.addLayout(qr_lay)

        # 6. 按钮
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok)
        btn_box.accepted.connect(self.accept)
        root.addWidget(btn_box)

        # 7. 加载样式
        self.load_qss()

    # ---------- 私有方法 ----------
    def load_avatar(self):
        avatar_path = "img/avatar.jpg"
        pixmap = QPixmap(avatar_path)
        if pixmap.isNull():
            self.avatar.setText("头像")
            return

        size = self.avatar.width()
        rounded = QPixmap(size, size)
        rounded.fill(Qt.transparent)

        painter = QPainter(rounded)
        painter.setRenderHint(QPainter.Antialiasing)
        path = QPainterPath()
        path.addRoundedRect(0, 0, size, size, size//2, size//2)
        painter.setClipPath(path)
        painter.drawPixmap(0, 0, size, size, pixmap)
        painter.end()

        self.avatar.setPixmap(rounded)


    def load_qr(self, label: QLabel, path: str, alt: str):
        path = resource_path(path)
        pixmap = QPixmap(path)
        if pixmap.isNull():
            label.setText(f"{alt}\n加载失败")
            label.setAlignment(Qt.AlignCenter)
            return
        label.setPixmap(pixmap.scaled(label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))
    def load_qss(self):
        qss_path = Path(__file__).resolve().parent.parent / "css" / "about_style.qss"
        if qss_path.exists():
            with open(qss_path, 'r', encoding='utf-8') as f:
                self.setStyleSheet(f.read())


class TaskRunner(QObject):
    task_completed = Signal(str, bool, str)
    task_progress = Signal(str, int, int)
    task_stopped = Signal(str)
    log_message = Signal(str, str)  # 新增日志信号

    def __init__(self, task_name, steps,auto_skip_image_timeout=False,timeout=10,instant_click=False,move_duration=0.1):
        super().__init__()
        self.task_name = task_name
        self.steps = steps
        self.is_running = False
        self.current_step = 0
        self.repeat_count = 0
        self.max_repeat = 1  # 默认执行1次

        self.auto_skip_image_timeout = auto_skip_image_timeout
        self.timeout = timeout  # 用户设置的超时时间

        self.instant_click = instant_click        # 是否跳过移动动画
        self.default_move_duration = move_duration  # 全局移动动画时长

        self._excel_cycle = None
        self._excel_cache = {}   # 路径->(wb, ws, rows)

    def set_repeat_count(self, count):
        self.max_repeat = count

    def execute_mouse_click(self, params):
        image_path = params.get("image_path", "")
        click_type = params.get("click_type", "左键单击")
        offset_x = params.get("offset_x", 0)
        offset_y = params.get("offset_y", 0)
        confidence = params.get("confidence", 0.8)
        timeout = self.timeout
        move_duration = params.get("move_duration", self.default_move_duration)

        if not image_path:
            if self.auto_skip_image_timeout:
                self.log_message.emit(self.task_name, "⚠️ 图片路径为空，跳过此步骤")
                return
            else:
                raise ValueError("image_path 不能为空")

        print(f"[DEBUG] 开始定位图片: {image_path}")

        def find_image_center():
            start = time.time()
            while True:
                pos = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
                if pos:
                    return pos
                if time.time() - start > timeout:
                    return None
                time.sleep(0.2)

        center = find_image_center()
        if center is None:
            if self.auto_skip_image_timeout:
                self.log_message.emit(self.task_name, f"⚠️ 在 {timeout}s 内未找到图片: {os.path.basename(image_path)}，自动跳过")
                return  # ✅ 跳过，不抛异常
            else:
                raise RuntimeError(f"在 {timeout}s 内未找到图片: {image_path}")

        target_x = center.x + offset_x
        target_y = center.y + offset_y

        if not self.instant_click:
            try:
                pyautogui.moveTo(target_x, target_y, duration=move_duration)
            except Exception as e:
                if self.auto_skip_image_timeout:
                    self.log_message.emit(self.task_name, f"⚠️ 鼠标移动失败，跳过: {e}")
                    return
                raise
        else:
            pyautogui.moveTo(target_x, target_y, duration=0)  # 瞬移

        click_map = {
            "左键单击": pyautogui.click,
            "左键双击": pyautogui.doubleClick,
            "右键单击": pyautogui.rightClick,
            "中键单击": pyautogui.middleClick,
        }
        if click_type not in click_map:
            if self.auto_skip_image_timeout:
                self.log_message.emit(self.task_name, f"⚠️ 不支持的点击类型: {click_type}，跳过")
                return
            else:
                raise ValueError(f"不支持的 click_type: {click_type}")

        click_map[click_type](target_x, target_y)
        print(f"[DEBUG] 已完成 {click_type} 操作")
    def run(self):
        self.is_running = True
        self.current_step = 0
        total_steps = len(self.steps)
        self.repeat_count = 0

        self.log_message.emit(self.task_name, f"🚀 开始执行任务: {self.task_name}, 共 {total_steps} 个步骤")

        try:
            while self.repeat_count < self.max_repeat and self.is_running:
                self.repeat_count += 1
                if self.max_repeat > 1:
                    self.log_message.emit(self.task_name, f"🔄 第 {self.repeat_count}/{self.max_repeat} 次执行")

                for i, step in enumerate(self.steps):
                    if not self.is_running:
                        self.log_message.emit(self.task_name, "⏹️ 任务被中断")
                        break

                    self.current_step = i
                    self.task_progress.emit(self.task_name, i + 1, total_steps)

                    # 执行步骤
                    step_type = step.get("type", "")
                    params = step.get("params", {})
                    delay = step.get("delay", 0)

                    # 简化日志显示
                    if step_type == "鼠标点击":
                        image_name = os.path.basename(params.get("image_path", ""))
                        click_type = params.get("click_type", "左键单击")
                        self.log_message.emit(self.task_name, f"📝 执行步骤 {i + 1}/{total_steps}: {step_type}")
                        self.log_message.emit(self.task_name, f"🖼️ 图片: {image_name}, 点击类型: {click_type}")
                    else:
                        self.log_message.emit(self.task_name, f"📝 执行步骤 {i + 1}/{total_steps}: {step_type}")
                        self.log_message.emit(self.task_name, f"⚙️ 参数: {json.dumps(params, ensure_ascii=False)}")

                    if step_type == "鼠标点击":
                        self.execute_mouse_click(params)
                    elif step_type == "文本输入":
                        self.execute_keyboard_input(params)
                    elif step_type == "等待":
                        self.execute_wait(params)
                    elif step_type == "截图":
                        self.execute_screenshot(params)
                    elif step_type == "拖拽":
                        self.execute_drag(params)
                    elif step_type == "鼠标滚轮":
                        self.execute_mouse_scroll(params)
                    elif step_type == "键盘热键":
                        self.execute_hotkey(params)
                    else:
                        self.log_message.emit(self.task_name, f"⚠️ 未知步骤类型: {step_type}")

                    # 步骤间延时
                    if delay > 0:
                        self.log_message.emit(self.task_name, f"⏱️ 步骤延时: {delay}秒")
                        time.sleep(delay)

                if not self.is_running:
                    break

            success = self.is_running
            message = "✅ 任务完成" if success else "⏹️ 任务被中断"
            self.log_message.emit(self.task_name, message)
            self.task_completed.emit(self.task_name, success, message)
        except Exception as e:
            error_msg = f"❌ 任务执行出错: {str(e)}"
            self.log_message.emit(self.task_name, error_msg)
            self.task_completed.emit(self.task_name, False, error_msg)
        finally:
            self.is_running = False

    def stop(self):
        self.log_message.emit(self.task_name, "⏹️ 停止任务")
        self.is_running = False
        self.task_stopped.emit(self.task_name)
    def chinese_qixi(self,year: int) -> date:
        """
        计算指定年份的七夕节（农历七月初七）的公历日期
        使用近似算法，误差在±1天内

        Args:
            year: 要计算的年份

        Returns:
            该年份七夕节的公历日期
        """
        # 扩展的年份对照表（2000-2030年）
        table = {
            2000: date(2000, 8, 6), 2001: date(2001, 8, 25), 2002: date(2002, 8, 15),
            2003: date(2003, 8, 4), 2004: date(2004, 8, 22), 2005: date(2005, 8, 11),
            2006: date(2006, 7, 31), 2007: date(2007, 8, 19), 2008: date(2008, 8, 7),
            2009: date(2009, 8, 26), 2010: date(2010, 8, 16), 2011: date(2011, 8, 6),
            2012: date(2012, 8, 23), 2013: date(2013, 8, 13), 2014: date(2014, 8, 2),
            2015: date(2015, 8, 20), 2016: date(2016, 8, 9), 2017: date(2017, 8, 28),
            2018: date(2018, 8, 17), 2019: date(2019, 8, 7), 2020: date(2020, 8, 25),
            2021: date(2021, 8, 14), 2022: date(2022, 8, 4), 2023: date(2023, 8, 22),
            2024: date(2024, 8, 10), 2025: date(2025, 8, 1), 2026: date(2026, 8, 19),
            2027: date(2027, 8, 8), 2028: date(2028, 7, 28), 2029: date(2029, 8, 16),
            2030: date(2030, 8, 5)
        }

        # 如果在已知年份范围内，直接返回表中日期
        if year in table:
            return table[year]

        # 对于表外的年份，使用近似算法计算
        # 基础年份选择2023年，七夕日期为8月22日
        base_year = 2023
        base_date = date(base_year, 8, 22)

        # 计算与基础年份的差异（考虑农历年的平均长度）
        year_diff = year - base_year
        # 农历年平均长度约为29.53天×12个月 = 354.36天
        days_diff = round(year_diff * 354.36 - year_diff * 365.25)

        # 计算预估日期
        estimated_date = base_date + timedelta(days=days_diff)

        # 调整到8月附近（七夕通常在7月底到8月底之间）
        if estimated_date.month < 7:
            estimated_date += timedelta(days=30)
        elif estimated_date.month > 9:
            estimated_date -= timedelta(days=30)

        return estimated_date

    # def execute_mouse_click(self, params):
    #     AutoClicker().execute_mouse_click(params)
    #     self.log_message.emit(self.task_name, "🖱️ 鼠标点击操作完成")

    def execute_mouse_scroll(self, params):
        direction = params.get("direction", "向下滚动")
        clicks = params.get("clicks", 3)

        self.log_message.emit(self.task_name,
                              f"🖱 鼠标滚轮 {direction} {clicks} 格（当前位置）")

        try:
            scroll_amount = clicks * 120 if direction == "向下滚动" else -clicks * 120
            pyautogui.scroll(scroll_amount)
            self.log_message.emit(self.task_name, "✅ 滚轮完成")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ 滚轮出错: {str(e)}")
            raise

    def execute_hotkey(self, params):
        hotkey = params.get("hotkey", "ctrl+c")
        delay = params.get("delay_ms", 100)

        self.log_message.emit(self.task_name, f"⌨ 热键 {hotkey.upper()} 执行")

        try:
            pyautogui.hotkey(*hotkey.split("+"))
            if delay > 0:
                time.sleep(delay / 1000.0)
            self.log_message.emit(self.task_name, "✅ 热键完成")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ 热键出错: {str(e)}")
            raise

    def execute_keyboard_input(self, params):
        # 1. 纯文本优先
        text = params.get("text", "").strip()
        if not text:
            # 2. 动态纪念日文案
            love_str = params.get("love_date")
            if love_str:
                love_dt = datetime.fromisoformat(love_str)
                today = date.today()
                today_1314 = datetime.combine(today, time(13, 14))

                delta = today_1314 - love_dt
                days, sec = delta.days, delta.seconds
                hours, rem = divmod(sec, 3600)
                minutes, secs = divmod(rem, 60)
                duration = f"{days}天{hours}时{minutes}分{secs}秒"

                year_start = datetime(today.year, 1, 1, 13, 14)
                count = (today_1314 - year_start).days + 1

                # 特殊节日
                is_xmas = (love_dt.month, love_dt.day) == (12, 25)
                special = ""
                if today == date(today.year, 12, 25):
                    special = "\n圣诞快乐，Merry Christmas！"
                elif today == date(today.year, 2, 14):
                    special = "\n情人节快乐！"
                elif today == self.chinese_qixi(today.year):
                    special = "\n七夕快乐，鹊桥相会！"

                today_str = today.strftime("%Y年%m月%d日")
                if is_xmas:
                    text = (f"今天是{today_str}第{count}个1314，我们已相恋{duration}，"
                            f"从圣诞夜一直走到今天，未来也要一起闪耀！🎄❤{special}")
                else:
                    text = (f"今天是{today_str}第{count}个1314，"
                            f"我们已经相恋了{duration}，爱你❤{special}")
            else:
                # 3. 否则从 Excel 取
                excel_path = params.get("excel_path", "").strip()
                if not excel_path or not os.path.isfile(excel_path):
                    raise FileNotFoundError("未指定或找不到 Excel 文件")

                sheet_id = params.get("sheet", "0")
                col_index = int(params.get("col", 0))
                mode = params.get("mode", "顺序")

                # === 关键：使用 (文件, 表, 列) 作为缓存键 ===
                cache_key = (excel_path, str(sheet_id), col_index)

                # 1. 检查是否已缓存 workbook（避免重复打开）
                wb_cache_key = excel_path
                if wb_cache_key not in self._excel_cache:
                    wb = openpyxl.load_workbook(excel_path, data_only=True)
                    try:
                        ws = wb[int(sheet_id)] if str(sheet_id).isdigit() else wb[sheet_id]
                    except Exception:
                        ws = wb.worksheets[0]
                    rows = list(ws.iter_rows(values_only=True))
                    self._excel_cache[wb_cache_key] = (wb, ws, rows)
                _, _, rows = self._excel_cache[wb_cache_key]

                if not rows:
                    raise ValueError("Excel 表无数据")

                cells = [row[col_index] for row in rows if len(row) > col_index and row[col_index] is not None]
                if not cells:
                    raise ValueError("指定列为空")

                # === 2. 使用 cache_key 管理 cycle ===
                if mode == "顺序":
                    # 初始化类变量（如果还没创建）
                    if not hasattr(self, '_excel_cycle_dict'):
                        self._excel_cycle_dict = {}

                    # 如果该 (文件, 表, 列) 组合没有 cycle，创建一个
                    if cache_key not in self._excel_cycle_dict:
                        self._excel_cycle_dict[cache_key] = itertools.cycle(cells)

                    text = next(self._excel_cycle_dict[cache_key])

                else:  # 随机
                    text = random.choice(cells)

        self._send_text(str(text))

    def _send_text(self, text: str):
        """真正执行文本输入的公共逻辑"""
        self.log_message.emit(self.task_name, f"⌨️ 文本输入: '{text}'")
        try:
            import pyperclip, pyautogui, time
            pyperclip.copy(text)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            self.log_message.emit(self.task_name, "✅ 文本输入完成")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ 文本输入出错: {str(e)}")
            raise
    def execute_wait(self, params):
        seconds = params.get("seconds", 0)
        if seconds > 0:
            self.log_message.emit(self.task_name, f"⏱️ 等待 {seconds}秒")
            try:
                time.sleep(seconds)
            except Exception as e:
                self.log_message.emit(self.task_name, f"❌ 等待操作出错: {str(e)}")
                raise

    def execute_screenshot(self, params):
        save_path = params.get("save_path", "")
        region = params.get("region", None)

        self.log_message.emit(self.task_name, f"📸 截图保存到: {save_path}")

        try:
            if region:
                x, y, width, height = region
                self.log_message.emit(self.task_name, f"🖼️ 截图区域: x={x}, y={y}, width={width}, height={height}")
                screenshot = pyautogui.screenshot(region=(x, y, width, height))
            else:
                self.log_message.emit(self.task_name, "🖼️ 全屏截图")
                screenshot = pyautogui.screenshot()

            screenshot.save(save_path)
            self.log_message.emit(self.task_name, "✅ 截图保存成功")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ 截图操作出错: {str(e)}")
            raise

    def execute_drag(self, params):
        start_x = params.get("start_x", 0)
        start_y = params.get("start_y", 0)
        end_x = params.get("end_x", 0)
        end_y = params.get("end_y", 0)
        duration = params.get("duration", 1.0)

        self.log_message.emit(self.task_name,
                              f"↔️ 从 ({start_x}, {start_y}) 拖拽到 ({end_x}, {end_y}), 时长: {duration}秒")

        try:
            pyautogui.moveTo(start_x, start_y)
            pyautogui.dragTo(end_x, end_y, duration=duration, button='left')
            self.log_message.emit(self.task_name, "✅ 拖拽操作完成")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ 拖拽操作出错: {str(e)}")
            raise


class StepConfigDialog(QDialog):
    def __init__(self, step_data=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("配置步骤")
        self.setMinimumWidth(500)
        self.setWindowIcon(ATIcon.icon())

        layout = QVBoxLayout(self)

        # 步骤类型选择
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("步骤类型:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(["鼠标点击", "文本输入", "等待", "截图", "拖拽", "鼠标滚轮", '键盘热键'])
        type_layout.addWidget(self.type_combo)
        layout.addLayout(type_layout)

        # 参数配置区域
        self.params_stack = QWidget()
        self.params_layout = QVBoxLayout(self.params_stack)
        self.params_layout.setContentsMargins(0, 10, 0, 0)

        # 创建不同步骤类型的参数面板
        self.mouse_click_panel = self.create_mouse_click_panel()
        self.keyboard_input_panel = self.create_keyboard_input_panel()
        self.wait_panel = self.create_wait_panel()
        self.screenshot_panel = self.create_screenshot_panel()
        self.drag_panel = self.create_drag_panel()
        self.scroll_panel = self.create_mouse_scroll_panel()
        self.hot_keyboard_panel = self.create_hot_keyboard_panel()


        # 添加到堆栈
        self.params_layout.addWidget(self.mouse_click_panel)
        self.params_layout.addWidget(self.keyboard_input_panel)
        self.params_layout.addWidget(self.wait_panel)
        self.params_layout.addWidget(self.screenshot_panel)
        self.params_layout.addWidget(self.drag_panel)
        self.params_layout.addWidget(self.scroll_panel)
        self.params_layout.addWidget(self.hot_keyboard_panel)

        layout.addWidget(self.params_stack)

        # 延时设置
        delay_layout = QHBoxLayout()
        delay_layout.addWidget(QLabel("步骤执行后延时(秒):"))
        self.delay_spin = QSpinBox()
        self.delay_spin.setRange(0, 3600)
        self.delay_spin.setValue(0)
        delay_layout.addWidget(self.delay_spin)
        layout.addLayout(delay_layout)

        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        # 连接信号
        self.type_combo.currentIndexChanged.connect(self.update_params_panel)

        # 初始化UI
        self.update_params_panel()

        # 如果有传入步骤数据，填充表单
        if step_data:
            self.load_step_data(step_data)

    # 1. 新增极简滚轮面板
    def create_mouse_scroll_panel(self):
        panel = QWidget()
        layout = QFormLayout(panel)

        # 方向
        self.scroll_direction_combo = QComboBox()
        self.scroll_direction_combo.addItems(["向上滚动", "向下滚动"])
        layout.addRow("滚动方向:", self.scroll_direction_combo)

        # 格数
        self.scroll_clicks_spin = QSpinBox()
        self.scroll_clicks_spin.setRange(1, 100)
        self.scroll_clicks_spin.setValue(3)
        layout.addRow("滚动格数:", self.scroll_clicks_spin)

        return panel

    def create_hot_keyboard_panel(self):
        panel = QWidget()
        layout = QFormLayout(panel)

        # 热键选择下拉框
        self.hotkey_combo = QComboBox()
        self.hotkey_combo.addItems([
            "Ctrl+A  全选",
            "Ctrl+C  复制",
            "Ctrl+V  粘贴",
            "Ctrl+X  剪切",
            "Ctrl+Z  撤销",
            "Ctrl+Y  重做",
            "Ctrl+S  保存",
            "Ctrl+F  查找"
        ])
        layout.addRow("热键:", self.hotkey_combo)

        # 额外延迟（ms）
        self.hotkey_delay_spin = QSpinBox()
        self.hotkey_delay_spin.setRange(0, 5000)
        self.hotkey_delay_spin.setValue(100)
        self.hotkey_delay_spin.setSuffix(" ms")
        layout.addRow("执行后延时:", self.hotkey_delay_spin)

        return panel

    def capture_region(self):
        parent = self.parent()
        parent.hide()
        self.hide()

        self.overlay = RegionCaptureOverlay()
        self.overlay.finished.connect(self.on_region_done)
        self.overlay.show()

    def on_region_done(self, geo: QRect):

        # 先关闭覆盖层窗口（关键修复！）
        if hasattr(self, 'overlay') and self.overlay is not None:
            self.overlay.close()  # 或者 self.overlay.hide()
            self.overlay.deleteLater()  # 可选，帮助 Qt 彻底清理
            self.overlay = None  # 可选，避免野指针

        parent = self.parent()
        if geo.isNull():
            print("❌ 用户未选择有效区域")
            parent.show()
            self.show()
            return


        pixmap = QApplication.primaryScreen().grabWindow(
            0, geo.x(), geo.y(), geo.width(), geo.height()
        )

        img_dir = os.path.join(os.getcwd(), "img")
        # img_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "img")
        os.makedirs(img_dir, exist_ok=True)
        file_name = datetime.now().strftime("%Y%m%d_%H%M%S") + ".png"
        file_path = os.path.join(img_dir, file_name)

        if pixmap.save(file_path, "PNG"):
            self.image_path_edit.setText(file_path)
            QMessageBox.information(self, "框选截图成功", f"已保存：{file_name}")
            # 直接调用 add_step_to_table
            step_data = self.get_step_data()
            parent.add_step_to_table(step_data)
            # 添加到当前任务配置
            if parent.current_task and parent.current_task in parent.tasks:
                parent.tasks[parent.current_task]["steps"].append(step_data)


        else:
            QMessageBox.warning(self, "失败", "截图保存失败！")

        parent.show()
        self.show()


    def create_mouse_click_panel(self):
        panel = QWidget()
        layout = QGridLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # 图片路径
        layout.addWidget(QLabel("图片路径:"), 0, 0)
        self.image_path_edit = QLineEdit()
        layout.addWidget(self.image_path_edit, 0, 1)
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self.browse_image)
        layout.addWidget(browse_btn, 0, 2)

        # >>> 新增：一键录制按钮
        record_btn = QPushButton("框选截图")
        record_btn.clicked.connect(self.capture_region)
        layout.addWidget(record_btn, 0, 3)

        # 点击类型
        layout.addWidget(QLabel("点击类型:"), 1, 0)
        self.click_type_combo = QComboBox()
        self.click_type_combo.addItems(["左键单击", "左键双击", "右键单击", "中键单击"])
        layout.addWidget(self.click_type_combo, 1, 1, 1, 2)

        # 偏移量
        layout.addWidget(QLabel("X偏移:"), 2, 0)
        self.offset_x_spin = QSpinBox()
        self.offset_x_spin.setRange(-1000, 1000)
        layout.addWidget(self.offset_x_spin, 2, 1)

        layout.addWidget(QLabel("Y偏移:"), 2, 2)
        self.offset_y_spin = QSpinBox()
        self.offset_y_spin.setRange(-1000, 1000)
        layout.addWidget(self.offset_y_spin, 2, 3)

        # 识别设置
        layout.addWidget(QLabel("识别精度(0-1):"), 3, 0)
        self.confidence_spin = QDoubleSpinBox()
        self.confidence_spin.setRange(0.5, 1.0)
        self.confidence_spin.setValue(0.8)
        self.confidence_spin.setSingleStep(0.05)
        layout.addWidget(self.confidence_spin, 3, 1)

        layout.addWidget(QLabel("超时时间(秒):"), 3, 2)
        self.timeout_spin = QDoubleSpinBox()
        self.timeout_spin.setRange(0.1, 60)
        self.timeout_spin.setSingleStep(0.1)
        self.timeout_spin.setValue(1.0)
        self.timeout_spin.setDecimals(1)
        layout.addWidget(self.timeout_spin, 3, 3)

        return panel


    def generate_love_text(self):
        from datetime import datetime, date, time
        love_dt = self.love_datetime_edit.dateTime().toPython()  # 用户选的时刻
        today = date.today()
        today_1314 = datetime.combine(today, time(13, 14))  # 今天 13:14

        # 相恋时长（精确到秒）
        delta = today_1314 - love_dt
        days = delta.days
        sec = delta.seconds
        hours, rem = divmod(sec, 3600)
        minutes, secs = divmod(rem, 60)
        duration = f"{days}天{hours}时{minutes}分{secs}秒"

        # 今年第几个 13:14
        year_start_1314 = datetime(today.year, 1, 1, 13, 14)
        count = (today_1314 - year_start_1314).days + 1

        # 特殊节日
        year = today.year
        is_xmas = (love_dt.month, love_dt.day) == (12, 25)
        special = None
        if is_xmas:
            special = "我们的爱情从圣诞夜点亮，愿它像圣诞树一样永远闪耀！"
        elif today == date(year, 2, 14):
            special = "情人节快乐！"
        elif today == self.chinese_qixi(year):
            special = "七夕快乐，鹊桥相会！"
        elif today == date(year, 12, 25):
            special = "圣诞快乐，Merry Christmas！"

        today_str = f"{today.year}年{today.month}月{today.day}日"
        if is_xmas:
            text = (f"今天是{today_str}第{count}个1314，"
                    f"我们已相恋{duration}，"
                    f"从圣诞夜一直走到今天，未来也要一起闪耀！🎄❤")
        else:
            text = (f"今天是{today_str}第{count}个1314，"
                    f"我们已经相恋了{duration}，爱你❤")
            if special:
                text += f"\n{special}"

        self.text_edit.setPlainText(text)
    def create_keyboard_input_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # 1. 原始文本输入（多行）
        layout.addWidget(QLabel("输入文本（留空则用 Excel 或纪念日）:"))

        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlaceholderText("在此输入固定文本...\n留空则自动从 Excel 或纪念日生成内容")
        self.text_edit.setMaximumHeight(80)
        self.text_edit.setLineWrapMode(QPlainTextEdit.WidgetWidth)

        layout.addWidget(self.text_edit)


        # -------- 新增纪念日区域 --------
        love_group = QWidget()
        h_layout = QHBoxLayout(love_group)  # 横向布局

        # 1. 启用复选框
        self.use_love_checkbox = QCheckBox("启用纪念日")
        self.use_love_checkbox.setChecked(False)  # 默认不启用
        h_layout.addWidget(self.use_love_checkbox)

        # 2. 标签
        h_layout.addWidget(QLabel("时间:"))

        # 3. 时间选择器
        self.love_datetime_edit = QDateTimeEdit()
        self.love_datetime_edit.setCalendarPopup(True)
        self.love_datetime_edit.setDisplayFormat("yyyy-MM-dd hh:mm:ss")
        self.love_datetime_edit.setDateTime(QDateTime(QDate(2022, 12, 25), QTime(7, 0, 0)))
        # 可选：默认禁用，直到 checkbox 勾选
        self.love_datetime_edit.setEnabled(False)
        self.use_love_checkbox.toggled.connect(self.love_datetime_edit.setEnabled)

        h_layout.addWidget(self.love_datetime_edit)

        # 4. 生成按钮
        gen_btn = QPushButton("生成文案")
        gen_btn.clicked.connect(self.generate_love_text)
        h_layout.addWidget(gen_btn)

        # 可选：设置拉伸，防止挤压
        h_layout.addStretch()

        # 将 group 添加到主 layout
        layout.addWidget(love_group)


        # 3. Excel 区域
        excel_group = QWidget()
        g = QVBoxLayout(excel_group)

        # 文件选择
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("Excel 文件:"))
        self.excel_path_edit = QLineEdit()
        btn = QPushButton("浏览")
        btn.clicked.connect(lambda: self.excel_path_edit.setText(
            QFileDialog.getOpenFileName(filter="*.xlsx")[0]))
        h1.addWidget(self.excel_path_edit)
        h1.addWidget(btn)
        g.addLayout(h1)

        # 工作表
        h2 = QHBoxLayout()
        h2.addWidget(QLabel("工作表(名称或序号):"))
        self.sheet_edit = QLineEdit("0")
        h2.addWidget(self.sheet_edit)
        g.addLayout(h2)

        # 列
        h3 = QHBoxLayout()
        h3.addWidget(QLabel("列(首列=0):"))
        self.col_spin = QSpinBox()
        self.col_spin.setValue(0)
        h3.addWidget(self.col_spin)
        g.addLayout(h3)

        # 读取模式
        h4 = QHBoxLayout()
        h4.addWidget(QLabel("读取模式:"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["顺序", "随机"])
        h4.addWidget(self.mode_combo)
        g.addLayout(h4)

        layout.addWidget(excel_group)
        return panel


    def create_wait_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)

        layout.addWidget(QLabel("等待时间(秒):"))
        self.wait_spin = QSpinBox()
        self.wait_spin.setRange(1, 3600)
        self.wait_spin.setValue(5)
        layout.addWidget(self.wait_spin)

        return panel

    def create_screenshot_panel(self):
        panel = QWidget()
        layout = QGridLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # 保存路径
        layout.addWidget(QLabel("保存路径:"), 0, 0)
        self.screenshot_path_edit = QLineEdit()
        layout.addWidget(self.screenshot_path_edit, 0, 1)
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self.browse_save_path)
        layout.addWidget(browse_btn, 0, 2)

        # 截图区域
        layout.addWidget(QLabel("截图区域(可选):"), 1, 0)

        layout.addWidget(QLabel("X:"), 2, 0)
        self.screenshot_x_spin = QSpinBox()
        self.screenshot_x_spin.setRange(0, 10000)
        layout.addWidget(self.screenshot_x_spin, 2, 1)

        layout.addWidget(QLabel("Y:"), 2, 2)
        self.screenshot_y_spin = QSpinBox()
        self.screenshot_y_spin.setRange(0, 10000)
        layout.addWidget(self.screenshot_y_spin, 2, 3)

        layout.addWidget(QLabel("宽度:"), 3, 0)
        self.screenshot_width_spin = QSpinBox()
        self.screenshot_width_spin.setRange(1, 10000)
        self.screenshot_width_spin.setValue(800)
        layout.addWidget(self.screenshot_width_spin, 3, 1)

        layout.addWidget(QLabel("高度:"), 3, 2)
        self.screenshot_height_spin = QSpinBox()
        self.screenshot_height_spin.setRange(1, 10000)
        self.screenshot_height_spin.setValue(600)
        layout.addWidget(self.screenshot_height_spin, 3, 3)

        return panel

    def create_drag_panel(self):
        panel = QWidget()
        layout = QGridLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # 起点坐标
        layout.addWidget(QLabel("起点坐标:"), 0, 0)

        layout.addWidget(QLabel("X:"), 1, 0)
        self.drag_start_x_spin = QSpinBox()
        self.drag_start_x_spin.setRange(0, 10000)
        layout.addWidget(self.drag_start_x_spin, 1, 1)

        layout.addWidget(QLabel("Y:"), 1, 2)
        self.drag_start_y_spin = QSpinBox()
        self.drag_start_y_spin.setRange(0, 10000)
        layout.addWidget(self.drag_start_y_spin, 1, 3)

        # 终点坐标
        layout.addWidget(QLabel("终点坐标:"), 2, 0)

        layout.addWidget(QLabel("X:"), 3, 0)
        self.drag_end_x_spin = QSpinBox()
        self.drag_end_x_spin.setRange(0, 10000)
        layout.addWidget(self.drag_end_x_spin, 3, 1)

        layout.addWidget(QLabel("Y:"), 3, 2)
        self.drag_end_y_spin = QSpinBox()
        self.drag_end_y_spin.setRange(0, 10000)
        layout.addWidget(self.drag_end_y_spin, 3, 3)

        # 拖拽时间
        layout.addWidget(QLabel("拖拽时间(秒):"), 4, 0)
        self.drag_duration_spin = QDoubleSpinBox()
        self.drag_duration_spin.setRange(0.1, 10.0)
        self.drag_duration_spin.setValue(1.0)
        self.drag_duration_spin.setSingleStep(0.1)
        layout.addWidget(self.drag_duration_spin, 4, 1)

        return panel

    def browse_image(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择图片", "", "图片文件 (*.png *.jpg *.bmp)"
        )
        if file_path:
            self.image_path_edit.setText(file_path)

    def browse_save_path(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存截图", "", "PNG图片 (*.png)"
        )
        if file_path:
            if not file_path.lower().endswith('.png'):
                file_path += '.png'
            self.screenshot_path_edit.setText(file_path)

    def update_params_panel(self):
        # 隐藏所有面板
        for i in range(self.params_layout.count()):
            widget = self.params_layout.itemAt(i).widget()
            if widget:
                widget.hide()

        # 显示当前选中的面板
        step_type = self.type_combo.currentText()
        if step_type == "鼠标点击":
            self.mouse_click_panel.show()
        elif step_type == "文本输入":
            self.keyboard_input_panel.show()
        elif step_type == "等待":
            self.wait_panel.show()
        elif step_type == "截图":
            self.screenshot_panel.show()
        elif step_type == "拖拽":
            self.drag_panel.show()
        elif step_type == "鼠标滚轮":
            self.scroll_panel.show()
        elif step_type == "键盘热键":
            self.hot_keyboard_panel.show()

    def load_step_data(self, step_data):
        step_type = step_data.get("type", "")
        self.type_combo.setCurrentText(step_type)

        # 设置延时
        self.delay_spin.setValue(step_data.get("delay", 0))

        # 设置参数
        params = step_data.get("params", {})
        if step_type == "鼠标点击":
            self.image_path_edit.setText(params.get("image_path", ""))
            self.click_type_combo.setCurrentText(params.get("click_type", "左键单击"))
            self.offset_x_spin.setValue(params.get("offset_x", 0))
            self.offset_y_spin.setValue(params.get("offset_y", 0))
            self.confidence_spin.setValue(params.get("confidence", 0.8))
            self.timeout_spin.setValue(params.get("timeout", 10))
        elif step_type == "文本输入":
            self.text_edit.setPlainText(params.get("text", ""))
            self.excel_path_edit.setText(params.get("excel_path", ""))
            self.sheet_edit.setText(str(params.get("sheet", "0")))
            self.col_spin.setValue(int(params.get("col", 0)))
            # 确保下拉框里能找到对应文本
            mode = params.get("mode", "顺序")
            if mode in ["顺序", "随机"]:
                self.mode_combo.setCurrentText(mode)
            else:
                self.mode_combo.setCurrentIndex(0)  # 默认顺序
        elif step_type == "等待":
            self.wait_spin.setValue(params.get("seconds", 5))
        elif step_type == "截图":
            self.screenshot_path_edit.setText(params.get("save_path", ""))
            region = params.get("region", [0, 0, 0, 0])
            if len(region) == 4:
                self.screenshot_x_spin.setValue(region[0])
                self.screenshot_y_spin.setValue(region[1])
                self.screenshot_width_spin.setValue(region[2])
                self.screenshot_height_spin.setValue(region[3])
        elif step_type == "拖拽":
            self.drag_start_x_spin.setValue(params.get("start_x", 0))
            self.drag_start_y_spin.setValue(params.get("start_y", 0))
            self.drag_end_x_spin.setValue(params.get("end_x", 0))
            self.drag_end_y_spin.setValue(params.get("end_y", 0))
            self.drag_duration_spin.setValue(params.get("duration", 1.0))
        elif step_type == "鼠标滚轮":
            self.scroll_direction_combo.setCurrentText(params.get("direction", "向下滚动"))
            self.scroll_clicks_spin.setValue(params.get("clicks", 3))
        elif step_type == "键盘热键":
            hotkey_map = {
                "ctrl+a": "Ctrl+A  全选",
                "ctrl+c": "Ctrl+C  复制",
                "ctrl+v": "Ctrl+V  粘贴",
                "ctrl+x": "Ctrl+X  剪切",
                "ctrl+z": "Ctrl+Z  撤销",
                "ctrl+y": "Ctrl+Y  重做",
                "ctrl+s": "Ctrl+S  保存",
                "ctrl+f": "Ctrl+F  查找"
            }
            key_str = params.get("hotkey", "").lower()
            self.hotkey_combo.setCurrentText(hotkey_map.get(key_str, "Ctrl+C  复制"))
            self.hotkey_delay_spin.setValue(params.get("delay_ms", 100))

    def get_step_data(self):
        step_type = self.type_combo.currentText()
        params = {}

        if step_type == "鼠标点击":
            params = {
                "image_path": self.image_path_edit.text(),
                "click_type": self.click_type_combo.currentText(),
                "offset_x": self.offset_x_spin.value(),
                "offset_y": self.offset_y_spin.value(),
                "confidence": self.confidence_spin.value(),
                "timeout": self.timeout_spin.value()
            }
        elif step_type == "文本输入":
            use_love = self.use_love_checkbox.isChecked()
            love_date_str = ""
            if use_love:
                love_date_str = self.love_datetime_edit.dateTime().toPython().isoformat()

            params = {
                "text": self.text_edit.toPlainText().strip(),
                "excel_path": self.excel_path_edit.text().strip(),
                "sheet": self.sheet_edit.text().strip(),
                "col": self.col_spin.value(),
                "mode": self.mode_combo.currentText(),
                "love_date": love_date_str  # 只有启用时才传值
            }
        elif step_type == "等待":
            params = {
                "seconds": self.wait_spin.value()
            }
        elif step_type == "截图":
            params = {
                "save_path": self.screenshot_path_edit.text(),
                "region": [
                    self.screenshot_x_spin.value(),
                    self.screenshot_y_spin.value(),
                    self.screenshot_width_spin.value(),
                    self.screenshot_height_spin.value()
                ]
            }
        elif step_type == "拖拽":
            params = {
                "start_x": self.drag_start_x_spin.value(),
                "start_y": self.drag_start_y_spin.value(),
                "end_x": self.drag_end_x_spin.value(),
                "end_y": self.drag_end_y_spin.value(),
                "duration": self.drag_duration_spin.value()
            }
        elif step_type == "鼠标滚轮":
            params = {
                "direction": self.scroll_direction_combo.currentText(),
                "clicks": self.scroll_clicks_spin.value()
            }
        elif step_type == "键盘热键":
            hotkey_text = self.hotkey_combo.currentText()  # 例如 "Ctrl+C  复制"
            key_only = hotkey_text.split()[0]  # 取 "Ctrl+C"
            params = {
                "hotkey": key_only.lower(),  # 统一存小写，如 "ctrl+c"
                "delay_ms": self.hotkey_delay_spin.value()
            }
        params["step_time"] = datetime.now().strftime("%H:%M:%S")
        return {
            "type": step_type,
            "params": params,
            "delay": self.delay_spin.value()
        }

    def chinese_qixi(self,year: int) -> date:
        """
        计算指定年份的七夕节（农历七月初七）的公历日期
        使用近似算法，误差在±1天内

        Args:
            year: 要计算的年份

        Returns:
            该年份七夕节的公历日期
        """
        # 扩展的年份对照表（2000-2030年）
        table = {
            2000: date(2000, 8, 6), 2001: date(2001, 8, 25), 2002: date(2002, 8, 15),
            2003: date(2003, 8, 4), 2004: date(2004, 8, 22), 2005: date(2005, 8, 11),
            2006: date(2006, 7, 31), 2007: date(2007, 8, 19), 2008: date(2008, 8, 7),
            2009: date(2009, 8, 26), 2010: date(2010, 8, 16), 2011: date(2011, 8, 6),
            2012: date(2012, 8, 23), 2013: date(2013, 8, 13), 2014: date(2014, 8, 2),
            2015: date(2015, 8, 20), 2016: date(2016, 8, 9), 2017: date(2017, 8, 28),
            2018: date(2018, 8, 17), 2019: date(2019, 8, 7), 2020: date(2020, 8, 25),
            2021: date(2021, 8, 14), 2022: date(2022, 8, 4), 2023: date(2023, 8, 22),
            2024: date(2024, 8, 10), 2025: date(2025, 8, 1), 2026: date(2026, 8, 19),
            2027: date(2027, 8, 8), 2028: date(2028, 7, 28), 2029: date(2029, 8, 16),
            2030: date(2030, 8, 5)
        }

        # 如果在已知年份范围内，直接返回表中日期
        if year in table:
            return table[year]

        # 对于表外的年份，使用近似算法计算
        # 基础年份选择2023年，七夕日期为8月22日
        base_year = 2023
        base_date = date(base_year, 8, 22)

        # 计算与基础年份的差异（考虑农历年的平均长度）
        year_diff = year - base_year
        # 农历年平均长度约为29.53天×12个月 = 354.36天
        days_diff = round(year_diff * 354.36 - year_diff * 365.25)

        # 计算预估日期
        estimated_date = base_date + timedelta(days=days_diff)

        # 调整到8月附近（七夕通常在7月底到8月底之间）
        if estimated_date.month < 7:
            estimated_date += timedelta(days=30)
        elif estimated_date.month > 9:
            estimated_date -= timedelta(days=30)

        return estimated_date


class TaskItemWidget(QWidget):
    def __init__(self, name, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.task_name = name
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 5, 10, 5)
        layout.setSpacing(8)

        # 任务名称标签 - 设置为透明
        self.name_label = QLabel(name)
        self.name_label.setFont(QFont("Arial", 10, QFont.Medium))
        self.name_label.setMinimumWidth(150)
        self.name_label.setStyleSheet("background: transparent;")  # 设置透明背景

        # 状态标签
        self.status_label = QLabel("已停止")
        self.status_label.setFont(QFont("Arial", 9))
        self.status_label.setStyleSheet("background: transparent;")  # 设置透明背景

        # 操作按钮 - 添加emoji
        self.start_btn = QPushButton("▶️")
        self.start_btn.setToolTip("开始任务")
        self.start_btn.setFixedSize(28, 28)

        self.stop_btn = QPushButton("⏹️")
        self.stop_btn.setToolTip("停止任务")
        self.stop_btn.setFixedSize(28, 28)
        # self.stop_btn.setEnabled(False)

        self.delete_btn = QPushButton("🗑️")
        self.delete_btn.setToolTip("删除任务")
        self.delete_btn.setFixedSize(28, 28)

        # 添加到布局
        layout.addWidget(self.name_label)
        layout.addWidget(self.status_label)
        layout.addStretch()
        layout.addWidget(self.start_btn)
        layout.addWidget(self.stop_btn)
        layout.addWidget(self.delete_btn)

        # 连接信号
        self.start_btn.clicked.connect(self.start_task)
        self.stop_btn.clicked.connect(self.stop_task)
        self.delete_btn.clicked.connect(lambda: self.parent.delete_task(name))

    def start_task(self):
        self.status_label.setText("运行中")
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        # 更新主界面状态
        if self.parent:
            self.parent.task_status.setText("运行中")
            self.parent.start_current_task()

    def stop_task(self):
        self.status_label.setText("已停止")
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        # 更新主界面状态
        if self.parent:
            self.parent.task_status.setText("已停止")
            self.parent.stop_current_task()


class AutomationUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("自动化任务管理器")

        self.setGeometry(100, 100, 1000, 550)  # 减少高度

        # 应用设置
        self.settings = QSettings("MyCompany", "AutomationManager")
        self.load_settings()

        # 存储任务配置
        self.tasks = {}
        self.current_task = None
        self.task_runner = None
        self.task_thread = None
        self.scheduled_timers = {}  # 存储定时任务的计时器
        # 热键监听器
        self.hotkey_listener = None

        self.setup_hotkey_listener()

        # 创建主布局
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 使用分割器实现可调整宽度的任务列表
        self.splitter = QSplitter(Qt.Horizontal)

        # 左侧任务列表区域
        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.StyledPanel)
        left_panel.setMinimumWidth(280)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(10, 10, 10, 10)

        # 任务列表标题和新建按钮 - 添加emoji
        title_layout = QHBoxLayout()
        title_label = QLabel("📋 任务列表")
        title_label.setFont(QFont("Arial", 11, QFont.Bold))
        title_layout.addWidget(title_label)
        title_layout.addStretch()

        self.new_task_btn = QPushButton("➕ 新建任务")
        self.new_task_btn.setFixedSize(100, 32)
        title_layout.addWidget(self.new_task_btn)

        left_layout.addLayout(title_layout)

        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        left_layout.addWidget(separator)

        # 任务列表
        self.task_list = QListWidget()
        self.task_list.setMinimumHeight(200)
        # 优化hover样式
        self.task_list.setStyleSheet("""
            QListWidget::item:hover {
                background-color: #e0e0e0;
            }
        """)

        # 设置右键菜单
        self.task_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.task_list.customContextMenuRequested.connect(self.show_context_menu)

        # 添加已有任务
        self.load_all_configs("config")

        left_layout.addWidget(self.task_list)

        # 日志区域 - 新增现代化日志记录
        log_group = QGroupBox("📝 执行日志")
        log_layout = QVBoxLayout()

        # 添加清空日志按钮
        log_header_layout = QHBoxLayout()
        log_header_layout.addWidget(QLabel("执行日志:"))
        log_header_layout.addStretch()
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.setFixedSize(80, 24)
        self.clear_log_btn.clicked.connect(self.clear_log)
        log_header_layout.addWidget(self.clear_log_btn)

        log_layout.addLayout(log_header_layout)

        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(200)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        left_layout.addWidget(log_group)

        # 右侧配置区域
        right_panel = QFrame()
        right_panel.setFrameShape(QFrame.StyledPanel)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(15, 15, 15, 15)

        # 任务信息组
        task_info_group = QGroupBox("ℹ️ 任务信息")
        task_info_layout = QFormLayout()
        task_info_layout.setLabelAlignment(Qt.AlignRight)
        task_info_layout.setSpacing(10)

        self.task_name = QLineEdit()
        self.task_name.setPlaceholderText("输入任务名称")
        self.task_name.setFont(QFont("Arial", 10))

        self.task_status = QLabel("未运行")

        task_info_layout.addRow("任务名称:", self.task_name)
        task_info_layout.addRow("当前状态:", self.task_status)

        task_info_group.setLayout(task_info_layout)

        # 定时设置组 - 修改为横向布局
        schedule_group = QGroupBox("⏰ 定时设置")
        schedule_layout = QGridLayout()
        schedule_layout.setSpacing(10)

        # 执行方式
        schedule_layout.addWidget(QLabel("执行方式:"), 0, 0)
        self.schedule_enable = QComboBox()
        self.schedule_enable.addItems(["立即执行", "定时执行"])
        self.schedule_enable.setMinimumWidth(100)
        schedule_layout.addWidget(self.schedule_enable, 0, 1)

        # 执行时间
        schedule_layout.addWidget(QLabel("执行时间:"), 0, 2)
        self.schedule_time = QTimeEdit(QTime.currentTime())
        self.schedule_time.setDisplayFormat("HH:mm:ss")
        self.schedule_time.setMinimumWidth(90)
        schedule_layout.addWidget(self.schedule_time, 0, 3)

        # 重复间隔
        schedule_layout.addWidget(QLabel("重复间隔:"), 1, 0)
        self.repeat_interval = QSpinBox()
        self.repeat_interval.setRange(0, 1440)
        self.repeat_interval.setValue(10)
        self.repeat_interval.setMinimumWidth(60)
        schedule_layout.addWidget(self.repeat_interval, 1, 1)
        schedule_layout.addWidget(QLabel("分钟"), 1, 2)

        # 重复次数
        schedule_layout.addWidget(QLabel("重复次数:"), 1, 3)
        self.repeat_count = QComboBox()
        self.repeat_count.addItems(["1", "3", "5", "10", "无限"])
        self.repeat_count.setCurrentIndex(0)
        schedule_layout.addWidget(self.repeat_count, 1, 4)

        # 应用定时设置按钮
        self.apply_schedule_btn = QPushButton("应用定时设置")
        schedule_layout.addWidget(self.apply_schedule_btn, 2, 0, 1, 5)

        schedule_group.setLayout(schedule_layout)

        # 步骤配置区域
        steps_group = QGroupBox("⚙️ 操作步骤配置")
        steps_layout = QVBoxLayout()
        steps_layout.setSpacing(10)

        # 步骤表格 - 设置列宽可拖拽
        self.steps_table = QTableWidget(0, 4)
        self.steps_table.setHorizontalHeaderLabels(["类型", "描述", "参数", "延时(秒)"])
        self.steps_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)  # 可拖拽调整列宽
        self.steps_table.horizontalHeader().setStretchLastSection(True)
        self.steps_table.verticalHeader().setVisible(False)
        self.steps_table.setSelectionBehavior(QAbstractItemView.SelectRows)

        # 添加示例数据
        # self.populate_steps_table()

        # 步骤操作按钮 - 添加emoji和快捷键
        step_btn_layout = QHBoxLayout()
        self.add_step_btn = QPushButton("➕ 添加步骤 (A)")
        self.add_step_btn.setShortcut(QKeySequence("Ctrl+A"))
        self.edit_step_btn = QPushButton("✏️ 编辑步骤 (E)")
        self.edit_step_btn.setShortcut(QKeySequence("Ctrl+E"))
        self.remove_step_btn = QPushButton("➖ 删除步骤 (Del)")
        self.remove_step_btn.setShortcut(QKeySequence.Delete)  # 确保删除按钮的快捷键为 Delete
        self.copy_step_btn = QPushButton("📋 复制步骤")
        self.copy_step_btn.setShortcut(QKeySequence("Ctrl+C"))  # 新增：设置复制按钮的快捷键为 Ctrl+C
        self.move_up_btn = QPushButton("⬆️ 上移 (↑)")
        self.move_up_btn.setShortcut(QKeySequence("Ctrl+Up"))
        self.move_down_btn = QPushButton("⬇️ 下移 (↓)")
        self.move_down_btn.setShortcut(QKeySequence("Ctrl+Down"))

        step_btn_layout.addWidget(self.add_step_btn)
        step_btn_layout.addWidget(self.edit_step_btn)
        step_btn_layout.addWidget(self.copy_step_btn)
        step_btn_layout.addWidget(self.remove_step_btn)
        step_btn_layout.addStretch()
        step_btn_layout.addWidget(self.move_up_btn)
        step_btn_layout.addWidget(self.move_down_btn)


        steps_layout.addWidget(self.steps_table)
        steps_layout.addLayout(step_btn_layout)
        steps_group.setLayout(steps_layout)

        # 操作按钮组 - 添加emoji
        action_btn_layout = QHBoxLayout()
        self.start_current_btn = QPushButton("▶️ 开始当前任务")
        self.stop_current_btn = QPushButton("⏹️ 停止当前任务")
        self.stop_current_btn.setEnabled(False)
        self.save_btn = QPushButton("💾 保存配置")

        action_btn_layout.addWidget(self.start_current_btn)
        action_btn_layout.addWidget(self.stop_current_btn)
        action_btn_layout.addStretch()
        action_btn_layout.addWidget(self.save_btn)

        # 添加到右侧布局
        right_layout.addWidget(task_info_group)
        right_layout.addWidget(schedule_group)
        right_layout.addWidget(steps_group)
        right_layout.addLayout(action_btn_layout)

        # 添加左右面板到分割器
        self.splitter.addWidget(left_panel)
        self.splitter.addWidget(right_panel)

        # 恢复分割器位置
        splitter_sizes = self.settings.value("splitterSizes")
        if splitter_sizes:
            splitter_sizes = [int(s) for s in splitter_sizes]
            self.splitter.setSizes(splitter_sizes)
        else:
            self.splitter.setSizes([280, 700])

        # 添加日志区域可拖拽调整高度
        self.log_splitter = QSplitter(Qt.Vertical)
        self.log_splitter.addWidget(self.task_list)
        self.log_splitter.addWidget(log_group)
        left_layout.insertWidget(2, self.log_splitter)
        self.log_splitter.setSizes([300, 150])

        # 添加到主布局
        main_layout.addWidget(self.splitter)

        self.setCentralWidget(main_widget)

        # 创建菜单栏
        self.create_menus()

        # 连接信号
        self.task_list.currentItemChanged.connect(self.task_selected)
        self.new_task_btn.clicked.connect(self.create_new_task)
        self.start_current_btn.clicked.connect(self.start_current_task)
        self.stop_current_btn.clicked.connect(self.stop_current_task)
        self.add_step_btn.clicked.connect(self.add_step)
        self.edit_step_btn.clicked.connect(self.edit_step)
        self.remove_step_btn.clicked.connect(self.remove_step)
        self.move_up_btn.clicked.connect(self.move_step_up)
        self.move_down_btn.clicked.connect(self.move_step_down)
        self.save_btn.clicked.connect(self.save_task_config)
        self.apply_schedule_btn.clicked.connect(self.apply_schedule)
        self.copy_step_btn.clicked.connect(self.copy_step)

        # 应用当前主题
        self.apply_theme(self.current_theme)

        # 检测系统主题
        self.detect_system_theme()

        # 创建系统托盘图标
        self.create_system_tray()
    def setup_hotkey_listener(self):
        """启动 Esc 热键监听"""
        self.hotkey_listener = HotkeyListener(self)
        self.hotkey_listener.hotkey_activated.connect(self.on_esc_pressed)
        self.hotkey_listener.start()  # 启动线程
    @Slot()
    def on_esc_pressed(self):
        """响应 Esc 键（在主线程执行）"""
        if self.task_runner and self.task_thread and self.task_thread.is_alive():
            self.stop_current_task()
            self.statusBar().showMessage("🛑 Esc 被按下，任务已停止", 2000)


    def load_all_configs(self, config_dir="config"):
        """
        扫描 config_dir 内所有 *.json 并加载为任务
        """
        config_dir = resource_path(config_dir)
        if not os.path.isdir(config_dir):
            os.makedirs(config_dir, exist_ok=True)
            return

        for fname in os.listdir(config_dir):
            if not fname.endswith(".json"):
                continue
            path = os.path.join(config_dir, fname)
            try:
                if path:
                    try:
                        with open(path, 'r') as f:
                            task_config = json.load(f)

                        task_name = task_config.get("name", "导入的任务")
                        self.add_task(task_name)
                        self.tasks[task_name] = task_config

                        # 选中新导入的任务
                        for i in range(self.task_list.count()):
                            item = self.task_list.item(i)
                            widget = self.task_list.itemWidget(item)
                            if widget and widget.task_name == task_name:
                                self.task_list.setCurrentItem(item)
                                break
                        # QMessageBox.information(self, "导入成功", "任务配置已导入")
                    except Exception as e:
                        QMessageBox.critical(self, "导入失败", f"导入配置时出错: {str(e)}")

            except Exception as e:
                print(f"加载配置 {path} 失败：{e}")
    def create_system_tray(self):
        """创建系统托盘图标"""
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))

        # 创建托盘菜单
        tray_menu = QMenu()

        show_action = QAction("显示窗口", self)
        show_action.triggered.connect(self.show)
        tray_menu.addAction(show_action)

        hide_action = QAction("隐藏窗口", self)
        hide_action.triggered.connect(self.hide)
        tray_menu.addAction(hide_action)

        tray_menu.addSeparator()

        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close)
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()

        # 连接信号
        self.tray_icon.messageClicked.connect(self.tray_message_clicked)

    def tray_icon_activated(self, reason):
        """托盘图标被激活时的处理"""
        if reason == QSystemTrayIcon.DoubleClick:
            if self.isVisible():
                self.hide()
            else:
                self.showNormal()
                self.activateWindow()

    def tray_message_clicked(self):
        """托盘消息被点击时的处理"""
        self.showNormal()
        self.activateWindow()

    def closeEvent(self, event):
        """重写关闭事件，实现最小化到托盘"""
        if self.tray_icon.isVisible():
            self.hide()
            event.ignore()
        else:
            # 保存设置
            self.save_settings()
            # 停止所有定时器
            for timer in self.scheduled_timers.values():
                timer.stop()
            event.accept()

    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
        self.log_text.appendPlainText(f"[{time.strftime('%H:%M:%S')}] 日志已清空")

    def create_menus(self):
        menu_bar = self.menuBar()

        # === 新增：设置菜单 ===
        settings_menu = menu_bar.addMenu("⚙️ 设置")
        # 主容器
        settings_widget = QWidget()
        settings_layout = QVBoxLayout(settings_widget)
        settings_layout.setContentsMargins(8, 4, 8, 4)
        settings_layout.setSpacing(6)

        # 1. 自动跳过 + 超时时间（纵向）
        # --- 自动跳过复选框 ---
        self.auto_skip_checkbox = QCheckBox("图片查找超时后自动跳过")
        self.auto_skip_checkbox.setChecked(False)

        # --- 超时时间（水平布局）---
        timeout_layout = QHBoxLayout()
        timeout_label = QLabel("超时时间:")
        self.timeout_spinbox = QDoubleSpinBox()
        self.timeout_spinbox.setRange(0, 600)
        self.timeout_spinbox.setSingleStep(0.5)
        self.timeout_spinbox.setValue(3)
        self.timeout_spinbox.setSuffix(" 秒")
        self.timeout_spinbox.setFixedWidth(80)
        timeout_layout.addWidget(timeout_label)
        timeout_layout.addWidget(self.timeout_spinbox)
        timeout_layout.addStretch()

        # 2. 鼠标移动设置（水平布局）
        mouse_layout = QHBoxLayout()
        self.instant_click_checkbox = QCheckBox("直接点击")
        self.instant_click_checkbox.setChecked(False)

        self.move_duration_spinbox = QDoubleSpinBox()
        self.move_duration_spinbox.setRange(0.0, 10.0)
        self.move_duration_spinbox.setSingleStep(0.1)
        self.move_duration_spinbox.setValue(0.1)
        self.move_duration_spinbox.setDecimals(1)
        self.move_duration_spinbox.setSuffix(" 秒")
        self.move_duration_spinbox.setFixedWidth(80)
        self.move_duration_spinbox.setEnabled(False)  # 默认禁用，由 checkbox 控制

        # 连接 checkbox 控制 spinbox 启用状态
        def on_instant_click_toggled(checked):
            self.move_duration_spinbox.setEnabled(not checked)

        self.instant_click_checkbox.toggled.connect(on_instant_click_toggled)

        mouse_layout.addWidget(self.instant_click_checkbox)
        mouse_layout.addWidget(self.move_duration_spinbox)
        mouse_layout.addStretch()

        # 添加到主布局
        settings_layout.addWidget(self.auto_skip_checkbox)
        settings_layout.addLayout(timeout_layout)
        settings_layout.addLayout(mouse_layout)

        # 包装为菜单项
        action = QWidgetAction(settings_menu)
        action.setDefaultWidget(settings_widget)
        settings_menu.addAction(action)

        # 文件菜单
        file_menu = menu_bar.addMenu("📁 文件")
        new_action = QAction("📝 新建任务", self)
        save_action = QAction( "💾 保存配置", self)
        export_action = QAction( "📤 导出配置", self)
        import_action = QAction( "📥 导入配置", self)
        exit_action = QAction( "🚪 退出", self)

        file_menu.addAction(new_action)
        file_menu.addAction(save_action)
        file_menu.addSeparator()
        file_menu.addAction(export_action)
        file_menu.addAction(import_action)
        file_menu.addSeparator()
        file_menu.addAction(exit_action)
        # 关键：给菜单本身设置样式表
        file_menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)

        # 编辑菜单
        edit_menu = menu_bar.addMenu("✏️ 编辑")
        add_step_action = QAction( "➕ 添加步骤", self)
        edit_step_action = QAction( "✏️ 编辑步骤", self)
        remove_step_action = QAction( "➖ 删除步骤", self)
        copy_step_action = QAction("📋 复制步骤", self)

        edit_menu.addAction(add_step_action)
        edit_menu.addAction(edit_step_action)
        edit_menu.addAction(copy_step_action)
        edit_menu.addAction(remove_step_action)

        edit_menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)

        # 主题菜单（位于编辑和帮助之间）
        theme_menu = menu_bar.addMenu("🎨 主题")
        theme_menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)

        self.light_theme_action = QAction("☀️ 明亮主题", self)
        self.light_theme_action.setCheckable(True)
        self.light_theme_action.triggered.connect(lambda: self.switch_theme("light"))

        self.dark_theme_action = QAction("🌙 暗黑主题", self)
        self.dark_theme_action.setCheckable(True)
        self.dark_theme_action.triggered.connect(lambda: self.switch_theme("dark"))

        self.system_theme_action = QAction("🔄 跟随系统", self)
        self.system_theme_action.setCheckable(True)
        self.system_theme_action.triggered.connect(lambda: self.switch_theme("system"))

        theme_menu.addAction(self.light_theme_action)
        theme_menu.addAction(self.dark_theme_action)
        theme_menu.addAction(self.system_theme_action)

        # 设置当前主题选中状态
        if self.current_theme == "light":
            self.light_theme_action.setChecked(True)
        elif self.current_theme == "dark":
            self.dark_theme_action.setChecked(True)
        else:
            self.system_theme_action.setChecked(True)

        # 帮助菜单
        help_menu = menu_bar.addMenu("❓ 帮助")
        about_action = QAction("ℹ️ 关于", self)
        docs_action = QAction("📚 使用文档", self)

        help_menu.addAction(docs_action)
        help_menu.addAction(about_action)
        help_menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)

        # 连接菜单信号
        new_action.triggered.connect(self.create_new_task)
        save_action.triggered.connect(self.save_task_config)
        export_action.triggered.connect(self.export_config)
        import_action.triggered.connect(self.import_config)
        exit_action.triggered.connect(self.close)
        add_step_action.triggered.connect(self.add_step)
        edit_step_action.triggered.connect(self.edit_step)
        remove_step_action.triggered.connect(self.remove_step)
        docs_action.triggered.connect(self.show_docs)
        about_action.triggered.connect(self.show_about)
        copy_step_action.triggered.connect(self.copy_step)

    def show_docs(self):
        """显示使用文档对话框"""
        dialog = QDialog(self)
        dialog.setWindowTitle("使用文档")
        dialog.setMinimumWidth(500)
        dialog.setMinimumHeight(500)

        layout = QVBoxLayout(dialog)

        text = QTextEdit()
        text.setReadOnly(True)
        text.setHtml("""
            <h2>自动化任务管理器使用文档</h2>
            <p>欢迎使用自动化任务管理器！本工具可以帮助您自动化执行重复的计算机操作。</p>

            <h3>基本功能</h3>
            <ul>
                <li><b>创建任务</b>：点击"新建任务"按钮创建新任务</li>
                <li><b>添加步骤</b>：在任务中添加鼠标点击、文本输入、等待等操作步骤</li>
                <li><b>定时执行</b>：设置任务在特定时间自动执行，将执行方式改为定时执行，并点击应用定时设置</li>
                <li><b>执行日志</b>：查看任务执行过程中的详细日志</li>
            </ul>

            <h3>配置说明</h3>
            <p>配置任务时，请确保：</p>
            <ul>
                <li>图片路径正确不含有中文，且图片在屏幕上可见</li>
                <li>设置合适的识别精度和超时时间</li>
                <li>为需要等待的操作添加适当的延时</li>
            </ul>

            <h3>QQ交流群</h3>
            <p>加入我们的QQ交流群获取更多帮助：<b>1057721699</b></p>

            <h3>常见问题</h3>
            <p><b>Q: 为什么找不到图片？</b><br>
            A: 请确保图片在屏幕上可见，且识别精度设置合适（建议0.8-0.9）</p>

            <p><b>Q: 任务执行失败怎么办？</b><br>
            A: 查看执行日志中的错误信息，调整步骤参数后重试</p>
            <p><b>Q: 13:14如何计算的？</b><br>
            A: 无论用户什么时候点击按钮，文案中的“相恋时间”都以 今天 13:14 为截止点计算。以确保定时在13：14发送的逻辑</p>
        """)

        layout.addWidget(text)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)

        dialog.exec()

    def show_about(self):
        # 任意窗口里
        AboutDialog(self).exec()


    def add_task(self, name):
        # 创建自定义列表项
        item = QListWidgetItem(self.task_list)
        item_widget = TaskItemWidget(name, self)
        item.setSizeHint(QSize(0, 45))  # 固定高度确保按钮完全显示
        self.task_list.addItem(item)
        self.task_list.setItemWidget(item, item_widget)

        # 初始化任务配置
        self.tasks[name] = {
            "name": name,
            "schedule": {
                "enable": "立即执行",
                "time": QTime.currentTime().toString("HH:mm:ss"),
                "interval": 10,
                "repeat": "1"
            },
            "steps": []
        }

        # 应用当前主题样式
        self.apply_button_style(item_widget)

        # 选中新添加的任务
        if self.task_list.count() == 1:
            self.task_list.setCurrentItem(item)

    def create_new_task(self):
        name = f"新任务 {self.task_list.count() + 1}"
        self.add_task(name)
        ts = time.strftime("%H:%M:%S")
        # 日志带 emoji
        self.log_text.appendPlainText(f"[{ts}] ✅ [{name}] 已创建！")

    def duplicate_task(self, name):
        new_name = f"{name} 副本"
        self.add_task(new_name)

        # 复制任务配置
        if name in self.tasks:
            self.tasks[new_name] = self.tasks[name].copy()
            self.tasks[new_name]["name"] = new_name
        # 日志带 emoji
        ts = time.strftime("%H:%M:%S")
        self.log_text.appendPlainText(f"[{ts}] 📋 [{name}] → [{new_name}] 已复制！")

    def rename_task(self, name):
        """重命名任务：同步所有内部结构与 UI，保证顺序一致。"""
        if name not in self.tasks:
            return

        # 1. 弹窗获取新名称
        new_name, ok = QInputDialog.getText(
            self, "重命名任务", "请输入新名称：", text=name
        )
        if not ok:
            return
        new_name = new_name.strip()
        if not new_name:
            QMessageBox.warning(self, "提示", "任务名称不能为空！")
            return
        if new_name == name:
            return
        if new_name in self.tasks:
            QMessageBox.warning(self, "提示", f"任务“{new_name}”已存在。")
            return

        # 2. 找到旧 item 的行号
        row = -1
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget and widget.task_name == name:
                row = i
                break
        if row == -1:  # 理论上不会发生
            return

        # 3. 创建新 item，并插回原位置
        new_item = QListWidgetItem()
        new_widget = TaskItemWidget(new_name, self)
        new_item.setSizeHint(QSize(0, 45))

        # 4. 复制任务数据
        self.tasks[new_name] = self.tasks[name].copy()
        self.tasks[new_name]["name"] = new_name

        # 5. 替换 UI：先插新的，再删旧的
        self.task_list.insertItem(row, new_item)
        self.task_list.setItemWidget(new_item, new_widget)
        self.task_list.takeItem(row + 1)  # 原来的那行现在是 row+1
        if name in self.scheduled_timers:  # 如果之前有时钟，一起迁移
            self.scheduled_timers[new_name] = self.scheduled_timers.pop(name)

        # 6. 选中新任务并保持焦点
        self.task_list.setCurrentItem(new_item)
        self.apply_button_style(new_widget)

        # 7. 彻底删除旧任务
        del self.tasks[name]
        self.on_log_message(name, f"📝 重命名：{name} → {new_name}")
    def delete_task(self, name):

        row = -1
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget and widget.task_name == name:
                row = i
                break

        if row >= 0:
            self.task_list.takeItem(row)
            if name in self.tasks:
                # 如果任务有定时器，先停止
                if name in self.scheduled_timers:
                    self.scheduled_timers[name].stop()
                    del self.scheduled_timers[name]
                del self.tasks[name]
            print(f"已删除任务: {name}")
            self.on_log_message(name, f"🗑️ 已删除任务：{name}")

    def task_selected(self, current, previous):
        if current:
            widget = self.task_list.itemWidget(current)
            if widget:
                task_name = widget.task_name
                self.current_task = task_name
                self.task_name.setText(task_name)
                self.task_status.setText(widget.status_label.text())

                # 更新按钮状态
                if widget.status_label.text() == "运行中":
                    self.start_current_btn.setEnabled(False)
                    self.stop_current_btn.setEnabled(True)
                else:
                    self.start_current_btn.setEnabled(True)
                    self.stop_current_btn.setEnabled(False)

                # 加载任务配置
                self.load_task_config(task_name)

    def load_task_config(self, task_name):
        if task_name in self.tasks:
            task_config = self.tasks[task_name]

            # 加载定时设置
            self.schedule_enable.setCurrentText(task_config["schedule"]["enable"])
            self.schedule_time.setTime(QTime.fromString(task_config["schedule"]["time"], "HH:mm:ss"))
            self.repeat_interval.setValue(task_config["schedule"]["interval"])
            self.repeat_count.setCurrentText(task_config["schedule"]["repeat"])

            # 加载步骤
            self.steps_table.setRowCount(0)
            for step in task_config["steps"]:
                self.add_step_to_table(step)

    def add_step_to_table(self, step):
        row = self.steps_table.rowCount()
        self.steps_table.insertRow(row)

        self.steps_table.setItem(row, 0, QTableWidgetItem(step["type"]))
        # self.steps_table.setItem(row, 1, QTableWidgetItem(StepTableHelper.desc_of(step)))
        w = StepTableHelper.widget_of(step)
        self.steps_table.setCellWidget(row, 1, w)
        self.steps_table.setRowHeight(row, max(StepTableHelper.IMG_HEIGHT + 4, 24))
        self.steps_table.verticalHeader().setDefaultSectionSize(
            StepTableHelper.FIXED_ROW_HEIGHT
        )
        self.steps_table.horizontalHeader().setStretchLastSection(True)

        # 格式化参数显示
        params_text = ""
        if step["type"] == "鼠标点击":
            params_text = f"图片: {os.path.basename(step['params'].get('image_path', ''))}"
        elif step["type"] == "文本输入":
            params_text = f"文本: {step['params'].get('text', 'excel表内容')}"
        elif step["type"] == "等待":
            params_text = f"等待: {step['params'].get('seconds', 0)}秒"
        elif step["type"] == "截图":
            params_text = f"保存到: {step['params'].get('save_path', '')}"
        elif step["type"] == "鼠标滚轮":
            params_text = f"鼠标滚轮: {step['params'].get('direction', '向下滚动')},{step['params'].get('clicks', '3')}格"
        elif step["type"] == "键盘热键":
            hotkey = step["params"].get("hotkey", "ctrl+c").upper()
            delay = step["params"].get("delay_ms", 100)
            params_text = f"键盘热键: {hotkey}, 延时 {delay} ms"
        elif step["type"] == "拖拽":
            params_text = f"从({step['params'].get('start_x', 0)},{step['params'].get('start_y', 0)})到({step['params'].get('end_x', 0)},{step['params'].get('end_y', 0)})"

        self.steps_table.setItem(row, 2, QTableWidgetItem(params_text))
        self.steps_table.setItem(row, 3, QTableWidgetItem(str(step.get("delay", 0))))
        self.steps_table.resizeColumnToContents(1)  # 列宽按内容自适应

    def start_current_task(self):
        if not self.current_task:
            return

        # 获取任务配置
        task_config = self.tasks.get(self.current_task, {})
        steps = task_config.get("steps", [])

        if not steps:
            QMessageBox.warning(self, "无法启动", "当前任务没有配置任何步骤")
            return
        auto_skip = self.auto_skip_checkbox.isChecked()  # ✅ 读取 QCheckBox 状态
        timeout = self.timeout_spinbox.value()  # 获取用户设置的超时时间
        instant_click = self.instant_click_checkbox.isChecked()
        move_duration = self.move_duration_spinbox.value() if not instant_click else 0.0

        # 创建任务运行器
        self.task_runner = TaskRunner(self.current_task, steps,
                                      auto_skip_image_timeout=auto_skip,
                                      timeout=timeout,
                                      instant_click=instant_click,
                                      move_duration=move_duration)

        # 设置重复次数
        repeat_text = self.repeat_count.currentText()
        if repeat_text == "无限":
            self.task_runner.set_repeat_count(9999)  # 设置一个很大的数表示无限
        else:
            try:
                count = int(repeat_text)
                self.task_runner.set_repeat_count(count)
            except:
                self.task_runner.set_repeat_count(1)

        # 连接信号
        self.task_runner.task_completed.connect(self.on_task_completed)
        self.task_runner.task_progress.connect(self.on_task_progress)
        self.task_runner.log_message.connect(self.on_log_message)  # 连接日志信号

        # 在单独线程中运行任务
        self.task_thread = threading.Thread(target=self.task_runner.run)
        self.task_thread.daemon = True
        self.task_thread.start()

        # 更新UI状态
        self.start_current_btn.setEnabled(False)
        self.stop_current_btn.setEnabled(True)
        self.task_status.setText("运行中")

        # 更新任务列表中的状态
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget and widget.task_name == self.current_task:
                widget.status_label.setText("运行中")
                widget.start_btn.setEnabled(False)
                widget.stop_btn.setEnabled(True)
                break

        # 新增：任务开始后最小化窗口
        self.showMinimized()

    def stop_current_task(self):
        if self.task_runner and self.task_runner.is_running:
            self.task_runner.stop()

            # 更新UI状态
            self.start_current_btn.setEnabled(True)
            self.stop_current_btn.setEnabled(False)
            self.task_status.setText("已停止")

            # 更新任务列表中的状态
            for i in range(self.task_list.count()):
                item = self.task_list.item(i)
                widget = self.task_list.itemWidget(item)
                if widget and widget.task_name == self.current_task:
                    widget.status_label.setText("已停止")
                    widget.start_btn.setEnabled(True)
                    widget.stop_btn.setEnabled(False)
                    break

    def closeEvent(self, event):
        # 清理热键监听
        if self.hotkey_listener and self.hotkey_listener.isRunning():
            self.hotkey_listener.stop()

        super().closeEvent(event)

    def on_task_completed(self, task_name, success, message):
        # 新增：任务完成后恢复窗口显示
        self.showNormal()

        # 更新UI状态
        self.start_current_btn.setEnabled(True)
        self.stop_current_btn.setEnabled(False)
        self.task_status.setText("已停止" if success else "已中断")

        # 更新任务列表中的状态
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget and widget.task_name == task_name:
                widget.status_label.setText("已停止" if success else "已中断")
                widget.start_btn.setEnabled(True)
                widget.stop_btn.setEnabled(False)
                break

        # 记录日志
        self.log_text.appendPlainText(f"[{time.strftime('%H:%M:%S')}] {message}")

    def on_task_progress(self, task_name, current, total):
        self.task_status.setText(f"运行中 ({current}/{total})")

    def on_log_message(self, task_name, message):
        """处理日志消息"""
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        log_entry = f"[{timestamp}] [{task_name}] {message}"
        self.log_text.appendPlainText(log_entry)

        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def show_context_menu(self, pos):
        # 获取点击位置的item
        item = self.task_list.itemAt(pos)
        if not item:
            return

        # 创建上下文菜单
        menu = QMenu(self)

        # 获取任务名称
        widget = self.task_list.itemWidget(item)
        task_name = widget.task_name if widget else ""

        # 添加菜单项
        rename_action = menu.addAction(QIcon.fromTheme("edit-rename"), "✏️ 重命名")
        duplicate_action = menu.addAction(QIcon.fromTheme("edit-copy"), "📋 创建副本")
        menu.addSeparator()
        delete_action = menu.addAction(QIcon.fromTheme("edit-delete"), "🗑️ 删除任务")
        menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)
        # 显示菜单并获取选择
        action = menu.exec(self.task_list.mapToGlobal(pos))

        # 处理选择
        if action == rename_action:
            self.rename_task(task_name)
        elif action == duplicate_action:
            self.duplicate_task(task_name)
        elif action == delete_action:
            self.delete_task(task_name)

    def switch_theme(self, theme):
        if theme == "system":
            self.detect_system_theme()
        else:
            self.current_theme = theme
            self.apply_theme(theme)
            self.settings.setValue("theme", theme)

        # 更新主题菜单选中状态
        self.light_theme_action.setChecked(self.current_theme == "light")
        self.dark_theme_action.setChecked(self.current_theme == "dark")
        self.system_theme_action.setChecked(theme == "system")

        # 更新任务列表按钮样式
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget:
                self.apply_button_style(widget)

    def detect_system_theme(self):
        """检测系统主题设置"""
        try:
            # 尝试检测系统是否处于暗黑模式
            # 这里只是一个示例，实际实现需要根据操作系统进行适配
            # 在Windows上可以使用注册表，在macOS上可以使用NSAppearance
            # 这里简化为使用系统设置中的值
            dark_mode = self.settings.value("systemDarkMode", False, type=bool)
            self.current_theme = "dark" if dark_mode else "light"
        except:
            self.current_theme = "light"

        self.apply_theme(self.current_theme)
        self.settings.setValue("theme", "system")

    def apply_button_style(self, widget):
        """应用按钮样式到任务项控件"""
        if self.current_theme == "light":
            widget.start_btn.setStyleSheet(self.light_button_style("start"))
            widget.stop_btn.setStyleSheet(self.light_button_style("stop"))
            widget.delete_btn.setStyleSheet(self.light_button_style("delete"))
            widget.status_label.setStyleSheet("color: #888; background: transparent;")
        else:
            widget.start_btn.setStyleSheet(self.dark_button_style("start"))
            widget.stop_btn.setStyleSheet(self.dark_button_style("stop"))
            widget.delete_btn.setStyleSheet(self.dark_button_style("delete"))
            widget.status_label.setStyleSheet("color: #aaa; background: transparent;")

    def light_button_style(self, btn_type):
        """明亮主题按钮样式"""
        base_style = """
            QPushButton {
                border-radius: 14px;
                padding: 0px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
        """

        if btn_type == "start":
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f5f5f5, stop:1 #e0e0e0);
            """
        elif btn_type == "stop":
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f5f5f5, stop:1 #e0e0e0);
            """
        else:  # delete
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f5f5f5, stop:1 #e0e0e0);
            """

    def dark_button_style(self, btn_type):
        """暗黑主题按钮样式"""
        base_style = """
            QPushButton {
                border-radius: 14px;
                padding: 0px;
            }
            QPushButton:hover {
                background-color: #505050;
            }
            QPushButton:pressed {
                background-color: #404040;
            }
        """

        if btn_type == "start":
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #505050, stop:1 #404040);
            """
        elif btn_type == "stop":
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #505050, stop:1 #404040);
            """
        else:  # delete
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #505050, stop:1 #404040);
            """

    def apply_theme(self, theme):
        if theme == "light":
            # 明亮主题
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #f5f7fa;
                }
                QWidget {
                    background-color: #f5f7fa;
                }
                QGroupBox {
                    font-weight: bold;
                    border: 1px solid #d1d5db;
                    border-radius: 6px;
                    margin-top: 1ex;
                    background-color: white;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    top: -1px;    
                    padding: 0 5px;
                    background-color: transparent;
                    color: #333;
                }
                QListWidget, QTableWidget, QLineEdit, QComboBox, QTimeEdit, QSpinBox, QPlainTextEdit {
                    background-color: white;
                    border: 1px solid #d1d5db;
                    color: #333;
                }
                QHeaderView::section {
                    background-color: #f0f0f0;
                    color: #333;
                    border: none;
                    border-bottom: 1px solid #d1d5db;
                }
                QPushButton {
                    color: #333;
                }
                QLabel {
                    color: #333;
                }
                QCheckBox {
                    color: #333;
                    spacing: 8px;
                }
                QCheckBox::indicator {
                    width: 10px;
                    height: 10px;
                    border: 2px solid #999;
                    border-radius: 4px;
                    background: #fff;
                }
                QCheckBox::indicator:checked {
                    background: #4CAF50;
                    border: 2px solid #388E3C;
                }
                QCheckBox::indicator:hover {
                    border-color: #4CAF50;
                }
                QListWidget::item:selected {
                    background-color: #e3f2fd;
                }
            """)

            # 应用按钮样式
            self.start_current_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.stop_current_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.save_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.new_task_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
            """)

            self.apply_schedule_btn.setStyleSheet("""
QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 #f5f5f5, stop:0.5 #e8f5e9, stop:1 #c8e6c9);
    border: 1px solid #a5d6a7;
    color: #004d40;
    padding: 5px;
    border-radius: 4px;
}
QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                stop:0 #e8f5e9, stop:0.5 #c8e6c9, stop:1 #a5d6a7);
}
            """)

            self.clear_log_btn.setStyleSheet("""
                QPushButton {
                    background-color: #9E9E9E;
                    border: 1px solid #757575;
                    color: white;
                    padding: 2px 5px;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #757575;
                }
            """)
        else:
            # 暗黑主题
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #2d2d30;
                }
                QWidget {
                    background-color: #2d2d30;
                    color: #dcdcdc;
                }
                QGroupBox {
                    font-weight: bold;
                    border: 1px solid #3f3f46;
                    border-radius: 6px;
                    margin-top: 1ex;
                    background-color: #252526;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    top: -1px;    
                    padding: 0 5px;
                    background-color: transparent;
                    color: #dcdcdc;
                }
                QListWidget, QTableWidget, QLineEdit, QComboBox, QTimeEdit, QSpinBox, QPlainTextEdit {
                    background-color: #252526;
                    border: 1px solid #3f3f46;
                    color: #dcdcdc;
                    selection-background-color: #3e3e40;
                }
                QHeaderView::section {
                    background-color: #3e3e40;
                    color: #dcdcdc;
                    border: none;
                    border-bottom: 1px solid #3f3f46;
                }
                QLabel {
                    color: #dcdcdc;
                }
                               QCheckBox {
                    color: #dcdcdc;
                    spacing: 8px;
                }
                QCheckBox::indicator {
                    width: 10px;
                    height: 10px;
                    border: 2px solid #666;
                    border-radius: 4px;
                    background: #2d2d30;
                }
                QCheckBox::indicator:checked {
                    background: #4CAF50;
                    border: 2px solid #388E3C;
                }
                QCheckBox::indicator:hover {
                    border-color: #4CAF50;
                }
                QListWidget::item:selected {
                    background-color: #3e3e40;
                }
            """)

            # 应用按钮样式
            self.start_current_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.stop_current_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.save_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.new_task_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
            """)

            self.clear_log_btn.setStyleSheet("""
                QPushButton {
                    background-color: #9E9E9E;
                    border: 1px solid #757575;
                    color: white;
                    padding: 2px 5px;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #757575;
                }
            """)

    def load_settings(self):
        # 加载主题设置
        self.current_theme = self.settings.value("theme", "light")

    def save_settings(self):
        # 保存分割器位置
        self.settings.setValue("splitterSizes", self.splitter.sizes())

    def closeEvent(self, event):
        self.save_settings()
        # 停止所有定时器
        for timer in self.scheduled_timers.values():
            timer.stop()
        event.accept()

    def add_step(self):
        dialog = StepConfigDialog(parent=self)
        if dialog.exec() == QDialog.Accepted:
            step_data = dialog.get_step_data()
            self.add_step_to_table(step_data)

            # 添加到当前任务配置
            if self.current_task and self.current_task in self.tasks:
                self.tasks[self.current_task]["steps"].append(step_data)

    def copy_step(self):
        """复制当前选中的步骤"""
        selected_row = self.steps_table.currentRow()
        if selected_row < 0:
            QMessageBox.information(self, "提示", "请先选中一条步骤再复制。")
            return

        # 取出原步骤数据
        src_step = self.tasks[self.current_task]["steps"][selected_row]
        # 深拷贝，避免后续修改互相影响
        new_step = deepcopy(src_step)

        # 直接追加到表格和任务配置
        self.add_step_to_table(new_step)
        self.tasks[self.current_task]["steps"].append(new_step)
    def edit_step(self):
        selected_row = self.steps_table.currentRow()
        if selected_row < 0:
            return

        # 获取当前步骤数据
        step_data = self.tasks[self.current_task]["steps"][selected_row]
        dialog = StepConfigDialog(step_data,parent=self)
        if dialog.exec() == QDialog.Accepted:
            new_step_data = dialog.get_step_data()

            # 更新表格
            self.steps_table.setItem(selected_row, 0, QTableWidgetItem(new_step_data["type"]))

            # 格式化参数显示
            params_text = ""
            params = new_step_data["params"]
            if new_step_data["type"] == "鼠标点击":
                img_path = new_step_data['params'].get('image_path', '')
                img_name = os.path.basename(img_path)  # 去掉目录，只剩文件名
                params_text = f"图片: {img_name}"
            elif new_step_data["type"] == "文本输入":
                # 优先显示纯文本
                txt = params.get("text", "")
                if txt:
                    params_text = f"文本: {txt}"
                else:
                    # Excel 模式
                    mode = params.get("mode", "顺序")
                    path = os.path.basename(params.get("excel_path", ""))
                    sheet = params.get("sheet", "0")
                    col = params.get("col", 0)
                    params_text = f"Excel({mode}) {path}|{sheet}|列{col}"
            elif new_step_data["type"] == "等待":
                params_text = f"等待: {new_step_data['params'].get('seconds', 0)}秒"
            elif new_step_data["type"] == "截图":
                params_text = f"保存到: {new_step_data['params'].get('save_path', '')}"
            elif new_step_data["type"] == "拖拽":
                params_text = f"从({new_step_data['params'].get('start_x', 0)},{new_step_data['params'].get('start_y', 0)})到({new_step_data['params'].get('end_x', 0)},{new_step_data['params'].get('end_y', 0)})"

            self.steps_table.setItem(selected_row, 2, QTableWidgetItem(params_text))
            self.steps_table.setItem(selected_row, 3, QTableWidgetItem(str(new_step_data.get("delay", 0))))

            # 更新任务配置
            self.tasks[self.current_task]["steps"][selected_row] = new_step_data

    def remove_step(self):
        selected_row = self.steps_table.currentRow()
        if selected_row >= 0:
            self.steps_table.removeRow(selected_row)

            # 从任务配置中移除
            if self.current_task and self.current_task in self.tasks:
                self.tasks[self.current_task]["steps"].pop(selected_row)

    def move_step_up(self):
        selected_row = self.steps_table.currentRow()
        if selected_row > 0:
            # 移动表格行
            self.steps_table.insertRow(selected_row - 1)
            for col in range(self.steps_table.columnCount()):
                # 移动 QTableWidgetItem
                self.steps_table.setItem(selected_row - 1, col, self.steps_table.takeItem(selected_row + 1, col))
                # 移动 cellWidget
                widget = self.steps_table.cellWidget(selected_row + 1, col)
                if widget:
                    self.steps_table.setCellWidget(selected_row - 1, col, widget)
            self.steps_table.removeRow(selected_row + 1)
            self.steps_table.setCurrentCell(selected_row - 1, 0)

            # 移动任务配置中的步骤
            if self.current_task and self.current_task in self.tasks:
                steps = self.tasks[self.current_task]["steps"]
                steps.insert(selected_row - 1, steps.pop(selected_row))

    def move_step_down(self):
        selected_row = self.steps_table.currentRow()
        if selected_row >= 0 and selected_row < self.steps_table.rowCount() - 1:
            # 移动表格行
            self.steps_table.insertRow(selected_row + 2)
            for col in range(self.steps_table.columnCount()):
                # 移动 QTableWidgetItem
                self.steps_table.setItem(selected_row + 2, col, self.steps_table.takeItem(selected_row, col))
                # 移动 cellWidget
                widget = self.steps_table.cellWidget(selected_row, col)
                if widget:
                    self.steps_table.setCellWidget(selected_row + 2, col, widget)
            self.steps_table.removeRow(selected_row)
            self.steps_table.setCurrentCell(selected_row + 1, 0)

            # 移动任务配置中的步骤
            if self.current_task and self.current_task in self.tasks:
                steps = self.tasks[self.current_task]["steps"]
                steps.insert(selected_row + 1, steps.pop(selected_row))

    def save_task_config(self):
        if not self.current_task:
            return

        # 更新任务名称
        new_name = self.task_name.text().strip()
        if new_name and new_name != self.current_task:
            # 更新任务列表
            for i in range(self.task_list.count()):
                item = self.task_list.item(i)
                widget = self.task_list.itemWidget(item)
                if widget and widget.task_name == self.current_task:
                    widget.task_name = new_name
                    widget.name_label.setText(new_name)

                    # 更新任务配置
                    task_config = self.tasks.pop(self.current_task)
                    task_config["name"] = new_name
                    self.tasks[new_name] = task_config
                    self.current_task = new_name
                    break

        # 更新定时设置
        if self.current_task in self.tasks:
            self.tasks[self.current_task]["schedule"] = {
                "enable": self.schedule_enable.currentText(),
                "time": self.schedule_time.time().toString("HH:mm:ss"),
                "interval": self.repeat_interval.value(),
                "repeat": self.repeat_count.currentText()
            }
        self.export_config()
        # QMessageBox.information(self, "保存成功", "任务配置已保存")

    def export_config(self):
        if not self.current_task:
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出配置", "", "JSON文件 (*.json)"
        )
        if file_path:
            if not file_path.lower().endswith('.json'):
                file_path += '.json'

            if self.current_task in self.tasks:
                with open(file_path, 'w') as f:
                    json.dump(self.tasks[self.current_task], f, indent=4)
                QMessageBox.information(self, "导出成功", "任务配置已导出")

    def import_config(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "导入配置", "", "JSON文件 (*.json)"
        )
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    task_config = json.load(f)

                task_name = task_config.get("name", "导入的任务")
                self.add_task(task_name)
                self.tasks[task_name] = task_config
                self.steps_table.setRowCount(0)
                for step in task_config["steps"]:
                    self.add_step_to_table(step)

                # 选中新导入的任务
                for i in range(self.task_list.count()):
                    item = self.task_list.item(i)
                    widget = self.task_list.itemWidget(item)
                    if widget and widget.task_name == task_name:
                        self.task_list.setCurrentItem(item)
                        break

                QMessageBox.information(self, "导入成功", "任务配置已导入")
            except Exception as e:
                QMessageBox.critical(self, "导入失败", f"导入配置时出错: {str(e)}")

    def apply_schedule(self):
        """应用定时设置"""
        if not self.current_task:
            return

        task_name = self.current_task

        # 如果任务已有定时器，先停止
        if task_name in self.scheduled_timers:
            self.scheduled_timers[task_name].stop()
            del self.scheduled_timers[task_name]

        # 获取定时设置
        schedule_type = self.schedule_enable.currentText()
        if schedule_type == "立即执行":
            # 不需要定时器
            return

        # 定时执行
        schedule_time = self.schedule_time.time()
        interval_minutes = self.repeat_interval.value()
        repeat_count = self.repeat_count.currentText()

        # 计算第一次执行的时间
        now = QTime.currentTime()
        first_run = QTime(schedule_time.hour(), schedule_time.minute(), schedule_time.second())

        # 如果当前时间已超过设定时间，则明天执行
        if first_run < now:
            first_run = first_run.addSecs(24 * 3600)  # 加一天

        # 计算延迟时间（毫秒）
        delay_ms = now.msecsTo(first_run)

        # 创建定时器
        timer = QTimer(self)
        timer.setSingleShot(True)  # 第一次执行是单次

        # 连接定时器信号
        def run_task():
            # 执行任务
            self.start_current_task()

            # 如果不是无限循环，减少重复次数
            if repeat_count != "无限":
                try:
                    count = int(repeat_count)
                    if count > 1:
                        # 设置间隔定时器
                        interval_timer = QTimer(self)
                        interval_timer.setInterval(interval_minutes * 60 * 1000)  # 分钟转毫秒
                        interval_timer.timeout.connect(
                            lambda: self.run_scheduled_task(task_name, interval_timer, count - 1))
                        interval_timer.start()
                    # 保存定时器引用
                    self.scheduled_timers[task_name] = interval_timer
                except:
                    pass
            else:
                # 无限循环
                interval_timer = QTimer(self)
                interval_timer.setInterval(interval_minutes * 60 * 1000)  # 分钟转毫秒
                interval_timer.timeout.connect(self.start_current_task)
                interval_timer.start()
                # 保存定时器引用
                self.scheduled_timers[task_name] = interval_timer

        timer.timeout.connect(run_task)

        # 启动定时器
        timer.start(delay_ms)

        # 保存定时器引用
        self.scheduled_timers[task_name] = timer

        # 显示提示信息
        self.log_text.appendPlainText(
            f"[{time.strftime('%H:%M:%S')}] 已设置定时任务: {task_name} 将在 {first_run.toString('HH:mm:ss')} 执行")
        QMessageBox.information(self, "定时成功",  f"[{time.strftime('%H:%M:%S')}] 已设置定时任务: {task_name} 将在 {first_run.toString('HH:mm:ss')} 执行\n请保持桌面处于从不熄屏状态")
    def run_scheduled_task(self, task_name, timer, count):
        """执行定时任务（带计数）"""
        if count <= 0:
            timer.stop()
            if task_name in self.scheduled_timers:
                del self.scheduled_timers[task_name]
            return

        # 执行任务
        self.start_current_task()

        # 减少计数
        if count > 1:
            # 设置下一次执行
            QTimer.singleShot(0, lambda: self.run_scheduled_task(task_name, timer, count - 1))
        else:
            timer.stop()
            if task_name in self.scheduled_timers:
                del self.scheduled_timers[task_name]


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AutomationUI()
    window.setWindowIcon(ATIcon.icon())
    window.show()
    sys.exit(app.exec())

    # app = QApplication(sys.argv)  # 必须初始化
    # ok = ATIcon.pixmap().save("icon.ico", "ICO")
    # if ok:
    #     print("✅ icon.ico 已生成！")
    # else:
    #     print("❌ 保存失败")