#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import shutil
import tempfile
import configparser
from pathlib import Path
from datetime import datetime
from itertools import count

import fitz  # PyMuPDF
from PIL import Image
from PIL.ImageQt import ImageQt

from PyQt6.QtCore import (QCoreApplication, QMetaObject, QSize, Qt, QThread, pyqtSignal, QObject, QRect, QSizeF, QUrl)
from PyQt6.QtGui import QIcon, QPainter, QPageSize, QPageLayout, QColor, QDropEvent, QDragEnterEvent
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog, QPrinterInfo
from PyQt6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialogButtonBox, 
    QDoubleSpinBox, QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QPushButton, QSizePolicy, QSlider, 
    QSpacerItem, QTextEdit, QToolButton, QVBoxLayout, QWidget, QMessageBox, 
    QFileDialog, QDialog, QStyleFactory
)

# ==========================================
# 样式表定义
# ==========================================
DARK_QSS = """
QWidget { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #1c1e26, stop:1 #2a2d3a); color: #e0e0e0; font-family: 'Segoe UI'; font-size: 13px; }
QLabel#status { font-weight: bold; font-size: 15px; color: #4cafef; padding: 6px; border-radius: 8px; background-color: #2a2d3a; }
QLabel#warning { color: #ff4d4d; font-weight: bold; font-size: 13px; padding: 8px; border: 1px solid #ff4d4d; border-radius: 8px; background-color: #2a1c1c; }
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox { background-color: #2a2d3a; border: 1px solid #3a3f4a; border-radius: 8px; color: #e0e0e0; padding: 6px; }
QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus { border: 1px solid #4cafef; }
QListWidget { background-color: #2a2d3a; border: 1px solid #3a3f4a; border-radius: 8px; padding: 5px; }
QPushButton { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #3a3f4a, stop:1 #4a4f5a); border: none; border-radius: 8px; color: #fff; padding: 8px 12px; font-weight: bold; }
QPushButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4cafef, stop:1 #5cc6ff); }
QPushButton#apply { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4cafef, stop:1 #5cc6ff); color: #fff; }
QPushButton#apply:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #5cc6ff, stop:1 #6cd8ff); }
"""

# ==========================================
# UI 界面代码 (移除了自动弹窗 Dock, 添加了高级按钮)
# ==========================================
class Ui_PrintForm(object):
    def setupUi(self, PrintForm):
        PrintForm.setObjectName(u"PrintForm")
        PrintForm.resize(700, 550)
        PrintForm.setAcceptDrops(True)
        PrintForm.setWindowIcon(QIcon.fromTheme("printer"))
        
        self.horizontalLayout = QHBoxLayout(PrintForm)
        
        self.listWidget = QListWidget(PrintForm)
        self.listWidget.setObjectName(u"listWidget")
        self.horizontalLayout.addWidget(self.listWidget)

        self.verticalLayout_2 = QVBoxLayout()
        
        self.btn_add_file = QPushButton("添加文件")
        self.btn_add_current_pdf = QPushButton("一键添加同目录PDF")
        self.btn_add_folder_pdf = QPushButton("选择文件夹添加PDF")
        self.btn_clean_pdf = QPushButton("一键清理同目录PDF")
        self.btn_clean_pdf.setStyleSheet("QPushButton { background: #5c2b29; color: #ffcccc; } QPushButton:hover { background: #e04338; color: white; }")
        
        self.btn_clear = QPushButton("清空列表")
        self.btn_preview = QPushButton("打印预览")
        self.btn_print = QPushButton("直接打印")
        self.btn_about = QPushButton("关于")
        self.btn_advanced = QPushButton("高级设定 (路径粘贴/自定义纸张)")
        self.btn_advanced.setObjectName("apply")

        self.label_8 = QLabel("准备就绪......")
        self.label_8.setObjectName("status")

        self.verticalLayout_2.addWidget(self.btn_add_file)
        self.verticalLayout_2.addWidget(self.btn_add_current_pdf)
        self.verticalLayout_2.addWidget(self.btn_add_folder_pdf)
        self.verticalLayout_2.addWidget(self.btn_clean_pdf)
        self.verticalLayout_2.addWidget(self.btn_clear)
        self.verticalLayout_2.addWidget(self.btn_advanced)
        self.verticalLayout_2.addWidget(self.btn_preview)
        self.verticalLayout_2.addWidget(self.btn_print)
        self.verticalLayout_2.addWidget(self.btn_about)
        self.verticalLayout_2.addWidget(self.label_8)

        self.groupBox = QGroupBox("打印选项")
        self.gridLayout = QGridLayout(self.groupBox)
        
        self.label_6 = QLabel("打印机选择")
        self.comboBox_5 = QComboBox() # 打印机列表
        self.gridLayout.addWidget(self.label_6, 0, 0)
        self.gridLayout.addWidget(self.comboBox_5, 0, 1)

        self.label = QLabel("纸张大小")
        self.comboBox = QComboBox()
        self.comboBox.addItems(["A4", "A4两版", "A4三版", "A4四版-2*2", "A4六版-2*3", "A5", "A5两版", "A5四版-2*2", "自定义纸张"])
        self.gridLayout.addWidget(self.label, 1, 0)
        self.gridLayout.addWidget(self.comboBox, 1, 1)

        self.label_4 = QLabel("双面打印")
        self.comboBox_3 = QComboBox()
        self.comboBox_3.addItems(["单面打印", "长边翻转", "短边翻转", "自动选择"])
        self.gridLayout.addWidget(self.label_4, 2, 0)
        self.gridLayout.addWidget(self.comboBox_3, 2, 1)

        self.label_7 = QLabel("页面方向")
        self.comboBox_6 = QComboBox()
        self.comboBox_6.addItems(["自动旋转", "纵向", "横向"])
        self.gridLayout.addWidget(self.label_7, 3, 0)
        self.gridLayout.addWidget(self.comboBox_6, 3, 1)

        self.label_5 = QLabel("居中方式")
        self.comboBox_4 = QComboBox()
        self.comboBox_4.addItems(["水平居中", "靠右居中", "垂直两端", "无"])
        self.gridLayout.addWidget(self.label_5, 4, 0)
        self.gridLayout.addWidget(self.comboBox_4, 4, 1)

        self.label_2 = QLabel("打印份数")
        self.doubleSpinBox = QDoubleSpinBox()
        self.doubleSpinBox.setDecimals(0)
        self.doubleSpinBox.setMinimum(1)
        self.doubleSpinBox.setMaximum(1000)
        self.gridLayout.addWidget(self.label_2, 5, 0)
        self.gridLayout.addWidget(self.doubleSpinBox, 5, 1)

        self.label_3 = QLabel("打印分辨率")
        self.comboBox_2 = QComboBox()
        self.comboBox_2.addItems(["150dpi", "300dpi", "600dpi"])
        self.comboBox_2.setCurrentIndex(1)
        self.gridLayout.addWidget(self.label_3, 6, 0)
        self.gridLayout.addWidget(self.comboBox_2, 6, 1)

        self.checkBox_2 = QCheckBox("合并页面")
        self.checkBox = QCheckBox("灰度打印")
        self.gridLayout.addWidget(self.checkBox_2, 7, 0)
        self.gridLayout.addWidget(self.checkBox, 7, 1)

        self.verticalLayout_2.addWidget(self.groupBox)

        self.toolButton = QToolButton()
        self.toolButton.setText("使用系统对话框进行打印..")
        self.toolButton.setAutoRaise(True)
        self.toolButton.setStyleSheet("color: #4cafef;")
        self.verticalLayout_2.addWidget(self.toolButton, 0, Qt.AlignmentFlag.AlignHCenter)

        self.horizontalLayout.addLayout(self.verticalLayout_2)
        self.horizontalLayout.setStretch(0, 6)
        self.horizontalLayout.setStretch(1, 4)
        PrintForm.setWindowTitle("洛璃打印助手")

# ==========================================
# 高级选项弹窗 (路径粘贴 & 自定义纸张)
# ==========================================
class AdvancedDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("高级设定")
        self.setStyleSheet(DARK_QSS)
        self.setMinimumSize(400, 300)
        
        layout = QVBoxLayout(self)
        
        # 路径粘贴区
        layout.addWidget(QLabel("<b>批量路径粘贴：</b> (一行一个绝对路径)"))
        self.textEdit = QTextEdit()
        self.textEdit.setPlaceholderText("粘贴完整文件路径到此处...")
        layout.addWidget(self.textEdit)
        
        # 自定义纸张区
        layout.addWidget(QLabel("<b>自定义纸张大小 (选择“自定义纸张”时生效):</b>"))
        size_layout = QHBoxLayout()
        size_layout.addWidget(QLabel("宽(mm):"))
        self.width_input = QLineEdit("210")
        size_layout.addWidget(self.width_input)
        size_layout.addWidget(QLabel("高(mm):"))
        self.height_input = QLineEdit("297")
        size_layout.addWidget(self.height_input)
        layout.addLayout(size_layout)
        
        # 按钮
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

# ==========================================
# PDF 临时清理管理窗口
# ==========================================
class PdfManagerWindow(QWidget):
    def __init__(self, temp_dir, files, original_dir):
        super().__init__()
        self.temp_dir = temp_dir
        self.files = files
        self.original_dir = original_dir
        self.setStyleSheet(DARK_QSS)
        self.setWindowTitle("PDF 暂存站管理")
        self.setMinimumSize(450, 350)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"以下文件已移入临时区：\n{self.temp_dir}"))
        
        self.list_widget = QListWidget()
        for f in self.files:
            self.list_widget.addItem(f.name)
        layout.addWidget(self.list_widget)

        btn_layout = QHBoxLayout()
        btn_restore = QPushButton("恢复所有")
        btn_restore.clicked.connect(self.restore_all)
        btn_delete = QPushButton("彻底粉碎")
        btn_delete.setStyleSheet("background: #e04338;")
        btn_delete.clicked.connect(self.delete_all)
        btn_cancel = QPushButton("保留并关闭")
        btn_cancel.clicked.connect(self.close)
        
        btn_layout.addWidget(btn_restore)
        btn_layout.addWidget(btn_delete)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def restore_all(self):
        for f in self.files:
            dest = self.original_dir / f.name
            if not dest.exists(): shutil.move(str(f), str(dest))
        try: self.temp_dir.rmdir()
        except: pass
        QMessageBox.information(self, "完成", "文件已恢复。")
        self.close()

    def delete_all(self):
        try:
            shutil.rmtree(self.temp_dir)
            QMessageBox.information(self, "完成", "已彻底清理。")
            self.close()
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

# ==========================================
# 统一打印核心渲染引擎 (剥离原先混乱的多线程逻辑)
# ==========================================
class PrintEngine:
    def __init__(self, main_win, printer: QPrinter, file_paths: list):
        self.win = main_win
        self.printer = printer
        self.paths = list(file_paths)
        self.paper = self.win.ui.comboBox.currentText()
        self.paper_index = self.win.ui.comboBox.currentIndex()
        self.dpi = int(self.win.ui.comboBox_2.currentText()[:3])
        self.orientation_index = self.win.ui.comboBox_6.currentIndex()
        self.alignment = self.win.ui.comboBox_4.currentText()
        self.inter = 2 if self.win.ui.comboBox_4.currentIndex() in (0, 2, 3) else 1
        self.custom_w = self._safe_positive_float(self.win.custom_width, 210.0)
        self.custom_h = self._safe_positive_float(self.win.custom_height, 297.0)
        self.grayscale = self.win.ui.checkBox.isChecked()

    @staticmethod
    def _safe_positive_float(value, default):
        try:
            parsed = float(str(value).strip())
        except (TypeError, ValueError):
            return default
        return parsed if parsed > 0 else default

    def _get_paint_rect(self, painter):
        page_rect = self.printer.pageRect(QPrinter.Unit.DevicePixel)
        if page_rect.width() > 0 and page_rect.height() > 0:
            return QRect(
                int(page_rect.x()),
                int(page_rect.y()),
                max(1, int(page_rect.width())),
                max(1, int(page_rect.height())),
            )

        rect = painter.viewport()
        if rect.width() > 0 and rect.height() > 0:
            return QRect(
                int(rect.x()),
                int(rect.y()),
                max(1, int(rect.width())),
                max(1, int(rect.height())),
            )

        return QRect(0, 0, max(self.w_dpx, 1), max(self.h_dpx, 1))

    def setup_printer(self):
        """閰嶇疆鎵撳嵃鏈哄弬鏁?"""
        self.printer.setResolution(self.dpi)
        self.printer.setFullPage(False)
        if 'A4' in self.paper:
            paper_w_mm, paper_h_mm = 210.0, 297.0
        elif 'A5' in self.paper:
            paper_w_mm, paper_h_mm = 148.0, 210.0
        else:
            paper_w_mm, paper_h_mm = self.custom_w, self.custom_h

        self.printer.setPageSize(QPageSize(QSizeF(paper_w_mm, paper_h_mm), QPageSize.Unit.Millimeter))
        self.h_dpx, self.w_dpx = self.calc_size(self.dpi, paper_w_mm, paper_h_mm)

        self.printer.setPrintRange(QPrinter.PrintRange.AllPages)
        color_mode = QPrinter.ColorMode.GrayScale if self.win.ui.checkBox.isChecked() else QPrinter.ColorMode.Color
        self.printer.setColorMode(color_mode)

        duplex_idx = self.win.ui.comboBox_3.currentIndex()
        if duplex_idx == 1: self.printer.setDuplex(QPrinter.DuplexMode.DuplexLongSide)
        elif duplex_idx == 2: self.printer.setDuplex(QPrinter.DuplexMode.DuplexShortSide)
        elif duplex_idx == 3: self.printer.setDuplex(QPrinter.DuplexMode.DuplexAuto)
        else: self.printer.setDuplex(QPrinter.DuplexMode.DuplexNone)
        if self.orientation_index == 1:
            self.printer.setPageOrientation(QPageLayout.Orientation.Portrait)
        elif self.orientation_index == 2:
            self.printer.setPageOrientation(QPageLayout.Orientation.Landscape)
        
        self.printer.setCopyCount(int(self.win.ui.doubleSpinBox.value()))

    def execute(self):
        """鎵ц瀹為檯缁樺埗锛屽彲琚洿鎺ユ墦鍗板拰棰勮鍏卞悓璋冪敤"""
        self.setup_printer()
        if not self.paths:
            return
        if self.paper_index == 1:
            images = self.load_all_images(self.paths)
            self.print_a4_separated(images)
            return

        painter = QPainter()
        if not painter.begin(self.printer):
            raise RuntimeError("Printer painter initialization failed.")

        try:
            rect = self._get_paint_rect(painter)
            first_page = True

            if self.win.ui.checkBox_2.isChecked():
                images = self.load_all_images(self.paths)
                for img in images:
                    if not first_page:
                        if not self.printer.newPage():
                            break
                        rect = self._get_paint_rect(painter)
                    self.draw_single_image(img, rect, painter)
                    first_page = False
            else:
                for path in self.paths:
                    p = Path(path)
                    if p.suffix.lower() == '.pdf':
                        images = []
                        self.extract_pdf_images(p, images)
                        for img in images:
                            if not first_page:
                                if not self.printer.newPage():
                                    break
                                rect = self._get_paint_rect(painter)
                            self.draw_single_image(img, rect, painter)
                            first_page = False
                    else:
                        try:
                            with Image.open(p) as img:
                                pil_img = img.copy()
                                if not first_page:
                                    if not self.printer.newPage():
                                        break
                                    rect = self._get_paint_rect(painter)
                                self.draw_single_image(pil_img, rect, painter)
                                first_page = False
                        except Exception as e:
                            print(f"瑙ｆ瀽鍥剧墖澶辫触 {path}: {e}")
        finally:
            if painter.isActive():
                painter.end()

    def calc_size(self, dpi, w_mm, h_mm):
        h_dpx = max(1, int((h_mm / 25.4) * dpi))
        w_dpx = max(1, int((w_mm / 25.4) * dpi))
        return h_dpx, w_dpx

    def extract_pdf_images(self, path, container):
        with fitz.open(path) as doc:
            for page in doc:
                pix = page.get_pixmap(dpi=self.dpi)
                mode = "RGBA" if pix.alpha else "RGB"
                img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
                container.append(img)
        return container

    def load_all_images(self, paths):
        images = []
        for path in paths:
            p = Path(path)
            if p.suffix.lower() == '.pdf':
                self.extract_pdf_images(p, images)
            else:
                try:
                    with Image.open(p) as img: images.append(img.copy())
                except: pass
        return images

    def draw_single_image(self, pil_img, rect, painter):
        if pil_img is None:
            return

        if self.grayscale:
            pil_img = pil_img.convert("L").convert("RGBA")
        elif pil_img.mode != "RGBA":
            pil_img = pil_img.convert("RGBA")

        w, h = pil_img.size
        if w <= 0 or h <= 0:
            return

        rect_x = int(rect.x())
        rect_y = int(rect.y())
        view_w = max(1, int(rect.width()))
        view_h = max(1, int(rect.height()))

        if self.orientation_index == 0:
            if (view_h > view_w and w > h) or (view_h < view_w and w < h):
                pil_img = pil_img.transpose(Image.Transpose.ROTATE_90)
                w, h = pil_img.size

        scale = min(view_w / w, view_h / h)
        draw_w = max(1, int(w * scale))
        draw_h = max(1, int(h * scale))

        align_idx = self.win.ui.comboBox_4.currentIndex()
        if align_idx == 1:
            x = rect_x + max(view_w - draw_w, 0)
        else:
            x = rect_x + max((view_w - draw_w) // 2, 0)
        y = rect_y + max((view_h - draw_h) // 2, 0)

        print_area = QRect(int(x), int(y), int(draw_w), int(draw_h))
        q_img = ImageQt(pil_img)
        painter.drawImage(print_area, q_img)

    def print_a4_separated(self, images):
        if len(images) % 2 != 0:
            images.append(None)

        painter = QPainter()
        if not painter.begin(self.printer):
            raise RuntimeError("Printer painter initialization failed.")

        try:
            rect = self._get_paint_rect(painter)
            first_page = True
            for i in range(0, len(images), 2):
                if not first_page:
                    if not self.printer.newPage():
                        break
                    rect = self._get_paint_rect(painter)
                self.join_and_draw(images[i], images[i + 1], painter, rect)
                first_page = False
        finally:
            if painter.isActive():
                painter.end()

    def join_and_draw(self, img1, img2, painter, rect):
        if img2 is None: img2 = Image.new('RGB', img1.size, 'white')
        
        def process_half(img):
            if img.size[0] <= 0 or img.size[1] <= 0:
                return Image.new('RGB', (max(self.w_dpx, 1), max(int(self.h_dpx / 2), 1)), 'white')
            if img.size[1] / img.size[0] > 1:
                img = img.transpose(Image.Transpose.ROTATE_90)
            target_h = max(1, int(self.h_dpx / 2))
            ratio = target_h / img.size[1]
            target_w = max(1, int(img.size[0] * ratio))
            if target_w > self.w_dpx:
                target_w = self.w_dpx
                ratio = self.w_dpx / img.size[0]
                target_h = max(1, int(img.size[1] * ratio))
            return img.resize((target_w, target_h))

        img1, img2 = process_half(img1), process_half(img2)
        half_h = int(self.h_dpx / 2)
        merged = Image.new('RGB', (self.w_dpx, self.h_dpx), 'white')
        
        x1 = int((self.w_dpx - img1.size[0]) / self.inter) if img1.size[0] < self.w_dpx else 0
        y1 = int((half_h - img1.size[1]) / 2) if img1.size[1] < half_h else 0
        x2 = int((self.w_dpx - img2.size[0]) / self.inter) if img2.size[0] < self.w_dpx else 0
        y2 = int((half_h - img2.size[1]) / 2) if img2.size[1] < half_h else 0
        
        merged.paste(img1, (x1, y1))
        merged.paste(img2, (x2, half_h + y2))
        self.draw_single_image(merged, rect, painter)


# ==========================================
# 后台工作线程 (防止 UI 卡死)
# ==========================================
class PrintWorker(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(str, str)

    def __init__(self, engine: PrintEngine):
        super().__init__()
        self.engine = engine

    def run(self):
        try:
            self.progress.emit("渲染发送中，请稍候...", "#e0e0e0")
            self.engine.execute()
            self.progress.emit("文件已成功发送至打印机！", "#4cafef")
        except Exception as e:
            self.progress.emit(f"打印错误: {e}", "#ff4d4d")
        finally:
            self.finished.emit()


# ==========================================
# 主窗口入口
# ==========================================
class LuoLiPrintAssistant(QWidget):
    def __init__(self):
        super().__init__()
        self.ui = Ui_PrintForm()
        self.ui.setupUi(self)
        self.setStyleSheet(DARK_QSS)

        self.custom_width = "210"
        self.custom_height = "297"
        self.config_path = Path(__file__).resolve().parent / "printConfig.ini"
        
        self.file_paths = []
        self.load_printers()
        self.load_config()
        self.bind_events()

    def bind_events(self):
        self.ui.btn_add_file.clicked.connect(self.get_file)
        self.ui.btn_add_current_pdf.clicked.connect(self.add_current_dir_pdfs)
        self.ui.btn_add_folder_pdf.clicked.connect(self.add_folder_pdfs)
        self.ui.btn_clean_pdf.clicked.connect(self.clean_pdfs_to_temp)
        self.ui.btn_clear.clicked.connect(self.clear_files)
        
        self.ui.btn_advanced.clicked.connect(self.open_advanced_dialog)
        self.ui.btn_preview.clicked.connect(self.run_preview)
        self.ui.btn_print.clicked.connect(self.run_print_task)
        self.ui.btn_about.clicked.connect(self.open_about_dialog)
        self.ui.toolButton.clicked.connect(self.run_system_dialog)
        
        self.ui.comboBox.currentIndexChanged.connect(self.check_paper_type)

    def check_paper_type(self):
        is_a4 = self.ui.comboBox.currentText() == 'A4'
        self.ui.comboBox_6.setEnabled(is_a4)
        if not is_a4: self.ui.comboBox_6.setCurrentIndex(0)

    def load_printers(self):
        printers = QPrinterInfo.availablePrinters()
        self.ui.comboBox_5.addItems([p.printerName() for p in printers])

    # ===== 文件加载管理 =====
    def clear_files(self):
        self.file_paths.clear()
        self.ui.listWidget.clear()
        self.update_status("已清空列表。", "#e0e0e0")

    def show_list_widget(self, path):
        if path not in self.file_paths:
            self.file_paths.append(path)
            self.ui.listWidget.addItem(QListWidgetItem(Path(path).name))
        self.update_status(f"当前待打印文件数: {len(self.file_paths)}")

    def get_file(self):
        res, _ = QFileDialog.getOpenFileNames(self, '选择文件', '', 'Images & PDF (*.pdf *.jpg *.png *.jpeg *.bmp)')
        for p in res: self.show_list_widget(p)

    def add_current_dir_pdfs(self):
        for f in Path.cwd().glob("*.pdf"): self.show_list_widget(str(f))

    def add_folder_pdfs(self):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if d:
            for f in Path(d).glob("*.pdf"): self.show_list_widget(str(f))

    def clean_pdfs_to_temp(self):
        pdfs = list(Path.cwd().glob("*.pdf"))
        if not pdfs: return QMessageBox.information(self, "提示", "无 PDF 文件")
        
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        temp_dir = Path(tempfile.gettempdir()) / f"PDF_Backup_{ts}"
        temp_dir.mkdir(parents=True)
        
        moved = []
        for p in pdfs:
            dest = temp_dir / p.name
            shutil.move(str(p), str(dest))
            moved.append(dest)
            
        self.pdf_mgr = PdfManagerWindow(temp_dir, moved, Path.cwd())
        self.pdf_mgr.show()

    def open_advanced_dialog(self):
        dlg = AdvancedDialog(self)
        dlg.width_input.setText(self.custom_width)
        dlg.height_input.setText(self.custom_height)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.custom_width = dlg.width_input.text()
            self.custom_height = dlg.height_input.text()
            lines = dlg.textEdit.toPlainText().splitlines()
            for ln in lines:
                if ln.strip() and Path(ln.strip()).exists():
                    self.show_list_widget(ln.strip())

    # ===== 拖拽支持 =====
    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls(): e.acceptProposedAction()

    def dropEvent(self, e: QDropEvent):
        valid = {'.pdf', '.jpg', '.jpeg', '.png', '.bmp'}
        for url in e.mimeData().urls():
            path = Path(url.toLocalFile())
            if path.is_dir():
                for f in path.rglob("*"):
                    if f.suffix.lower() in valid: self.show_list_widget(str(f))
            elif path.suffix.lower() in valid:
                self.show_list_widget(str(path))

    # ===== 核心打印流 =====
    def update_status(self, text, color="#4cafef"):
        self.ui.label_8.setText(text)
        self.ui.label_8.setStyleSheet(f"color: {color}; font-weight: bold; font-size: 15px; padding: 6px; border-radius: 8px; background-color: #2a2d3a;")

    def open_about_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("关于")
        dlg.setStyleSheet(DARK_QSS)

        layout = QVBoxLayout(dlg)
        label = QLabel('打印助手由 <a href="https://hiluoli.cn/print-assist">hiluoli.cn</a> @Chenianlaocu开发。')
        label.setWordWrap(True)
        label.setTextFormat(Qt.TextFormat.RichText)
        label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        label.setOpenExternalLinks(True)
        layout.addWidget(label)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        btn_box.accepted.connect(dlg.accept)
        layout.addWidget(btn_box)

        dlg.exec()

    def _prepare_printer(self) -> QPrinter:
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer_name = self.ui.comboBox_5.currentText()
        if printer_name: printer.setPrinterName(printer_name)
        return printer

    def run_preview(self):
        if not self.file_paths:
            return self.update_status("No printable files.", "#ff4d4d")

        printer = self._prepare_printer()
        dialog = QPrintPreviewDialog(printer, self)
        dialog.setStyleSheet("QDialog { background: #e0e0e0; color: black; }")

        def _paint_preview(preview_printer):
            try:
                engine = PrintEngine(self, preview_printer, list(self.file_paths))
                engine.execute()
            except Exception as e:
                self.update_status(f"Preview failed: {e}", "#ff4d4d")

        dialog.paintRequested.connect(_paint_preview)
        dialog.exec()

    def run_system_dialog(self):
        if not self.file_paths:
            return self.update_status("No printable files.", "#ff4d4d")

        printer = self._prepare_printer()
        dialog = QPrintDialog(printer, self)
        if dialog.exec():
            self._run_print_now(printer)

    def run_print_task(self):
        if not self.file_paths:
            return self.update_status("No printable files.", "#ff4d4d")

        printer = self._prepare_printer()
        self._run_print_now(printer)

    def _run_print_now(self, printer):
        try:
            self.update_status("Printing...", "#e0e0e0")
            QCoreApplication.processEvents()
            engine = PrintEngine(self, printer, list(self.file_paths))
            engine.execute()
            self.update_status("Print job sent successfully.", "#4cafef")
        except Exception as e:
            self.update_status(f"Print failed: {e}", "#ff4d4d")
            QMessageBox.critical(self, "Print Error", str(e))

    def load_config(self):
        c = configparser.ConfigParser()
        try:
            c.read(self.config_path, encoding="utf-8")
        except Exception:
            return

        if not c.has_section('Print'):
            return

        self.ui.doubleSpinBox.setValue(c.getint('Print', 'Series', fallback=1))
        self.ui.comboBox.setCurrentIndex(c.getint('Print', 'Paper', fallback=0))
        self.ui.comboBox_2.setCurrentIndex(c.getint('Print', 'Dpi', fallback=1))
        self.ui.comboBox_3.setCurrentIndex(c.getint('Print', 'Double', fallback=0))
        self.ui.comboBox_4.setCurrentIndex(c.getint('Print', 'Center', fallback=0))
        self.ui.comboBox_5.setCurrentText(c.get('Print', 'PrintName', fallback=''))
        self.ui.comboBox_6.setCurrentIndex(c.getint('Print', 'PageDirection', fallback=0))
        self.ui.checkBox.setChecked(c.getboolean('Print', 'Color', fallback=False))
        self.ui.checkBox_2.setChecked(c.getboolean('Print', 'Mergebox', fallback=False))

    def closeEvent(self, e):
        c = configparser.ConfigParser()
        c['Print'] = {
            'Series': str(int(self.ui.doubleSpinBox.value())),
            'Paper': str(self.ui.comboBox.currentIndex()),
            'Dpi': str(self.ui.comboBox_2.currentIndex()),
            'Double': str(self.ui.comboBox_3.currentIndex()),
            'Center': str(self.ui.comboBox_4.currentIndex()),
            'PrintName': self.ui.comboBox_5.currentText(),
            'PageDirection': str(self.ui.comboBox_6.currentIndex()),
            'Color': str(self.ui.checkBox.isChecked()),
            'Mergebox': str(self.ui.checkBox_2.isChecked())
        }

        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                c.write(f)
        except Exception as ex:
            print(f"save config failed: {ex}")

        e.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create('Fusion'))
    win = LuoLiPrintAssistant()
    win.show()
    sys.exit(app.exec())
