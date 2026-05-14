"""Main desktop window."""

import os
import subprocess
import sys
import threading
from pathlib import Path

from PyQt5.QtCore import (
    QSize,
    QTimer,
    Qt,
    pyqtSignal,
)
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QListWidgetItem,
    QMainWindow,
    QVBoxLayout,
    QWidget,
)

from qfluentwidgets import (
    BodyLabel,
    CaptionLabel,
    ComboBox,
    ElevatedCardWidget,
    SimpleCardWidget,
    FluentIcon as FIF,
    InfoBar,
    LineEdit,
    ListWidget,
    PrimaryPushButton,
    ProgressBar,
    PushButton,
    StrongBodyLabel,
    SubtitleLabel,
    Theme,
    TitleLabel,
    TransparentToolButton,
    setTheme,
    setThemeColor,
)

from converters.exceptions import ConversionCancelled


class PdfQueueListWidget(ListWidget):
    """Drag-and-drop PDF queue."""

    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setAlternatingRowColors(False)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event):
        if not event.mimeData().hasUrls():
            super().dropEvent(event)
            return

        paths = []
        for url in event.mimeData().urls():
            if not url.isLocalFile():
                continue
            local_path = Path(url.toLocalFile())
            if local_path.is_dir():
                paths.extend(str(path) for path in sorted(local_path.rglob("*.pdf")))
            elif local_path.suffix.lower() == ".pdf":
                paths.append(str(local_path))

        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
            return

        super().dropEvent(event)


class MainWindow(QMainWindow):
    """PDF conversion desktop window."""

    progress_signal = pyqtSignal(float)
    status_signal = pyqtSignal(str)
    current_task_signal = pyqtSignal(str)
    queue_item_status_signal = pyqtSignal(str, str)
    queue_item_progress_signal = pyqtSignal(str, float)
    output_file_signal = pyqtSignal(str)
    done_signal = pyqtSignal()

    MODE_TABLE = "table"
    MODE_WORD = "word"

    def __init__(self):
        super().__init__()
        self.selected_pdf_paths = []
        self.file_statuses = {}
        self.output_manual = False
        self.manual_output_target = ""
        self.is_running = False
        self.stop_requested = False
        self.current_task_name = ""
        self.current_output_path = ""
        self._dpi_value = 300
        self._output_files = []

        # 禁止 QFluentWidgets 默认的窗口拖拽行为
        self.setProperty("isDraggable", False)

        self.setWindowTitle("PDF 批量转换工作台")
        self.resize(1380, 860)
        self.setMinimumSize(1180, 760)

        setTheme(Theme.LIGHT)
        setThemeColor(QColor(56, 189, 248))

        self._init_ui()
        self._connect_signals()
        self._sync_mode_ui()
        self._refresh_queue()
        self._refresh_output_display()
        self._refresh_action_state()

    def _init_ui(self):
        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)
        self.setStyleSheet(
            """
            QWidget#centralWidget {
                background-color: #eef2f7;
            }
            ElevatedCardWidget, SimpleCardWidget {
                background-color: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 16px;
            }
            QFrame[section="flat"] {
                background-color: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
            """
        )

        root = QVBoxLayout(central)
        root.setContentsMargins(20, 10, 20, 16)
        root.setSpacing(12)

        root.addWidget(self._build_header())
        root.addWidget(self._build_control_panel(), 0)
        root.addWidget(self._build_queue_panel(), 1)

    def _build_header(self):
        card = SimpleCardWidget(self)
        layout = QHBoxLayout(card)
        layout.setContentsMargins(16, 5, 16, 5)
        layout.setSpacing(10)

        title_wrap = QVBoxLayout()
        title_wrap.setSpacing(1)

        self.title_label = TitleLabel("PDF 批量转换")
        self.title_label.setStyleSheet("color: #0f172a; font-size: 17px; font-weight: 700;")
        self.subtitle_label = BodyLabel("拖拽建队列，顺序执行，支持 OCR 转 Excel / Word")
        self.subtitle_label.setStyleSheet("color: #94a3b8; font-size: 11px;")
        title_wrap.addWidget(self.title_label)
        title_wrap.addWidget(self.subtitle_label)

        layout.addLayout(title_wrap, 1)

        header_info = QVBoxLayout()
        header_info.setSpacing(4)
        self.header_summary_label = CaptionLabel("队列 0 个文件  ·  当前模式 Excel 表格  ·  状态 空闲")
        self.header_summary_label.setStyleSheet("color: #94a3b8; font-size: 11px;")
        header_info.addWidget(self.header_summary_label)
        layout.addLayout(header_info)

        return card

    def _build_queue_panel(self):
        card = SimpleCardWidget(self)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(8)

        # 标题行
        header = QHBoxLayout()
        title = SubtitleLabel("任务队列")
        title.setStyleSheet("color: #0f172a; font-size: 17px; font-weight: 650;")
        header.addWidget(title)
        header.addStretch()
        layout.addLayout(header)

        # 工具栏
        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)
        self.add_files_btn = PrimaryPushButton("添加文件")
        self.add_folder_btn = PushButton("添加文件夹")
        self.move_up_btn = PushButton("上移")
        self.move_down_btn = PushButton("下移")
        self.add_files_btn.setMinimumHeight(32)
        self.add_folder_btn.setMinimumHeight(32)
        self.move_up_btn.setMinimumHeight(30)
        self.move_down_btn.setMinimumHeight(30)
        self.add_files_btn.setMinimumWidth(108)
        self.add_folder_btn.setMinimumWidth(108)
        self.move_up_btn.setMinimumWidth(82)
        self.move_down_btn.setMinimumWidth(82)
        self.remove_btn = TransparentToolButton(FIF.DELETE)
        self.remove_btn.setToolTip("删除选中")
        self.clear_btn = TransparentToolButton(FIF.BROOM)
        self.clear_btn.setToolTip("清理队列")
        toolbar.addWidget(self.add_files_btn)
        toolbar.addWidget(self.add_folder_btn)
        toolbar.addWidget(self.move_up_btn)
        toolbar.addWidget(self.move_down_btn)
        toolbar.addWidget(self.remove_btn)
        toolbar.addWidget(self.clear_btn)
        toolbar.addStretch()
        layout.addLayout(toolbar)

        # 左右并排：输入PDF列表 + 输出文件列表
        lists_row = QHBoxLayout()
        lists_row.setSpacing(14)

        # 左侧：选择的PDF
        input_col = QVBoxLayout()
        input_col.setSpacing(6)
        input_header = QHBoxLayout()
        input_header.setContentsMargins(0, 0, 0, 0)
        input_title = CaptionLabel("输入文件")
        input_title.setStyleSheet("color: #475569; font-size: 12px; font-weight: 600;")
        self.input_badge = CaptionLabel("0 个文件")
        self.input_badge.setStyleSheet(
            """
            background-color: #dbeafe;
            color: #1d4ed8;
            border: 1px solid #93c5fd;
            border-radius: 12px;
            padding: 3px 8px;
            font-weight: 600;
            font-size: 11px;
            """
        )
        input_header.addWidget(input_title)
        input_header.addStretch()
        input_header.addWidget(self.input_badge)
        input_header_widget = QWidget()
        input_header_widget.setFixedHeight(32)
        input_header_widget.setLayout(input_header)
        input_col.addWidget(input_header_widget)

        self.queue_list = PdfQueueListWidget(self)
        self.queue_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.queue_list.setIconSize(QSize(0, 0))
        self.queue_list.setStyleSheet(
            """
            ListWidget {
                background-color: #f1f5f9;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
                padding: 8px;
                color: #0f172a;
                font-size: 14px;
            }
            ListWidget::item {
                background-color: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                padding: 6px 10px;
                margin: 3px 2px;
            }
            ListWidget::item:selected {
                background-color: #dbeafe;
                border: 1px solid #60a5fa;
            }
            """
        )
        input_col.addWidget(self.queue_list, 1)
        lists_row.addLayout(input_col, 1)

        # 中间箭头（带脉冲动画）
        arrow_container = QWidget()
        arrow_container.setFixedWidth(40)
        arrow_container_layout = QVBoxLayout(arrow_container)
        arrow_container_layout.setContentsMargins(0, 0, 0, 0)
        arrow_container_layout.setAlignment(Qt.AlignCenter)
        self.arrow_label = QLabel("▸▸▸")
        self.arrow_label.setAlignment(Qt.AlignCenter)
        self.arrow_label.setFixedWidth(30)
        self.arrow_label.setStyleSheet("color: #94a3b8; font-size: 22px; font-weight: 700;")
        arrow_container_layout.addWidget(self.arrow_label)
        lists_row.addWidget(arrow_container)

        # 箭头动画 timer
        self._arrow_timer = QTimer(self)
        self._arrow_timer.setInterval(500)
        self._arrow_step = 0
        self._arrow_timer.timeout.connect(self._animate_arrow)

        # 右侧：输出文件
        output_col = QVBoxLayout()
        output_col.setSpacing(6)

        output_header = QHBoxLayout()
        output_header.setContentsMargins(0, 0, 0, 0)
        output_title = CaptionLabel("输出文件")
        output_title.setStyleSheet("color: #475569; font-size: 12px; font-weight: 600;")
        self.output_badge = CaptionLabel("0 个文件")
        self.output_badge.setStyleSheet(
            """
            background-color: #dcfce7;
            color: #15803d;
            border: 1px solid #86efac;
            border-radius: 12px;
            padding: 3px 8px;
            font-weight: 600;
            font-size: 11px;
            """
        )
        output_header.addWidget(output_title)
        output_header.addStretch()
        output_header.addWidget(self.output_badge)
        self.clear_output_btn = TransparentToolButton(FIF.BROOM)
        self.clear_output_btn.setToolTip("清理输出列表")
        output_header.addWidget(self.clear_output_btn)
        output_header_widget = QWidget()
        output_header_widget.setFixedHeight(32)
        output_header_widget.setLayout(output_header)
        output_col.addWidget(output_header_widget)

        self.output_list = ListWidget(self)
        self.output_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.output_list.setIconSize(QSize(0, 0))
        self.output_list.setStyleSheet(
            """
            ListWidget {
                background-color: #f1f5f9;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
                padding: 8px;
                color: #0f172a;
                font-size: 14px;
            }
            ListWidget::item {
                background-color: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                padding: 6px 10px;
                margin: 3px 2px;
            }
            ListWidget::item:selected {
                background-color: #dbeafe;
                border: 1px solid #60a5fa;
            }
            """
        )
        output_col.addWidget(self.output_list, 1)

        lists_row.addLayout(output_col, 1)
        layout.addLayout(lists_row, 1)

        # 底部摘要
        self.queue_summary_label = CaptionLabel("当前队列为空，请先添加 PDF 文件。支持拖入多个 PDF 或文件夹。")
        self.queue_summary_label.setWordWrap(True)
        self.queue_summary_label.setStyleSheet("color: #64748b; font-size: 12px;")
        layout.addWidget(self.queue_summary_label)

        return card

    def _build_control_panel(self):
        card = SimpleCardWidget(self)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(10)

        header = QHBoxLayout()
        title = SubtitleLabel("任务控制")
        title.setStyleSheet("color: #0f172a; font-size: 17px; font-weight: 650;")
        hint = CaptionLabel("按队列顺序执行")
        hint.setStyleSheet("color: #94a3b8;")
        header.addWidget(title)
        header.addStretch()
        header.addWidget(hint)
        layout.addLayout(header)

        # 状态行（直接放在卡片内，不再套 info_strip）
        status_row = QHBoxLayout()
        status_row.setContentsMargins(0, 0, 0, 0)
        status_row.setSpacing(12)
        self.status_badge = QLabel("空闲")
        self.status_badge.setAlignment(Qt.AlignCenter)
        self.status_badge.setFixedHeight(24)
        self.status_badge.setMinimumWidth(60)
        self.status_badge.setStyleSheet(
            """
            QLabel {
                background-color: #dbeafe;
                border: 1px solid #93c5fd;
                border-radius: 12px;
                color: #1d4ed8;
                padding: 5px 10px;
                font-size: 12px;
                font-weight: 700;
            }
            """
        )
        status_text_col = QVBoxLayout()
        status_text_col.setContentsMargins(0, 0, 0, 0)
        status_text_col.setSpacing(2)
        self.current_file_label = CaptionLabel("当前任务：未开始")
        self.current_file_label.setMinimumWidth(0)
        self.current_file_label.setStyleSheet("color: #0f172a; font-size: 13px; font-weight: 600;")
        self.stats_summary_label = QLabel(
            '<span style="color:#15803d;font-weight:600;">成功 0</span>'
            ' · <span style="color:#dc2626;font-weight:600;">失败 0</span>'
            ' · <span style="color:#64748b;font-weight:600;">待处理 0</span>'
            ' · <span style="color:#1d4ed8;font-weight:600;">处理中 0</span>'
        )
        self.stats_summary_label.setStyleSheet("font-size: 12px;")
        status_text_col.addWidget(self.current_file_label)
        status_text_col.addWidget(self.stats_summary_label)
        status_row.addWidget(self.status_badge)
        status_row.addLayout(status_text_col, 1)
        layout.addLayout(status_row)

        # 分隔线
        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setStyleSheet("background-color: #e2e8f0; max-height: 1px; border: none;")
        layout.addWidget(divider)

        # 设置标题
        settings_title = CaptionLabel("转换设置")
        settings_title.setStyleSheet("color: #64748b; font-size: 12px; font-weight: 600;")
        layout.addWidget(settings_title)

        settings_grid = QGridLayout()
        settings_grid.setContentsMargins(0, 0, 0, 0)
        settings_grid.setHorizontalSpacing(16)
        settings_grid.setVerticalSpacing(10)

        mode_label = self._form_label("转换模式")
        self.mode_combo = ComboBox()
        self.mode_combo.addItems(["PDF 表格转 Excel", "PDF 转 Word"])
        self.mode_combo.setMinimumHeight(36)
        settings_grid.addWidget(mode_label, 0, 0)
        settings_grid.addWidget(self.mode_combo, 0, 1)

        quality_label = self._form_label("质量预设")
        self.quality_combo = ComboBox()
        self.quality_combo.addItems(["高质量，速度慢", "低质量，速度快"])
        self.quality_combo.setCurrentIndex(0)
        self.quality_combo.setMinimumHeight(36)
        settings_grid.addWidget(quality_label, 0, 2)
        settings_grid.addWidget(self.quality_combo, 0, 3)

        output_label = self._form_label("输出位置")
        self.out_btn = PushButton("选择输出")
        self.out_btn.setMinimumHeight(36)
        self.out_btn.setMinimumWidth(110)
        settings_grid.addWidget(output_label, 1, 0)
        settings_grid.addWidget(self.out_btn, 1, 1)

        self.out_edit = LineEdit()
        self.out_edit.setReadOnly(True)
        self.out_edit.setMinimumHeight(36)
        self.out_edit.setPlaceholderText("默认输出到源文件目录，或手动指定统一输出目录...")
        settings_grid.addWidget(self.out_edit, 1, 2, 1, 2)
        settings_grid.setColumnStretch(1, 1)
        settings_grid.setColumnStretch(3, 1)
        layout.addLayout(settings_grid)

        action_row = QHBoxLayout()
        action_row.setSpacing(12)
        self.start_btn = PrimaryPushButton("开始转换")
        self.start_btn.setMinimumHeight(38)
        self.start_btn.setStyleSheet(
            """
            PrimaryPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #3b82f6, stop:1 #6366f1);
                color: white;
                font-size: 15px;
                font-weight: 700;
                border-radius: 14px;
                border: none;
            }
            PrimaryPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #60a5fa, stop:1 #818cf8);
            }
            PrimaryPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #2563eb, stop:1 #4f46e5);
            }
            PrimaryPushButton:disabled {
                background-color: #e2e8f0;
                color: #94a3b8;
                border: none;
            }
            """
        )
        self.stop_btn = PushButton("停止")
        self.stop_btn.setMinimumHeight(38)
        self.stop_btn.setMinimumWidth(100)
        self.stop_btn.setStyleSheet(
            """
            PushButton {
                background-color: #fef2f2;
                color: #dc2626;
                font-size: 15px;
                font-weight: 700;
                border-radius: 14px;
                border: 1px solid #fecaca;
            }
            PushButton:hover {
                background-color: #fee2e2;
            }
            PushButton:pressed {
                background-color: #fecaca;
            }
            PushButton:disabled {
                background-color: #f1f5f9;
                color: #94a3b8;
                border: 1px solid #e2e8f0;
            }
            """
        )
        action_row.addWidget(self.start_btn, 1)
        action_row.addWidget(self.stop_btn, 1)
        layout.addLayout(action_row)

        progress_row = QHBoxLayout()
        progress_row.setSpacing(8)
        progress_title = CaptionLabel("总进度")
        progress_title.setStyleSheet("color: #475569; font-size: 12px; font-weight: 600;")
        self.progress_text_label = StrongBodyLabel("0%")
        self.progress_text_label.setStyleSheet("color: #0f172a; font-size: 13px;")
        progress_row.addWidget(progress_title)
        progress_row.addStretch()
        progress_row.addWidget(self.progress_text_label)
        layout.addLayout(progress_row)

        self.progress_bar = ProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(6)
        layout.addWidget(self.progress_bar)

        return card

    def _create_metric_card(self, value, label, accent):
        card = QFrame(self)
        card.setStyleSheet(
            f"""
            QFrame {{
                background-color: transparent;
                border: none;
            }}
            """
        )
        layout = QVBoxLayout(card)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        value_label = QLabel(value)
        value_label.setStyleSheet(f"color: {accent}; font-size: 20px; font-weight: 700;")
        text_label = QLabel(label)
        text_label.setStyleSheet("color: #94a3b8; font-size: 12px; font-weight: 600;")
        layout.addWidget(value_label)
        layout.addWidget(text_label)

        return {"card": card, "value": value_label, "label": text_label}

    def _flat_section(self):
        card = QFrame(self)
        card.setProperty("section", "flat")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)
        return {"card": card, "layout": layout}

    def _info_strip(self):
        card = QFrame(self)
        card.setStyleSheet(
            """
            QFrame {
                background-color: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
            """
        )
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)
        return {"card": card, "layout": layout}

    def _form_label(self, title):
        label = CaptionLabel(title)
        label.setMinimumWidth(72)
        label.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        label.setStyleSheet("color: #94a3b8; font-size: 12px; font-weight: 600;")
        return label

    def _simple_field(self, title):
        card = QFrame(self)
        card.setStyleSheet(
            """
            QFrame {
                background-color: transparent;
                border: none;
            }
            """
        )
        layout = QVBoxLayout(card)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        label = CaptionLabel(title)
        label.setStyleSheet("color: #475569; font-size: 12px; font-weight: 600;")
        layout.addWidget(label)
        return {"card": card, "layout": layout, "label": label}

    def _section_caption(self, title):
        label = CaptionLabel(title)
        label.setStyleSheet("color: #475569; font-size: 12px; font-weight: 600;")
        return label

    def _connect_signals(self):
        self.mode_combo.currentIndexChanged.connect(self._on_mode_change)
        self.quality_combo.currentIndexChanged.connect(self._on_quality_change)
        self.add_files_btn.clicked.connect(self._select_pdf_files)
        self.add_folder_btn.clicked.connect(self._select_folder)
        self.move_up_btn.clicked.connect(lambda: self._move_selected(-1))
        self.move_down_btn.clicked.connect(lambda: self._move_selected(1))
        self.remove_btn.clicked.connect(self._remove_selected_files)
        self.clear_btn.clicked.connect(self._clear_queue)
        self.out_btn.clicked.connect(self._select_output)
        self.start_btn.clicked.connect(self._start_conversion)
        self.stop_btn.clicked.connect(self._request_stop)
        self.queue_list.files_dropped.connect(self._add_pdf_paths)
        self.queue_list.itemSelectionChanged.connect(self._refresh_action_state)
        self.progress_signal.connect(self._update_progress)
        self.status_signal.connect(self._update_status)
        self.current_task_signal.connect(self._update_current_task)
        self.queue_item_status_signal.connect(self._set_queue_item_status)
        self.queue_item_progress_signal.connect(self._set_queue_item_progress)
        self.output_file_signal.connect(self._add_output_file)
        self.clear_output_btn.clicked.connect(self._clear_output_list)
        self.done_signal.connect(self._on_done)

    def _current_mode(self):
        return self.MODE_WORD if self.mode_combo.currentIndex() == 1 else self.MODE_TABLE

    def _output_suffix_for_mode(self, mode):
        return ".docx" if mode == self.MODE_WORD else ".xlsx"

    def _default_output_path(self, pdf_path, mode=None, output_dir=None):
        suffix = self._output_suffix_for_mode(mode or self._current_mode())
        output_path = Path(pdf_path).with_suffix(suffix)
        if output_dir:
            output_path = Path(output_dir) / output_path.name
        return str(output_path)

    def _resolve_output_path(self, pdf_path, mode, output_target=None, multi=False):
        if multi:
            return self._default_output_path(pdf_path, mode=mode, output_dir=output_target)
        if output_target:
            target = Path(output_target)
            suffix = self._output_suffix_for_mode(mode)
            if target.suffix.lower() != suffix:
                target = target.with_suffix(suffix)
            return str(target)
        return self._default_output_path(pdf_path, mode=mode)

    def _batch_output_paths(self, pdf_paths, mode, output_target=None):
        """预计算所有输出路径，同名文件全部追加父目录名。"""
        multi = len(pdf_paths) > 1
        suffix = self._output_suffix_for_mode(mode)

        # 第一遍：检测哪些 stem 有重复
        from collections import Counter
        stems = [Path(p).stem for p in pdf_paths]
        stem_counts = Counter(stems)
        dup_stems = {s for s, c in stem_counts.items() if c > 1}

        # 第二遍：生成路径，重名的全部加父目录名
        result = []
        used = set()
        for pdf_path in pdf_paths:
            if multi and output_target:
                stem = Path(pdf_path).stem
                if stem in dup_stems:
                    parent_name = Path(pdf_path).parent.name
                    out = Path(output_target) / f"{stem}_{parent_name}{suffix}"
                    # 极端情况：父目录名也相同，追加数字
                    counter = 2
                    while str(out) in used or out.exists():
                        out = Path(output_target) / f"{stem}_{parent_name}_{counter}{suffix}"
                        counter += 1
                else:
                    out = Path(output_target) / f"{stem}{suffix}"
                result.append(str(out))
            else:
                out = self._resolve_output_path(pdf_path, mode, output_target=output_target, multi=multi)
                result.append(out)
            used.add(result[-1])
        return result

    def _check_cancelled(self):
        if self.stop_requested:
            raise ConversionCancelled("用户停止了任务")

    def _queue_status_meta(self, status):
        meta = {
            "waiting": ("待处理", "#64748b", "#f1f5f9", "#e2e8f0"),
            "running": ("处理中", "#1d4ed8", "#dbeafe", "#93c5fd"),
            "success": ("已完成", "#15803d", "#dcfce7", "#86efac"),
            "failed": ("失败", "#dc2626", "#fee2e2", "#fca5a5"),
            "stopped": ("已停止", "#b45309", "#fef3c7", "#fde68a"),
        }
        return meta.get(status, meta["waiting"])

    def _shorten_path(self, text, max_len=52):
        if len(text) <= max_len:
            return text
        part = max_len // 2 - 3
        return f"{text[:part]}...{text[-part:]}"

    def _create_queue_item_widget(self, index, pdf_path, status):
        path = Path(pdf_path)
        status_text, fg, bg, border = self._queue_status_meta(status)

        widget = QWidget(self.queue_list)
        widget.setFixedHeight(62)
        outer = QVBoxLayout(widget)
        outer.setContentsMargins(8, 4, 8, 4)
        outer.setSpacing(4)

        root = QHBoxLayout()
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        index_label = QLabel(f"{index:02d}")
        index_label.setAlignment(Qt.AlignCenter)
        index_label.setFixedSize(30, 30)
        index_label.setStyleSheet(
            """
            QLabel {
                color: #1d4ed8;
                background-color: #dbeafe;
                border: 1px solid #93c5fd;
                border-radius: 15px;
                font-size: 12px;
                font-weight: 700;
            }
            """
        )

        text_col = QVBoxLayout()
        text_col.setContentsMargins(0, 0, 0, 0)
        text_col.setSpacing(2)

        name_label = StrongBodyLabel(self._shorten_path(path.name, max_len=34))
        name_label.setMinimumWidth(0)
        name_label.setToolTip(path.name)
        name_label.setStyleSheet("color: #0f172a; font-size: 13px;")
        folder_label = CaptionLabel(self._shorten_path(str(path.parent), max_len=42))
        folder_label.setMinimumWidth(0)
        folder_label.setStyleSheet("color: #64748b; font-size: 11px;")
        folder_label.setToolTip(str(path.parent))
        text_col.addWidget(name_label)
        text_col.addWidget(folder_label)

        status_label = QLabel(status_text)
        status_label.setAlignment(Qt.AlignCenter)
        status_label.setFixedSize(68, 24)
        status_label.setStyleSheet(
            f"""
            QLabel {{
                color: {fg};
                background-color: {bg};
                border: 1px solid {border};
                border-radius: 12px;
                padding: 0;
                font-size: 11px;
                font-weight: 700;
            }}
            """
        )

        root.addWidget(index_label, 0, Qt.AlignVCenter)
        root.addLayout(text_col, 1)
        root.addWidget(status_label, 0, Qt.AlignVCenter)
        outer.addLayout(root)

        # 进度条
        progress_bar = ProgressBar()
        progress_bar.setFixedHeight(4)
        progress_bar.setValue(100 if status == "success" else 0)
        progress_bar.setVisible(status in ("running", "success"))
        outer.addWidget(progress_bar)

        # 保存进度条引用
        widget.setProperty("progress_bar", progress_bar)
        widget.setToolTip(pdf_path)
        return widget

    def _update_task_metrics(self):
        paths = self.selected_pdf_paths
        counts = {"waiting": 0, "running": 0, "success": 0, "failed": 0, "stopped": 0}
        for path in paths:
            status = self.file_statuses.get(path, "waiting")
            if status in counts:
                counts[status] += 1
            else:
                counts["waiting"] += 1

        if hasattr(self, "stats_summary_label"):
            self.stats_summary_label.setText(
                f'<span style="color:#15803d;font-weight:600;">成功 {counts["success"]}</span>'
                f' · <span style="color:#dc2626;font-weight:600;">失败 {counts["failed"]}</span>'
                f' · <span style="color:#64748b;font-weight:600;">待处理 {counts["waiting"]}</span>'
                f' · <span style="color:#1d4ed8;font-weight:600;">处理中 {counts["running"]}</span>'
            )

    def _set_queue_item_status(self, pdf_path, status):
        if pdf_path not in self.selected_pdf_paths:
            return
        self.file_statuses[pdf_path] = status
        self._refresh_queue()

    def _set_queue_item_progress(self, pdf_path, progress):
        """更新单个文件的进度条（不刷新整个队列）。"""
        if pdf_path not in self.selected_pdf_paths:
            return
        index = self.selected_pdf_paths.index(pdf_path)
        item = self.queue_list.item(index)
        if not item:
            return
        widget = self.queue_list.itemWidget(item)
        if not widget:
            return
        progress_bar = widget.property("progress_bar")
        if progress_bar:
            progress_bar.setVisible(True)
            progress_bar.setValue(int(progress * 100))

    def _refresh_queue(self):
        self.file_statuses = {path: self.file_statuses.get(path, "waiting") for path in self.selected_pdf_paths}
        self.queue_list.clear()
        for index, pdf_path in enumerate(self.selected_pdf_paths, 1):
            item = QListWidgetItem()
            item.setToolTip(pdf_path)
            item.setData(Qt.UserRole, pdf_path)
            item.setSizeHint(QSize(0, 68))
            self.queue_list.addItem(item)
            status = self.file_statuses.get(pdf_path, "waiting")
            self.queue_list.setItemWidget(item, self._create_queue_item_widget(index, pdf_path, status))

        count = len(self.selected_pdf_paths)
        self.input_badge.setText(f"{count} 个文件")
        if count == 0:
            summary = "当前队列为空，请先添加 PDF 文件。支持拖入多个 PDF 或文件夹。"
        elif count == 1:
            summary = f"当前已添加 1 个文件，输出文件将自动按当前模式生成。"
        else:
            summary = f"当前已添加 {count} 个文件，可拖拽补充，也可通过上移/下移调整执行顺序。"
        self.queue_summary_label.setText(summary)
        self._refresh_header_summary()
        self._update_task_metrics()
        self._refresh_action_state()

    def _refresh_output_display(self):
        multi = len(self.selected_pdf_paths) > 1
        if multi:
            self.out_btn.setText("选择输出目录")
            self.out_edit.setPlaceholderText("默认输出到各 PDF 原目录，或指定统一输出目录...")
        else:
            self.out_btn.setText("选择输出路径")
            self.out_edit.setPlaceholderText("单文件可单独指定输出文件...")

        if self.output_manual and self.manual_output_target:
            text = self.manual_output_target
            if not multi and self.selected_pdf_paths:
                text = self._resolve_output_path(
                    self.selected_pdf_paths[0],
                    self._current_mode(),
                    output_target=self.manual_output_target,
                    multi=False,
                )
                self.manual_output_target = text
            self.out_edit.setText(text)
            self.out_edit.setToolTip(text)
            return

        if not self.selected_pdf_paths:
            self.out_edit.clear()
            self.out_edit.setToolTip("")
            return

        if len(self.selected_pdf_paths) == 1:
            text = self._default_output_path(self.selected_pdf_paths[0])
        else:
            parents = {str(Path(path).parent) for path in self.selected_pdf_paths}
            text = f"自动输出到：{parents.pop()}" if len(parents) == 1 else "自动输出到各 PDF 所在目录"
        self.out_edit.setText(text)
        self.out_edit.setToolTip(text)

    def _refresh_action_state(self):
        has_files = bool(self.selected_pdf_paths)
        selected_rows = sorted(self.queue_list.row(item) for item in self.queue_list.selectedItems())
        has_selection = bool(selected_rows)
        idle = not self.is_running

        self.add_files_btn.setEnabled(idle)
        self.add_folder_btn.setEnabled(idle)
        self.move_up_btn.setEnabled(idle and has_selection and selected_rows[0] > 0)
        self.move_down_btn.setEnabled(idle and has_selection and selected_rows[-1] < len(self.selected_pdf_paths) - 1)
        self.remove_btn.setEnabled(idle and has_selection)
        self.clear_btn.setEnabled(idle and has_files)
        self.mode_combo.setEnabled(idle)
        self.quality_combo.setEnabled(idle)
        self.out_btn.setEnabled(idle and has_files)
        self.start_btn.setEnabled(idle and has_files)
        self.stop_btn.setEnabled(self.is_running)

    def _add_output_file(self, output_path):
        """向输出文件列表添加一个已完成的文件，格式与输入文件一致。"""
        if not hasattr(self, "_output_files"):
            self._output_files = []
        self._output_files.append(output_path)
        index = len(self._output_files)
        path = Path(output_path)

        item = QListWidgetItem()
        item.setData(Qt.UserRole, output_path)
        item.setSizeHint(QSize(0, 68))
        self.output_list.addItem(item)

        widget = QWidget(self.output_list)
        widget.setFixedHeight(62)
        root = QHBoxLayout(widget)
        root.setContentsMargins(8, 0, 8, 0)
        root.setSpacing(12)

        index_label = QLabel(f"{index:02d}")
        index_label.setAlignment(Qt.AlignCenter)
        index_label.setFixedSize(30, 30)
        index_label.setStyleSheet(
            """
            QLabel {
                color: #15803d;
                background-color: #dcfce7;
                border: 1px solid #86efac;
                border-radius: 15px;
                font-size: 12px;
                font-weight: 700;
            }
            """
        )

        text_col = QVBoxLayout()
        text_col.setContentsMargins(0, 0, 0, 0)
        text_col.setSpacing(3)

        name_label = StrongBodyLabel(self._shorten_path(path.name, max_len=28))
        name_label.setMinimumWidth(0)
        name_label.setToolTip(str(path))
        name_label.setStyleSheet("color: #0f172a; font-size: 14px;")
        folder_label = CaptionLabel(self._shorten_path(str(path.parent), max_len=36))
        folder_label.setMinimumWidth(0)
        folder_label.setStyleSheet("color: #64748b; font-size: 12px;")
        folder_label.setToolTip(str(path.parent))
        text_col.addWidget(name_label)
        text_col.addWidget(folder_label)

        open_file_btn = TransparentToolButton(FIF.DOCUMENT)
        open_file_btn.setToolTip("打开文件")
        open_file_btn.setFixedSize(30, 30)
        open_file_btn.clicked.connect(lambda checked, p=output_path: self._open_output_file(p))

        open_folder_btn = TransparentToolButton(FIF.FOLDER)
        open_folder_btn.setToolTip("打开文件夹")
        open_folder_btn.setFixedSize(30, 30)
        open_folder_btn.clicked.connect(lambda checked, p=output_path: self._open_output_folder(p))

        root.addWidget(index_label, 0, Qt.AlignVCenter)
        root.addLayout(text_col, 1)
        root.addWidget(open_file_btn, 0, Qt.AlignVCenter)
        root.addWidget(open_folder_btn, 0, Qt.AlignVCenter)

        self.output_list.setItemWidget(item, widget)
        self.output_badge.setText(f"{len(self._output_files)} 个文件")

    def _open_output_file(self, file_path):
        if not file_path or not Path(file_path).exists():
            InfoBar.warning("提示", "文件不存在。", parent=self, duration=3000)
            return
        if sys.platform == "darwin":
            subprocess.Popen(["open", file_path])
        elif sys.platform == "win32":
            os.startfile(file_path)
        else:
            subprocess.Popen(["xdg-open", file_path])

    def _open_output_folder(self, file_path):
        if not file_path:
            return
        folder = str(Path(file_path).parent)
        if not Path(folder).exists():
            InfoBar.warning("提示", "文件夹不存在。", parent=self, duration=3000)
            return
        if sys.platform == "darwin":
            subprocess.Popen(["open", folder])
        elif sys.platform == "win32":
            subprocess.Popen(["explorer", folder])
        else:
            subprocess.Popen(["xdg-open", folder])

    def _clear_output_list(self):
        self.output_list.clear()
        self._output_files = []
        self.output_badge.setText("0 个文件")

    def _animate_arrow(self):
        """箭头脉冲动画：在运行时循环切换样式。"""
        frames = [ "▸", "▸▸", "▸▸▸"]
        colors = ["#3b82f6", "#60a5fa", "#93c5fd", "#60a5fa", "#3b82f6"]
        self._arrow_step = (self._arrow_step + 1) % len(frames)
        self.arrow_label.setText(frames[self._arrow_step])
        self.arrow_label.setStyleSheet(
            f"color: {colors[self._arrow_step]}; font-size: 22px; font-weight: 700;"
        )

    def _start_arrow_animation(self):
        self._arrow_step = 0
        self._arrow_timer.start()

    def _stop_arrow_animation(self):
        self._arrow_timer.stop()
        self.arrow_label.setText("→")
        self.arrow_label.setStyleSheet("color: #94a3b8; font-size: 22px; font-weight: 700;")

    def _update_progress(self, val):
        self.progress_bar.setValue(int(val * 100))
        self.progress_text_label.setText(f"{int(val * 100)}%")

    def _update_status(self, status):
        label_map = {
            "Idle": "空闲",
            "Ready": "就绪",
            "Running": "运行中",
            "Stopping": "停止中",
            "Stopped": "已停止",
            "Finished": "已完成",
            "Error": "异常",
        }
        display = label_map.get(status, status)
        self.status_badge.setText(display)
        self._refresh_header_summary(status_text=display)
        palette = {
            "Idle": ("#1d4ed8", "#dbeafe", "#93c5fd"),
            "Ready": ("#6d28d9", "#ede9fe", "#c4b5fd"),
            "Running": ("#15803d", "#dcfce7", "#86efac"),
            "Stopping": ("#b45309", "#fef3c7", "#fde68a"),
            "Stopped": ("#dc2626", "#fee2e2", "#fca5a5"),
            "Finished": ("#15803d", "#dcfce7", "#86efac"),
            "Error": ("#dc2626", "#fee2e2", "#fca5a5"),
        }
        fg, bg, border = palette.get(status, palette["Idle"])
        self.status_badge.setStyleSheet(
            f"""
            QLabel {{
                background-color: {bg};
                border: 1px solid {border};
                border-radius: 12px;
                color: {fg};
                padding: 5px 10px;
                font-size: 12px;
                font-weight: 700;
            }}
            """
        )

    def _update_current_task(self, text):
        raw_text = (text or "未开始").strip()
        display_text = self._shorten_path(raw_text.replace("\n", "  →  "), max_len=86)
        self.current_file_label.setText(f"当前任务：{display_text}")
        self.current_file_label.setToolTip(raw_text)

    def _on_mode_change(self, _index):
        self._sync_mode_ui()

    def _refresh_header_summary(self, status_text=None):
        mode_text = "Word 文档" if self._current_mode() == self.MODE_WORD else "Excel 表格"
        status = status_text or self.status_badge.text()
        count = len(self.selected_pdf_paths)
        self.header_summary_label.setText(
            f"队列 {count} 个文件  ·  当前模式 {mode_text}  ·  状态 {status}"
        )

    def _sync_mode_ui(self):
        is_word = self._current_mode() == self.MODE_WORD
        if is_word:
            self.title_label.setText("PDF 转 Word")
            self.subtitle_label.setText("批量 OCR 识别，输出可编辑 DOCX")
        else:
            self.title_label.setText("PDF 表格转 Excel")
            self.subtitle_label.setText("批量表格识别，输出结构化 Excel")
        self._refresh_header_summary()
        self._refresh_output_display()

    def _on_quality_change(self, index):
        self._dpi_value = 300 if index == 0 else 200

    def _select_pdf_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "添加 PDF 文件",
            "",
            "PDF Files (*.pdf);;All Files (*)",
        )
        if paths:
            self._add_pdf_paths(paths)

    def _select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择 PDF 文件夹", "")
        if not folder:
            return
        paths = [str(path) for path in sorted(Path(folder).rglob("*.pdf"))]
        if not paths:
            InfoBar.warning("提示", "所选文件夹中没有找到 PDF 文件。", parent=self, duration=3000)
            return
        self._add_pdf_paths(paths)

    def _add_pdf_paths(self, paths):
        normalized = []
        existed = set(self.selected_pdf_paths)
        for raw_path in paths:
            path = Path(raw_path)
            if path.is_dir():
                for child in sorted(path.rglob("*.pdf")):
                    child_str = str(child)
                    if child_str not in existed:
                        existed.add(child_str)
                        normalized.append(child_str)
            elif path.suffix.lower() == ".pdf":
                path_str = str(path)
                if path_str not in existed:
                    existed.add(path_str)
                    normalized.append(path_str)

        if not normalized:
            return

        self.selected_pdf_paths.extend(normalized)
        for path in normalized:
            self.file_statuses[path] = "waiting"
        self.output_manual = False
        self.manual_output_target = ""
        self._refresh_queue()
        self._refresh_output_display()
        self._update_status("Ready")

    def _remove_selected_files(self):
        selected_paths = {item.data(Qt.UserRole) for item in self.queue_list.selectedItems()}
        if not selected_paths:
            return
        self.selected_pdf_paths = [path for path in self.selected_pdf_paths if path not in selected_paths]
        self.file_statuses = {path: status for path, status in self.file_statuses.items() if path in self.selected_pdf_paths}
        self.output_manual = False
        self.manual_output_target = ""
        self._refresh_queue()
        self._refresh_output_display()

    def _move_selected(self, direction):
        selected_rows = sorted(self.queue_list.row(item) for item in self.queue_list.selectedItems())
        if not selected_rows:
            return
        if direction < 0 and selected_rows[0] == 0:
            return
        if direction > 0 and selected_rows[-1] == len(self.selected_pdf_paths) - 1:
            return

        items = list(self.selected_pdf_paths)
        if direction < 0:
            for row in selected_rows:
                items[row - 1], items[row] = items[row], items[row - 1]
            new_rows = [row - 1 for row in selected_rows]
        else:
            for row in reversed(selected_rows):
                items[row + 1], items[row] = items[row], items[row + 1]
            new_rows = [row + 1 for row in selected_rows]

        self.selected_pdf_paths = items
        self._refresh_queue()
        for row in new_rows:
            item = self.queue_list.item(row)
            if item:
                item.setSelected(True)
        self._refresh_action_state()

    def _clear_queue(self):
        self.selected_pdf_paths = []
        self.file_statuses = {}
        self.output_manual = False
        self.manual_output_target = ""
        self.current_task_name = ""
        self.current_output_path = ""
        self._update_current_task("未开始")
        self._refresh_queue()
        self._refresh_output_display()
        self._update_progress(0)
        self._update_status("Idle")

    def _select_output(self):
        if not self.selected_pdf_paths:
            return

        if len(self.selected_pdf_paths) > 1:
            default_dir = str(Path(self.selected_pdf_paths[0]).parent)
            path = QFileDialog.getExistingDirectory(self, "选择输出目录", default_dir)
            if path:
                self.output_manual = True
                self.manual_output_target = path
                self._refresh_output_display()
            return

        default_path = self.out_edit.text() or self._default_output_path(self.selected_pdf_paths[0])
        if self._current_mode() == self.MODE_WORD:
            output_filter = "Word Files (*.docx);;All Files (*)"
        else:
            output_filter = "Excel Files (*.xlsx);;All Files (*)"
        path, _ = QFileDialog.getSaveFileName(self, "选择输出文件", default_path, output_filter)
        if path:
            self.output_manual = True
            self.manual_output_target = path
            self._refresh_output_display()

    def _request_stop(self):
        if not self.is_running or self.stop_requested:
            return
        self.stop_requested = True
        self.stop_btn.setEnabled(False)
        self.stop_btn.setText("停止中...")
        self.status_signal.emit("Stopping")

    def _start_conversion(self):
        if not self.selected_pdf_paths:
            InfoBar.warning("提示", "请先添加至少一个 PDF 文件。", parent=self, duration=3000)
            return

        # 判断是继续还是全新开始
        is_resume = any(
            self.file_statuses.get(p) in ("stopped", "waiting")
            for p in self.selected_pdf_paths
        ) and any(
            self.file_statuses.get(p) in ("success", "failed", "stopped")
            for p in self.selected_pdf_paths
        )

        self.is_running = True
        self.stop_requested = False

        if is_resume:
            # 继续模式：只处理 waiting 和 stopped 的文件
            pending_paths = [
                p for p in self.selected_pdf_paths
                if self.file_statuses.get(p) in ("waiting", "stopped")
            ]
            # 把 stopped 的改回 waiting
            for p in pending_paths:
                self.file_statuses[p] = "waiting"
        else:
            # 全新开始
            pending_paths = list(self.selected_pdf_paths)
            self.progress_bar.setValue(0)
            self.progress_text_label.setText("0%")
            self.output_list.clear()
            self._output_files = []
            self.output_badge.setText("0 个文件")
            self.file_statuses = {path: "waiting" for path in self.selected_pdf_paths}

        self.current_task_name = ""
        self.current_output_path = ""
        self._update_current_task("准备启动批量任务...")
        self.status_signal.emit("Running")
        self.start_btn.setText("转换中...")
        self._start_arrow_animation()
        self._refresh_queue()
        self._refresh_action_state()

        mode = self._current_mode()
        output_target = self.manual_output_target if self.output_manual else None
        dpi = self._dpi_value

        thread = threading.Thread(
            target=self._run_conversion,
            args=(mode, pending_paths, output_target, dpi),
            daemon=True,
        )
        thread.start()

    def _run_conversion(self, mode, pdf_paths, output_target, dpi):
        total = len(pdf_paths)
        success_count = 0
        failed = []
        stopped = False

        try:
            if mode == self.MODE_WORD:
                from converters.pdf_to_word import convert_pdf_to_word

                converter = convert_pdf_to_word
            else:
                from converters.table_to_excel import convert_table_pdf_to_excel

                converter = convert_table_pdf_to_excel

            output_paths = self._batch_output_paths(pdf_paths, mode, output_target)
            for index, (pdf_path, out_path) in enumerate(zip(pdf_paths, output_paths)):
                self._check_cancelled()
                self.current_task_name = Path(pdf_path).name
                self.current_output_path = out_path
                self.queue_item_status_signal.emit(pdf_path, "running")
                self.current_task_signal.emit(f"[{index + 1}/{total}] {Path(pdf_path).name}\n{out_path}")
                try:
                    kwargs = {
                        "log_callback": lambda msg: None,
                        "progress_callback": (
                            lambda val, current=index, p=pdf_path: (
                                self.progress_signal.emit((current + val) / total),
                                self.queue_item_progress_signal.emit(p, val),
                            )
                        ),
                        "cancel_check": self._check_cancelled,
                    }
                    converter(pdf_path, out_path, dpi=dpi, **kwargs)
                    success_count += 1
                    self.queue_item_status_signal.emit(pdf_path, "success")
                    self.output_file_signal.emit(out_path)
                    self.progress_signal.emit((index + 1) / total)
                except ConversionCancelled:
                    stopped = True
                    self.queue_item_status_signal.emit(pdf_path, "stopped")
                    self.progress_signal.emit(index / total if total else 0)
                    break
                except Exception as e:
                    failed.append((pdf_path, str(e)))
                    self.queue_item_status_signal.emit(pdf_path, "failed")
                    self.progress_signal.emit((index + 1) / total)

            if stopped:
                self.status_signal.emit("Stopped")
            elif failed:
                self.status_signal.emit("Finished")
            else:
                self.status_signal.emit("Finished")
        except ConversionCancelled:
            self.status_signal.emit("Stopped")
        except Exception as e:
            self.status_signal.emit("Error")
        finally:
            self.done_signal.emit()

    def _on_done(self):
        self.is_running = False
        self.stop_requested = False
        self._stop_arrow_animation()
        self.stop_btn.setText("停止")
        if not self.selected_pdf_paths and self.status_badge.text() == "就绪":
            self.status_signal.emit("Idle")
        self._update_task_metrics()

        # 判断是否有未完成的任务
        has_pending = any(
            self.file_statuses.get(p) in ("waiting", "stopped")
            for p in self.selected_pdf_paths
        )
        if has_pending:
            self.start_btn.setText("继续转换")
        else:
            self.start_btn.setText("开始转换")

        self.current_file_label.setText(
            self.current_file_label.text() if self.current_file_label.text().strip() else "未开始"
        )
        self._refresh_action_state()
