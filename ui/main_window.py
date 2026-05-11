"""Main desktop window."""

import threading
from pathlib import Path

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QFileDialog, QHBoxLayout, QMainWindow, QVBoxLayout, QWidget

from qfluentwidgets import (
    BodyLabel,
    CaptionLabel,
    CardWidget,
    ComboBox,
    InfoBar,
    LineEdit,
    PrimaryPushButton,
    ProgressBar,
    PushButton,
    SpinBox,
    SubtitleLabel,
    TextEdit,
    Theme,
    setTheme,
    setThemeColor,
)


class MainWindow(QMainWindow):
    """PDF conversion desktop window."""

    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(float)
    done_signal = pyqtSignal()

    MODE_TABLE = "table"
    MODE_WORD = "word"

    def __init__(self):
        super().__init__()
        self.output_manual = False
        self.setWindowTitle("PDF 转换工具")
        self.resize(900, 700)
        self.setMinimumSize(800, 620)

        setTheme(Theme.DARK)
        setThemeColor(QColor(59, 130, 246))

        self._init_ui()
        self._connect_signals()

    def _init_ui(self):
        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)
        self.setStyleSheet("#centralWidget { background-color: #1A1A2E; }")

        root = QVBoxLayout(central)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(20)

        header = QHBoxLayout()
        self.title_label = SubtitleLabel("📄 PDF 转换工具")
        self.title_label.setStyleSheet("color: white; font-size: 24px; font-weight: bold;")
        header.addWidget(self.title_label)
        header.addStretch()
        self.subtitle_label = CaptionLabel("表格转 Excel · PDF 转 Word")
        self.subtitle_label.setStyleSheet("color: #94A3B8;")
        header.addWidget(self.subtitle_label, alignment=Qt.AlignBottom)
        root.addLayout(header)

        control_card = CardWidget(self)
        control_layout = QVBoxLayout(control_card)
        control_layout.setContentsMargins(24, 20, 24, 20)
        control_layout.setSpacing(16)

        mode_row = QHBoxLayout()
        mode_label = BodyLabel("功能:")
        mode_label.setStyleSheet("color: white; font-weight: 500;")
        mode_label.setFixedWidth(120)
        self.mode_combo = ComboBox()
        self.mode_combo.addItems(["PDF 表格转 Excel", "PDF 转 Word"])
        mode_row.addWidget(mode_label)
        mode_row.addWidget(self.mode_combo)
        control_layout.addLayout(mode_row)

        pdf_row = QHBoxLayout()
        self.pdf_btn = PushButton("📂 选择 PDF")
        self.pdf_btn.setFixedWidth(120)
        self.pdf_edit = LineEdit()
        self.pdf_edit.setPlaceholderText("请选择需要转换的 PDF 文件...")
        self.pdf_edit.setReadOnly(True)
        pdf_row.addWidget(self.pdf_btn)
        pdf_row.addWidget(self.pdf_edit)
        control_layout.addLayout(pdf_row)

        out_row = QHBoxLayout()
        self.out_btn = PushButton("💾 保存路径")
        self.out_btn.setFixedWidth(120)
        self.out_edit = LineEdit()
        self.out_edit.setPlaceholderText("自动生成，或手动选择...")
        self.out_edit.setReadOnly(True)
        out_row.addWidget(self.out_btn)
        out_row.addWidget(self.out_edit)
        control_layout.addLayout(out_row)

        settings_row = QHBoxLayout()
        settings_row.setSpacing(20)

        quality_layout = QHBoxLayout()
        quality_layout.setSpacing(10)
        self.quality_label = BodyLabel("质量模式:")
        self.quality_label.setStyleSheet("color: white; font-weight: 500;")
        self.quality_combo = ComboBox()
        self.quality_combo.addItems(["高精度 (300 DPI)", "极速 (200 DPI)"])
        self.quality_combo.setCurrentIndex(0)
        quality_layout.addWidget(self.quality_label)
        quality_layout.addWidget(self.quality_combo, stretch=1)

        dpi_layout = QHBoxLayout()
        dpi_layout.setSpacing(10)
        self.dpi_label = BodyLabel("DPI:")
        self.dpi_label.setStyleSheet("color: white; font-weight: 500;")
        self.dpi_spin = SpinBox()
        self.dpi_spin.setRange(100, 600)
        self.dpi_spin.setValue(300)
        dpi_layout.addWidget(self.dpi_label)
        dpi_layout.addWidget(self.dpi_spin, stretch=1)

        self.start_btn = PrimaryPushButton("🚀 开始转换")
        self.start_btn.setFixedHeight(34)
        self.start_btn.setStyleSheet("""
            PrimaryPushButton {
                background-color: #3B82F6;
                color: white;
                font-size: 15px;
                font-weight: bold;
                border-radius: 6px;
                border: 1px solid #2563EB;
            }
            PrimaryPushButton:hover {
                background-color: #2563EB;
            }
            PrimaryPushButton:pressed {
                background-color: #1D4ED8;
            }
            PrimaryPushButton:disabled {
                background-color: #475569;
                color: #94A3B8;
                border: 1px solid #334155;
            }
        """)

        settings_row.addLayout(quality_layout, 1)
        settings_row.addLayout(dpi_layout, 1)
        settings_row.addWidget(self.start_btn, 1)
        control_layout.addLayout(settings_row)

        root.addWidget(control_card)

        self.progress_bar = ProgressBar()
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setValue(0)
        root.addWidget(self.progress_bar)

        log_card = CardWidget(self)
        log_layout = QVBoxLayout(log_card)
        log_layout.setContentsMargins(16, 12, 16, 12)

        self.log_box = TextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("""
            TextEdit {
                background-color: #0A0A0F;
                color: #A7F3D0;
                font-family: 'Consolas', 'Menlo', monospace;
                font-size: 13px;
                border: none;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        log_layout.addWidget(self.log_box)
        root.addWidget(log_card, stretch=1)

        self._log("✨ 欢迎使用 PDF 转换工具！")
        self._log("✨ 请选择功能和 PDF 文件后点击 [开始转换]。\n" + "─" * 50)
        self._sync_mode_ui()

    def _connect_signals(self):
        self.mode_combo.currentIndexChanged.connect(self._on_mode_change)
        self.pdf_btn.clicked.connect(self._select_pdf)
        self.out_btn.clicked.connect(self._select_output)
        self.start_btn.clicked.connect(self._start_conversion)
        self.quality_combo.currentIndexChanged.connect(self._on_quality_change)
        self.log_signal.connect(self._log)
        self.progress_signal.connect(self._update_progress)
        self.done_signal.connect(self._on_done)

    def _current_mode(self):
        return self.MODE_WORD if self.mode_combo.currentIndex() == 1 else self.MODE_TABLE

    def _output_suffix(self):
        return ".docx" if self._current_mode() == self.MODE_WORD else ".xlsx"

    def _output_filter(self):
        if self._current_mode() == self.MODE_WORD:
            return "Word Files (*.docx);;All Files (*)"
        return "Excel Files (*.xlsx);;All Files (*)"

    def _default_output_path(self, pdf_path):
        return str(Path(pdf_path).with_suffix(self._output_suffix()))

    def _sync_mode_ui(self):
        is_word = self._current_mode() == self.MODE_WORD
        self.title_label.setText("📄 PDF 转 Word" if is_word else "📄 PDF 表格转 Excel")
        self.subtitle_label.setText("OCR 识别 · 生成可编辑 DOCX" if is_word else "智能 OCR · 扫描件支持 · 极速生成")
        self.quality_label.setText("OCR 模式:" if is_word else "质量模式:")
        self.quality_combo.setEnabled(True)
        self.dpi_spin.setEnabled(True)
        self.quality_label.setEnabled(True)
        self.dpi_label.setEnabled(True)
        if self.pdf_edit.text() and not self.output_manual:
            self.out_edit.setText(self._default_output_path(self.pdf_edit.text()))

    def _on_mode_change(self, _index):
        self._sync_mode_ui()

    def _select_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 PDF 文件", "", "PDF Files (*.pdf);;All Files (*)"
        )
        if path:
            self.pdf_edit.setText(path)
            self.output_manual = False
            self.out_edit.setText(self._default_output_path(path))

    def _select_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "保存文件", "", self._output_filter()
        )
        if path:
            suffix = self._output_suffix()
            if Path(path).suffix.lower() != suffix:
                path = str(Path(path).with_suffix(suffix))
            self.output_manual = True
            self.out_edit.setText(path)

    def _on_quality_change(self, index):
        self.dpi_spin.setValue(300 if index == 0 else 200)

    def _log(self, msg):
        self.log_box.append(msg)
        cursor = self.log_box.textCursor()
        cursor.movePosition(cursor.End)
        self.log_box.setTextCursor(cursor)

    def _update_progress(self, val):
        self.progress_bar.setValue(int(val * 100))

    def _on_done(self):
        self.start_btn.setEnabled(True)
        self.start_btn.setText("开始转换")

    def _start_conversion(self):
        pdf_path = self.pdf_edit.text()
        out_path = self.out_edit.text()
        mode = self._current_mode()
        dpi = self.dpi_spin.value()

        if not pdf_path:
            InfoBar.warning("提示", "请先选择 PDF 文件！", parent=self, duration=3000)
            return
        if not out_path:
            InfoBar.warning("提示", "请选择输出路径！", parent=self, duration=3000)
            return

        self.start_btn.setEnabled(False)
        self.start_btn.setText("转换中...")
        self.progress_bar.setValue(0)
        self.log_box.clear()
        self._log(f"🚀 开始任务，文件: {Path(pdf_path).name}\n" + "─" * 50)

        thread = threading.Thread(
            target=self._run_conversion,
            args=(mode, pdf_path, out_path, dpi),
            daemon=True,
        )
        thread.start()

    def _run_conversion(self, mode, pdf_path, out_path, dpi):
        try:
            kwargs = {
                "log_callback": lambda msg: self.log_signal.emit(msg),
                "progress_callback": lambda val: self.progress_signal.emit(val),
            }
            if mode == self.MODE_WORD:
                self.log_signal.emit("⏳ 正在加载 PDF 转 Word 引擎...")
                from converters.pdf_to_word import convert_pdf_to_word

                convert_pdf_to_word(pdf_path, out_path, dpi=dpi, **kwargs)
            else:
                self.log_signal.emit("⏳ 正在加载表格转 Excel 引擎...")
                from converters.table_to_excel import convert_table_pdf_to_excel

                convert_table_pdf_to_excel(pdf_path, out_path, dpi=dpi, **kwargs)
            self.log_signal.emit("\n✅ 转换完成！")
        except Exception as e:
            self.log_signal.emit(f"\n❌ 发生错误: {str(e)}")
        finally:
            self.done_signal.emit()
