# PDF 转换工具

一个基于 Python 的桌面应用，当前包含两个独立功能：

- **PDF 表格转 Excel**：沿用原有 RapidOCR + 表格重构逻辑，支持扫描件表格提取到 `.xlsx`。
- **PDF 转 Word**：使用 RapidOCR 识别 PDF 页面，生成可编辑 `.docx`。

## 核心特性

### PDF 表格转 Excel

- **高精度 OCR 识别**：采用 PP-OCRv4 神经网络模型（ONNX Runtime 推理）。
- **智能图像预处理**：自动进行 CLAHE 对比度增强和去噪。
- **双质量模式**：
  - **高精度模式 (300 DPI)**：适合模糊或复杂扫描件。
  - **极速模式 (200 DPI)**：适合清晰文档或批量处理。
- **智能数值安全校验**：自动校验 `数量 × 单价 = 金额`，可疑行在 Excel 中黄色高亮。
- **动态表格重构**：自动分析表头结构计算列宽边界。

### PDF 转 Word

- **统一 OCR 路线**：渲染 PDF 页面并用 RapidOCR 识别为可编辑文字。
- **独立转换模块**：不影响原有表格转 Excel 逻辑。
- **说明**：当前 Word 输出优先保证文字可编辑，复杂表格和版式不会 100% 还原。

## 项目结构

```text
TableConversion/
├── main.py                     # 程序入口和 CLI 分发
├── converters/
│   ├── table_to_excel.py        # PDF 表格转 Excel
│   └── pdf_to_word.py           # PDF 转 Word
└── ui/
    └── main_window.py           # 桌面界面
```

## 快速开始

### 环境要求

- Python 3.10+

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行桌面程序

```bash
python main.py
```

### 命令行使用

旧用法保持兼容，默认仍是 PDF 表格转 Excel：

```bash
python main.py 你的文件.pdf -o 输出文件.xlsx
python main.py 你的文件.pdf --dpi 400
```

也可以显式指定功能：

```bash
python main.py table 你的文件.pdf -o 输出文件.xlsx
python main.py word 你的文件.pdf -o 输出文件.docx
python main.py word 扫描件.pdf -o 输出文件.docx --dpi 300
```

## 打包与分发 (Windows .exe)

本项目已配置 GitHub Actions 自动构建。将代码推送到 GitHub 后，在 `Actions` 页面即可下载打包好的免安装 `.zip` 压缩包。

本地打包（需要 Windows 环境）：

```bash
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --onedir --windowed --name "TableConversion" --icon "app.ico" --hidden-import "PyQt5.QtPrintSupport" --hidden-import "qfluentwidgets" --hidden-import "rapidocr" --hidden-import "onnxruntime" --hidden-import "openpyxl" --hidden-import "pdfplumber" --hidden-import "fitz" --hidden-import "docx" --hidden-import "shapely" --hidden-import "ui.main_window" --hidden-import "converters.table_to_excel" --hidden-import "converters.pdf_to_word" --collect-all "qfluentwidgets" --collect-all "rapidocr" --collect-all "onnxruntime" --collect-all "pdfplumber" --collect-all "shapely" --copy-metadata "PyQt-Fluent-Widgets" --copy-metadata "rapidocr" --copy-metadata "onnxruntime" --copy-metadata "PyMuPDF" --copy-metadata "python-docx" main.py
```

打包完成后，将 `dist/TableConversion` 整个文件夹发给用户即可。

## 技术栈

| 组件 | 技术 |
|------|------|
| OCR 引擎 | RapidOCR (PP-OCRv4 ONNX) |
| 图像预处理 | OpenCV |
| PDF 表格解析 | pdfplumber |
| Excel 输出 | openpyxl |
| PDF 转 Word | PyMuPDF + RapidOCR + python-docx |
| GUI 框架 | PyQt5 + PyQt-Fluent-Widgets |
| 打包 | PyInstaller + GitHub Actions |
