# PDF 表格转 Excel 智能提取工具 📊

一个基于 Python 的现代化桌面应用，将 PDF 扫描件中的表格高精度转换为可编辑的 Excel (`.xlsx`) 文件。

## ✨ 核心特性

- **高精度 OCR 识别**：采用 PP-OCRv4 神经网络模型（ONNX Runtime 推理），中文识别准确率极高。
- **智能图像预处理**：自动对扫描件进行 CLAHE 自适应对比度增强 + 非局部均值去噪，大幅提升模糊文档的识别率。
- **现代化 GUI 界面**：Slate 暗黑主题设计，界面优雅、响应迅速，支持拖拽选择文件一键转换。
- **双质量模式**：
  - 🎯 **高精度模式 (300 DPI)**：高分辨率扫描，适合模糊或复杂的扫描件。
  - ⚡ **极速模式 (200 DPI)**：更快的处理速度，适合清晰文档或批量处理。
- **智能数值安全校验**：自动进行 `数量 × 单价 = 金额` 交叉比对，可疑数据在 Excel 中用**黄色高亮**标记，绝不篡改原始数据。
- **动态表格重构**：自动分析表头结构计算列宽边界，精准还原原有的表格排版。

## 🚀 快速开始

### 环境要求
- Python 3.10+

### 安装依赖
```bash
git clone <repository-url>
cd TableConversion
pip install -r requirements.txt
```

### 运行程序
**GUI 桌面模式 (推荐):**
```bash
python main.py
```

**CLI 命令行模式:**
```bash
python main.py 你的文件.pdf -o 输出文件.xlsx

# 指定 DPI
python main.py 你的文件.pdf --dpi 400
```

## 📦 打包与分发 (Windows .exe)

本项目已配置 **GitHub Actions 自动构建**。将代码推送到 GitHub 后，在 `Actions` 页面即可下载打包好的免安装 `.zip` 压缩包。

本地打包（需要 Windows 环境）：
```bash
pip install pyinstaller
pyinstaller --noconfirm --onedir --windowed --name "TableConversion" --collect-all "customtkinter" --collect-all "rapidocr_onnxruntime" --collect-all "onnxruntime" main.py
```
打包完成后，将 `dist/TableConversion` 整个文件夹发给用户即可。

## 🔧 技术栈

| 组件 | 技术 |
|------|------|
| OCR 引擎 | RapidOCR (PP-OCRv4 ONNX) |
| 图像预处理 | OpenCV (CLAHE + NLMeans) |
| PDF 解析 | pdfplumber |
| Excel 输出 | openpyxl |
| GUI 框架 | CustomTkinter |
| 打包 | PyInstaller + GitHub Actions |
