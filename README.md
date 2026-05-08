# PDF 表格转 Excel 智能提取工具 📊

这是一个基于 Python 编写的现代化桌面应用程序，旨在将包含表格的 PDF 文件（尤其是复杂的图片扫描件）高精度地转换为可编辑的 Excel (`.xlsx`) 文件。

## ✨ 核心特性

- **双引擎随心切换**：
  - 🎯 **高精度模式 (PaddleOCR)**：采用强大的百度 PP-OCRv5 模型，具有极高的中文识别准确率，专治各种模糊、歪斜的复杂扫描件。
  - ⚡ **极速模式 (RapidOCR)**：基于 ONNX 极速推理，十几秒即可完成多页文档处理，适合常规清晰文档。
- **现代化 GUI 界面**：采用 Slate 暗黑主题设计，界面优雅、响应迅速，小白也能一键操作。
- **智能数值安全校验**：不盲目篡改数据！独创的“透明校验”架构会在后台进行 `数量 × 单价 = 金额` 的交叉比对。对于识别有偏差的账目，会在生成的 Excel 中用**显眼的黄色背景**高亮标记，保障财务数据的绝对安全。
- **动态表格重构**：自动分析表头结构计算列宽边界，精准还原原有的表格排版。

## 🚀 快速运行 (开发者指南)

### 环境要求
- Python 3.10+
- 推荐使用虚拟环境 (conda 或 venv)

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
# 使用高精度模式转换
python main.py 你的文件.pdf -o 输出文件.xlsx

# 使用极速模式转换
python main.py 你的文件.pdf --engine rapid
```

## 📦 打包与分发 (Windows .exe)

本项目已配置了完整的 **GitHub Actions 自动化构建流程**。
你不需要在本地折腾繁琐的打包环境，只需将代码推送到 GitHub，即可在 `Actions` 页面自动获取打包好的免安装免环境 `.zip` 压缩包。

如果你坚持要在本地打包：
1. 强烈建议在 Windows 环境下操作。
2. 安装打包工具：`pip install pyinstaller`
3. 执行打包命令（请务必使用 `--onedir` 文件夹模式，避免解压过慢）：
```bash
pyinstaller --noconfirm --onedir --windowed --name "PDF表格转Excel工具" --hidden-import="paddle" --hidden-import="paddleocr" --hidden-import="rapidocr_onnxruntime" --collect-all "customtkinter" --collect-all "paddleocr" main.py
```
打包完成后，将 `dist/PDF表格转Excel工具` 整个文件夹发给客户即可。

## ⚠️ 注意事项（离线环境必读）

当你首次使用 **高精度模式 (PaddleOCR)** 时，程序会在后台自动从百度官方服务器下载最新的 `PP-OCRv5` 模型，并缓存到你的电脑中（路径一般为：`C:\Users\你的用户名\.paddlex\official_models\`）。
**如果你要把打包好的程序发给没有互联网连接的客户，请务必将该目录下的模型文件手动拷贝到客户电脑的对应目录中。**
