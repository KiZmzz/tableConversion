"""
PDF 表格转 Excel 工具（支持扫描件）
使用 RapidOCR (PP-OCRv4 ONNX) 进行中文识别 + 坐标聚类提取表格

依赖: pip install pdfplumber openpyxl rapidocr onnxruntime opencv-python-headless PyQt-Fluent-Widgets
用法: python main.py [PDF文件] [-o 输出文件]
"""

import os, sys, re, argparse
from pathlib import Path
import threading

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QFileDialog)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor

from qfluentwidgets import (
    PrimaryPushButton, PushButton, LineEdit, ComboBox, SpinBox,
    TextEdit, CardWidget, ProgressBar, InfoBar,
    SubtitleLabel, StrongBodyLabel, BodyLabel, CaptionLabel,
    FluentIcon, setTheme, setThemeColor, Theme
)

import cv2
import numpy as np
import pdfplumber
from rapidocr import RapidOCR
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


def detect_h_lines(img_path, min_len=60):
    """检测水平线以确定行位置"""
    img = cv2.imread(img_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (min_len, 1))
    h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel)
    h_sum = np.sum(h_lines, axis=1)
    threshold = max(np.max(h_sum) * 0.3, 1)
    positions = np.where(h_sum > threshold)[0]

    rows = []
    group = [positions[0]] if len(positions) > 0 else []
    for p in positions[1:]:
        if p - group[-1] <= 8:
            group.append(p)
        else:
            rows.append(int(np.mean(group)))
            group = [p]
    if group:
        rows.append(int(np.mean(group)))
    return rows


def preprocess_image(img_path):
    """图像预处理：增强对比度 + 去噪，提升扫描件 OCR 精度"""
    img = cv2.imread(img_path)
    if img is None:
        return img_path
    
    # 转灰度
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # CLAHE 自适应直方图均衡化（大幅提升扫描件对比度）
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)
    
    # 轻度去噪（保留文字边缘）
    denoised = cv2.fastNlMeansDenoising(enhanced, h=10, templateWindowSize=7, searchWindowSize=21)
    
    # 保存预处理后的图像
    processed_path = img_path.replace('.png', '_enhanced.png')
    cv2.imwrite(processed_path, denoised)
    return processed_path


def full_page_ocr(ocr_engine, img_path):
    """整页 OCR，返回带位置信息的结果"""
    items = []
    result = ocr_engine(img_path)
    if not result or not result.txts: return items
    for i in range(len(result.txts)):
        box = result.boxes[i]
        text = result.txts[i]
        conf = result.scores[i]
        y_center = (float(box[0][1]) + float(box[2][1])) / 2
        x_center = (float(box[0][0]) + float(box[2][0])) / 2
        items.append({
            'text': text.strip(),
            'y': y_center,
            'x': x_center,
            'conf': float(conf)
        })
    return items


def assign_to_rows(items, h_lines):
    """将 OCR 文本块分配到最近的表格行"""
    if not h_lines or not items:
        return {}

    # 行区间：h_lines[i] ~ h_lines[i+1]
    rows = {}
    for item in items:
        y = item['y']
        row_idx = None
        for i in range(len(h_lines) - 1):
            if h_lines[i] <= y <= h_lines[i + 1]:
                row_idx = i
                break
        if row_idx is None:
            # 在最后一条线之下
            if y > h_lines[-1]:
                continue
            # 在第一条线之上
            if y < h_lines[0]:
                continue
        if row_idx is not None:
            if row_idx not in rows:
                rows[row_idx] = []
            rows[row_idx].append(item)
    return rows

def assign_to_columns(row_items, col_boundaries):
    """将一行中的文本块分配到列"""
    cells = {}
    for item in row_items:
        x = item['x']
        text = item['text'].strip()
        # 处理类似 "根根根根" 的 OCR 叠字错误
        if len(text) > 1 and len(set(text.replace(' ', ''))) == 1:
            text = text.replace(' ', '')[0]
            
        col_idx = len(col_boundaries) - 2  # 默认最后一列
        for c in range(len(col_boundaries) - 1):
            if col_boundaries[c] <= x < col_boundaries[c + 1]:
                col_idx = c
                break

        if col_idx in cells:
            # 同列合并：按 x 排序
            existing = cells[col_idx]
            # 去重：如果识别到的文本完全一致（例如合并单元格导致的重复），则不再重复添加
            if text == existing['text'].strip() or text in existing['text'].split():
                pass
            elif item['x'] < existing['x']:
                cells[col_idx] = {
                    'text': text + ' ' + existing['text'],
                    'x': item['x']
                }
            else:
                cells[col_idx] = {
                    'text': existing['text'] + ' ' + text,
                    'x': existing['x']
                }
        else:
            cells[col_idx] = {'text': text, 'x': item['x']}

    return {k: v['text'] for k, v in cells.items()}


def clean_number(s):
    """清洗 OCR 识别出的数字字符串，修正常见 OCR 错误
    返回 (float_value, cleaned_str) 或 (None, original_str)
    """
    if not s or not isinstance(s, str):
        return None, str(s) if s else ''
    s = s.strip()
    if not s:
        return None, ''
    
    # 常见 OCR 替换：冒号→点，中文逗号→空，空格→空
    cleaned = s.replace('：', '.').replace(':', '.').replace('，', '').replace(',', '').replace(' ', '')
    # 移除多余的点（例如 "1..16" → "1.16", "6. 5." → "6.5"）
    # 先处理连续两个点
    while '..' in cleaned:
        cleaned = cleaned.replace('..', '.')
    # 移除末尾的点
    cleaned = cleaned.rstrip('.')
    # 如果有多个小数点，只保留第一个
    parts = cleaned.split('.')
    if len(parts) > 2:
        cleaned = parts[0] + '.' + ''.join(parts[1:])
    
    # 移除非数字字符（除了小数点和负号）
    result = ''
    for ch in cleaned:
        if ch.isdigit() or ch == '.' or (ch == '-' and not result):
            result += ch
    
    if not result or result == '.' or result == '-':
        return None, s
    
    try:
        val = float(result)
        return val, result
    except ValueError:
        return None, s


def post_process_rows(rows, log_func=None):
    """对提取的表格行进行后处理：
    1. 清洗所有数值字段中的 OCR 错误（冒号→点等客观修正）
    2. 缺失值：仅当只缺一个值时，用另外两个推算补充
    3. 交叉校验 数量×单价=金额，不匹配的行标记为可疑（不自动修正）
    返回 suspicious_rows: set of row indices that need manual review
    """
    def log(msg):
        if log_func:
            log_func(msg)
    
    suspicious = set()
    fixed_count = 0
    
    for i, row in enumerate(rows):
        qty_raw = str(row.get(2, '')).strip()
        price_raw = str(row.get(4, '')).strip()
        amount_raw = str(row.get(5, '')).strip()
        
        qty_val, qty_clean = clean_number(qty_raw)
        price_val, price_clean = clean_number(price_raw)
        amount_val, amount_clean = clean_number(amount_raw)
        
        # 用清洗后的值更新（只是修正格式，如冒号→点）
        if qty_val is not None:
            row[2] = qty_clean
        if price_val is not None:
            row[4] = price_clean
        if amount_val is not None:
            row[5] = amount_clean
        
        # 统计有几个值是有效的
        have = sum(1 for v in [qty_val, price_val, amount_val] if v is not None)
        
        if have == 3:
            # 三个值都有，做交叉校验
            expected = qty_val * price_val
            if qty_val > 0 and price_val > 0 and abs(expected - amount_val) > max(0.5, amount_val * 0.01):
                # 不匹配 → 标记为可疑，不修改
                name = str(row.get(0, ''))[:15]
                log(f"  ⚠️ 可疑: {name} | 数量={qty_val}×单价={price_val}={expected:.2f}, 但金额={amount_val}")
                suspicious.add(i)
        
        elif have == 2:
            # 只缺一个值，安全地补充
            if qty_val is None and price_val and amount_val and price_val > 0:
                qty = amount_val / price_val
                row[2] = str(int(round(qty))) if abs(qty - round(qty)) < 0.02 else f"{qty:.2f}"
                log(f"  🔧 补充数量: → {row[2]}")
                fixed_count += 1
            elif price_val is None and qty_val and amount_val and qty_val > 0:
                price = amount_val / qty_val
                row[4] = str(int(price)) if price == int(price) else f"{price:.2f}"
                log(f"  🔧 补充单价: → {row[4]}")
                fixed_count += 1
            elif amount_val is None and qty_val and price_val:
                amount = qty_val * price_val
                row[5] = str(int(amount)) if amount == int(amount) else f"{amount:.2f}"
                log(f"  🔧 补充金额: → {row[5]}")
                fixed_count += 1
        
        elif have <= 1:
            # 缺太多数值，标记可疑
            if any(v is not None and v > 0 for v in [qty_val, price_val, amount_val]):
                suspicious.add(i)
    
    return suspicious, fixed_count


def create_excel(rows, header, output_path, contract_info=None, suspicious_rows=None):
    """生成带格式的 Excel，可疑行用黄色高亮标记"""
    wb = Workbook()
    ws = wb.active
    ws.title = "购销合同"

    header_font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
    data_font = Font(name='微软雅黑', size=10)
    alt_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
    warn_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # 黄色高亮
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center')
    right = Alignment(horizontal='right', vertical='center')
    left = Alignment(vertical='center', wrap_text=True)

    if suspicious_rows is None:
        suspicious_rows = set()

    cur = 1
    num_cols = len(header)

    # 合同信息
    if contract_info:
        for txt, font in [
            (contract_info.get('company', ''), Font(name='微软雅黑', size=16, bold=True)),
            ('购 销 合 同', Font(name='微软雅黑', size=14, bold=True)),
        ]:
            ws.merge_cells(f'A{cur}:{get_column_letter(num_cols)}{cur}')
            c = ws.cell(row=cur, column=1, value=txt)
            c.font = font
            c.alignment = center
            cur += 1

        for key, label in [('contract_no', '合同编号'), ('date', '日期'),
                           ('supplier', '供方'), ('buyer', '需方')]:
            val = contract_info.get(key, '')
            if val:
                ws.merge_cells(f'A{cur}:{get_column_letter(num_cols)}{cur}')
                ws.cell(row=cur, column=1,
                        value=f'{label}：{val}').font = Font(name='微软雅黑', size=11)
                cur += 1
        cur += 1

    # 表头
    for ci, name in enumerate(header, 1):
        c = ws.cell(row=cur, column=ci, value=name)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border
    
    # 合并 A 和 B 列表头 (对应物品名称)
    ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=2)
    header_row = cur
    cur += 1

    # 数据
    for ri, row in enumerate(rows):
        is_suspicious = ri in suspicious_rows
        for ci in range(num_cols):
            val = row.get(ci, '') if isinstance(row, dict) else (row[ci] if ci < len(row) else '')
            c = ws.cell(row=cur, column=ci + 1, value=val)
            c.font = data_font
            c.border = border

            if ci in (2, 3, 4, 5) and val:
                try:
                    num = float(str(val).replace(' ', '').replace('，', '').replace(',', ''))
                    if num.is_integer():
                        c.value = int(num)
                        c.number_format = '0'
                    else:
                        c.value = num
                        c.number_format = '0.##'
                    c.alignment = right if ci >= 4 else center
                except ValueError:
                    c.alignment = center if ci in (2, 3) else left
            else:
                c.alignment = left

            # 可疑行：黄色高亮（优先级高于斑马纹）
            if is_suspicious:
                c.fill = warn_fill
            elif ri % 2 == 1:
                c.fill = alt_fill
        cur += 1

    for i, w in enumerate([14, 40, 8, 6, 12, 12, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f'A{header_row + 1}'
    wb.save(output_path)


def process_pdf(pdf_path, output_path, dpi=300, log_callback=None, progress_callback=None):
    def log(msg):
        print(msg)
        if log_callback:
            log_callback(msg)

    log(f"\n{'=' * 55}")
    log(f"  PDF 表格转 Excel 工具")
    log(f"{'=' * 55}")
    log(f"\n📄 输入：{pdf_path}")
    log(f"📊 输出：{output_path}")

    if not os.path.exists(pdf_path):
        log(f"\n❌ 找不到文件 '{pdf_path}'")
        raise FileNotFoundError(f"找不到文件 '{pdf_path}'")

    log("\n🎯 初始化 OCR 引擎 (RapidOCR PP-OCRv4)...")
    ocr = RapidOCR()
    scale = dpi / 200.0  # 相对 200 DPI 的缩放比

    # 默认列边界（基于 200 DPI）
    # 列: 物品名称 | 规格型号 | 数量 | 单位 | 含税单价 | 金额 | 备注
    col_bounds_200 = [0, 250, 750, 900, 1000, 1200, 1360, 1600]

    # PDF 转图片 + OCR
    log(f"\n📷 步骤 1/3：PDF 转图片 (DPI={dpi})...")
    if progress_callback: progress_callback(0.1)
    page_data = []
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            tmp = f"_tmp_page_{i}.png"
            page.to_image(resolution=dpi).save(tmp)
            log(f"  ✓ 第 {i + 1}/{total_pages} 页")
            page_data.append(tmp)
            if progress_callback: progress_callback(0.1 + 0.3 * ((i + 1) / total_pages))

    # 动态检测列边界
    log(f"\n🔍 分析表头以动态计算列边界...")
    if page_data:
        first_page_img = page_data[0]
        result = full_page_ocr(ocr, first_page_img)
        if result:
            headers = {}
            for item in result:
                x_center = item['x']
                y_center = item['y']
                text = item['text']
                # 只在页面上半部分找表头
                if y_center > 800 * scale: continue
                for k in ['物品名称', '规格', '数量', '单位', '单价', '金额', '备注']:
                    if k in text and k not in headers:
                        headers[k] = x_center
            
            log(f"  检测到表头: {list(headers.keys())}")
            # 只需要 '单价' 和 '金额' 就能推算出所有边界
            if '单价' in headers and '金额' in headers:
                x_price = headers['单价']
                x_amount = headers['金额']
                x_unit = headers.get('单位', x_price - 150 * scale)
                x_qty = headers.get('数量', x_unit - 150 * scale)
                x_remark = headers.get('备注', x_amount + 150 * scale)
                x_name = headers.get('物品名称', 400 * scale)
                # 规格列：物品名称表头的右侧就是规格列的开始
                # 物品名称表头中心在 x_name，但实际名称数据在更左边
                # 所以名称/规格的分界点应该在表头中心偏左的位置
                b0 = 0
                b1 = x_name - 150 * scale               # 物品名称 | 规格型号 的分界
                b2 = x_qty - 80 * scale                  # 规格型号 | 数量 的分界
                b3 = (x_qty + x_unit) / 2                # 数量 | 单位
                b4 = (x_unit + x_price) / 2              # 单位 | 单价
                b5 = (x_price + x_amount) / 2            # 单价 | 金额
                b6 = (x_amount + x_remark) / 2           # 金额 | 备注
                b7 = 10000 * scale
                # 将像素坐标转回 200 DPI 基准
                col_bounds_200 = sorted([v / scale for v in [b0, b1, b2, b3, b4, b5, b6, b7]])
                log(f"  ✅ 成功应用动态边界！")
            else:
                log(f"  未找齐关键表头，使用默认边界。")

    log(f"\n🔍 步骤 2/3：OCR + 表格提取...")
    all_rows = []
    contract_info = {}
    header = ['物品名称', '', '数量', '单位', '含税单价', '金额', '备注']
    col_bounds = [int(x * scale) for x in col_bounds_200]

    for pi, img_path in enumerate(page_data):
        log(f"\n  --- 第 {pi + 1} 页 ---")

        # 检测水平线
        h_lines = detect_h_lines(img_path, min_len=int(60 * scale))
        log(f"  水平线：{len(h_lines)} 条")

        # 图像预处理（增强对比度 + 去噪）
        enhanced_path = preprocess_image(img_path)

        # 整页 OCR
        items = full_page_ocr(ocr, enhanced_path)
        log(f"  OCR 文本块：{len(items)} 个")

        # 清理预处理临时文件
        if enhanced_path != img_path and os.path.exists(enhanced_path):
            os.remove(enhanced_path)

        # 提取合同信息（第一页）
        if pi == 0 and h_lines:
            table_start_y = 0
            for y in h_lines:
                if y > int(280 * scale):
                    table_start_y = y
                    break
            for item in items:
                t = item['text']
                if item['y'] < table_start_y:
                    if re.search(r'A\d{8,}', t):
                        m = re.search(r'A\d+', t)
                        contract_info['contract_no'] = m.group()
                    if '日期' in t and '/' in t:
                        contract_info['date'] = t.replace('日期：', '').replace('日期:', '').strip()
                    if '供方' in t:
                        contract_info['supplier'] = t.replace('供方：', '').replace('供方:', '').strip()
                    if '需方' in t:
                        contract_info['buyer'] = t.replace('需方：', '').replace('需方:', '').strip()
                    if '有限公司' in t and '供方' not in t and '需方' not in t:
                        if len(t) > len(contract_info.get('company', '')):
                            contract_info['company'] = t

        # 分配到行
        row_groups = assign_to_rows(items, h_lines)

        # 确定表格数据行范围
        if pi == 0:
            # 第一页：跳过表头区域 (前几行是标题/信息)
            # 找到包含 "物品名称" 的行，下一行开始是数据
            header_row_idx = None
            for ridx, ritems in row_groups.items():
                texts = ' '.join(it['text'] for it in ritems)
                if '物品名称' in texts or '含税单价' in texts:
                    header_row_idx = ridx
                    break
            data_start = (header_row_idx + 1) if header_row_idx is not None else 2
        else:
            data_start = 0

        # 转换为表格行
        page_rows = []
        for ridx in sorted(row_groups.keys()):
            if ridx < data_start:
                continue
            row_items = row_groups[ridx]

            # 跳过页脚和合同条款
            texts = ' '.join(it['text'] for it in row_items)
            if '第' in texts and '页' in texts:
                continue
            if any(k in texts for k in ['合计', '营计', '总计', '合汁']) and len(texts.replace(' ', '')) <= 15:
                # 保留 OCR 原本的合计行，然后停止提取（过滤掉后面的无关页脚）
                row_dict = assign_to_columns(row_items, col_bounds)
                page_rows.append(row_dict)
                break  # 合计行及其后面不再有表格数据
            # 跳过合同条款
            skip_keywords = ['总计金额', '交货地点', '运输方式', '验收标准',
                             '保修期', '结算方式', '违约责任', '解决合同',
                             '本合同', '以上产品', '单位名称', '单位地址',
                             '代理人', '电话', '传真', '开户行', '账号',
                             '需方需求', '供方负担', '签字盖章']
            if any(kw in texts for kw in skip_keywords):
                continue

            # 分配到列
            row_dict = assign_to_columns(row_items, col_bounds)

            # 只保留有意义的行（至少有2个非空单元格，且包含名称或金额数据）
            non_empty = sum(1 for v in row_dict.values() if v.strip())
            has_name = bool(row_dict.get(0, '').strip() or row_dict.get(1, '').strip())
            has_money = bool(row_dict.get(4, '').strip() or row_dict.get(5, '').strip())
            if non_empty >= 2 and (has_name or has_money):
                page_rows.append(row_dict)

        all_rows.extend(page_rows)
        log(f"  ✓ 提取 {len(page_rows)} 行数据")
        if progress_callback: progress_callback(0.4 + 0.4 * ((pi + 1) / total_pages))

    log(f"\n  📋 共 {len(all_rows)} 行")

    # 后处理：OCR 数值清洗 + 交叉校验 (数量×单价=金额)
    log(f"\n🔧 数值校验与修正...")
    suspicious_rows, fixed_count = post_process_rows(all_rows, log_func=log)
    log(f"  ✅ 补充了 {fixed_count} 处缺失数据")
    if suspicious_rows:
        log(f"  ⚠️ 发现 {len(suspicious_rows)} 行可疑数据（将在 Excel 中用黄色标记）")

    # 预览
    log(f"\n  前5行预览：")
    for i, row in enumerate(all_rows[:5]):
        cells = []
        for c in range(7):
            val = row.get(c, '')[:18]
            cells.append(f'{val:18s}')
        log(f"  {i + 1}: {'|'.join(cells)}")

    # 生成 Excel
    log(f"\n📝 步骤 3/3：生成 Excel...")
    if progress_callback: progress_callback(0.9)
    create_excel(all_rows, header, output_path, contract_info, suspicious_rows=suspicious_rows)

    # 清理
    for f in page_data:
        try:
            os.remove(f)
        except OSError:
            pass

    if progress_callback: progress_callback(1.0)
    log(f"\n{'=' * 55}")
    log(f"  ✅ 完成！输出：{os.path.abspath(output_path)}")
    log(f"{'=' * 55}\n")


class MainWindow(QMainWindow):
    """PDF 表格转 Excel - Fluent Design 主窗口"""

    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(float)
    done_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF 表格转 Excel")
        self.resize(900, 680)
        self.setMinimumSize(800, 600)

        # 深色模式
        setTheme(Theme.DARK)
        setThemeColor(QColor(59, 130, 246))  # #3B82F6 蓝色主题

        self._init_ui()
        self._connect_signals()

    # ── UI 构建 ──────────────────────────────────────────
    def _init_ui(self):
        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)
        # 只针对中心面板设置深色背景，防止破坏按钮等控件的自带样式
        self.setStyleSheet("#centralWidget { background-color: #1A1A2E; }")
        
        root = QVBoxLayout(central)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(20)

        # ─ 标题栏 ─
        header = QHBoxLayout()
        title = SubtitleLabel("📄 PDF 表格转 Excel")
        title.setStyleSheet("color: white; font-size: 24px; font-weight: bold;")
        header.addWidget(title)
        header.addStretch()
        subtitle = CaptionLabel("智能 OCR · 扫描件支持 · 极速生成")
        subtitle.setStyleSheet("color: #94A3B8;")
        header.addWidget(subtitle, alignment=Qt.AlignBottom)
        root.addLayout(header)

        # ─ 控制面板卡片 (合并文件和转换设置) ─
        control_card = CardWidget(self)
        control_layout = QVBoxLayout(control_card)
        control_layout.setContentsMargins(24, 20, 24, 20)
        control_layout.setSpacing(16)

        # 行 1: PDF 选择
        pdf_row = QHBoxLayout()
        self.pdf_btn = PushButton("📂 选择 PDF")
        self.pdf_btn.setFixedWidth(120)
        self.pdf_edit = LineEdit()
        self.pdf_edit.setPlaceholderText("请选择需要转换的 PDF 文件...")
        self.pdf_edit.setReadOnly(True)
        pdf_row.addWidget(self.pdf_btn)
        pdf_row.addWidget(self.pdf_edit)
        control_layout.addLayout(pdf_row)

        # 行 2: 输出路径
        out_row = QHBoxLayout()
        self.out_btn = PushButton("💾 保存路径")
        self.out_btn.setFixedWidth(120)
        self.out_edit = LineEdit()
        self.out_edit.setPlaceholderText("自动生成，或手动选择...")
        self.out_edit.setReadOnly(True)
        out_row.addWidget(self.out_btn)
        out_row.addWidget(self.out_edit)
        control_layout.addLayout(out_row)

        # 行 3: 质量、DPI、开始转换 (三等分)
        settings_row = QHBoxLayout()
        settings_row.setSpacing(20)
        
        # 1. 质量模式
        quality_layout = QHBoxLayout()
        quality_layout.setSpacing(10)
        quality_label = BodyLabel("质量模式:")
        quality_label.setStyleSheet("color: white; font-weight: 500;")
        self.quality_combo = ComboBox()
        self.quality_combo.addItems(["高精度 (300 DPI)", "极速 (200 DPI)"])
        self.quality_combo.setCurrentIndex(0)
        quality_layout.addWidget(quality_label)
        quality_layout.addWidget(self.quality_combo, stretch=1)
        
        # 2. DPI
        dpi_layout = QHBoxLayout()
        dpi_layout.setSpacing(10)
        dpi_label = BodyLabel("DPI:")
        dpi_label.setStyleSheet("color: white; font-weight: 500;")
        self.dpi_spin = SpinBox()
        self.dpi_spin.setRange(100, 600)
        self.dpi_spin.setValue(300)
        dpi_layout.addWidget(dpi_label)
        dpi_layout.addWidget(self.dpi_spin, stretch=1)

        # 3. 开始转换按钮
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

        # ─ 进度条 ─
        self.progress_bar = ProgressBar()
        self.progress_bar.setFixedHeight(4)
        self.progress_bar.setValue(0)
        root.addWidget(self.progress_bar)

        # ─ 日志区域 ─
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

        # 欢迎消息
        self._log("✨ 欢迎使用 PDF 表格转 Excel 工具！")
        self._log("✨ 请在上方选择文件后点击 [开始转换]。\n" + "─" * 50)

    # ── 信号连接 ─────────────────────────────────────────
    def _connect_signals(self):
        self.pdf_btn.clicked.connect(self._select_pdf)
        self.out_btn.clicked.connect(self._select_output)
        self.start_btn.clicked.connect(self._start_conversion)
        self.quality_combo.currentIndexChanged.connect(self._on_quality_change)
        self.log_signal.connect(self._log)
        self.progress_signal.connect(self._update_progress)
        self.done_signal.connect(self._on_done)

    # ── 槽函数 ───────────────────────────────────────────
    def _select_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 PDF 文件", "", "PDF Files (*.pdf);;All Files (*)"
        )
        if path:
            self.pdf_edit.setText(path)
            self.out_edit.setText(str(Path(path).with_suffix('.xlsx')))

    def _select_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "保存 Excel 文件", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if path:
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
            args=(pdf_path, out_path, dpi),
            daemon=True
        )
        thread.start()

    def _run_conversion(self, pdf_path, out_path, dpi):
        try:
            process_pdf(
                pdf_path, out_path, dpi=dpi,
                log_callback=lambda msg: self.log_signal.emit(msg),
                progress_callback=lambda val: self.progress_signal.emit(val)
            )
            self.log_signal.emit("\n✅ 转换完成！")
        except Exception as e:
            self.log_signal.emit(f"\n❌ 发生错误: {str(e)}")
        finally:
            self.done_signal.emit()


def main():
    if len(sys.argv) > 1:
        # 命令行模式
        parser = argparse.ArgumentParser(description='PDF 表格转 Excel（支持扫描件）')
        parser.add_argument('pdf_file', nargs='?', default='Scan.pdf', help='PDF 文件路径')
        parser.add_argument('-o', '--output', default=None, help='输出 Excel 文件路径')
        parser.add_argument('--dpi', type=int, default=300, help='OCR 分辨率（默认300）')
        args = parser.parse_args()
        output = args.output or f"{Path(args.pdf_file).stem}.xlsx"
        process_pdf(args.pdf_file, output, args.dpi)
    else:
        # GUI 模式
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec_())

if __name__ == '__main__':
    main()

