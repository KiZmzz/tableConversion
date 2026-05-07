"""
PDF 表格转 Excel 工具（支持扫描件）
使用 RapidOCR 进行中文 OCR + 坐标聚类提取表格

依赖: pip install pdfplumber openpyxl rapidocr-onnxruntime opencv-python-headless
用法: python main.py [PDF文件] [-o 输出文件]
"""

import os, sys, re, argparse
from pathlib import Path

import cv2
import numpy as np
import pdfplumber
from rapidocr_onnxruntime import RapidOCR
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


def full_page_ocr(ocr_engine, img_path):
    """整页 OCR，返回带位置信息的结果"""
    result, _ = ocr_engine(img_path)
    items = []
    if not result:
        return items
    for box, text, conf in result:
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


def create_excel(rows, header, output_path, contract_info=None):
    """生成带格式的 Excel"""
    wb = Workbook()
    ws = wb.active
    ws.title = "购销合同"

    header_font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
    data_font = Font(name='微软雅黑', size=10)
    alt_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center')
    right = Alignment(horizontal='right', vertical='center')
    left = Alignment(vertical='center', wrap_text=True)

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

            if ri % 2 == 1:
                c.fill = alt_fill
        cur += 1

    # 合计
    ws.merge_cells(f'A{cur}:E{cur}')
    c = ws.cell(row=cur, column=1, value='合计')
    c.font = Font(name='微软雅黑', size=11, bold=True)
    c.alignment = center
    c.border = border
    for col in range(2, num_cols + 1):
        ws.cell(row=cur, column=col).border = border
    sc = ws.cell(row=cur, column=6)
    sc.value = f'=SUM(F{header_row + 1}:F{cur - 1})'
    sc.font = Font(name='微软雅黑', size=11, bold=True)
    sc.alignment = right
    sc.border = border
    sc.number_format = '0.##'

    for i, w in enumerate([14, 40, 8, 6, 12, 12, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f'A{header_row + 1}'
    wb.save(output_path)


def process_pdf(pdf_path, output_path, dpi=300):
    print(f"\n{'=' * 55}")
    print(f"  PDF 表格转 Excel 工具")
    print(f"{'=' * 55}")
    print(f"\n📄 输入：{pdf_path}")
    print(f"📊 输出：{output_path}")

    if not os.path.exists(pdf_path):
        print(f"\n❌ 找不到文件 '{pdf_path}'")
        sys.exit(1)

    ocr = RapidOCR(text_score=0.1, box_score_thresh=0.1)
    scale = dpi / 200.0  # 相对 200 DPI 的缩放比

    # 列边界（基于 200 DPI，会按 scale 缩放）
    # 列: 物品名称 | 规格型号 | 数量 | 单位 | 含税单价 | 金额 | 备注
    col_bounds_200 = [0, 250, 750, 900, 1000, 1200, 1360, 1600]

    # PDF 转图片 + OCR
    print(f"\n📷 步骤 1/3：PDF 转图片 (DPI={dpi})...")
    page_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            tmp = f"_tmp_page_{i}.png"
            page.to_image(resolution=dpi).save(tmp)
            print(f"  ✓ 第 {i + 1}/{len(pdf.pages)} 页")
            page_data.append(tmp)

    print(f"\n🔍 步骤 2/3：OCR + 表格提取...")
    all_rows = []
    contract_info = {}
    header = ['物品名称', '', '数量', '单位', '含税单价', '金额', '备注']
    col_bounds = [int(x * scale) for x in col_bounds_200]

    for pi, img_path in enumerate(page_data):
        print(f"\n  --- 第 {pi + 1} 页 ---")

        # 检测水平线
        h_lines = detect_h_lines(img_path, min_len=int(60 * scale))
        print(f"  水平线：{len(h_lines)} 条")

        # 整页 OCR
        items = full_page_ocr(ocr, img_path)
        print(f"  OCR 文本块：{len(items)} 个")

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
            if '合计' in texts and len(texts.replace(' ', '')) <= 10:
                break  # 合计行后面不再有表格数据
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

        # 后处理：反算数量并填充推断的单位
        for row in page_rows:
            # 尝试通过 金额 / 含税单价 计算数量
            qty_str = str(row.get(2, '')).strip()
            # 如果数量为空，或者包含字母（OCR误认小数字为字母如 'ew', 'uy', 'G'）
            if not qty_str or not qty_str.replace('.', '').isdigit():
                try:
                    price_str = str(row.get(4, '')).replace(' ', '').replace(',', '')
                    amount_str = str(row.get(5, '')).replace(' ', '').replace(',', '')
                    
                    if price_str and amount_str:
                        price = float(price_str)
                        amount = float(amount_str)
                        if price > 0:
                            qty = amount / price
                            # 如果计算出来的数量接近整数，则取整
                            if abs(qty - round(qty)) < 0.01:
                                row[2] = str(int(round(qty)))
                            else:
                                row[2] = f"{qty:.2f}"
                except ValueError:
                    pass

        all_rows.extend(page_rows)
        print(f"  ✓ 提取 {len(page_rows)} 行数据")

    print(f"\n  📋 共 {len(all_rows)} 行")

    # 预览
    print(f"\n  前5行预览：")
    for i, row in enumerate(all_rows[:5]):
        cells = []
        for c in range(7):
            val = row.get(c, '')[:18]
            cells.append(f'{val:18s}')
        print(f"  {i + 1}: {'|'.join(cells)}")

    # 生成 Excel
    print(f"\n📝 步骤 3/3：生成 Excel...")
    create_excel(all_rows, header, output_path, contract_info)

    # 清理
    for f in page_data:
        try:
            os.remove(f)
        except OSError:
            pass

    print(f"\n{'=' * 55}")
    print(f"  ✅ 完成！输出：{os.path.abspath(output_path)}")
    print(f"{'=' * 55}\n")


def main():
    parser = argparse.ArgumentParser(description='PDF 表格转 Excel（支持扫描件）')
    parser.add_argument('pdf_file', nargs='?', default='Scan.pdf', help='PDF 文件路径')
    parser.add_argument('-o', '--output', default=None, help='输出 Excel 文件路径')
    parser.add_argument('--dpi', type=int, default=300, help='OCR 分辨率（默认300）')
    args = parser.parse_args()
    output = args.output or f"{Path(args.pdf_file).stem}.xlsx"
    process_pdf(args.pdf_file, output, args.dpi)


if __name__ == '__main__':
    main()
