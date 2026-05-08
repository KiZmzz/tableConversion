import pdfplumber
from rapidocr_onnxruntime import RapidOCR

ocr = RapidOCR()

def get_dynamic_bounds(img_path):
    result, _ = ocr(img_path)
    headers = {}
    if not result:
        return None
    for box, text, conf in result:
        x_center = (float(box[0][0]) + float(box[2][0])) / 2
        y_center = (float(box[0][1]) + float(box[2][1])) / 2
        if y_center > 1000: continue # Only look at top part for headers
        for k in ['物品名称', '规格', '数量', '单位', '单价', '金额', '备注']:
            if k in text and k not in headers:
                headers[k] = x_center
    
    print(f"Headers found: {headers}")
    if '数量' in headers and '单价' in headers and '金额' in headers:
        b0 = 0
        b1 = headers.get('规格', headers.get('物品名称', 0) + 200) if '规格' in headers else 250
        
        x_qty = headers['数量']
        x_unit = headers.get('单位', x_qty + 100)
        x_price = headers['单价']
        x_amount = headers['金额']
        x_remark = headers.get('备注', x_amount + 200)

        b2 = x_qty - 60
        b3 = (x_qty + x_unit) / 2
        b4 = (x_unit + x_price) / 2
        b5 = (x_price + x_amount) / 2
        b6 = (x_amount + x_remark) / 2
        b7 = 10000

        # Ensure monotonic
        bounds = [b0, b1, b2, b3, b4, b5, b6, b7]
        return sorted(bounds)
    return None

with pdfplumber.open("Scan.pdf") as pdf:
    pdf.pages[0].to_image(resolution=200).save("_tmp.png")
    bounds = get_dynamic_bounds("_tmp.png")
    print("Dynamic Bounds (200 DPI):", bounds)
    print("Old Bounds (200 DPI):    ", [0, 250, 750, 900, 1000, 1200, 1360, 1600])
