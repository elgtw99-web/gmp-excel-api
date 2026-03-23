from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from copy import copy
import json, shutil, os, tempfile, io

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), '配方表_模板.xlsx')
SUPPLIER_DB_PATH = os.path.join(os.path.dirname(__file__), 'supplier_db.json')

with open(SUPPLIER_DB_PATH, 'r', encoding='utf-8') as f:
    SUPPLIER_DB = json.load(f)

def lookup_supplier(inci):
    if not inci:
        return {}
    key = inci.strip().upper()
    if key in SUPPLIER_DB:
        return SUPPLIER_DB[key]
    for k, v in SUPPLIER_DB.items():
        if key in k or k in key:
            return v
    return {}

def copy_row_style(ws, src_row, dst_row, max_col=12):
    for c in range(1, max_col + 1):
        src = ws.cell(row=src_row, column=c)
        dst = ws.cell(row=dst_row, column=c)
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE_PATH)})

@app.route('/generate-excel', methods=['POST', 'OPTIONS'])
def generate_excel():
    if request.method == 'OPTIONS':
        return '', 204

    data = request.get_json()
    if not data:
        return jsonify({'error': '無效的請求資料'}), 400

    product_name = data.get('product_name', '新產品')
    batch_size_g = float(data.get('batch_size_g', 62500))
    batch_no = data.get('batch_no', '')
    date_str = data.get('date', '')
    ingredients = data.get('ingredients', [])
    process_steps = data.get('process_steps', [])

    if not ingredients:
        return jsonify({'error': '沒有原料資料'}), 400

    try:
        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        output_path = tmp.name

        shutil.copy(TEMPLATE_PATH, output_path)
        wb = load_workbook(output_path)
        src_ws = wb.worksheets[0]
        ws = wb.copy_worksheet(src_ws)
        ws.title = product_name[:31]

        # 填入基本資訊
        ws['F2'] = product_name
        ws['F3'] = batch_size_g
        ws['J2'] = batch_no
        ws['J3'] = date_str

        # 清空原料列（6到16）
        for r in range(6, 17):
            for c in range(1, 13):
                if c not in (9, 10):
                    ws.cell(row=r, column=c).value = None

        total = len(ingredients)
        default_rows = 11
        start_row = 6

        # 若原料超過預設列數，插入列
        if total > default_rows:
            extra = total - default_rows
            insert_at = start_row + default_rows
            ws.insert_rows(insert_at, extra)
            for new_r in range(insert_at, insert_at + extra):
                copy_row_style(ws, insert_at - 1, new_r)
                ws.cell(row=new_r, column=9).value = f'=SUM(H{new_r}*$F$3)'
                ws.cell(row=new_r, column=10).value = 'Gm'

        # 填入原料資料
        last_phase = None
        for i, ing in enumerate(ingredients):
            row = start_row + i
            phase = ing.get('phase', 'A')
            inci = ing.get('inci', '')

            # 若沒有提供廠商資訊，自動查詢
            company = ing.get('company', '') or lookup_supplier(inci).get('company', '')
            supplier_code = ing.get('supplierCode', '') or lookup_supplier(inci).get('supplierCode', '')
            product_code = ing.get('productCode', '') or lookup_supplier(inci).get('productCode', '')
            batch_no_raw = ing.get('batchNo', '')
            item_name = ing.get('name', inci)

            ws.cell(row=row, column=1).value = phase if phase != last_phase else None
            ws.cell(row=row, column=2).value = company
            ws.cell(row=row, column=3).value = supplier_code
            ws.cell(row=row, column=4).value = product_code
            ws.cell(row=row, column=5).value = batch_no_raw
            ws.cell(row=row, column=6).value = item_name
            ws.cell(row=row, column=7).value = inci
            ws.cell(row=row, column=8).value = ing.get('percentage', 0)
            ws.cell(row=row, column=9).value = f'=SUM(H{row}*$F$3)'
            ws.cell(row=row, column=10).value = 'Gm'
            last_phase = phase

        # Total 列
        total_row = start_row + total
        ws.cell(row=total_row, column=6).value = 'Total:'
        ws.cell(row=total_row, column=8).value = f'=SUM(H{start_row}:H{total_row-1})'
        ws.cell(row=total_row, column=9).value = f'=SUM(I{start_row}:I{total_row-1})'
        ws.cell(row=total_row, column=10).value = 'Gm'

        # 流程說明
        if process_steps:
            flow_row = total_row + 2
            ws.cell(row=flow_row, column=2).value = '流    程:'
            for j, step in enumerate(process_steps):
                ws.cell(row=flow_row + j, column=4).value = step

        wb.save(output_path)

        safe_name = product_name.replace('/', '_').replace('\\', '_')
        return send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'{safe_name}_秤料單.xlsx'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        try:
            if os.path.exists(output_path):
                os.unlink(output_path)
        except:
            pass

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
