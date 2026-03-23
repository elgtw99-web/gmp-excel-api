from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from copy import copy
import json, shutil, os, tempfile

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=False)

TEMPLATE_PATH    = os.path.join(os.path.dirname(__file__), '配方表_空白_.xlsx')
SUPPLIER_DB_PATH = os.path.join(os.path.dirname(__file__), 'supplier_db.json')

with open(SUPPLIER_DB_PATH, 'r', encoding='utf-8') as f:
    SUPPLIER_DB = json.load(f)

def lookup_supplier(inci):
    if not inci: return {}
    key = inci.strip().upper()
    if key in SUPPLIER_DB: return SUPPLIER_DB[key]
    for k, v in SUPPLIER_DB.items():
        if key in k or k in key: return v
    return {}

def copy_row_style(ws, src_row, dst_row, max_col=12):
    for c in range(1, max_col + 1):
        src = ws.cell(row=src_row, column=c)
        dst = ws.cell(row=dst_row, column=c)
        if src.has_style:
            dst.font       = copy(src.font)
            dst.border     = copy(src.border)
            dst.fill       = copy(src.fill)
            dst.alignment  = copy(src.alignment)
            dst.number_format = src.number_format

PHASE_KEYWORDS = {
    'A': ['water','aqua','glycerin','propanediol','butylene glycol','sodium hyaluronate',
          'niacinamide','allantoin','betaine','xanthan gum','carbomer','hydroxyethylcellulose',
          'algin','tranexamic acid','panthenol','dipotassium glycyrrhizinate','diglycerin',
          'sodium chloride','inositol','sodium lactate','hyaluronate'],
    'B': ['oil','butter','wax','silicone','dimethicone','cetyl','stearyl','palmitate',
          'oleate','myristate','caprylic','squalane','jojoba','olive','sunflower',
          'tocopherol','emulsifier','peg-','polysorbate','cetearyl','isononyl',
          'ethylhexyl','hydrogenated','lanolin','beeswax'],
    'C': ['phenoxyethanol','chlorphenesin','1,2-hexanediol','ethylhexylglycerin',
          'methylparaben','fragrance','parfum','hydroxyacetophenone','benzyl',
          'caprylyl glycol','sodium benzoate','potassium sorbate','dehydroacetic']
}

def infer_phase(inci, given_phase):
    if given_phase and given_phase.strip().upper() in ('A','B','C','D'):
        return given_phase.strip().upper()
    inci_lower = inci.lower()
    for phase, kws in PHASE_KEYWORDS.items():
        for kw in kws:
            if kw in inci_lower:
                return phase
    return 'A'

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE_PATH)})

@app.route('/generate-excel', methods=['POST', 'OPTIONS'])
def generate_excel():
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp

    data = request.get_json()
    if not data: return jsonify({'error': '無效請求'}), 400

    product_name  = data.get('product_name', '新產品')
    batch_size_g  = float(data.get('batch_size_g', 62500))
    batch_no      = data.get('batch_no', '')
    date_str      = data.get('date', '')
    ingredients   = data.get('ingredients', [])
    process_steps = data.get('process_steps', [])

    if not ingredients: return jsonify({'error': '沒有原料資料'}), 400

    output_path = None
    try:
        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        output_path = tmp.name

        shutil.copy(TEMPLATE_PATH, output_path)
        wb = load_workbook(output_path)
        ws = wb.active
        ws.title = product_name[:31]

        # 填表頭
        ws['F2'] = product_name
        ws['F3'] = batch_size_g
        ws['J2'] = batch_no
        ws['J3'] = date_str

        # 分組
        grouped = {'A': [], 'B': [], 'C': [], 'D': []}
        for ing in ingredients:
            ph = infer_phase(ing.get('inci',''), ing.get('phase',''))
            grouped[ph].append(ing)

        # 模板預設行數
        A_START, A_DEFAULT = 6, 2    # 列6-7
        B_START, B_DEFAULT = 8, 4    # 列8-11
        C_START, C_DEFAULT = 12, 3   # 列12-14
        TOTAL_ROW = 15

        need_A = max(len(grouped['A']), 1)
        need_B = max(len(grouped['B']), 1)
        need_C = max(len(grouped['C']) + len(grouped['D']), 1)

        extra_A = max(need_A - A_DEFAULT, 0)
        extra_B = max(need_B - B_DEFAULT, 0)
        extra_C = max(need_C - C_DEFAULT, 0)

        # 從後往前插入，避免行號位移
        if extra_C > 0:
            ins = C_START + C_DEFAULT
            ws.insert_rows(ins, extra_C)
            for r in range(ins, ins + extra_C):
                copy_row_style(ws, ins - 1, r)

        if extra_B > 0:
            ins = B_START + B_DEFAULT
            ws.insert_rows(ins, extra_B)
            for r in range(ins, ins + extra_B):
                copy_row_style(ws, ins - 1, r)

        if extra_A > 0:
            ins = A_START + A_DEFAULT
            ws.insert_rows(ins, extra_A)
            for r in range(ins, ins + extra_A):
                copy_row_style(ws, ins - 1, r)

        # 實際起始列（插入後重新計算）
        aA = A_START
        aB = B_START + extra_A
        aC = C_START + extra_A + extra_B
        aT = TOTAL_ROW + extra_A + extra_B + extra_C
        aPK = aT + 1

        # 清空原料區
        for r in range(aA, aT):
            for c in range(1, 13):
                ws.cell(row=r, column=c).value = None

        def fill_group(label, ings, start):
            for i, ing in enumerate(ings):
                row = start + i
                inci = ing.get('inci', '')
                name = ing.get('name', inci)
                pct  = ing.get('percentage', None)
                sup  = lookup_supplier(inci)

                ws.cell(row=row, column=1).value = label if i == 0 else None
                ws.cell(row=row, column=2).value = ing.get('company','')       or sup.get('company','')       or None
                ws.cell(row=row, column=3).value = ing.get('supplierCode','')  or sup.get('supplierCode','')  or None
                ws.cell(row=row, column=4).value = ing.get('productCode','')   or sup.get('productCode','')   or None
                ws.cell(row=row, column=5).value = ing.get('batchNo','') or None
                ws.cell(row=row, column=6).value = name or None
                ws.cell(row=row, column=7).value = inci or None
                if pct is not None:
                    ws.cell(row=row, column=8).value = pct
                ws.cell(row=row, column=9).value  = f'=SUM(H{row}*$F$3)'
                ws.cell(row=row, column=10).value = 'Gm'

        fill_group('A', grouped['A'], aA)
        fill_group('B', grouped['B'], aB)
        fill_group('C', grouped['C'] + grouped['D'], aC)

        # 合計
        ws.cell(row=aT, column=9).value  = f'=SUM(I{aA}:I{aT-1})'
        ws.cell(row=aT, column=10).value = 'Gm'

        # 包材
        ws.cell(row=aPK, column=2).value = '包材部分'

        # 流程說明（找到模板中的流程列）
        flow_row = None
        for r in range(aPK, ws.max_row + 1):
            v = ws.cell(row=r, column=2).value
            if v and '流' in str(v):
                flow_row = r
                break
        if flow_row is None:
            flow_row = aPK + 6

        if process_steps:
            steps = process_steps[:5]
        else:
            steps = []
            if grouped['A'] or grouped['B']:
                steps.append('A項混合均勻後加入B項攪拌均勻呈均質液體')
            if grouped['C']:
                steps.append('C項依序加入(A+B)中攪拌均勻')
            if grouped['D']:
                steps.append('D項依序加入混合均勻')

        ws.cell(row=flow_row, column=2).value = '流    程:'
        for j, step in enumerate(steps):
            ws.cell(row=flow_row + j, column=4).value = step

        wb.save(output_path)

        safe = product_name.replace('/', '_').replace('\\', '_')
        resp = send_file(
            output_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'{safe}_秤料單.xlsx'
        )
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    finally:
        try:
            if output_path and os.path.exists(output_path):
                os.unlink(output_path)
        except:
            pass

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
