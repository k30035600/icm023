# -*- coding: utf-8 -*-
"""
KB023 고객관리 - Flask 서버
prebargaining.xlsx를 서버에서 추가(수정) 저장합니다.
"""
import os
from datetime import datetime, timedelta

from flask import Flask, request, jsonify, send_from_directory
import openpyxl

app = Flask(__name__, static_folder='.', static_url_path='')
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PREBARGAINING_FILE = os.path.join(BASE_DIR, 'kb023_prebargaining.xlsx')

# Excel 시리얼: YYYYMMDD(정수) -> Excel 날짜 시리얼
def yyyymmdd_to_excel_date(n):
    if n is None or not isinstance(n, (int, float)):
        return None
    n = int(n)
    if n < 19000101 or n > 29991231:
        return None
    y = n // 10000
    m = (n // 100) % 100
    d = n % 100
    try:
        dt = datetime(y, m, d)
        epoch = datetime(1899, 12, 30)
        return (dt - epoch).days
    except ValueError:
        return None

@app.route('/')
def index():
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/<path:path>')
def static_file(path):
    return send_from_directory(BASE_DIR, path)

@app.route('/api/prebargaining', methods=['POST'])
def prebargaining_append():
    """기존 kb023_prebargaining.xlsx에 행 추가(추가·수정)."""
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({'ok': False, 'error': 'JSON 없음'}), 400
        headers = data.get('headers') or ['계약자명', '주민등록번호', '구분', '기록일자', '기록시간', '문/답', '세부내용']
        rows_to_append = data.get('rowsToAppend') or []
        if not rows_to_append:
            return jsonify({'ok': True, 'appended': 0})

        if os.path.isfile(PREBARGAINING_FILE):
            wb = openpyxl.load_workbook(PREBARGAINING_FILE)
            ws = wb.active
            existing_rows = ws.max_row
            if existing_rows < 1:
                ws.append(headers)
                start_row = 2
            else:
                start_row = existing_rows + 1
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'prebargaining'
            ws.append(headers)
            start_row = 2

        for row_dict in rows_to_append:
            row_values = []
            for h in headers:
                v = row_dict.get(h)
                if h == '기록일자' and v is not None:
                    serial = yyyymmdd_to_excel_date(v)
                    row_values.append(serial if serial is not None else v)
                elif h == '기록시간' and v is not None and isinstance(v, (int, float)):
                    row_values.append(float(v))
                else:
                    row_values.append(v if v is not None else '')
            ws.append(row_values)

        wb.save(PREBARGAINING_FILE)
        wb.close()
        return jsonify({'ok': True, 'appended': len(rows_to_append)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
