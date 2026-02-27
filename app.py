# -*- coding: utf-8 -*-
"""
KB023 고객관리 - Flask 서버
information/insurance/accident/counsel xlsx CRUD 및 accident/counsel 추가 저장.
"""
import logging
import os
import re
from datetime import datetime, date

# 204/304 등 정상 요청 로그를 터미널에 찍지 않음 (에러만 출력)
logging.getLogger('werkzeug').setLevel(logging.ERROR)

from flask import Flask, request, jsonify, send_from_directory
import openpyxl

app = Flask(__name__, static_folder='.', static_url_path='')
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INFORMATION_FILE = os.path.join(BASE_DIR, 'kb023_information.xlsx')
INSURANCE_FILE   = os.path.join(BASE_DIR, 'kb023_insurance.xlsx')
ACCIDENT_FILE    = os.path.join(BASE_DIR, 'kb023_accident.xlsx')
COUNSEL_FILE     = os.path.join(BASE_DIR, 'kb023_counse.xlsx')

# ---------------------------------------------------------------------------
# 유틸리티
# ---------------------------------------------------------------------------

def normalize_jumin(v):
    if v is None:
        return ''
    s = re.sub(r'\D', '', str(v).strip())
    return s[:13] if s else ''

def str_to_yyyymmdd(s):
    """'yy/mm/dd' 또는 'yyyy/mm/dd' 문자열 → YYYYMMDD 정수 (실패 시 None)."""
    if not s or not isinstance(s, str):
        return None
    parts = re.sub(r'\D+', ' ', s.strip()).split()
    if len(parts) < 3:
        return None
    try:
        y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
        if y < 100:
            y += 2000 if y < 50 else 1900
        if 1 <= m <= 12 and 1 <= d <= 31:
            return y * 10000 + m * 100 + d
    except (ValueError, TypeError):
        pass
    return None

def yyyymmdd_to_excel_date(n):
    """YYYYMMDD 정수 → Excel 날짜 시리얼 (실패 시 None)."""
    if n is None or not isinstance(n, (int, float)):
        return None
    n = int(n)
    if n < 19000101 or n > 29991231:
        return None
    y, m, d = n // 10000, (n // 100) % 100, n % 100
    try:
        return (datetime(y, m, d) - datetime(1899, 12, 30)).days
    except ValueError:
        return None

def _date_value_for_sheet(h, v):
    """배정일자·종료일자·납입만기 문자열을 Excel 날짜 시리얼로 변환."""
    if v is None or v == '':
        return ''
    if h in ('배정일자', '종료일자', '납입만기') and isinstance(v, str):
        n = str_to_yyyymmdd(v)
        if n is not None:
            serial = yyyymmdd_to_excel_date(n)
            return serial if serial is not None else v
    return v

def _amount_to_number(v):
    """'1,234' 형태의 금액 문자열을 float으로 변환. 변환 불가 시 원값 반환."""
    if isinstance(v, str) and re.match(r'^[\d,]+$', v):
        try:
            return float(v.replace(',', ''))
        except ValueError:
            pass
    return v

def _load_or_create_wb(file_path, default_headers, sheet_title='Sheet1'):
    """xlsx 로드(헤더 포함) 또는 신규 워크북+헤더 행 생성. (wb, ws, headers) 반환."""
    if os.path.isfile(file_path):
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        headers = [str(cell.value).strip() if cell.value is not None else '' for cell in ws[1]]
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_title
        headers = list(default_headers)
        ws.append(headers)
    return wb, ws, headers

def _delete_rows_by_key(ws, headers, key_name, key_jumin):
    """ws에서 계약자명+주민번호가 일치하는 행을 모두 삭제. 삭제 건수 반환."""
    name_col = next((i for i, h in enumerate(headers) if h == '계약자명'), None)
    jumin_col = next((i for i, h in enumerate(headers) if h in ('주민등록번호', '주민번호')), None)
    if name_col is None or jumin_col is None:
        return 0
    to_delete = [
        r for r in range(2, ws.max_row + 1)
        if (ws.cell(row=r, column=name_col + 1).value or '').strip() == key_name
        and normalize_jumin(ws.cell(row=r, column=jumin_col + 1).value) == key_jumin
    ]
    for r in reversed(to_delete):
        ws.delete_rows(r, 1)
    return len(to_delete)

def _build_log_row(headers, row_dict):
    """accident/counsel 행 값 배열 생성 (기록일자·기록시간 변환 포함)."""
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
    return row_values

# ---------------------------------------------------------------------------
# 라우트
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return send_from_directory(BASE_DIR, 'index.html')

@app.route('/favicon.ico')
def favicon():
    """favicon 요청 시 파일 없으면 204 반환하여 404 로그 방지."""
    path = os.path.join(BASE_DIR, 'favicon.ico')
    if os.path.isfile(path):
        return send_from_directory(BASE_DIR, 'favicon.ico')
    return '', 204

@app.route('/<path:path>')
def static_file(path):
    return send_from_directory(BASE_DIR, path)

# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------

@app.route('/api/information', methods=['POST'])
def api_information():
    """information xlsx: action=insert|update|delete."""
    try:
        data = request.get_json(force=True) or {}
        action = (data.get('action') or '').strip()
        if action not in ('insert', 'update', 'delete'):
            return jsonify({'ok': False, 'error': 'action은 insert/update/delete 중 하나여야 합니다.'}), 400

        default_hdrs = data.get('headers') or ['계약자명', '주민번호', '배정일자', '종료일자', '문자전송', 'DB종류', '연락처']
        wb, ws, headers = _load_or_create_wb(INFORMATION_FILE, default_hdrs)
        row_index = data.get('rowIndex')
        row_dict  = data.get('row') or {}

        if action == 'delete':
            if row_index is None:
                return jsonify({'ok': False, 'error': '삭제 시 rowIndex 필요'}), 400
            excel_row = 2 + int(row_index)
            if excel_row < 2 or excel_row > ws.max_row:
                return jsonify({'ok': False, 'error': '유효하지 않은 행'}), 400
            ws.delete_rows(excel_row, 1)
        elif action == 'update':
            if row_index is None:
                return jsonify({'ok': False, 'error': '수정 시 rowIndex 필요'}), 400
            excel_row = 2 + int(row_index)
            if excel_row < 2 or excel_row > ws.max_row:
                return jsonify({'ok': False, 'error': '유효하지 않은 행'}), 400
            for col, h in enumerate(headers, 1):
                ws.cell(row=excel_row, column=col, value=_date_value_for_sheet(h, row_dict.get(h, '')))
        else:
            ws.append([_date_value_for_sheet(h, row_dict.get(h, '')) for h in headers])

        wb.save(INFORMATION_FILE)
        wb.close()
        return jsonify({'ok': True, 'action': action})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/insurance', methods=['POST'])
def api_insurance():
    """insurance xlsx: action=insert|update|delete. update/delete는 key로 행 식별."""
    try:
        data = request.get_json(force=True) or {}
        action = (data.get('action') or '').strip()
        if action not in ('insert', 'update', 'delete'):
            return jsonify({'ok': False, 'error': 'action은 insert/update/delete 중 하나여야 합니다.'}), 400

        key_raw   = data.get('key') or {}
        key_name  = (key_raw.get('계약자명') or '').strip()
        key_jumin = normalize_jumin(key_raw.get('주민번호') or key_raw.get('주민등록번호'))

        default_hdrs = ['계약자명', '주민번호', '상품종류', '상품명', '증권번호', '납입만기', '월보험료', '납입보험료', '총보험료', '해지환급금', '대출금']
        wb, ws, headers = _load_or_create_wb(INSURANCE_FILE, default_hdrs)
        if '보험종류' in headers and '상품종류' not in headers:
            headers = ['상품종류' if h == '보험종류' else h for h in headers]

        row_dict = data.get('row') or {}

        def _make_row(src):
            return [_amount_to_number(_date_value_for_sheet(h, src.get(h, '') or (src.get('상품종류', '') if h == '보험종류' else ''))) for h in headers]

        def row_matches(r):
            return (r.get('계약자명') or '').strip() == key_name and \
                   normalize_jumin(r.get('주민번호') or r.get('주민등록번호')) == key_jumin

        if action == 'delete':
            _delete_rows_by_key(ws, headers, key_name, key_jumin)
        elif action == 'update':
            updated = False
            for r in range(2, ws.max_row + 1):
                row_obj = {h: ws.cell(row=r, column=c).value for c, h in enumerate(headers, 1)}
                if row_matches(row_obj):
                    merged = {h: row_dict.get(h, row_obj.get(h, '')) for h in headers}
                    for c, h in enumerate(headers, 1):
                        ws.cell(row=r, column=c, value=_amount_to_number(_date_value_for_sheet(h, merged[h])))
                    updated = True
                    break
            if not updated:
                ws.append(_make_row(row_dict))
        else:
            ws.append(_make_row(row_dict))

        wb.save(INSURANCE_FILE)
        wb.close()
        return jsonify({'ok': True, 'action': action})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/accident', methods=['POST'])
def accident_append():
    """accident xlsx: action=deleteRow(행 인덱스) | delete(키) | append(기본)."""
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({'ok': False, 'error': 'JSON 없음'}), 400
        action = (data.get('action') or '').strip()

        if action == 'deleteRow':
            row_index = data.get('rowIndex')
            if row_index is None:
                return jsonify({'ok': False, 'error': '삭제할 행의 rowIndex가 필요합니다.'}), 400
            if not os.path.isfile(ACCIDENT_FILE):
                return jsonify({'ok': True, 'deleted': 0})
            wb = openpyxl.load_workbook(ACCIDENT_FILE)
            ws = wb.active
            excel_row = 2 + int(row_index)
            if excel_row < 2 or excel_row > ws.max_row:
                wb.close()
                return jsonify({'ok': False, 'error': '유효하지 않은 행 인덱스입니다.'}), 400
            ws.delete_rows(excel_row, 1)
            wb.save(ACCIDENT_FILE)
            wb.close()
            return jsonify({'ok': True, 'deleted': 1})

        if action == 'delete':
            key_raw   = data.get('key') or {}
            key_name  = (key_raw.get('계약자명') or '').strip()
            key_jumin = normalize_jumin(key_raw.get('주민번호') or key_raw.get('주민등록번호'))
            if not os.path.isfile(ACCIDENT_FILE):
                return jsonify({'ok': True, 'deleted': 0})
            wb = openpyxl.load_workbook(ACCIDENT_FILE)
            ws = wb.active
            headers = [str(ws.cell(row=1, column=c).value).strip() if ws.cell(row=1, column=c).value else '' for c in range(1, ws.max_column + 1)]
            deleted = _delete_rows_by_key(ws, headers, key_name, key_jumin)
            wb.save(ACCIDENT_FILE)
            wb.close()
            return jsonify({'ok': True, 'deleted': deleted})

        # append
        default_hdrs  = ['계약자명', '주민번호', '구분', '사고일자', '사고내용', '입력일자', '기록일자', '기록시간', '문/답', '세부내용']
        headers       = data.get('headers') or default_hdrs
        rows_to_append = data.get('rowsToAppend') or []
        if not rows_to_append:
            return jsonify({'ok': True, 'appended': 0})

        today_serial = yyyymmdd_to_excel_date(
            date.today().year * 10000 + date.today().month * 100 + date.today().day
        )
        wb, ws, _ = _load_or_create_wb(ACCIDENT_FILE, headers, sheet_title='accident')
        for row_dict in rows_to_append:
            row_dict = dict(row_dict)
            if '입력일자' in headers and (row_dict.get('입력일자') is None or row_dict.get('입력일자') == ''):
                row_dict['입력일자'] = today_serial
            ws.append(_build_log_row(headers, row_dict))
        wb.save(ACCIDENT_FILE)
        wb.close()
        return jsonify({'ok': True, 'appended': len(rows_to_append)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/counsel', methods=['POST'])
def counsel_append():
    """counsel xlsx: action=delete(키) | append(기본)."""
    try:
        data = request.get_json(force=True)
        if not data:
            return jsonify({'ok': False, 'error': 'JSON 없음'}), 400
        action = (data.get('action') or '').strip()

        if action == 'delete':
            key_raw   = data.get('key') or {}
            key_name  = (key_raw.get('계약자명') or '').strip()
            key_jumin = normalize_jumin(key_raw.get('주민번호') or key_raw.get('주민등록번호'))
            if not os.path.isfile(COUNSEL_FILE):
                return jsonify({'ok': True, 'deleted': 0})
            wb = openpyxl.load_workbook(COUNSEL_FILE)
            ws = wb.active
            headers = [str(ws.cell(row=1, column=c).value).strip() if ws.cell(row=1, column=c).value else '' for c in range(1, ws.max_column + 1)]
            deleted = _delete_rows_by_key(ws, headers, key_name, key_jumin)
            wb.save(COUNSEL_FILE)
            wb.close()
            return jsonify({'ok': True, 'deleted': deleted})

        # append
        default_hdrs   = ['계약자명', '주민번호', '기록일자', '기록시간', '문/답', '세부내용']
        headers        = data.get('headers') or default_hdrs
        rows_to_append = data.get('rowsToAppend') or []
        if not rows_to_append:
            return jsonify({'ok': True, 'appended': 0})

        wb, ws, _ = _load_or_create_wb(COUNSEL_FILE, headers, sheet_title='counsel')
        for row_dict in rows_to_append:
            ws.append(_build_log_row(headers, row_dict))
        wb.save(COUNSEL_FILE)
        wb.close()
        return jsonify({'ok': True, 'appended': len(rows_to_append)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
