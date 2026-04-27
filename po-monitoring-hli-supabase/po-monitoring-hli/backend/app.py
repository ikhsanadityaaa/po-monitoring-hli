from flask import Flask, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
import pandas as pd
import re
import os
from datetime import datetime, date, timedelta
import io
from sqlalchemy import func, text
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)
CORS(app, resources={r"/api/*": {
    "origins": "*",
    "methods": ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    "allow_headers": ["Content-Type", "Authorization", "Accept"]
}})

_db_url = os.environ.get('DATABASE_URL', '')
if _db_url:
    if _db_url.startswith('postgres://'):
        _db_url = _db_url.replace('postgres://', 'postgresql://', 1)
    app.config['SQLALCHEMY_DATABASE_URI'] = _db_url
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'pool_pre_ping': True, 'pool_recycle': 300, 'pool_size': 5, 'max_overflow': 10,
    }
else:
    _inst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'instance')
    os.makedirs(_inst, exist_ok=True)
    app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{_inst}/po_database.db'

app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
db = SQLAlchemy(app)

class POData(db.Model):
    __tablename__ = 'po_data'
    id = db.Column(db.Integer, primary_key=True)
    po_number = db.Column(db.String(50), index=True)
    item_no = db.Column(db.String(50))
    po_item_detail = db.Column(db.Text)
    item_code = db.Column(db.String(50))
    po_item_type = db.Column(db.String(100))
    supplier = db.Column(db.String(200))
    vendor_name_smro = db.Column(db.String(200))
    qty = db.Column(db.Float)
    unit = db.Column(db.String(20))
    price = db.Column(db.Float)
    amount = db.Column(db.Float)
    currency = db.Column(db.String(10))
    po_date = db.Column(db.Date)
    purchase_member = db.Column(db.String(200))
    request_delivery = db.Column(db.Date)
    delivery_plan_date = db.Column(db.Date)
    remarks = db.Column(db.Text)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)

class SOData(db.Model):
    __tablename__ = 'so_data'
    id = db.Column(db.Integer, primary_key=True)
    so_number = db.Column(db.String(50), index=True)
    so_item = db.Column(db.String(100))
    so_status = db.Column(db.String(50))
    operation_unit_name = db.Column(db.String(200))
    vendor_name = db.Column(db.String(200))
    customer_po_number = db.Column(db.String(200))
    delivery_memo = db.Column(db.Text)
    product_name = db.Column(db.Text)
    so_qty = db.Column(db.Float)
    sales_unit = db.Column(db.String(20))
    sales_price = db.Column(db.Float)
    sales_amount = db.Column(db.Float)
    currency = db.Column(db.String(10))
    purchasing_price = db.Column(db.Float)
    purchasing_amount = db.Column(db.Float)
    so_create_date = db.Column(db.Date)
    delivery_possible_date = db.Column(db.Date)
    matched_po_number = db.Column(db.String(50))
    delivery_plan_date = db.Column(db.Date)
    remarks = db.Column(db.Text)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)

class UploadLog(db.Model):
    __tablename__ = 'upload_log'
    id = db.Column(db.Integer, primary_key=True)
    file_type = db.Column(db.String(50))
    filename = db.Column(db.String(255))
    records_count = db.Column(db.Integer)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)

class DeleteRequest(db.Model):
    __tablename__ = 'delete_request'
    id = db.Column(db.Integer, primary_key=True)
    ref_type = db.Column(db.String(10))
    ref_number = db.Column(db.String(100))
    reason = db.Column(db.Text)
    requested_at = db.Column(db.DateTime, default=datetime.utcnow)
    is_hidden = db.Column(db.Boolean, default=True)

with app.app_context():
    db.create_all()
    print('DB schema ready.')

CLOSED_STATUSES = {
    'Delivery Completed', 'SO Cancel',
    'Approval Apply', 'Approval Complete Step', 'Approval Reject'
}

EXCLUDED_OP_UNITS = {'HLI GREEN POWER (CONSUMABLE)'}

PO_HLI_RE = re.compile(r'(?:[A-Za-z]{1,4}[-]?)?(\d{7,})(?:-(\d+))?')

# ─── Indonesian Public Holidays (extend as needed) ────────────────────────────
# Add dates in YYYY-MM-DD format. These are skipped in business day calculations.
INDONESIAN_HOLIDAYS = set([
    # 2024
    '2024-01-01', '2024-02-08', '2024-02-09', '2024-02-10',
    '2024-03-11', '2024-03-29', '2024-04-08', '2024-04-09', '2024-04-10',
    '2024-04-11', '2024-04-14', '2024-04-15', '2024-05-01', '2024-05-09',
    '2024-05-23', '2024-06-01', '2024-06-17', '2024-06-18',
    '2024-08-17', '2024-09-16', '2024-12-25', '2024-12-26',
    # 2025
    '2025-01-01', '2025-01-27', '2025-01-28', '2025-01-29',
    '2025-03-28', '2025-03-29', '2025-03-30', '2025-03-31',
    '2025-04-18', '2025-05-01', '2025-05-12', '2025-05-29',
    '2025-06-01', '2025-06-06', '2025-08-17', '2025-09-05',
    '2025-12-25', '2025-12-26',
    # 2026
    '2026-01-01', '2026-01-16', '2026-01-17',
    '2026-03-19', '2026-03-20', '2026-04-02', '2026-04-03',
    '2026-05-01', '2026-05-14', '2026-05-21', '2026-05-26',
    '2026-06-01', '2026-08-17', '2026-12-25',
])

def _holiday_set():
    result = set()
    for s in INDONESIAN_HOLIDAYS:
        try:
            result.add(date.fromisoformat(s))
        except Exception:
            pass
    return result

_HOLIDAY_DATES = _holiday_set()

def count_business_days(start_date, end_date):
    """Count business days (Mon–Fri, excluding Indonesian public holidays) between two dates.
    Returns positive int. start_date is the SO create date, end_date is today."""
    if start_date is None or end_date is None:
        return None
    if start_date > end_date:
        # negative aging means future date
        return -count_business_days(end_date, start_date)
    count = 0
    current = start_date
    while current <= end_date:
        if current.weekday() < 5 and current not in _HOLIDAY_DATES:  # Mon=0 … Fri=4
            count += 1
        current += timedelta(days=1)
    # subtract 1 so that same-day = 0 business days elapsed
    return max(0, count - 1)

def business_days_remaining(from_date, to_date):
    """Business days from today to a target date. Negative = overdue."""
    if from_date is None or to_date is None:
        return None
    if from_date <= to_date:
        return count_business_days(from_date, to_date)
    else:
        return -count_business_days(to_date, from_date)


def _normalize_item_no(item_no):
    if item_no is None:
        return set()
    s = str(item_no).strip()
    variants = {s}
    if s.endswith('.0'):
        s = s[:-2]
        variants.add(s)
    try:
        n = int(float(s))
        variants.add(str(n))
        variants.add(f"{n:02d}")
        variants.add(f"{n:03d}")
    except (ValueError, OverflowError):
        pass
    return variants

def extract_po_hli(val):
    if not val:
        return []
    text = str(val).strip()
    result = set()
    for m in PO_HLI_RE.finditer(text):
        po_num  = m.group(1)
        item_no = m.group(2)
        result.add(po_num)
        if item_no:
            for item_var in _normalize_item_no(item_no):
                result.add(f"{po_num}-{item_var}")
    return list(result)

def open_so_filter():
    return db.or_(
        SOData.so_status.is_(None),
        SOData.so_status.notin_(list(CLOSED_STATUSES))
    )

def is_return_so_item(so_item):
    if not so_item:
        return False
    return str(so_item).strip().startswith('9')

def has_internal_po_ref(customer_po_number, delivery_memo):
    for field in [customer_po_number, delivery_memo]:
        if not field:
            continue
        text = str(field).strip()
        for token in re.split(r'[\s,;]+', text):
            token = token.strip()
            if token and token[0] == '2' and re.match(r'^2\d{6,}', token):
                return True
    return False

def so_is_countable(so_item, so_number=None, customer_po_number=None, delivery_memo=None):
    if is_return_so_item(so_item):
        return False
    if has_internal_po_ref(customer_po_number, delivery_memo):
        return False
    return True

def clean(val):
    if val is None: return None
    try:
        if pd.isna(val): return None
    except (TypeError, ValueError): pass
    s = str(val).strip()
    return None if s.lower() in ('nan', 'none', '') else s

def parse_date(val):
    if val is None: return None
    try:
        if pd.isna(val): return None
    except (TypeError, ValueError): pass
    try: return pd.to_datetime(val).date()
    except: return None

def safe_float(val, default=0.0):
    try:
        if pd.isna(val): return default
    except (TypeError, ValueError): pass
    try: return float(val)
    except: return default

def find_column(df, names):
    low = {c.lower().strip(): c for c in df.columns}
    for n in names:
        if n.lower().strip() in low: return low[n.lower().strip()]
    return None

def df_val(row, col):
    return row.get(col) if col else None

def get_aging_label(biz_days):
    """Classify aging bucket using business days."""
    if biz_days is None: return 'No Date'
    if biz_days >= 180: return '180+'
    if biz_days >= 90:  return '90-180'
    if biz_days >= 30:  return '30-90'
    return '0-30'

def so_dict(s):
    today = date.today()
    biz_days = count_business_days(s.so_create_date, today) if s.so_create_date else None
    return {
        'id': s.id, 'so_number': s.so_number, 'so_item': s.so_item,
        'so_status': s.so_status, 'operation_unit_name': s.operation_unit_name,
        'vendor_name': s.vendor_name, 'customer_po_number': s.customer_po_number,
        'delivery_memo': s.delivery_memo, 'product_name': s.product_name,
        'so_qty': s.so_qty, 'sales_price': s.sales_price, 'sales_amount': s.sales_amount,
        'purchasing_price': s.purchasing_price, 'purchasing_amount': s.purchasing_amount,
        'so_create_date': s.so_create_date.isoformat() if s.so_create_date else '',
        'delivery_possible_date': s.delivery_possible_date.isoformat() if s.delivery_possible_date else '',
        'delivery_plan_date': s.delivery_plan_date.isoformat() if s.delivery_plan_date else '',
        'remarks': s.remarks or '',
        'aging_days': biz_days,
        'aging_label': get_aging_label(biz_days)
    }

# ─── Build hidden set from delete requests ────────────────────────────────
def get_hidden_po_hli_keys():
    reqs = DeleteRequest.query.filter_by(ref_type='PO', is_hidden=True).all()
    return {r.ref_number for r in reqs}

def get_hidden_so_items():
    reqs = DeleteRequest.query.filter_by(ref_type='SO', is_hidden=True).all()
    return {r.ref_number for r in reqs}

def get_hidden_po_numbers():
    return get_hidden_po_hli_keys()

def po_hli_key(po_number, item_no):
    if not po_number:
        return None
    if item_no:
        try:
            n = int(float(str(item_no).strip()))
            norm = str(n)
        except (ValueError, OverflowError):
            norm = str(item_no).strip()
        return f"{po_number}-{norm}"
    return po_number

def is_po_hidden(po_number, item_no, hidden_keys):
    key = po_hli_key(po_number, item_no)
    if key and key in hidden_keys:
        return True
    if item_no:
        for var in _normalize_item_no(item_no):
            if f"{po_number}-{var}" in hidden_keys:
                return True
    if po_number in hidden_keys:
        return True
    return False


@app.route('/api/dashboard/stats', methods=['GET'])
def get_dashboard_stats():
    try:
        hidden_po = get_hidden_po_numbers()
        hidden_so = get_hidden_so_items()

        total_po_amount = db.session.query(func.sum(POData.amount)).scalar() or 0

        # ── FIX: total_so_count uses same filter as aging: open + so_is_countable ──
        today = date.today()
        total_so_count = 0
        for s in db.session.query(SOData).filter(open_so_filter()).all():
            if s.so_item in hidden_so or s.so_number in hidden_so:
                continue
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            total_so_count += 1

        po_numbers = get_po_hli_key_set()
        matched_set = build_matched_set()

        po_without_so_count = 0
        for p in POData.query.all():
            if is_po_hidden(p.po_number, p.item_no, hidden_po):
                continue
            op_unit = get_operation_unit(p.po_item_type, p.item_code)
            if op_unit in EXCLUDED_OP_UNITS:
                continue
            if not po_is_matched(p.po_number, p.item_no, matched_set):
                po_without_so_count += 1

        so_without_po_count = 0
        for s in db.session.query(SOData).filter(
                open_so_filter(),
                ~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))).all():
            if s.so_item in hidden_so or s.so_number in hidden_so:
                continue
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            candidates = extract_po_hli(s.customer_po_number) + extract_po_hli(s.delivery_memo)
            if not candidates or not any(c in po_numbers for c in candidates):
                so_without_po_count += 1

        monthly = {}
        for s in db.session.query(SOData).filter(open_so_filter()).all():
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            if s.so_item in hidden_so or s.so_number in hidden_so:
                continue
            d = s.so_create_date
            amt = s.sales_amount
            if d:
                k = d.strftime('%b %Y')
                if k not in monthly:
                    monthly[k] = {'month': k, 'so_count': 0, 'amount': 0.0, '_s': d.replace(day=1)}
                monthly[k]['so_count'] += 1
                monthly[k]['amount'] += round((amt or 0) / 1_000_000, 2)
        monthly_trend = sorted(monthly.values(), key=lambda x: x['_s'])
        for m in monthly_trend: del m['_s']

        top_vendors = [
            {'vendor': r[0], 'so_count': r[1], 'total_amount': round(r[2] or 0, 2)}
            for r in db.session.query(
                SOData.vendor_name, func.count(SOData.id), func.sum(SOData.sales_amount)
            ).filter(open_so_filter(), SOData.vendor_name.isnot(None))
             .group_by(SOData.vendor_name)
             .order_by(func.sum(SOData.sales_amount).desc()).limit(5).all()
        ]

        top_op_units = [
            {'op_unit': r[0], 'so_count': r[1], 'total_amount': round(r[2] or 0, 2)}
            for r in db.session.query(
                SOData.operation_unit_name, func.count(SOData.id), func.sum(SOData.sales_amount)
            ).filter(open_so_filter(), SOData.operation_unit_name.isnot(None))
             .group_by(SOData.operation_unit_name)
             .order_by(func.sum(SOData.sales_amount).desc()).limit(10).all()
        ]

        total_open_for_pct = total_so_count or 1
        so_status = [{'name': r[0], 'value': r[1],
            'percentage': round(r[1] / total_open_for_pct * 100, 1),
            'amount': round(r[2] or 0, 2)
        } for r in db.session.query(
            SOData.so_status, func.count(SOData.id), func.sum(SOData.sales_amount)
        ).filter(open_so_filter(), SOData.so_status.isnot(None))
         .group_by(SOData.so_status)
         .order_by(func.count(SOData.id).desc()).all()]

        monthly_by_status = {}
        all_months_set = set()
        for s_status, s_date, s_amt in db.session.query(
                SOData.so_status, SOData.so_create_date, SOData.sales_amount
        ).filter(open_so_filter()).all():
            st = s_status or 'Unknown'
            amt_v = float(s_amt or 0)
            if s_date:
                mk = s_date.strftime('%b %Y')
                all_months_set.add((s_date.replace(day=1), mk))
            else:
                mk = None
            if st not in monthly_by_status:
                monthly_by_status[st] = {'monthly': {}, 'total': 0, 'amount': 0.0}
            monthly_by_status[st]['total'] += 1
            monthly_by_status[st]['amount'] += amt_v
            if mk:
                monthly_by_status[st]['monthly'][mk] = monthly_by_status[st]['monthly'].get(mk, 0) + 1

        sorted_months = [mk for _, mk in sorted(all_months_set)]
        so_status_monthly = sorted(
            [{'name': st, 'monthly': d['monthly'], 'total': d['total'],
              'percentage': round(d['total'] / total_open_for_pct * 100, 1),
              'amount': round(d['amount'], 2)}
             for st, d in monthly_by_status.items()],
            key=lambda x: x['total'], reverse=True
        )

        total_open_so_amount = db.session.query(func.sum(SOData.sales_amount))\
                                          .filter(open_so_filter()).scalar() or 0

        po_date_range = db.session.query(func.min(POData.po_date), func.max(POData.po_date)).first()
        so_date_range = db.session.query(func.min(SOData.so_create_date), func.max(SOData.so_create_date)).first()

        return jsonify({
            'po_without_so': po_without_so_count,
            'so_without_po': so_without_po_count,
            'total_po_amount': float(total_po_amount),
            'total_so_count': total_so_count,
            'total_open_so_amount': float(total_open_so_amount),
            'monthly_trend': monthly_trend,
            'top_vendors': top_vendors,
            'top_op_units': top_op_units,
            'so_status': so_status,
            'so_status_monthly': so_status_monthly,
            'status_months': sorted_months,
            'po_date_range': {
                'min': po_date_range[0].isoformat() if po_date_range and po_date_range[0] else None,
                'max': po_date_range[1].isoformat() if po_date_range and po_date_range[1] else None,
            },
            'so_date_range': {
                'min': so_date_range[0].isoformat() if so_date_range and so_date_range[0] else None,
                'max': so_date_range[1].isoformat() if so_date_range and so_date_range[1] else None,
            },
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


def get_operation_unit(po_item_type, item_code):
    t = (po_item_type or '').strip().upper()
    has_code = bool(item_code and item_code.strip())
    if has_code:
        if t == 'MRO':
            return 'HLI GREEN POWER (CONSUMABLE)'
        else:
            return 'HLI GREEN POWER(BONDED AREA)'
    else:
        if t == 'EQUIPMENT':
            return 'HLI GREEN POWER'
        else:
            return 'HLI GREEN POWER(BONDED AREA)'


def build_matched_set():
    matched = set()
    for s in db.session.query(
            SOData.customer_po_number, SOData.delivery_memo, SOData.so_item,
            SOData.operation_unit_name).all():
        cust_po, memo, so_item, op_unit = s[0], s[1], s[2], s[3]
        if op_unit in EXCLUDED_OP_UNITS:
            continue
        if is_return_so_item(so_item):
            continue
        for ref in extract_po_hli(cust_po) + extract_po_hli(memo):
            matched.add(ref)
    return matched


@app.route('/api/debug/matching', methods=['GET'])
def debug_matching():
    po_number = request.args.get('po_number', '').strip()
    item_no   = request.args.get('item_no', '').strip() or None
    if not po_number:
        return jsonify({'error': 'Provide ?po_number=xxxx'}), 400

    matched_set = build_matched_set()
    item_variants = list(_normalize_item_no(item_no)) if item_no else []
    keys_checked = [po_number] + [f"{po_number}-{v}" for v in item_variants]
    hits = [k for k in keys_checked if k in matched_set]

    matching_so_rows = []
    for s in db.session.query(SOData).all():
        refs = extract_po_hli(s.customer_po_number) + extract_po_hli(s.delivery_memo)
        if any(r.startswith(po_number) for r in refs):
            matching_so_rows.append({
                'so_number': s.so_number,
                'so_item': s.so_item,
                'so_status': s.so_status,
                'customer_po_number': s.customer_po_number,
                'delivery_memo': s.delivery_memo,
                'extracted_refs': refs,
            })

    sample_in_matched = [m for m in matched_set if po_number in m][:20]

    return jsonify({
        'po_number': po_number,
        'item_no': item_no,
        'item_variants': item_variants,
        'keys_checked': keys_checked,
        'is_matched': bool(hits),
        'matched_by': hits,
        'sample_matched_set_entries': sorted(sample_in_matched),
        'so_rows_referencing_this_po': matching_so_rows,
    })


def po_is_matched(po_number, item_no, matched_set):
    if po_number in matched_set:
        return True
    if item_no:
        for item_var in _normalize_item_no(item_no):
            if f"{po_number}-{item_var}" in matched_set:
                return True
    return False


def get_po_hli_key_set():
    keys = set()
    for p in POData.query.with_entities(POData.po_number, POData.item_no).all():
        if p.po_number:
            keys.add(p.po_number)
            if p.item_no:
                for var in _normalize_item_no(p.item_no):
                    keys.add(f"{p.po_number}-{var}")
    return keys


@app.route('/api/data/po-without-so', methods=['GET'])
def get_po_without_so():
    try:
        matched_set = build_matched_set()
        hidden_po = get_hidden_po_numbers()
        today = date.today()

        seen_keys = set()
        result = []
        for p in POData.query.all():
            key = (p.po_number, p.item_no)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            if is_po_hidden(p.po_number, p.item_no, hidden_po):
                continue
            op_unit = get_operation_unit(p.po_item_type, p.item_code)
            if op_unit in EXCLUDED_OP_UNITS:
                continue
            if not po_is_matched(p.po_number, p.item_no, matched_set):
                # Business days remaining to request delivery
                days_remaining = business_days_remaining(today, p.request_delivery) if p.request_delivery else None
                result.append({
                    'id': p.id, 'po_no': p.po_number, 'item_no': p.item_no,
                    'item_code': p.item_code,
                    'po_item_type': p.po_item_type or '',
                    'operation_unit': op_unit,
                    'description': p.po_item_detail, 'qty': p.qty, 'unit': p.unit or '',
                    'price': p.price or 0, 'amount': p.amount,
                    'currency': p.currency, 'supplier': p.supplier,
                    'po_date': p.po_date.isoformat() if p.po_date else '',
                    'purchase_member': p.purchase_member or '',
                    'req_delivery': p.request_delivery.isoformat() if p.request_delivery else '',
                    'days_remaining': days_remaining,
                    'delivery_plan_date': p.delivery_plan_date.isoformat() if p.delivery_plan_date else '',
                    'remarks': p.remarks or ''
                })
        return jsonify(result)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/so-without-po', methods=['GET'])
def get_so_without_po():
    try:
        po_hli_keys = get_po_hli_key_set()
        hidden_so = get_hidden_so_items()
        result = []
        for s in db.session.query(SOData).filter(
                open_so_filter(),
                ~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))).all():
            if s.so_item in hidden_so or s.so_number in hidden_so:
                continue
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            candidates = extract_po_hli(s.customer_po_number) + extract_po_hli(s.delivery_memo)
            if not candidates or not any(c in po_hli_keys for c in candidates):
                result.append(so_dict(s))
        return jsonify(result)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/aging', methods=['GET'])
def get_aging_data():
    try:
        today = date.today()
        vendors = {}
        # ── FIX: use same so_is_countable filter as total_so_count ──
        for s in db.session.query(SOData).filter(open_so_filter()).all():
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            # Include even records without so_create_date — they land in '0-30' via get_aging_label(None)
            v = s.vendor_name or 'Unknown'
            if v not in vendors:
                vendors[v] = {'vendor': v, 'less_30': 0, 'days_30_90': 0,
                              'days_90_180': 0, 'more_180': 0, 'total_open': 0, 'sales_amount': 0.0}
            biz_days = count_business_days(s.so_create_date, today) if s.so_create_date else None
            label = get_aging_label(biz_days)
            if label == '0-30' or label == 'No Date':
                vendors[v]['less_30']     += 1
            elif label == '30-90':
                vendors[v]['days_30_90']  += 1
            elif label == '90-180':
                vendors[v]['days_90_180'] += 1
            else:
                vendors[v]['more_180']    += 1
            vendors[v]['total_open'] += 1
            vendors[v]['sales_amount'] += float(s.sales_amount or 0)
        return jsonify(sorted(vendors.values(), key=lambda x: x['total_open'], reverse=True))
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/aging-detail/<path:vendor_name>', methods=['GET'])
def get_aging_detail(vendor_name):
    try:
        bucket = request.args.get('bucket')
        today = date.today()
        sos = db.session.query(SOData).filter(
            open_so_filter(), SOData.vendor_name == vendor_name
        ).order_by(SOData.so_create_date.asc()).all()
        sos = [s for s in sos if so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo)]
        if bucket:
            bucket = bucket.strip().replace(' ', '+')
            sos = [s for s in sos if get_aging_label(
                count_business_days(s.so_create_date, today) if s.so_create_date else None
            ) == bucket]
        return jsonify([so_dict(s) for s in sos])
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/aging-detail-all', methods=['GET'])
def get_aging_detail_all():
    try:
        bucket = request.args.get('bucket')
        if bucket:
            bucket = bucket.strip().replace(' ', '+')
        today = date.today()
        sos = db.session.query(SOData).filter(
            open_so_filter()
        ).order_by(SOData.vendor_name.asc(), SOData.so_create_date.asc()).all()
        sos = [s for s in sos if so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo)]
        if bucket:
            sos = [s for s in sos if get_aging_label(
                count_business_days(s.so_create_date, today) if s.so_create_date else None
            ) == bucket]
        return jsonify([so_dict(s) for s in sos])
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/all-so', methods=['GET'])
def get_all_so():
    try:
        op_units      = request.args.getlist('op_unit')
        vendors       = request.args.getlist('vendor')
        statuses      = request.args.getlist('status')
        aging         = request.args.getlist('aging')
        so_items      = request.args.getlist('so_item')
        margin_filter = request.args.get('margin_filter', 'all')
        page          = max(1, int(request.args.get('page', 1)))
        per_page      = min(500, int(request.args.get('per_page', 20)))

        today = date.today()
        q = SOData.query.filter(open_so_filter())
        if op_units:  q = q.filter(SOData.operation_unit_name.in_(op_units))
        if vendors:   q = q.filter(SOData.vendor_name.in_(vendors))
        if statuses:  q = q.filter(SOData.so_status.in_(statuses))
        if so_items:  q = q.filter(SOData.so_item.in_(so_items))

        all_sos = q.order_by(SOData.so_create_date.asc()).all()

        if aging:
            def matches_aging(s):
                biz = count_business_days(s.so_create_date, today) if s.so_create_date else None
                return get_aging_label(biz) in aging
            all_sos = [s for s in all_sos if matches_aging(s)]

        if margin_filter in ('positive', 'negative'):
            def calc_margin(s):
                po_amt = (s.purchasing_price or 0) * (s.so_qty or 0)
                return (s.sales_amount or 0) - po_amt
            if margin_filter == 'negative':
                all_sos = [s for s in all_sos if calc_margin(s) < 0]
            else:
                all_sos = [s for s in all_sos if calc_margin(s) >= 0]

        total = len(all_sos)
        paged = all_sos[(page-1)*per_page : page*per_page]

        op_units_opts = sorted({s.operation_unit_name for s in all_sos if s.operation_unit_name})
        vendors_opts  = sorted({s.vendor_name for s in all_sos if s.vendor_name})
        statuses_opts = sorted({s.so_status for s in all_sos if s.so_status})

        return jsonify({
            'data': [so_dict(s) for s in paged],
            'total': total, 'page': page, 'per_page': per_page,
            'filters': {'op_units': list(op_units_opts), 'vendors': list(vendors_opts), 'statuses': list(statuses_opts)}
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/so-status-detail/<path:status>', methods=['GET'])
def get_so_status_detail(status):
    try:
        month = request.args.get('month')
        sos = SOData.query.filter_by(so_status=status).all()
        if month:
            filtered = []
            for s in sos:
                if s.so_create_date and s.so_create_date.strftime('%b %Y') == month:
                    filtered.append(s)
            sos = filtered
        return jsonify([so_dict(s) for s in sos])
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/so-status-detail-all', methods=['GET'])
def get_so_status_detail_all():
    try:
        month = request.args.get('month')
        if month:
            sos = [s for s in SOData.query.filter(open_so_filter()).all()
                   if s.so_create_date and s.so_create_date.strftime('%b %Y') == month]
        else:
            sos = SOData.query.filter(open_so_filter()).order_by(SOData.so_create_date.desc()).all()
        return jsonify([so_dict(s) for s in sos])
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/top-vendor-detail/<path:vendor_name>', methods=['GET'])
def get_top_vendor_detail(vendor_name):
    try:
        sos = db.session.query(SOData).filter(
            open_so_filter(), SOData.vendor_name == vendor_name).all()
        return jsonify([so_dict(s) for s in sos])
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ═══════════════════════════════════════════════════════════════════
# UPLOAD ENDPOINTS
# ═══════════════════════════════════════════════════════════════════

@app.route('/api/upload/po-list', methods=['POST'])
def upload_po_list():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        col_po     = find_column(df, ['po no.','po no','po number','po'])
        col_item   = find_column(df, ['item no.','item no','item number','no. item'])
        col_detail = find_column(df, ['short text','description','item description','po item detail','detail'])
        col_code   = find_column(df, ['material','material number','item code','material code'])
        col_type   = find_column(df, ['po item type','item type','type','po type'])
        col_sup    = find_column(df, ['supplier','vendor','supplier name'])
        col_qty    = find_column(df, ['qty.','qty','quantity'])
        col_unit   = find_column(df, ['unit','uom','unit of measure'])
        col_price  = find_column(df, ['net price','price','unit price'])
        col_amount = find_column(df, ['amount','total amount','total','net value'])
        col_curr   = find_column(df, ['currency','curr'])
        col_date   = find_column(df, ['po date','order date','tanggal po','document date'])
        col_member = find_column(df, ['purchase member','buyer','purchasing member','purchaser'])
        col_req    = find_column(df, ['request delivery date','delivery date','req delivery','req. del. date'])

        if not col_po:
            return jsonify({'error': 'PO Number column not found'}), 400

        db.session.query(POData).delete()
        count = 0
        seen = set()
        for _, row in df.iterrows():
            po_num = clean(df_val(row, col_po))
            if not po_num:
                continue
            item_no = clean(df_val(row, col_item))
            key = (po_num, item_no)
            if key in seen:
                continue
            seen.add(key)
            rec = POData(
                po_number=po_num,
                item_no=item_no,
                po_item_detail=clean(df_val(row, col_detail)),
                item_code=clean(df_val(row, col_code)),
                po_item_type=clean(df_val(row, col_type)),
                supplier=clean(df_val(row, col_sup)),
                qty=safe_float(df_val(row, col_qty)),
                unit=clean(df_val(row, col_unit)),
                price=safe_float(df_val(row, col_price)),
                amount=safe_float(df_val(row, col_amount)),
                currency=clean(df_val(row, col_curr)) or 'IDR',
                po_date=parse_date(df_val(row, col_date)),
                purchase_member=clean(df_val(row, col_member)),
                request_delivery=parse_date(df_val(row, col_req)),
            )
            db.session.add(rec)
            count += 1

        db.session.commit()
        db.session.add(UploadLog(file_type='po', filename=file.filename, records_count=count))
        db.session.commit()
        return jsonify({'message': f'PO List uploaded: {count} records saved'})
    except Exception as e:
        db.session.rollback()
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/upload/smro', methods=['POST'])
def upload_smro():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        col_so_num   = find_column(df, ['so number','so no','so no.','so','sales order','sales order number','no so','nomor so'])
        col_so_item  = find_column(df, ['so item no','item no','line','so line','so item'])
        col_status   = find_column(df, ['so status','status','order status'])
        col_op_unit  = find_column(df, ['operation unit name','op unit','client name','client','operation unit'])
        col_vendor   = find_column(df, ['vendor name','vendor','supplier'])
        col_cust_po  = find_column(df, ['customer po number','customer po','po ref','po reference'])
        col_memo     = find_column(df, ['delivery memo','memo','delivery note'])
        col_product  = find_column(df, ['product name','product','item name','description','material description'])
        col_qty      = find_column(df, ['so qty','qty','quantity','sales qty'])
        col_unit     = find_column(df, ['sales unit','unit','uom'])
        col_s_price  = find_column(df, ['sales price','unit price','price'])
        col_s_amount = find_column(df, ['sales amount(exclude tax)','sales amount','amount','total'])
        col_currency = find_column(df, ['currency','curr'])
        col_pur_price= find_column(df, ['purchasing price','po price','purchase price'])
        col_pur_amt  = find_column(df, ['purchasing amount','po amount','purchase amount'])
        col_date     = find_column(df, ['so create date','order date','so date','create date'])
        col_del_poss = find_column(df, ['delivery possible date','possible date','promised date','delivery possible'])

        if not col_so_num:
            return jsonify({'error': 'SO Number column not found'}), 400

        db.session.query(SOData).delete()
        count = 0
        for _, row in df.iterrows():
            so_num = clean(df_val(row, col_so_num))
            if not so_num:
                continue
            rec = SOData(
                so_number=so_num,
                so_item=clean(df_val(row, col_so_item)),
                so_status=clean(df_val(row, col_status)),
                operation_unit_name=clean(df_val(row, col_op_unit)),
                vendor_name=clean(df_val(row, col_vendor)),
                customer_po_number=clean(df_val(row, col_cust_po)),
                delivery_memo=clean(df_val(row, col_memo)),
                product_name=clean(df_val(row, col_product)),
                so_qty=safe_float(df_val(row, col_qty)),
                sales_unit=clean(df_val(row, col_unit)),
                sales_price=safe_float(df_val(row, col_s_price)),
                sales_amount=safe_float(df_val(row, col_s_amount)),
                currency=clean(df_val(row, col_currency)) or 'IDR',
                purchasing_price=safe_float(df_val(row, col_pur_price)),
                purchasing_amount=safe_float(df_val(row, col_pur_amt)),
                so_create_date=parse_date(df_val(row, col_date)),
                delivery_possible_date=parse_date(df_val(row, col_del_poss)),
            )
            db.session.add(rec)
            count += 1

        db.session.commit()
        db.session.add(UploadLog(file_type='smro', filename=file.filename, records_count=count))
        db.session.commit()
        return jsonify({'message': f'SMRO uploaded: {count} records saved'})
    except Exception as e:
        db.session.rollback()
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/so/<int:so_id>', methods=['PUT'])
def update_so(so_id):
    try:
        s = db.session.get(SOData, so_id)
        if not s:
            return jsonify({'error': 'SO not found'}), 404
        data = request.json
        if 'delivery_plan_date' in data:
            s.delivery_plan_date = parse_date(data['delivery_plan_date']) if data['delivery_plan_date'] else None
        if 'remarks' in data:
            s.remarks = data['remarks'] or None
        db.session.commit()
        return jsonify({'success': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/so/batch-upload', methods=['POST'])
def batch_upload_so():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        col_so_num  = find_column(df, ['SO Number', 'so number', 'so_number'])
        col_plan    = find_column(df, ['Delivery Plan Date', 'delivery_plan_date', 'plan date'])
        col_remarks = find_column(df, ['Remarks', 'remarks', 'remark'])

        if not col_so_num:
            return jsonify({'error': 'SO Number column not found'}), 400

        updated = 0
        for _, row in df.iterrows():
            so_num = clean(df_val(row, col_so_num))
            if not so_num:
                continue
            sos = SOData.query.filter_by(so_number=so_num).all()
            for s in sos:
                changed = False
                if col_plan:
                    v = clean(df_val(row, col_plan))
                    nd = parse_date(v) if v else None
                    if nd != s.delivery_plan_date:
                        s.delivery_plan_date = nd
                        changed = True
                if col_remarks:
                    v = clean(df_val(row, col_remarks))
                    if v != s.remarks:
                        s.remarks = v
                        changed = True
                if changed:
                    updated += 1

        db.session.commit()
        return jsonify({'message': f'Batch update complete: {updated} records updated', 'updated': updated})
    except Exception as e:
        db.session.rollback()
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


def _style_wb(ws, headers, num_cols=None):
    ws.append(headers)
    fill = PatternFill(start_color="7C3AED", end_color="7C3AED", fill_type="solid")
    for i, cell in enumerate(ws[1], 1):
        cell.fill = fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal='center')
        ws.column_dimensions[get_column_letter(i)].width = 20
    if num_cols:
        for c in num_cols:
            ws.column_dimensions[get_column_letter(c)].width = 18


@app.route('/api/export/all-so', methods=['GET'])
def export_all_so():
    try:
        op_units = request.args.getlist('op_unit')
        vendors  = request.args.getlist('vendor')
        statuses = request.args.getlist('status')
        q = SOData.query.filter(open_so_filter())
        if op_units: q = q.filter(SOData.operation_unit_name.in_(op_units))
        if vendors:  q = q.filter(SOData.vendor_name.in_(vendors))
        if statuses: q = q.filter(SOData.so_status.in_(statuses))
        sos = q.all()
        today = date.today()
        wb = Workbook(); ws = wb.active; ws.title = "SO List"
        _style_wb(ws, ['Aging (Biz Days)','SO Number','SO Item','Status','Op Unit','Vendor','Product',
                       'SO Qty','Sales Price','Sales Amount','PO Price','PO Amount',
                       'SO Date','Delivery Possible','Customer PO','Delivery Memo',
                       'Delivery Plan Date','Remarks'], num_cols=[8,9,10,11,12])
        for s in sos:
            biz = count_business_days(s.so_create_date, today) if s.so_create_date else None
            ws.append([get_aging_label(biz), s.so_number, s.so_item, s.so_status,
                s.operation_unit_name, s.vendor_name, s.product_name,
                s.so_qty or 0, s.sales_price or 0, s.sales_amount or 0,
                s.purchasing_price or 0, s.purchasing_amount or 0,
                s.so_create_date.isoformat() if s.so_create_date else '',
                s.delivery_possible_date.isoformat() if s.delivery_possible_date else '',
                s.customer_po_number or '', s.delivery_memo or '',
                s.delivery_plan_date.isoformat() if s.delivery_plan_date else '',
                s.remarks or ''])
        output = io.BytesIO(); wb.save(output); output.seek(0)
        return send_file(output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True, download_name=f"SO_List_{datetime.now().strftime('%Y%m%d')}.xlsx")
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/export/po-without-so', methods=['GET'])
def export_po_without_so():
    try:
        matched_set = build_matched_set()
        hidden_po = get_hidden_po_numbers()
        today = date.today()

        seen_keys = set()
        pos = []
        for p in POData.query.all():
            key = (p.po_number, p.item_no)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            if is_po_hidden(p.po_number, p.item_no, hidden_po):
                continue
            op_unit = get_operation_unit(p.po_item_type, p.item_code)
            if op_unit in EXCLUDED_OP_UNITS:
                continue
            if not po_is_matched(p.po_number, p.item_no, matched_set):
                pos.append((p, op_unit))

        wb = Workbook(); ws = wb.active; ws.title = "PO Without SO"
        _style_wb(ws, ['PO Number','PO Item Type','Item No','Item Code','Operation Unit','Description','Supplier',
                       'Qty','Unit','Price','Amount','Currency',
                       'PO Date','Purchase Member','Req. Delivery','Biz Days Remaining',
                       'Delivery Plan Date','Remarks'], num_cols=[8,10,11])
        for p, op_unit in pos:
            biz_rem = business_days_remaining(today, p.request_delivery) if p.request_delivery else ''
            ws.append([p.po_number, p.po_item_type or '', p.item_no or '', p.item_code or '', op_unit,
                p.po_item_detail, p.supplier,
                p.qty or 0, p.unit or '', p.price or 0, p.amount or 0, p.currency or 'IDR',
                p.po_date.isoformat() if p.po_date else '',
                p.purchase_member or '',
                p.request_delivery.isoformat() if p.request_delivery else '',
                biz_rem,
                p.delivery_plan_date.isoformat() if p.delivery_plan_date else '',
                p.remarks or ''])
        output = io.BytesIO(); wb.save(output); output.seek(0)
        return send_file(output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True, download_name=f"PO_Without_SO_{datetime.now().strftime('%Y%m%d')}.xlsx")
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/template/hide', methods=['GET'])
def download_hide_template():
    hide_type = request.args.get('type', 'PO').upper()
    wb = Workbook()
    ws = wb.active

    if hide_type == 'SO':
        ws.title = "Hide SO Template"
        headers = ['SO Number', 'Reason']
        ws.append(headers)
        fill = PatternFill(start_color="1D4ED8", end_color="1D4ED8", fill_type="solid")
        for i, cell in enumerate(ws[1], 1):
            cell.fill = fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(i)].width = 30 if i == 1 else 50
        ws.append(['9008988017-10', 'Reason why this SO should be hidden'])
        ws.append(['INSTRUCTIONS: Fill SO Number (format: SO_NUMBER-ITEM_NO or SO_NUMBER), and Reason (required)'])
    else:
        ws.title = "Hide PO HLI Template"
        headers = ['PO HLI Number (PO Number-Item No)', 'Reason']
        ws.append(headers)
        fill = PatternFill(start_color="7C3AED", end_color="7C3AED", fill_type="solid")
        for i, cell in enumerate(ws[1], 1):
            cell.fill = fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(i)].width = 35 if i == 1 else 50
        ws.append(['4502358819-10', 'Reason why this PO HLI should be hidden'])
        ws.append(['INSTRUCTIONS: Fill PO HLI Number with format PO_NUMBER-ITEM_NO (e.g. 4502358819-10), and Reason (required)'])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    fname = f"Template_Hide_{'SO' if hide_type == 'SO' else 'PO_HLI'}.xlsx"
    return send_file(output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=fname)


@app.route('/api/upload/hide-batch', methods=['POST'])
def upload_hide_batch():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        hide_type = request.form.get('type', 'PO').upper()

        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        if hide_type == 'SO':
            col_ref = find_column(df, ['SO Number', 'SO No', 'SO Item', 'SO Number-Item No'])
        else:
            col_ref = find_column(df, [
                'PO HLI Number (PO Number-Item No)', 'NO PO HLI (PO Number-Item No)', 'NO PO HLI', 'PO HLI',
                'PO Number-Item No', 'PO HLI Number', 'PO Number'
            ])

        col_reason = find_column(df, ['Reason', 'Alasan', 'Keterangan'])

        if not col_ref:
            return jsonify({'error': f'Reference number column not found. Available columns: {df.columns.tolist()}'}), 400
        if not col_reason:
            return jsonify({'error': f'Reason column not found. Available columns: {df.columns.tolist()}'}), 400

        success_count = 0
        skipped = []
        errors = []

        for idx, row in df.iterrows():
            ref_number = clean(df_val(row, col_ref))
            reason = clean(df_val(row, col_reason))

            if not ref_number or ref_number.upper().startswith('PETUNJUK') or ref_number.upper().startswith('INSTRUCTIONS'):
                continue
            if reason and (reason.lower().startswith('alasan kenapa') or reason.lower().startswith('reason why')):
                continue

            if not reason:
                errors.append(f"Row {idx+2}: Reason is empty for {ref_number}")
                continue

            existing = DeleteRequest.query.filter_by(
                ref_type=hide_type, ref_number=ref_number, is_hidden=True
            ).first()
            if existing:
                skipped.append(ref_number)
                continue

            req = DeleteRequest(
                ref_type=hide_type,
                ref_number=ref_number,
                reason=reason,
                is_hidden=True
            )
            db.session.add(req)
            success_count += 1

        db.session.commit()

        msg = f'{success_count} items hidden successfully'
        if skipped:
            msg += f'. {len(skipped)} already hidden: {", ".join(skipped[:5])}'
        if errors:
            msg += f'. {len(errors)} errors: {"; ".join(errors[:3])}'

        return jsonify({
            'message': msg,
            'hidden': success_count,
            'skipped': len(skipped),
            'errors': errors
        })
    except Exception as e:
        db.session.rollback()
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ═══════════════════════════════════════════════════════════════════
# DELETE REQUEST ENDPOINTS
# ═══════════════════════════════════════════════════════════════════

@app.route('/api/delete-requests', methods=['GET'])
def get_delete_requests():
    try:
        reqs = DeleteRequest.query.order_by(DeleteRequest.requested_at.desc()).all()
        return jsonify([{
            'id': r.id,
            'ref_type': r.ref_type,
            'ref_number': r.ref_number,
            'reason': r.reason,
            'requested_at': r.requested_at.isoformat(),
            'is_hidden': r.is_hidden,
        } for r in reqs])
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/delete-requests', methods=['POST'])
def create_delete_request():
    try:
        data = request.json
        ref_type = data.get('ref_type', '').upper()
        ref_number = (data.get('ref_number') or '').strip()
        reason = (data.get('reason') or '').strip()

        if ref_type not in ('PO', 'SO'):
            return jsonify({'error': 'ref_type must be PO or SO'}), 400
        if not ref_number:
            return jsonify({'error': 'Reference number is required'}), 400
        if not reason:
            return jsonify({'error': 'Reason is required'}), 400

        existing = DeleteRequest.query.filter_by(ref_type=ref_type, ref_number=ref_number, is_hidden=True).first()
        if existing:
            return jsonify({'error': f'{ref_type} {ref_number} is already hidden'}), 400

        req = DeleteRequest(
            ref_type=ref_type,
            ref_number=ref_number,
            reason=reason,
            is_hidden=True
        )
        db.session.add(req)
        db.session.commit()
        return jsonify({'success': True, 'id': req.id, 'message': f'{ref_type} {ref_number} hidden from dashboard'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500


@app.route('/api/delete-requests/<int:req_id>/restore', methods=['PUT'])
def restore_delete_request(req_id):
    try:
        req = db.session.get(DeleteRequest, req_id)
        if not req:
            return jsonify({'error': 'Request not found'}), 404
        req.is_hidden = False
        db.session.commit()
        return jsonify({'success': True, 'message': f'{req.ref_type} {req.ref_number} restored to dashboard'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("Backend: http://127.0.0.1:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
