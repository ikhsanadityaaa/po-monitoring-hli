from flask import Flask, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
import pandas as pd
import re
import os
from datetime import datetime, date
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

with app.app_context():
    db.create_all()
    print('DB schema ready.')

CLOSED_STATUSES = {
    'Delivery Completed', 'SO Cancel',
    'Approval Apply', 'Approval Complete Step', 'Approval Reject'
}

# Operation units yang dikecualikan dari matching PO HLI vs SO
EXCLUDED_OP_UNITS = {'HLI GREEN POWER (CONSUMABLE)'}

# Regex: optional short letter prefix, then 7+ digit PO number, optionally followed by -ItemNo
PO_HLI_RE = re.compile(r'(?:[A-Za-z]{1,4}[-]?)?(\d{7,})(?:-(\d+))?')

def _normalize_item_no(item_no):
    """
    Return a set of normalised variants of item_no to maximise matching.
    e.g. '010' -> {'010', '10'}, '10.0' -> {'10', '010'}, '10' -> {'10', '010'}
    """
    if item_no is None:
        return set()
    s = str(item_no).strip()
    variants = {s}
    # Remove trailing .0 (Excel stores integers as floats sometimes)
    if s.endswith('.0'):
        s = s[:-2]
        variants.add(s)
    # Integer form
    try:
        n = int(float(s))
        variants.add(str(n))        # plain  e.g. '10'
        variants.add(f"{n:02d}")    # 2-digit e.g. '10' (or '01' for n=1)
        variants.add(f"{n:03d}")    # 3-digit e.g. '010'
    except (ValueError, OverflowError):
        pass
    return variants

def extract_po_hli(val):
    """
    Extract all PO HLI reference keys from a free-text field.
    For '4570226161-10' emits: '4570226161', '4570226161-10', '4570226161-010'
    For '4570226161'    emits: '4570226161'
    """
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
    """Return SQLAlchemy filter clause for open (non-closed) SO records."""
    return db.or_(
        SOData.so_status.is_(None),
        SOData.so_status.notin_(list(CLOSED_STATUSES))
    )

def is_return_so_item(so_item):
    """
    Return True if this SO item should be excluded from PO HLI matching/counting.
    Rules:
      - SO Item that starts with '9' (return item), e.g. 9008988017-10
    """
    if not so_item:
        return False
    return str(so_item).strip().startswith('9')

def has_internal_po_ref(customer_po_number, delivery_memo):
    """
    Return True if the Customer PO Number or Delivery Memo contains a reference
    that starts with '2' (internal/non-HLI order), e.g. 2024154899-20 or 2024390384-10.
    These should NOT be counted as 'SO without PO HLI'.
    """
    for field in [customer_po_number, delivery_memo]:
        if not field:
            continue
        text = str(field).strip()
        # Check each token (space or comma separated)
        for token in re.split(r'[\s,;]+', text):
            token = token.strip()
            if token and token[0] == '2' and re.match(r'^2\d{6,}', token):
                return True
    return False

def so_is_countable(so_item, so_number=None, customer_po_number=None, delivery_memo=None):
    """Return True if this SO record should be counted in PO HLI without SO logic."""
    # Exclude return items (SO Item starts with 9)
    if is_return_so_item(so_item):
        return False
    # Exclude SO where Cust PO or Delivery Memo references start with '2'
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

def get_aging_label(days):
    if days is None: return 'No Date'
    if days >= 180: return '180+'
    if days >= 90: return '90-180'
    if days >= 30: return '30-90'
    return '0-30'

def so_dict(s):
    today = date.today()
    age_days = (today - s.so_create_date).days if s.so_create_date else None
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
        'aging_days': age_days,
        'aging_label': get_aging_label(age_days)
    }

@app.route('/api/dashboard/stats', methods=['GET'])
def get_dashboard_stats():
    try:
        po_count = db.session.query(func.count(POData.id)).scalar() or 0
        total_po_amount = db.session.query(func.sum(POData.amount)).scalar() or 0
        total_so_count = db.session.query(func.count(SOData.id)).filter(open_so_filter()).scalar() or 0

        po_numbers = {r[0] for r in db.session.query(POData.po_number).all() if r[0]}

        matched_set = build_matched_set()

        po_without_so_count = 0
        for p in POData.query.all():
            op_unit = get_operation_unit(p.po_item_type, p.item_code)
            if op_unit in EXCLUDED_OP_UNITS:
                continue
            if not po_is_matched(p.po_number, p.item_no, matched_set):
                po_without_so_count += 1

        so_without_po_count = 0
        for s in db.session.query(SOData).filter(
                open_so_filter(),
                ~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))).all():
            # Exclude return items (SO Item starts with 9) and internal refs (Cust PO/Memo starts with 2)
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            candidates = extract_po_hli(s.customer_po_number) + extract_po_hli(s.delivery_memo)
            if not candidates or not any(c in po_numbers for c in candidates):
                so_without_po_count += 1

        monthly = {}
        for d, amt in db.session.query(SOData.so_create_date, SOData.sales_amount)\
                                 .filter(open_so_filter()).all():
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

        # Top Operation Units
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

        # Date ranges
        po_date_range = db.session.query(func.min(POData.po_date), func.max(POData.po_date)).first()
        so_date_range = db.session.query(func.min(SOData.so_create_date), func.max(SOData.so_create_date)).first()

        return jsonify({
            'po_without_so': po_without_so_count,
            'so_without_po': so_without_po_count,
            'total_po_amount': total_po_amount,
            'total_open_so_amount': total_open_so_amount,
            'total_po_count': po_count,
            'total_so_count': total_so_count,
            'monthly_trend': monthly_trend,
            'top_vendors': top_vendors,
            'top_op_units': top_op_units,
            'so_status': so_status,
            'so_status_monthly': so_status_monthly,
            'status_months': sorted_months,
            'po_date_range': {
                'min': po_date_range[0].isoformat() if po_date_range[0] else None,
                'max': po_date_range[1].isoformat() if po_date_range[1] else None,
            },
            'so_date_range': {
                'min': so_date_range[0].isoformat() if so_date_range[0] else None,
                'max': so_date_range[1].isoformat() if so_date_range[1] else None,
            },
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

def build_matched_set():
    """
    Build set of PO references that already have a matching SO.
    Excludes:
      - SO rows from EXCLUDED_OP_UNITS
      - SO items starting with '9' (return items)
      - SO numbers starting with '2' (internal order type)
    The set contains both plain PO numbers AND combined "PO-ItemNo" strings
    so we can match at either level.
    """
    matched = set()
    for so_item, so_number, cpn, memo in db.session.query(
            SOData.so_item, SOData.so_number,
            SOData.customer_po_number, SOData.delivery_memo)\
            .filter(~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))).all():
        if not so_is_countable(so_item, customer_po_number=cpn, delivery_memo=memo):
            continue
        for n in extract_po_hli(cpn): matched.add(n)
        for n in extract_po_hli(memo): matched.add(n)
    return matched

def po_is_matched(po_number, item_no, matched_set):
    """
    Check if a PO line (po_number + item_no) is covered by any SO reference.
    Logic:
      - If matched_set contains the plain po_number  -> ALL items of that PO are matched
      - If matched_set contains 'po_number-item_no' (any normalised variant) -> this item matched
    """
    if not po_number:
        return False
    po_str = str(po_number).strip()

    # 1. Plain PO number match (SO references entire PO, all items covered)
    if po_str in matched_set:
        return True

    # 2. Combined PO-ItemNo match (SO references specific item)
    if item_no:
        for item_var in _normalize_item_no(item_no):
            if f"{po_str}-{item_var}" in matched_set:
                return True

    return False

def get_operation_unit(po_item_type, item_code):
    """
    Determine the Operation Unit based on PO Item Type and whether Item Code exists.
    Logic (from reference table):
    item_code ada:
      MRO      -> HLI GREEN POWER (CONSUMABLE)
      Equipment-> HLI GREEN POWER(BONDED AREA)
      ETC      -> HLI GREEN POWER(BONDED AREA)
    item_code tidak ada:
      MRO      -> HLI GREEN POWER(BONDED AREA)
      Equipment-> HLI GREEN POWER
      ETC      -> HLI GREEN POWER(BONDED AREA)
    """
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

@app.route('/api/data/po-without-so', methods=['GET'])
def get_po_without_so():
    try:
        matched_set = build_matched_set()
        today = date.today()
        result = []
        for p in POData.query.all():
            op_unit = get_operation_unit(p.po_item_type, p.item_code)
            # Exclude CONSUMABLE items from this list entirely
            if op_unit in EXCLUDED_OP_UNITS:
                continue
            if not po_is_matched(p.po_number, p.item_no, matched_set):
                days_remaining = (p.request_delivery - today).days if p.request_delivery else None
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
        po_numbers = {r[0] for r in db.session.query(POData.po_number).all() if r[0]}
        result = []
        for s in db.session.query(SOData).filter(
                open_so_filter(),
                ~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))).all():
            # Exclude return items (SO item starts with 9) and internal refs in Cust PO/Memo
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            candidates = extract_po_hli(s.customer_po_number) + extract_po_hli(s.delivery_memo)
            if not candidates or not any(c in po_numbers for c in candidates):
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
        for s in db.session.query(SOData).filter(open_so_filter(), SOData.so_create_date.isnot(None)).all():
            # Exclude return items (SO item starts with 9) and internal refs in Cust PO/Memo
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            v = s.vendor_name or 'Unknown'
            if v not in vendors:
                vendors[v] = {'vendor': v, 'less_30': 0, 'days_30_90': 0,
                              'days_90_180': 0, 'more_180': 0, 'total_open': 0, 'sales_amount': 0.0}
            age = (today - s.so_create_date).days
            if age < 30:      vendors[v]['less_30']     += 1
            elif age < 90:    vendors[v]['days_30_90']  += 1
            elif age < 180:   vendors[v]['days_90_180'] += 1
            else:             vendors[v]['more_180']    += 1
            vendors[v]['total_open'] += 1
            vendors[v]['sales_amount'] += float(s.sales_amount or 0)
        return jsonify(sorted(vendors.values(), key=lambda x: x['total_open'], reverse=True))
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/data/aging-detail/<path:vendor_name>', methods=['GET'])
def get_aging_detail(vendor_name):
    try:
        bucket = request.args.get('bucket')  # '0-30', '30-90', '90-180', '180+', or None for all
        today = date.today()
        sos = db.session.query(SOData).filter(
            open_so_filter(), SOData.vendor_name == vendor_name
        ).order_by(SOData.so_create_date.asc()).all()
        # Exclude return items and internal refs in Cust PO/Memo
        sos = [s for s in sos if so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo)]
        if bucket:
            bucket = bucket.strip().replace(' ', '+')  # fix URL encoding: + decoded as space
            sos = [s for s in sos if get_aging_label((today - s.so_create_date).days if s.so_create_date else None) == bucket]
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
            open_so_filter(), SOData.so_create_date.isnot(None)
        ).order_by(SOData.vendor_name.asc(), SOData.so_create_date.asc()).all()
        # Exclude return items and internal refs in Cust PO/Memo
        sos = [s for s in sos if so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo)]
        if bucket:
            sos = [s for s in sos if get_aging_label((today - s.so_create_date).days if s.so_create_date else None) == bucket]
        return jsonify([so_dict(s) for s in sos])
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/data/all-so', methods=['GET'])
def get_all_so():
    try:
        op_units  = request.args.getlist('op_unit')   # multi-value
        vendors   = request.args.getlist('vendor')    # multi-value
        statuses  = request.args.getlist('status')    # multi-value
        aging     = request.args.getlist('aging')
        so_items  = request.args.getlist('so_item')   # search by SO item numbers
        page      = max(1, int(request.args.get('page', 1)))
        per_page  = min(500, int(request.args.get('per_page', 20)))

        today = date.today()
        q = SOData.query
        if op_units:  q = q.filter(SOData.operation_unit_name.in_(op_units))
        if vendors:   q = q.filter(SOData.vendor_name.in_(vendors))
        if statuses:  q = q.filter(SOData.so_status.in_(statuses))
        if so_items:  q = q.filter(SOData.so_item.in_(so_items))

        all_sos = q.order_by(SOData.so_create_date.desc()).all()

        if aging:
            def matches_aging(s):
                age = (today - s.so_create_date).days if s.so_create_date else None
                return get_aging_label(age) in aging
            all_sos = [s for s in all_sos if matches_aging(s)]

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
        month = request.args.get('month')  # e.g. 'Jan 2024', optional
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
    """SO records filtered by optional month — for TOTAL row clicks in Status Distribution."""
    try:
        month = request.args.get('month')
        if month:
            sos = [s for s in SOData.query.all()
                   if s.so_create_date and s.so_create_date.strftime('%b %Y') == month]
        else:
            sos = SOData.query.order_by(SOData.so_create_date.desc()).all()
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

CHUNK_SIZE = 200

@app.route('/api/upload/po-list', methods=['POST'])
def upload_po_list():
    try:
        if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        # ── Header validation ──────────────────────────────────────────────
        # These are the key columns that MUST exist (using any of their aliases)
        REQUIRED_PO_COLS = {
            'PO Number':        ['PO No.','PO No','PO Number','PO'],
            'Item No':          ['Item No.','Item No','Item Number','No. Item'],
            'PO Item Type':     ['PO Item Type','Item Type','Type','PO Type'],
            'Supplier':         ['Supplier','Vendor','Supplier Name'],
            'Qty':              ['Qty.','Qty','Quantity'],
            'Amount':           ['Amount','Total Amount','Total'],
            'PO Date':          ['PO Date','Order Date','Tanggal PO'],
            'Request Delivery': ['Request Delivery Date','Delivery Date','Req Delivery'],
        }
        missing_required = []
        for friendly_name, aliases in REQUIRED_PO_COLS.items():
            if not find_column(df, aliases):
                missing_required.append(friendly_name)
        if len(missing_required) >= 3:
            return jsonify({
                'error': (
                    f'❌ File tidak valid — {len(missing_required)} kolom penting tidak ditemukan: '
                    f'{", ".join(missing_required)}. '
                    f'Pastikan Anda mengupload file HLI PO List yang benar, lalu cek kembali.'
                )
            }), 400
        # ── End header validation ──────────────────────────────────────────

        col_po   = find_column(df, ['PO No.','PO No','PO Number','PO'])
        if not col_po:
            return jsonify({'error': f'Kolom PO Number tidak ditemukan. Kolom: {df.columns.tolist()}'}), 400

        col_itemno = find_column(df, ['Item No.','Item No','Item Number','No. Item'])
        col_desc = find_column(df, ['PO Item Detail','Description','Item Description','Deskripsi'])
        col_item = find_column(df, ['Item Code','Material','Item No','Item'])
        col_itype = find_column(df, ['PO Item Type','Item Type','Type','PO Type'])
        col_supp = find_column(df, ['Supplier','Vendor','Supplier Name'])
        col_vndr = find_column(df, ['Vendor Name SMRO','Vendor Name'])
        col_qty  = find_column(df, ['Qty.','Qty','Quantity'])
        col_unit = find_column(df, ['Unit','UOM'])
        col_price= find_column(df, ['Price','Unit Price'])
        col_amt  = find_column(df, ['Amount','Total Amount','Total'])
        col_cur  = find_column(df, ['Currency','Curr'])
        col_pdt  = find_column(df, ['PO Date','Order Date','Tanggal PO'])
        col_pm   = find_column(df, ['Purchase Member','Purchasing Member','PIC','Buyer'])
        col_rdd  = find_column(df, ['Request Delivery Date','Delivery Date','Req Delivery'])

        # Build lookup of existing PO records (po_number + item_no) -> existing record
        existing_po = {}
        for p in POData.query.all():
            key = (p.po_number, p.item_no)
            existing_po[key] = p

        new_keys_in_file = set()
        count = 0
        for _, row in df.iterrows():
            po_num = clean(df_val(row, col_po))
            if not po_num: continue
            item_no = clean(df_val(row, col_itemno))
            key = (po_num, item_no)
            new_keys_in_file.add(key)

            new_data = {
                'po_number': po_num,
                'item_no': item_no,
                'po_item_detail': clean(df_val(row, col_desc)),
                'item_code': clean(df_val(row, col_item)),
                'po_item_type': clean(df_val(row, col_itype)),
                'supplier': clean(df_val(row, col_supp)),
                'vendor_name_smro': clean(df_val(row, col_vndr)),
                'qty': safe_float(df_val(row, col_qty)),
                'unit': clean(df_val(row, col_unit)),
                'price': safe_float(df_val(row, col_price)),
                'amount': safe_float(df_val(row, col_amt)),
                'currency': clean(df_val(row, col_cur)) or 'IDR',
                'po_date': parse_date(df_val(row, col_pdt)),
                'purchase_member': clean(df_val(row, col_pm)),
                'request_delivery': parse_date(df_val(row, col_rdd)),
                'uploaded_at': datetime.utcnow()
            }

            if key in existing_po:
                # Update existing, preserve remarks & delivery_plan_date
                existing = existing_po[key]
                for field, val in new_data.items():
                    setattr(existing, field, val)
                # remarks & delivery_plan_date already preserved (not in new_data update)
            else:
                new_rec = POData(**new_data)
                db.session.add(new_rec)

            count += 1
            if count % CHUNK_SIZE == 0:
                db.session.flush()

        # Delete records not in the new file
        keys_to_delete = set(existing_po.keys()) - new_keys_in_file
        if keys_to_delete:
            for key in keys_to_delete:
                db.session.delete(existing_po[key])

        db.session.add(UploadLog(file_type='PO', filename=file.filename, records_count=count))
        db.session.commit()
        return jsonify({'message': f'Berhasil upload {count} PO items', 'uploaded': count})
    except Exception as e:
        db.session.rollback(); import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload/smro', methods=['POST'])
def upload_smro():
    try:
        if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        # ── Header validation ──────────────────────────────────────────────
        REQUIRED_SMRO_COLS = {
            'SO Number':       ['SO Number','SO No','SO No.','SO','Sales Order','Sales Order Number','No SO','Nomor SO'],
            'SO Item':         ['SO Item No','Item No','Line','SO Line','SO Item'],
            'SO Status':       ['SO Status','Status','Order Status'],
            'Operation Unit':  ['Operation Unit Name','Op Unit','Client Name','Client','Operation Unit'],
            'Vendor Name':     ['Vendor Name','Vendor','Supplier'],
            'Customer PO':     ['Customer PO Number','Customer PO','PO Ref','PO Reference'],
            'Sales Amount':    ['Sales Amount(Exclude Tax)','Sales Amount','Amount','Total'],
            'SO Create Date':  ['SO Create Date','Order Date','SO Date','Create Date'],
        }
        missing_required = []
        for friendly_name, aliases in REQUIRED_SMRO_COLS.items():
            if not find_column(df, aliases):
                missing_required.append(friendly_name)
        if len(missing_required) >= 3:
            return jsonify({
                'error': (
                    f'❌ File tidak valid — {len(missing_required)} kolom penting tidak ditemukan: '
                    f'{", ".join(missing_required)}. '
                    f'Pastikan Anda mengupload file SMRO yang benar, lalu cek kembali.'
                )
            }), 400
        # ── End header validation ──────────────────────────────────────────

        col_so = find_column(df, ['SO Number','SO No','SO No.','SO','SO Item',
                                   'Sales Order','Sales Order Number','No SO','Nomor SO'])
        if not col_so:
            return jsonify({'error': f'Kolom SO Number tidak ditemukan. Kolom: {df.columns.tolist()}'}), 400

        col_soitem  = find_column(df, ['SO Item No','Item No','Line','SO Line','SO Item'])
        col_status  = find_column(df, ['SO Status','Status','Order Status'])
        col_opunit  = find_column(df, ['Operation Unit Name','Op Unit','Client Name','Client','Operation Unit'])
        col_vendor  = find_column(df, ['Vendor Name','Vendor','Supplier'])
        col_custpo  = find_column(df, ['Customer PO Number','Customer PO','PO Ref','PO Reference'])
        col_memo    = find_column(df, ['Delivery Memo','Memo','Delivery Note'])
        col_prod    = find_column(df, ['Product Name','Item Name','Description','Product'])
        col_qty     = find_column(df, ['SO Quantity','SO Qty','Qty','Quantity'])
        col_sunit   = find_column(df, ['Sales Unit','Unit','UOM'])
        col_sprice  = find_column(df, ['Sales Price(Exclude Tax)','Sales Price','Price','Unit Price'])
        col_samt    = find_column(df, ['Sales Amount(Exclude Tax)','Sales Amount','Amount','Total'])
        col_cur     = find_column(df, ['Currency','Curr'])
        col_pprice  = find_column(df, ['Purchasing Price','Purchase Price','PO Price'])
        col_pamt    = find_column(df, ['Purchasing Amount','Purchase Amount','PO Amount'])
        col_sodate  = find_column(df, ['SO Create Date','Order Date','SO Date','Create Date'])
        col_delposs = find_column(df, ['Delivery Possible Date','Possible Delivery Date','Est Delivery'])
        col_matchpo = find_column(df, ['Matched PO Number','Matched PO','PO HLI','PO HLI Number'])

        # Build lookup of existing SO records by so_item (unique key)
        existing_so = {}
        for s in SOData.query.all():
            if s.so_item:
                existing_so[s.so_item] = s

        new_soitems_in_file = set()
        count = 0
        for _, row in df.iterrows():
            so_val = clean(df_val(row, col_so))
            if not so_val: continue
            so_item_val = clean(df_val(row, col_soitem))
            new_soitems_in_file.add(so_item_val)

            new_data = {
                'so_number': so_val,
                'so_item': so_item_val,
                'so_status': clean(df_val(row, col_status)),
                'operation_unit_name': clean(df_val(row, col_opunit)),
                'vendor_name': clean(df_val(row, col_vendor)),
                'customer_po_number': clean(df_val(row, col_custpo)),
                'delivery_memo': clean(df_val(row, col_memo)),
                'product_name': clean(df_val(row, col_prod)),
                'so_qty': safe_float(df_val(row, col_qty)),
                'sales_unit': clean(df_val(row, col_sunit)),
                'sales_price': safe_float(df_val(row, col_sprice)),
                'sales_amount': safe_float(df_val(row, col_samt)),
                'currency': clean(df_val(row, col_cur)) or 'IDR',
                'purchasing_price': safe_float(df_val(row, col_pprice)),
                'purchasing_amount': safe_float(df_val(row, col_pamt)),
                'so_create_date': parse_date(df_val(row, col_sodate)),
                'delivery_possible_date': parse_date(df_val(row, col_delposs)),
                'matched_po_number': clean(df_val(row, col_matchpo)),
                'uploaded_at': datetime.utcnow()
            }

            if so_item_val and so_item_val in existing_so:
                # Update existing record, preserve remarks & delivery_plan_date
                existing = existing_so[so_item_val]
                preserved_remarks = existing.remarks
                preserved_plan_date = existing.delivery_plan_date
                for field, val in new_data.items():
                    setattr(existing, field, val)
                # Restore preserved fields
                existing.remarks = preserved_remarks
                existing.delivery_plan_date = preserved_plan_date
            else:
                new_rec = SOData(**new_data)
                db.session.add(new_rec)

            count += 1
            if count % CHUNK_SIZE == 0:
                db.session.flush()

        # Delete records not in the new file (by so_item)
        keys_to_delete = set(existing_so.keys()) - new_soitems_in_file
        if keys_to_delete:
            for key in keys_to_delete:
                db.session.delete(existing_so[key])

        db.session.add(UploadLog(file_type='SO', filename=file.filename, records_count=count))
        db.session.commit()
        return jsonify({'message': f'Berhasil upload {count} SO items', 'uploaded': count})
    except Exception as e:
        db.session.rollback(); import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/data/so/<int:so_id>', methods=['PUT'])
def update_so(so_id):
    try:
        data = request.json
        so = db.session.get(SOData, so_id)
        if not so: return jsonify({'error': 'Not found'}), 404
        if 'delivery_plan_date' in data: so.delivery_plan_date = parse_date(data['delivery_plan_date'])
        if 'remarks' in data: so.remarks = data['remarks']
        db.session.commit()
        return jsonify({'success': True})
    except Exception as e:
        db.session.rollback(); return jsonify({'error': str(e)}), 500

@app.route('/api/data/so/batch-upload', methods=['POST'])
def batch_upload_so():
    try:
        if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]
        col_so   = find_column(df, ['SO Number','SO No','SO Item'])
        col_plan = find_column(df, ['Delivery Plan Date','Plan Date'])
        col_rem  = find_column(df, ['Remarks','Remark'])
        updated = 0
        for _, row in df.iterrows():
            so_num = clean(df_val(row, col_so)) if col_so else None
            if not so_num: continue
            so = SOData.query.filter_by(so_number=so_num).first()
            if so:
                if col_plan:
                    d = parse_date(df_val(row, col_plan))
                    if d: so.delivery_plan_date = d
                if col_rem:
                    r = clean(df_val(row, col_rem))
                    if r: so.remarks = r
                updated += 1
        db.session.commit()
        return jsonify({'updated': updated})
    except Exception as e:
        db.session.rollback(); import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

def _style_wb(ws, headers, num_cols=None):
    ws.append(headers)
    fill = PatternFill(start_color="8B5CF6", end_color="8B5CF6", fill_type="solid")
    for i, cell in enumerate(ws[1], 1):
        cell.fill = fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal='center')
        ws.column_dimensions[get_column_letter(i)].width = 20
    if num_cols:
        for row in ws.iter_rows(min_row=2):
            for ci in num_cols:
                row[ci-1].number_format = '#,##0.00'

@app.route('/api/export/all-so', methods=['GET'])
def export_all_so():
    try:
        q = SOData.query
        op_units = request.args.getlist('op_unit')
        vendors  = request.args.getlist('vendor')
        statuses = request.args.getlist('status')
        if op_units: q = q.filter(SOData.operation_unit_name.in_(op_units))
        if vendors:  q = q.filter(SOData.vendor_name.in_(vendors))
        if statuses: q = q.filter(SOData.so_status.in_(statuses))
        sos = q.all()
        today = date.today()
        wb = Workbook(); ws = wb.active; ws.title = "SO List"
        _style_wb(ws, ['Aging','SO Number','SO Item','Status','Op Unit','Vendor','Product',
                       'SO Qty','Sales Price','Sales Amount','PO Price','PO Amount',
                       'SO Date','Delivery Possible','Customer PO','Delivery Memo',
                       'Delivery Plan Date','Remarks'], num_cols=[8,9,10,11,12])
        for s in sos:
            age = (today - s.so_create_date).days if s.so_create_date else None
            ws.append([get_aging_label(age), s.so_number, s.so_item, s.so_status,
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
        today = date.today()
        pos = []
        for p in POData.query.all():
            op_unit = get_operation_unit(p.po_item_type, p.item_code)
            if op_unit in EXCLUDED_OP_UNITS:
                continue
            if not po_is_matched(p.po_number, p.item_no, matched_set):
                pos.append((p, op_unit))
        wb = Workbook(); ws = wb.active; ws.title = "PO Without SO"
        _style_wb(ws, ['PO Number','PO Item Type','Item No','Item Code','Operation Unit','Description','Supplier',
                       'Qty','Unit','Price','Amount','Currency',
                       'PO Date','Purchase Member','Request Delivery','Days Remaining',
                       'Delivery Plan Date','Remarks'], num_cols=[8,10,11])
        for p, op_unit in pos:
            days_rem = (p.request_delivery - today).days if p.request_delivery else ''
            ws.append([p.po_number, p.po_item_type or '', p.item_no or '', p.item_code or '', op_unit,
                p.po_item_detail, p.supplier,
                p.qty or 0, p.unit or '', p.price or 0, p.amount or 0, p.currency or 'IDR',
                p.po_date.isoformat() if p.po_date else '',
                p.purchase_member or '',
                p.request_delivery.isoformat() if p.request_delivery else '',
                days_rem,
                p.delivery_plan_date.isoformat() if p.delivery_plan_date else '',
                p.remarks or ''])
        output = io.BytesIO(); wb.save(output); output.seek(0)
        return send_file(output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True, download_name=f"PO_Without_SO_{datetime.now().strftime('%Y%m%d')}.xlsx")
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print("Backend: http://127.0.0.1:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
