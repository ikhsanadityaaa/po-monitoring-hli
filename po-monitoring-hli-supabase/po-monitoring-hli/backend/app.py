from flask import Flask, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
import pandas as pd
import re
import os
import json
from datetime import datetime, date, timedelta
import io
from sqlalchemy import func, text
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ─── Indonesian Public Holidays (auto-generated, year-flexible) ───────────
# We use the `holidays` package to generate Indonesian national holidays for
# any year automatically — no need to hand-maintain a list when the year
# rolls over.  Government-announced "cuti bersama" / replacement days that
# the package doesn't know about live in `holiday_extras.json` next to this
# file; that file is a plain JSON array of "YYYY-MM-DD" strings the user can
# edit when SKB tahunan is published.
_HOLIDAY_CACHE = None
_HOLIDAY_CACHE_KEY = None

def _holiday_set():
    """Return cached set of Indonesian non-working public holidays.

    The cache covers a sliding ±10-year window around today, so once a new
    year arrives the holidays are picked up automatically without code
    changes."""
    global _HOLIDAY_CACHE, _HOLIDAY_CACHE_KEY
    today_year = date.today().year
    cache_key = today_year
    if _HOLIDAY_CACHE is not None and _HOLIDAY_CACHE_KEY == cache_key:
        return _HOLIDAY_CACHE

    years = list(range(today_year - 5, today_year + 11))
    try:
        import holidays as _holidays_pkg
        s = set(_holidays_pkg.country_holidays('ID', years=years).keys())
    except Exception:
        # If the package fails to import (e.g. dependency not installed),
        # fall back to weekends-only — the extras JSON still applies below.
        s = set()

    extras_path = os.path.join(os.path.dirname(__file__), 'holiday_extras.json')
    try:
        with open(extras_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # Accept either {"dates": [...]} or a bare list for forward-compat.
        items = data.get('dates', []) if isinstance(data, dict) else data
        for ds in items or []:
            try:
                s.add(date.fromisoformat(str(ds).strip()))
            except (ValueError, TypeError):
                pass
    except FileNotFoundError:
        pass
    except (OSError, json.JSONDecodeError):
        pass

    _HOLIDAY_CACHE = s
    _HOLIDAY_CACHE_KEY = cache_key
    return s

def is_workday(d):
    """Return True if date is a working day (Mon–Fri, not a public holiday)."""
    return d.weekday() < 5 and d not in _holiday_set()

def count_workdays(start, end):
    """Count working days between start and end (exclusive of end).
    Returns negative if end < start (overdue).
    """
    if start is None or end is None:
        return None
    if start == end:
        return 0
    if end > start:
        count = 0
        cur = start
        while cur < end:
            if is_workday(cur):
                count += 1
            cur += timedelta(days=1)
        return count
    else:
        # overdue — count negatively
        count = 0
        cur = end
        while cur < start:
            if is_workday(cur):
                count += 1
            cur += timedelta(days=1)
        return -count

def workdays_since(past_date, today=None):
    """Count working days from past_date to today (aging)."""
    if past_date is None:
        return None
    if today is None:
        today = date.today()
    return count_workdays(past_date, today)

def workdays_until(future_date, today=None):
    """Count working days from today to future_date (days remaining)."""
    if future_date is None:
        return None
    if today is None:
        today = date.today()
    return count_workdays(today, future_date)


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
    specification = db.Column(db.Text)
    product_id = db.Column(db.String(100))
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

# ─── NEW: Delete Request model ─────────────────────────────────────────────
class DeleteRequest(db.Model):
    __tablename__ = 'delete_request'
    id = db.Column(db.Integer, primary_key=True)
    ref_type = db.Column(db.String(10))       # 'PO' or 'SO'
    ref_number = db.Column(db.String(100))    # PO HLI number or SO number/item
    reason = db.Column(db.Text)
    requested_at = db.Column(db.DateTime, default=datetime.utcnow)
    is_hidden = db.Column(db.Boolean, default=True)  # True = hidden from dashboard

def _ensure_so_extra_columns():
    """Best-effort online migration: add `specification` and `product_id`
    columns to so_data if they don't exist yet.  `db.create_all()` only
    creates tables that don't exist; it never adds columns to an existing
    table.  Both SQLite and PostgreSQL support `ALTER TABLE ... ADD COLUMN
    IF NOT EXISTS` (SQLite >= 3.35), so a single statement works on both."""
    statements = [
        "ALTER TABLE so_data ADD COLUMN IF NOT EXISTS specification TEXT",
        "ALTER TABLE so_data ADD COLUMN IF NOT EXISTS product_id VARCHAR(100)",
    ]
    for sql in statements:
        try:
            db.session.execute(text(sql))
            db.session.commit()
        except Exception:
            db.session.rollback()
            # Older SQLite (<3.35) lacks IF NOT EXISTS — fall back and ignore
            # the duplicate-column error from a plain ADD COLUMN.
            try:
                db.session.execute(text(sql.replace(' IF NOT EXISTS', '')))
                db.session.commit()
            except Exception:
                db.session.rollback()


with app.app_context():
    db.create_all()
    _ensure_so_extra_columns()
    print('DB schema ready.')

CLOSED_STATUSES = {
    'Delivery Completed', 'SO Cancel',
    'Approval Apply', 'Approval Complete Step', 'Approval Reject'
}

EXCLUDED_OP_UNITS = {'HLI GREEN POWER (CONSUMABLE)'}

# ─── PO HLI extraction patterns ──────────────────────────────────────────
# Full PO HLI: 7+ digit number, optionally followed by `-<item line>`.  We
# only treat `-` as the item separator (the original convention) so that
# two adjacent PO numbers separated by space / `/` / `_` / `.` parse as
# *two* separate POs instead of being merged.  The trailing `(?!\d)`
# lookahead protects against truncating into a longer adjacent number
# (e.g. "4502342011-10245" still yields PO=4502342011 with no item, not a
# bogus item="1024").  Examples we capture from free-text fields:
#   "4502342011-10"          → 4502342011 + item 10
#   "4502342011"             → 4502342011 (bare)
#   "Po No 4502202743_..."   → 4502202743
#   "4502342011 4502342012"  → both numbers separately
PO_HLI_RE = re.compile(r'(\d{7,})(?:-(\d{1,4}))?(?!\d)')
# Short reference like "PO 626", "P.O #626", "PO-626", "po:626" — used to
# match against the *suffix* of full PO HLI keys in the PO table.
PO_SHORT_REF_RE = re.compile(
    r'\bP\s*\.?\s*O\s*\.?\s*[#:.\-]?\s*(\d{2,6})\b',
    re.IGNORECASE,
)

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
    """Return all candidate PO HLI keys (full PO and PO-item) found in `val`."""
    if not val:
        return []
    text = str(val).strip()
    result = set()
    for m in PO_HLI_RE.finditer(text):
        po_num  = m.group(1)
        item_no = m.group(2)
        # Skip leading-2 numbers — those are non-HLI internal PO refs the user
        # explicitly wants ignored (e.g. "2123456789").  Real HLI POs start
        # with 4/5/6 etc.  This also avoids accidentally matching dates that
        # happen to be 8+ digits (e.g. "20240105") when written without
        # separators.
        if po_num.startswith('2'):
            continue
        result.add(po_num)
        if item_no:
            for item_var in _normalize_item_no(item_no):
                result.add(f"{po_num}-{item_var}")
    return list(result)


def extract_po_short_refs(val):
    """Return short numeric references like '626' parsed from 'PO 626'.

    Used as a fallback to suffix-match against full PO HLI numbers in the PO
    table when the customer wrote the PO in shorthand form."""
    if not val:
        return []
    text = str(val).strip()
    refs = set()
    for m in PO_SHORT_REF_RE.finditer(text):
        n = m.group(1)
        # Avoid double-counting full POs that already came out of extract_po_hli.
        if len(n) >= 7:
            continue
        refs.add(n)
    return list(refs)


def so_has_matching_po_hli(s, po_hli_keys, po_suffix_index=None):
    """Return True if SO `s` references at least one PO HLI present in
    `po_hli_keys`.  Considers both the full-number extraction and the
    short-reference (suffix) fallback."""
    candidates = extract_po_hli(s.customer_po_number) + extract_po_hli(s.delivery_memo)
    if any(c in po_hli_keys for c in candidates):
        return True
    short_refs = extract_po_short_refs(s.customer_po_number) + extract_po_short_refs(s.delivery_memo)
    if not short_refs:
        return False
    if po_suffix_index is not None:
        return any(r in po_suffix_index for r in short_refs)
    # Fallback: linear suffix scan when caller didn't precompute the index.
    # Skip composite "po-item" keys so short refs like "10" don't falsely
    # match the trailing item-line of every PO HLI.
    return any(
        any(k.endswith(r) for k in po_hli_keys if '-' not in k)
        for r in short_refs
    )


def build_po_suffix_index(po_hli_keys):
    """Map every meaningful trailing-N-digit suffix (length 2..6) of every
    full PO HLI number in `po_hli_keys` to True, for O(1) suffix lookups."""
    idx = set()
    for k in po_hli_keys:
        # Only index pure-digit full PO numbers (skip composite "po-item" keys).
        if not k or '-' in k or not k.isdigit():
            continue
        for n in range(2, min(len(k), 7)):
            idx.add(k[-n:])
    return idx

def open_so_filter():
    return db.or_(
        SOData.so_status.is_(None),
        SOData.so_status.notin_(list(CLOSED_STATUSES))
    )


def parse_so_date_args(args=None):
    """Read date_year / date_from / date_to (with optional legacy `year`)
    from a request args object and normalize them.
    Returns (date_year_str, date_from_str, date_to_str)."""
    args = args if args is not None else request.args
    date_year = args.get('date_year', '')
    date_from = args.get('date_from', '')
    date_to   = args.get('date_to', '')
    if not date_year:
        legacy = args.get('year', '')
        if legacy and legacy != 'all':
            date_year = legacy
    return date_year, date_from, date_to


def apply_so_create_date_filter(query, date_year='', date_from='', date_to='', is_sqlite=None):
    """Apply SO Create Date filter to any query that references SOData."""
    if is_sqlite is None:
        is_sqlite = 'sqlite' in app.config['SQLALCHEMY_DATABASE_URI']
    if date_year:
        try:
            yr = int(date_year)
            if is_sqlite:
                return query.filter(func.strftime('%Y', SOData.so_create_date) == str(yr))
            return query.filter(func.extract('year', SOData.so_create_date) == yr)
        except (ValueError, TypeError):
            return query
    if date_from:
        query = query.filter(SOData.so_create_date >= date_from)
    if date_to:
        query = query.filter(SOData.so_create_date <= date_to)
    return query


def utc_isoformat(dt):
    """Serialize a (naive UTC) datetime as an ISO-8601 string with a trailing
    'Z' so JS Date() parses it as UTC and the browser converts to local time.
    Datetimes that already carry a timezone designator (Z, +HH:MM, or -HH:MM
    after the time portion) are returned unchanged."""
    if dt is None:
        return None
    s = dt.isoformat()
    tail = s[10:]  # everything after the date portion (skip the leading YYYY-MM-DD)
    if s.endswith('Z') or '+' in tail or '-' in tail:
        return s
    return s + 'Z'

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

def get_aging_label(workday_count):
    """Classify aging bucket based on working days. None (no date) → '180+' bucket."""
    if workday_count is None: return '180+'
    if workday_count >= 180: return '180+'
    if workday_count >= 90:  return '90-180'
    if workday_count >= 30:  return '30-90'
    return '0-30'

def so_dict(s):
    today = date.today()
    age_days = workdays_since(s.so_create_date, today)
    return {
        'id': s.id, 'so_number': s.so_number, 'so_item': s.so_item,
        'so_status': s.so_status, 'operation_unit_name': s.operation_unit_name,
        'vendor_name': s.vendor_name, 'customer_po_number': s.customer_po_number,
        'delivery_memo': s.delivery_memo, 'product_name': s.product_name,
        'specification': s.specification, 'product_id': s.product_id,
        'so_qty': s.so_qty, 'sales_price': s.sales_price, 'sales_amount': s.sales_amount,
        'purchasing_price': s.purchasing_price, 'purchasing_amount': s.purchasing_amount,
        'so_create_date': s.so_create_date.isoformat() if s.so_create_date else '',
        'delivery_possible_date': s.delivery_possible_date.isoformat() if s.delivery_possible_date else '',
        'delivery_plan_date': s.delivery_plan_date.isoformat() if s.delivery_plan_date else '',
        'remarks': s.remarks or '',
        'aging_days': age_days,
        'aging_label': get_aging_label(age_days)
    }

# ─── Build hidden set from delete requests ────────────────────────────────
def get_hidden_po_hli_keys():
    """Return set of hidden PO HLI keys. Format stored: 'po_number-item_no' or just 'po_number'."""
    reqs = DeleteRequest.query.filter_by(ref_type='PO', is_hidden=True).all()
    return {r.ref_number for r in reqs}

def get_hidden_so_items():
    """Return set of SO items/numbers that are hidden from dashboard."""
    reqs = DeleteRequest.query.filter_by(ref_type='SO', is_hidden=True).all()
    return {r.ref_number for r in reqs}

# Keep alias for backward compat in export
def get_hidden_po_numbers():
    return get_hidden_po_hli_keys()

def po_hli_key(po_number, item_no):
    """Generate PO HLI key: po_number-item_no (normalized item_no)."""
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
    """Check if this PO item is hidden. Matches by combined key or by po_number alone."""
    key = po_hli_key(po_number, item_no)
    if key and key in hidden_keys:
        return True
    # Also check all normalized variants of item_no
    if item_no:
        for var in _normalize_item_no(item_no):
            if f"{po_number}-{var}" in hidden_keys:
                return True
    # Check if po_number alone is hidden (legacy)
    if po_number in hidden_keys:
        return True
    return False


@app.route('/api/dashboard/stats', methods=['GET'])
def get_dashboard_stats():
    try:
        hidden_po = get_hidden_po_numbers()
        hidden_so = get_hidden_so_items()

        # SO Create Date filter (applied to every SO-based aggregate below).
        # PO-based metrics (total_po_amount, po_without_so_count) and the data
        # range / last_updated metadata are intentionally NOT filtered.
        date_year, date_from, date_to = parse_so_date_args()

        def so_q(*extra_filters):
            q = db.session.query(SOData).filter(*extra_filters) if extra_filters else db.session.query(SOData)
            return apply_so_create_date_filter(q, date_year, date_from, date_to)

        po_count = db.session.query(func.count(POData.id)).scalar() or 0
        total_po_amount = db.session.query(func.sum(POData.amount)).scalar() or 0

        # total_so_count: count only "countable" open SOs (same logic as aging)
        # excludes: return items (SO item starting with '9'), internal PO refs, excluded op units, hidden items
        total_so_count = 0
        for s in so_q(open_so_filter(),
                      ~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))).all():
            if s.so_item in hidden_so or s.so_number in hidden_so:
                continue
            if so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
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

        po_suffix_index = build_po_suffix_index(po_numbers)
        so_without_po_count = 0
        for s in so_q(open_so_filter(),
                      ~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))).all():
            if s.so_item in hidden_so or s.so_number in hidden_so:
                continue
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            if not so_has_matching_po_hli(s, po_numbers, po_suffix_index):
                so_without_po_count += 1

        # Monthly trend (open SO, sum of sales by month)
        monthly_q = apply_so_create_date_filter(
            db.session.query(SOData.so_create_date, SOData.sales_amount).filter(open_so_filter()),
            date_year, date_from, date_to,
        )
        monthly = {}
        for d, amt in monthly_q.all():
            if d:
                k = d.strftime('%b %Y')
                if k not in monthly:
                    monthly[k] = {'month': k, 'so_count': 0, 'amount': 0.0, '_s': d.replace(day=1)}
                monthly[k]['so_count'] += 1
                monthly[k]['amount'] += round((amt or 0) / 1_000_000, 2)
        monthly_trend = sorted(monthly.values(), key=lambda x: x['_s'])
        for m in monthly_trend: del m['_s']

        top_vendors_q = apply_so_create_date_filter(
            db.session.query(
                SOData.vendor_name, func.count(SOData.id), func.sum(SOData.sales_amount)
            ).filter(open_so_filter(), SOData.vendor_name.isnot(None)),
            date_year, date_from, date_to,
        ).group_by(SOData.vendor_name).order_by(func.sum(SOData.sales_amount).desc()).limit(5)
        top_vendors = [
            {'vendor': r[0], 'so_count': r[1], 'total_amount': round(r[2] or 0, 2)}
            for r in top_vendors_q.all()
        ]

        top_op_units_q = apply_so_create_date_filter(
            db.session.query(
                SOData.operation_unit_name, func.count(SOData.id), func.sum(SOData.sales_amount)
            ).filter(open_so_filter(), SOData.operation_unit_name.isnot(None)),
            date_year, date_from, date_to,
        ).group_by(SOData.operation_unit_name).order_by(func.sum(SOData.sales_amount).desc()).limit(10)
        top_op_units = [
            {'op_unit': r[0], 'so_count': r[1], 'total_amount': round(r[2] or 0, 2)}
            for r in top_op_units_q.all()
        ]

        total_open_for_pct = total_so_count or 1
        so_status_q = apply_so_create_date_filter(
            db.session.query(
                SOData.so_status, func.count(SOData.id), func.sum(SOData.sales_amount)
            ).filter(open_so_filter(), SOData.so_status.isnot(None)),
            date_year, date_from, date_to,
        ).group_by(SOData.so_status).order_by(func.count(SOData.id).desc())
        so_status = [{'name': r[0], 'value': r[1],
            'percentage': round(r[1] / total_open_for_pct * 100, 1),
            'amount': round(r[2] or 0, 2)
        } for r in so_status_q.all()]

        monthly_by_status_q = apply_so_create_date_filter(
            db.session.query(
                SOData.so_status, SOData.so_create_date, SOData.sales_amount
            ).filter(open_so_filter()),
            date_year, date_from, date_to,
        )
        monthly_by_status = {}
        all_months_set = set()
        for s_status, s_date, s_amt in monthly_by_status_q.all():
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

        total_open_so_amount = apply_so_create_date_filter(
            db.session.query(func.sum(SOData.sales_amount)).filter(open_so_filter()),
            date_year, date_from, date_to,
        ).scalar() or 0

        po_date_range = db.session.query(func.min(POData.po_date), func.max(POData.po_date)).first()
        so_date_range = db.session.query(func.min(SOData.so_create_date), func.max(SOData.so_create_date)).first()

        # Last updated: most recent upload timestamp
        last_upload = db.session.query(func.max(UploadLog.uploaded_at)).scalar()
        if not last_upload:
            # fallback: most recent SO or PO record
            last_so = db.session.query(func.max(SOData.uploaded_at)).scalar()
            last_po = db.session.query(func.max(POData.uploaded_at)).scalar()
            candidates = [x for x in [last_so, last_po] if x]
            last_upload = max(candidates) if candidates else None

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
            'last_updated': utc_isoformat(last_upload),
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
    """Build set of PO references that appear in ANY SO record (open or closed).
    We use ALL statuses here because a PO that has ever been linked to an SO
    (even if Delivery Completed or SO Cancel) should NOT appear as 'PO without SO'.
    """
    matched = set()
    for s in db.session.query(
            SOData.customer_po_number, SOData.delivery_memo, SOData.so_item,
            SOData.operation_unit_name).all():
        cust_po, memo, so_item, op_unit = s[0], s[1], s[2], s[3]
        # Skip excluded op units
        if op_unit in EXCLUDED_OP_UNITS:
            continue
        # Skip return items (SO Item starting with 9)
        if is_return_so_item(so_item):
            continue
        # Extract ALL PO references from Customer PO Number and Delivery Memo
        for ref in extract_po_hli(cust_po) + extract_po_hli(memo):
            matched.add(ref)
    return matched


@app.route('/api/debug/matching', methods=['GET'])
def debug_matching():
    """Debug endpoint — check why a specific PO HLI is not matched."""
    po_number = request.args.get('po_number', '').strip()
    item_no   = request.args.get('item_no', '').strip() or None
    if not po_number:
        return jsonify({'error': 'Provide ?po_number=xxxx'}), 400

    matched_set = build_matched_set()
    item_variants = list(_normalize_item_no(item_no)) if item_no else []
    keys_checked = [po_number] + [f"{po_number}-{v}" for v in item_variants]
    hits = [k for k in keys_checked if k in matched_set]

    # Find which SOs mention this PO
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

    # Sample of matched_set entries starting with this po_number
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


@app.route('/api/debug/so-fields', methods=['GET'])
def debug_so_fields():
    """Debug endpoint — inspect spec/product_id fill rate and a sample of SO data."""
    try:
        total = db.session.query(func.count(SOData.id)).scalar() or 0
        has_spec = db.session.query(func.count(SOData.id)).filter(
            SOData.specification.isnot(None), SOData.specification != ''
        ).scalar() or 0
        has_pid = db.session.query(func.count(SOData.id)).filter(
            SOData.product_id.isnot(None), SOData.product_id != ''
        ).scalar() or 0
        samples = db.session.query(
            SOData.so_item, SOData.product_name, SOData.specification, SOData.product_id
        ).limit(10).all()
        return jsonify({
            'total_so_records': total,
            'records_with_specification': has_spec,
            'records_with_product_id': has_pid,
            'spec_fill_pct': round(has_spec / total * 100, 1) if total else 0,
            'pid_fill_pct': round(has_pid / total * 100, 1) if total else 0,
            'sample_rows': [
                {'so_item': r[0], 'product_name': r[1], 'specification': r[2], 'product_id': r[3]}
                for r in samples
            ],
            'hint': (
                'If spec_fill_pct and pid_fill_pct are 0%, your SMRO Excel file likely uses '
                'different column headers. Re-upload SMRO after checking column names. '
                'Supported names: Specification|Spec|Specifications — Product ID|Product Id|'
                'Product Code|Material|Material No|Material Number|Material Code|SKU'
            )
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


def po_is_matched(po_number, item_no, matched_set):
    """Return True if this PO item has a matching SO."""
    if po_number in matched_set:
        return True
    if item_no:
        for item_var in _normalize_item_no(item_no):
            if f"{po_number}-{item_var}" in matched_set:
                return True
    return False


def get_po_hli_key_set():
    """Build set of all PO HLI keys from PO table: both 'po_number' and 'po_number-item_no'."""
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

        # ── FIX: Deduplicate by (po_number, item_no) — unique rows only ──
        seen_keys = set()
        result = []
        for p in POData.query.all():
            key = (p.po_number, p.item_no)
            if key in seen_keys:
                continue
            seen_keys.add(key)

            # Skip hidden — check by combined PO HLI key (po_number-item_no)
            if is_po_hidden(p.po_number, p.item_no, hidden_po):
                continue
            op_unit = get_operation_unit(p.po_item_type, p.item_code)
            if op_unit in EXCLUDED_OP_UNITS:
                continue
            if not po_is_matched(p.po_number, p.item_no, matched_set):
                days_remaining = workdays_until(p.request_delivery, today)
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
        # Apply the same SO Create Date filter the dashboard count uses, so the
        # KPI count and the detail modal stay consistent.
        date_year, date_from, date_to = parse_so_date_args()
        q = apply_so_create_date_filter(
            db.session.query(SOData).filter(
                open_so_filter(),
                ~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))
            ),
            date_year, date_from, date_to,
        )
        result = []
        po_suffix_index = build_po_suffix_index(po_hli_keys)
        for s in q.all():
            if s.so_item in hidden_so or s.so_number in hidden_so:
                continue
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            if not so_has_matching_po_hli(s, po_hli_keys, po_suffix_index):
                result.append(so_dict(s))
        return jsonify(result)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/aging', methods=['GET'])
def get_aging_data():
    """SO Aging by vendor.
    Uses IDENTICAL filtering logic as total_so_count in /api/dashboard/stats so
    the TOTAL row in the aging table always matches the KPI card:
      - open SO only (open_so_filter)
      - excludes EXCLUDED_OP_UNITS
      - excludes hidden SO items
      - excludes return items (so_item starts with '9')
      - excludes internal PO refs
      - SO records without so_create_date are bucketed as '180+' (not silently dropped)
    """
    try:
        today = date.today()
        hidden_so = get_hidden_so_items()
        vendors = {}

        for s in db.session.query(SOData).filter(
            open_so_filter(),
            ~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))
        ).all():
            # Apply same exclusions as total_so_count
            if s.so_item in hidden_so or s.so_number in hidden_so:
                continue
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue

            v = s.vendor_name or 'Unknown'
            if v not in vendors:
                vendors[v] = {'vendor': v, 'less_30': 0, 'days_30_90': 0,
                              'days_90_180': 0, 'more_180': 0, 'total_open': 0, 'sales_amount': 0.0}

            age = workdays_since(s.so_create_date, today) if s.so_create_date else None
            if age is None:
                # No SO create date — put in 180+ bucket (same as total_so_count which counts them)
                vendors[v]['more_180'] += 1
            elif age < 30:
                vendors[v]['less_30'] += 1
            elif age < 90:
                vendors[v]['days_30_90'] += 1
            elif age < 180:
                vendors[v]['days_90_180'] += 1
            else:
                vendors[v]['more_180'] += 1
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
        hidden_so = get_hidden_so_items()
        sos = db.session.query(SOData).filter(
            open_so_filter(),
            SOData.vendor_name == vendor_name,
            ~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))
        ).order_by(SOData.so_create_date.asc()).all()
        sos = [s for s in sos
               if s.so_item not in hidden_so and s.so_number not in hidden_so
               and so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo)]
        if bucket:
            bucket = bucket.strip().replace(' ', '+')
            sos = [s for s in sos if get_aging_label(workdays_since(s.so_create_date, today) if s.so_create_date else None) == bucket]
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
        hidden_so = get_hidden_so_items()
        sos = db.session.query(SOData).filter(
            open_so_filter(),
            ~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))
        ).order_by(SOData.vendor_name.asc(), SOData.so_create_date.asc()).all()
        sos = [s for s in sos
               if s.so_item not in hidden_so and s.so_number not in hidden_so
               and so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo)]
        if bucket:
            sos = [s for s in sos if get_aging_label(workdays_since(s.so_create_date, today) if s.so_create_date else None) == bucket]
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
        date_year     = request.args.get('date_year', '')
        date_from     = request.args.get('date_from', '')
        date_to       = request.args.get('date_to', '')
        page          = max(1, int(request.args.get('page', 1)))
        per_page      = min(500, int(request.args.get('per_page', 20)))

        today = date.today()
        q = SOData.query.filter(open_so_filter())
        if op_units:  q = q.filter(SOData.operation_unit_name.in_(op_units))
        if vendors:   q = q.filter(SOData.vendor_name.in_(vendors))
        if statuses:  q = q.filter(SOData.so_status.in_(statuses))
        if so_items:  q = q.filter(SOData.so_item.in_(so_items))
        is_sqlite = 'sqlite' in app.config['SQLALCHEMY_DATABASE_URI']
        if date_year:
            try:
                yr = int(date_year)
                if is_sqlite:
                    q = q.filter(func.strftime('%Y', SOData.so_create_date) == str(yr))
                else:
                    q = q.filter(func.extract('year', SOData.so_create_date) == yr)
            except ValueError:
                pass
        elif date_from or date_to:
            if date_from:
                q = q.filter(SOData.so_create_date >= date_from)
            if date_to:
                q = q.filter(SOData.so_create_date <= date_to)

        all_sos = q.order_by(SOData.so_create_date.asc()).all()

        if aging:
            def matches_aging(s):
                age = workdays_since(s.so_create_date, today)
                return get_aging_label(age) in aging
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
# DELETE REQUEST ENDPOINTS (soft-hide from dashboard)
# ═══════════════════════════════════════════════════════════════════

@app.route('/api/delete-requests', methods=['GET'])
def get_delete_requests():
    """Return all delete requests (both hidden and visible)."""
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
    """Request to hide a PO or SO from dashboard."""
    try:
        data = request.json
        ref_type = data.get('ref_type', '').upper()
        ref_number = (data.get('ref_number') or '').strip()
        reason = (data.get('reason') or '').strip()

        if ref_type not in ('PO', 'SO'):
            return jsonify({'error': 'ref_type harus PO atau SO'}), 400
        if not ref_number:
            return jsonify({'error': 'Reference number is required'}), 400
        if not reason:
            return jsonify({'error': 'Reason is required'}), 400

        # Check if already requested
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
        return jsonify({'success': True, 'id': req.id, 'message': f'{ref_type} {ref_number} successfully hidden from dashboard'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500


@app.route('/api/delete-requests/<int:req_id>/restore', methods=['PUT'])
def restore_delete_request(req_id):
    """Restore a hidden item back to dashboard."""
    try:
        req = db.session.get(DeleteRequest, req_id)
        if not req:
            return jsonify({'error': 'Request not found'}), 404
        req.is_hidden = False
        db.session.commit()
        return jsonify({'success': True, 'message': f'{req.ref_type} {req.ref_number} successfully restored to dashboard'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500


@app.route('/api/delete-requests/<int:req_id>', methods=['DELETE'])
def delete_request_permanently(req_id):
    """Permanently remove a delete request record."""
    try:
        req = db.session.get(DeleteRequest, req_id)
        if not req:
            return jsonify({'error': 'Request not found'}), 404
        db.session.delete(req)
        db.session.commit()
        return jsonify({'success': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500


# ═══════════════════════════════════════════════════════════════════
# UPLOAD ENDPOINTS
# ═══════════════════════════════════════════════════════════════════

CHUNK_SIZE = 200

@app.route('/api/upload/po-list', methods=['POST'])
def upload_po_list():
    try:
        if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

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
                    f'❌ Invalid file — {len(missing_required)} required columns not found: '
                    f'{", ".join(missing_required)}. '
                    f'Please make sure you are uploading the correct HLI PO List file and try again.'
                )
            }), 400

        col_po   = find_column(df, ['PO No.','PO No','PO Number','PO'])
        if not col_po:
            return jsonify({'error': f'PO Number column not found. Available columns: {df.columns.tolist()}'}), 400

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
                existing = existing_po[key]
                preserved_plan_date = existing.delivery_plan_date
                preserved_remarks = existing.remarks
                for field, val in new_data.items():
                    setattr(existing, field, val)
                existing.delivery_plan_date = preserved_plan_date
                existing.remarks = preserved_remarks
            else:
                new_rec = POData(**new_data)
                db.session.add(new_rec)

            count += 1
            if count % CHUNK_SIZE == 0:
                db.session.flush()

        # FIX: Do not delete old PO data (preserve history)
        # keys_to_delete logic removed

        db.session.add(UploadLog(file_type='PO', filename=file.filename, records_count=count))
        db.session.commit()
        return jsonify({'message': f'Berhasil upload {count} PO items', 'uploaded': count})
    except Exception as e:
        db.session.rollback(); import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/upload/smro', methods=['POST'])
def upload_smro():
    """
    SMRO upload — upsert only.
    - Records with SO Item matching new file → updated (remarks & delivery_plan_date preserved)
    - Records with SO Item NOT in new file → KEPT as-is (not deleted)
    - New SO Items in file not in DB → inserted
    """
    try:
        if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

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
                    f'❌ Invalid file — {len(missing_required)} required columns not found: '
                    f'{", ".join(missing_required)}. '
                    f'Please make sure you are uploading the correct SMRO file and try again.'
                )
            }), 400

        col_so = find_column(df, ['SO Number','SO No','SO No.','SO','SO Item',
                                   'Sales Order','Sales Order Number','No SO','Nomor SO'])
        if not col_so:
            return jsonify({'error': f'SO Number column not found. Available columns: {df.columns.tolist()}'}), 400

        col_soitem  = find_column(df, ['SO Item No','Item No','Line','SO Line','SO Item'])
        col_status  = find_column(df, ['SO Status','Status','Order Status'])
        col_opunit  = find_column(df, ['Operation Unit Name','Op Unit','Client Name','Client','Operation Unit'])
        col_vendor  = find_column(df, ['Vendor Name','Vendor','Supplier'])
        col_custpo  = find_column(df, ['Customer PO Number','Customer PO','PO Ref','PO Reference'])
        col_memo    = find_column(df, ['Delivery Memo','Memo','Delivery Note'])
        col_prod    = find_column(df, ['Product Name','Item Name','Description','Product'])
        col_spec    = find_column(df, ['Specification','Spec','Specifications','Product Specification'])
        col_pid     = find_column(df, ['Product ID','Product Id','Product Code','Material','Material No','Material Number','Material Code','SKU'])
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

        # Build lookup of existing SO records by so_item
        existing_so = {}
        for s in SOData.query.all():
            if s.so_item:
                existing_so[s.so_item] = s

        count = 0
        updated = 0
        inserted = 0

        for _, row in df.iterrows():
            so_val = clean(df_val(row, col_so))
            if not so_val: continue
            so_item_val = clean(df_val(row, col_soitem))

            new_data = {
                'so_number': so_val,
                'so_item': so_item_val,
                'so_status': clean(df_val(row, col_status)),
                'operation_unit_name': clean(df_val(row, col_opunit)),
                'vendor_name': clean(df_val(row, col_vendor)),
                'customer_po_number': clean(df_val(row, col_custpo)),
                'delivery_memo': clean(df_val(row, col_memo)),
                'product_name': clean(df_val(row, col_prod)),
                'specification': clean(df_val(row, col_spec)),
                'product_id': clean(df_val(row, col_pid)),
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
                # Update existing record — preserve remarks & delivery_plan_date
                existing = existing_so[so_item_val]
                preserved_remarks = existing.remarks
                preserved_plan_date = existing.delivery_plan_date
                for field, val in new_data.items():
                    setattr(existing, field, val)
                existing.remarks = preserved_remarks
                existing.delivery_plan_date = preserved_plan_date
                updated += 1
            else:
                new_rec = SOData(**new_data)
                db.session.add(new_rec)
                inserted += 1

            count += 1
            if count % CHUNK_SIZE == 0:
                db.session.flush()

        # ── KEY CHANGE: Do NOT delete records not in this file ──
        # Old records with different SO Items are preserved as-is.

        db.session.add(UploadLog(file_type='SO', filename=file.filename, records_count=count))
        db.session.commit()
        return jsonify({
            'message': f'Berhasil: {inserted} SO baru ditambahkan, {updated} SO diperbarui. Data lama yang tidak ada di file ini tetap dipertahankan.',
            'uploaded': count,
            'inserted': inserted,
            'updated': updated
        })
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
                       'PO Date','Purchase Member','Request Delivery','Days Remaining',
                       'Delivery Plan Date','Remarks'], num_cols=[8,10,11])
        for p, op_unit in pos:
            days_rem = workdays_until(p.request_delivery, today) if p.request_delivery else ''
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


@app.route('/api/template/hide', methods=['GET'])
def download_hide_template():
    """Download Excel template for batch hide requests."""
    hide_type = request.args.get('type', 'PO').upper()  # 'PO' or 'SO'
    wb = Workbook()
    ws = wb.active

    if hide_type == 'SO':
        ws.title = "Hide SO Template"
        headers = ['SO Number', 'Reason']
        ws.append(headers)
        # Style header
        fill = PatternFill(start_color="1D4ED8", end_color="1D4ED8", fill_type="solid")
        for i, cell in enumerate(ws[1], 1):
            cell.fill = fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(i)].width = 30 if i == 1 else 50
        # Example row
        ws.append(['9008988017-10', 'Reason why this SO should be hidden'])
        note_row = ws.append(['INSTRUCTIONS: Fill SO Number (format: SO_NUMBER-ITEM_NO or SO_NUMBER), and Reason (required)'])
    else:
        ws.title = "Hide PO HLI Template"
        headers = ['NO PO HLI (PO Number-Item No)', 'Reason']
        ws.append(headers)
        # Style header
        fill = PatternFill(start_color="7C3AED", end_color="7C3AED", fill_type="solid")
        for i, cell in enumerate(ws[1], 1):
            cell.fill = fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(i)].width = 35 if i == 1 else 50
        # Example row
        ws.append(['4502358819-10', 'Reason why this PO HLI should be hidden'])
        ws.append(['INSTRUCTIONS: Fill NO PO HLI with format PO_NUMBER-ITEM_NO (example: 4502358819-10), and Reason (required)'])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    fname = f"Template_Hide_{'SO' if hide_type == 'SO' else 'PO_HLI'}.xlsx"
    return send_file(output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True, download_name=fname)


@app.route('/api/upload/hide-batch', methods=['POST'])
def upload_hide_batch():
    """Process batch hide Excel file. Supports both PO HLI and SO hide templates."""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        hide_type = request.form.get('type', 'PO').upper()

        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        # Detect column names
        if hide_type == 'SO':
            col_ref = find_column(df, ['SO Number', 'SO No', 'SO Item', 'SO Number-Item No'])
        else:
            col_ref = find_column(df, [
                'NO PO HLI (PO Number-Item No)', 'NO PO HLI', 'PO HLI',
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

            # Skip header-like rows (instructions)
            if not ref_number or ref_number.upper().startswith('PETUNJUK') or ref_number.upper().startswith('INSTRUCTIONS'):
                continue
            # Skip example rows
            if reason and (reason.lower().startswith('alasan kenapa') or reason.lower().startswith('reason why')):
                continue

            if not reason:
                errors.append(f"Row {idx+2}: Reason is empty for {ref_number}")
                continue

            # Check if already hidden
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

        msg = f'{success_count} items successfully hidden'
        if skipped:
            msg += f'. {len(skipped)} were already hidden: {", ".join(skipped[:5])}'
        if errors:
            msg += f'. {len(errors)} error: {"; ".join(errors[:3])}'

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


@app.route('/api/completed/summary', methods=['GET'])
def completed_summary():
    try:
        year_filter = request.args.get('year', 'all')
        date_year   = request.args.get('date_year', '')
        date_from   = request.args.get('date_from', '')
        date_to     = request.args.get('date_to', '')
        is_sqlite = 'sqlite' in app.config['SQLALCHEMY_DATABASE_URI']

        q = db.session.query(SOData).filter(SOData.so_status == 'Delivery Completed')

        # Apply SO Create Date filter (date_year takes precedence over range,
        # and falls back to legacy `year` query param when present).
        effective_year = date_year or (year_filter if year_filter and year_filter != 'all' else '')
        if effective_year:
            try:
                yr = int(effective_year)
                if is_sqlite:
                    q = q.filter(func.strftime('%Y', SOData.so_create_date) == str(yr))
                else:
                    q = q.filter(func.extract('year', SOData.so_create_date) == yr)
            except ValueError:
                pass
        else:
            if date_from:
                q = q.filter(SOData.so_create_date >= date_from)
            if date_to:
                q = q.filter(SOData.so_create_date <= date_to)

        # Exclude consumable / non-revenue op units, matching every other
        # SOData query in the codebase (see /api/completed/margin-detail, etc.).
        rows = q.filter(~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))).all()

        def po_amt_of(s):
            v = float(s.purchasing_amount or 0)
            if v == 0 and s.purchasing_price:
                v = float(s.purchasing_price) * float(s.so_qty or 0)
            return v

        # Pre-compute per-row sales/purchase/margin once, then reuse.
        enriched = []
        for s in rows:
            po_amt = po_amt_of(s)
            sales = float(s.sales_amount or 0)
            enriched.append((s, po_amt, sales, sales - po_amt))

        # Monthly trend
        monthly = {}
        for s, po_amt, sales, _m in enriched:
            if not s.so_create_date:
                continue
            key = s.so_create_date.strftime('%Y-%m')
            if key not in monthly:
                monthly[key] = {'month': key, 'count': 0, 'sales_amount': 0.0, 'purchase_amount': 0.0}
            monthly[key]['count'] += 1
            monthly[key]['sales_amount'] += sales
            monthly[key]['purchase_amount'] += po_amt

        monthly_trend = sorted(monthly.values(), key=lambda x: x['month'])

        # Vendor summary (top 5 by sales)
        vendor_map = {}
        for s, po_amt, sales, m in enriched:
            v = s.vendor_name or 'Unknown'
            if v not in vendor_map:
                vendor_map[v] = {'vendor': v, 'count': 0, 'sales_amount': 0.0, 'purchase_amount': 0.0, 'margin': 0.0}
            vendor_map[v]['count'] += 1
            vendor_map[v]['sales_amount'] += sales
            vendor_map[v]['purchase_amount'] += po_amt
            vendor_map[v]['margin'] += m

        top_vendors = sorted(vendor_map.values(), key=lambda x: x['sales_amount'], reverse=True)[:5]

        # Margin distribution + totals (KPI cards)
        pos = neg = zero = 0
        total_sales = 0.0
        total_purchase = 0.0
        for _s, po_amt, sales, m in enriched:
            total_sales += sales
            total_purchase += po_amt
            if m > 0:
                pos += 1
            elif m < 0:
                neg += 1
            else:
                zero += 1

        # Top 20 items by sales amount (grouped by product / item label).
        # Also surface Specification + Product ID so the frontend table can
        # display them.  Group key prefers Product ID when present so the
        # same product across multiple SOs aggregates correctly even when
        # `product_name` differs slightly.
        item_map = {}
        for s, po_amt, sales, m in enriched:
            pid = (s.product_id or '').strip()
            label = s.product_name or s.so_item or 'Unknown'
            key = pid or label
            if key not in item_map:
                item_map[key] = {
                    'item': label,
                    'specification': s.specification or '',
                    'product_id': pid,
                    'count': 0, 'sales_amount': 0.0,
                    'purchase_amount': 0.0, 'margin': 0.0,
                }
            agg = item_map[key]
            agg['count'] += 1
            agg['sales_amount'] += sales
            agg['purchase_amount'] += po_amt
            agg['margin'] += m
            # Backfill spec from later rows if the first one was empty.
            if not agg['specification'] and s.specification:
                agg['specification'] = s.specification

        top_items = sorted(item_map.values(), key=lambda x: x['sales_amount'], reverse=True)[:20]

        # Worst-margin vendors: vendors with one or more negative-margin txns,
        # ranked by total negative margin (most negative first).
        neg_vendor_map = {}
        for s, po_amt, sales, m in enriched:
            if m >= 0:
                continue
            v = s.vendor_name or 'Unknown'
            if v not in neg_vendor_map:
                neg_vendor_map[v] = {
                    'vendor': v, 'margin': 0.0, 'count': 0,
                    'total_sales': 0.0, 'total_purchase': 0.0,
                }
            neg_vendor_map[v]['margin'] += m
            neg_vendor_map[v]['count'] += 1
            neg_vendor_map[v]['total_sales'] += sales
            neg_vendor_map[v]['total_purchase'] += po_amt

        worst_margin_vendors = sorted(neg_vendor_map.values(), key=lambda x: x['margin'])[:50]

        # Top 30 worst-margin transactions (UI scrolls within fixed-height box)
        neg_txns = [(s, po_amt, sales, m) for s, po_amt, sales, m in enriched if m < 0]
        neg_txns.sort(key=lambda x: x[3])  # most negative first
        worst_margin_transactions = []
        for s, po_amt, sales, m in neg_txns[:30]:
            pct = round(m / sales * 100, 1) if sales else None
            worst_margin_transactions.append({
                'so_item': s.so_item,
                'so_number': s.so_number,
                'item_code': (s.item_code if hasattr(s, 'item_code') and s.item_code else (s.so_item or '-')),
                'product': s.product_name or '-',
                'vendor': s.vendor_name or '-',
                'sales_amount': sales,
                'purchase_amount': po_amt,
                'margin': m,
                'margin_pct': pct,
                'count': 1,
                'date': s.so_create_date.isoformat() if s.so_create_date else None,
            })

        return jsonify({
            'total_count': len(rows),
            'total_sales': total_sales,
            'total_purchase': total_purchase,
            'total_margin': total_sales - total_purchase,
            'monthly_trend': monthly_trend,
            'top_vendors': top_vendors,
            'top_items': top_items,
            'worst_margin_vendors': worst_margin_vendors,
            'worst_margin_transactions': worst_margin_transactions,
            'margin_distribution': {
                'positive': pos,
                'negative': neg,
                'zero': zero
            }
        })

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500



@app.route('/api/completed/margin-detail', methods=['GET'])
def completed_margin_detail():
    """Return rows for a specific margin category (positive/negative/zero) for popup."""
    try:
        category = request.args.get('category', 'positive')  # positive|negative|zero
        date_from = request.args.get('date_from', '')
        date_to   = request.args.get('date_to', '')
        date_year = request.args.get('date_year', '')
        is_sqlite = 'sqlite' in app.config['SQLALCHEMY_DATABASE_URI']

        q = db.session.query(SOData).filter(SOData.so_status == 'Delivery Completed')
        if date_year:
            try:
                yr = int(date_year)
                if is_sqlite:
                    q = q.filter(func.strftime('%Y', SOData.so_create_date) == str(yr))
                else:
                    q = q.filter(func.extract('year', SOData.so_create_date) == yr)
            except ValueError:
                pass
        elif date_from or date_to:
            if date_from:
                q = q.filter(SOData.so_create_date >= date_from)
            if date_to:
                q = q.filter(SOData.so_create_date <= date_to)

        rows = q.filter(~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS))).all()

        def get_po_amt(s):
            po_amt = float(s.purchasing_amount or 0)
            if po_amt == 0 and s.purchasing_price:
                po_amt = float(s.purchasing_price) * float(s.so_qty or 0)
            return po_amt

        result = []
        for s in rows:
            po_amt = get_po_amt(s)
            m = float(s.sales_amount or 0) - po_amt
            if category == 'positive' and m <= 0:
                continue
            elif category == 'negative' and m >= 0:
                continue
            elif category == 'zero' and m != 0:
                continue
            result.append({
                'so_item': s.so_item,
                'so_number': s.so_number,
                'product': s.product_name or '-',
                'vendor': s.vendor_name or '-',
                'item_code': (s.item_code if hasattr(s, 'item_code') and s.item_code else '-'),
                'sales_amount': float(s.sales_amount or 0),
                'purchase_amount': po_amt,
                'margin': m,
                'margin_pct': round(m / float(s.sales_amount) * 100, 1) if s.sales_amount else None,
                'date': s.so_create_date.isoformat() if s.so_create_date else None,
                'so_status': s.so_status,
                'operation_unit_name': s.operation_unit_name,
            })

        result.sort(key=lambda x: x['margin'])
        return jsonify(result)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("Backend: http://127.0.0.1:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
