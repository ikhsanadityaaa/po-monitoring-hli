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

# Short-lived in-memory response cache for heavy Delivery Completed analytics.
# Keyed by request filters + cheap DB signature, so repeated page opens return
# immediately while uploads/changed rows naturally invalidate the cache.
_COMPLETED_SUMMARY_CACHE = {}
_COMPLETED_SUMMARY_CACHE_TTL_SECONDS = 300

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
    purchasing_currency = db.Column(db.String(10))
    purchasing_amount_idr = db.Column(db.Float)
    purchasing_amount_idr_cached_at = db.Column(db.DateTime)
    so_create_date = db.Column(db.Date)
    delivery_possible_date = db.Column(db.Date)
    matched_po_number = db.Column(db.String(50))
    delivery_plan_date = db.Column(db.Date)
    remarks = db.Column(db.Text)
    pic_name = db.Column(db.String(100))
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


class ExchangeRate(db.Model):
    """USD->IDR exchange rate per date. Auto-fetched via Frankfurter API on first need,
    or set manually via /api/exchange-rate endpoint."""
    __tablename__ = 'exchange_rate'
    id         = db.Column(db.Integer, primary_key=True)
    rate_date  = db.Column(db.Date, nullable=False, unique=True, index=True)
    usd_to_idr = db.Column(db.Float, nullable=False)
    source     = db.Column(db.String(50), default='manual')  # 'frankfurter'|'manual'|'fallback'
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class ProductIDDB(db.Model):
    """Database of Product ID → Category ID, downloaded from SAP."""
    __tablename__ = 'product_id_db'
    id            = db.Column(db.Integer, primary_key=True)
    product_id    = db.Column(db.String(100), unique=True, nullable=False, index=True)
    category_id   = db.Column(db.String(100))
    category_name = db.Column(db.String(255))
    product_name  = db.Column(db.Text)
    updated_at    = db.Column(db.DateTime, default=datetime.utcnow)


class MasterPIC(db.Model):
    """Master mapping: Category ID → PIC name. Updated manually via UI."""
    __tablename__ = 'master_pic'
    id            = db.Column(db.Integer, primary_key=True)
    category_id   = db.Column(db.String(100), unique=True, nullable=False, index=True)
    category_name = db.Column(db.String(255))
    pic_name      = db.Column(db.String(100))
    updated_at    = db.Column(db.DateTime, default=datetime.utcnow)


# ─── Exchange rate helpers ─────────────────────────────────────────────────
_RATE_CACHE = {}   # {date: float} in-process cache

def _fetch_rate_from_api(d):
    """Try Frankfurter (ECB data) for historical USD->IDR. Returns float or None."""
    try:
        import urllib.request, json as _json
        url = f"https://api.frankfurter.app/{d.isoformat()}?from=USD&to=IDR"
        with urllib.request.urlopen(url, timeout=6) as resp:
            data = _json.loads(resp.read())
        return float(data['rates']['IDR'])
    except Exception:
        return None

def _get_fallback_rate():
    last = ExchangeRate.query.order_by(ExchangeRate.rate_date.desc()).first()
    return last.usd_to_idr if last else 16000.0

def get_usd_to_idr(d, cache_only=False):
    """Return USD->IDR rate for date d.
    Order: in-memory cache -> DB exact -> (if not cache_only) API fetch -> DB nearest -> hardcoded fallback.

    Pass cache_only=True when calling inside a loop that has already called
    prefetch_exchange_rates() — this skips the expensive HTTP request path
    entirely, relying on the already-warmed in-memory cache and DB."""
    if d is None:
        return _get_fallback_rate()
    if d in _RATE_CACHE:
        return _RATE_CACHE[d]
    rec = ExchangeRate.query.filter_by(rate_date=d).first()
    if rec:
        _RATE_CACHE[d] = rec.usd_to_idr
        return rec.usd_to_idr
    if not cache_only and d <= date.today():
        rate = _fetch_rate_from_api(d)
        if rate:
            try:
                db.session.add(ExchangeRate(rate_date=d, usd_to_idr=rate, source='frankfurter'))
                db.session.commit()
            except Exception:
                db.session.rollback()
            _RATE_CACHE[d] = rate
            return rate
    # Nearest known rate (no HTTP call)
    nearest = ExchangeRate.query.order_by(
        func.abs(func.julianday(ExchangeRate.rate_date) - func.julianday(str(d)))
    ).first()
    if nearest:
        _RATE_CACHE[d] = nearest.usd_to_idr
        return nearest.usd_to_idr
    return _get_fallback_rate()


def prefetch_exchange_rates(dates, fetch_missing=True):
    """Warm the in-memory _RATE_CACHE for a collection of dates, minimising
    round-trips to the Frankfurter API.

    Algorithm:
    1. Skip dates already in _RATE_CACHE (already warm).
    2. Bulk-load all ExchangeRate rows from the DB in a single query and
       populate the cache — avoids N individual DB lookups.
    3. If fetch_missing=True, for any remaining dates (still not in cache) that
       are ≤ today, fetch from the Frankfurter API **sequentially** and persist
       to DB. For latency-sensitive dashboard endpoints, pass
       fetch_missing=False so page load never waits on external API calls.
    4. After optional API fetches, one final pass stores any date still missing
       in the in-process cache using the nearest DB rate / fallback (no HTTP).

    Call this once at the top of any endpoint that iterates over many rows and
    calls convert_to_idr(), then pass cache_only=True to get_usd_to_idr()
    inside the loop.
    """
    if not dates:
        return

    # 1. Filter to dates not already cached
    needed = {d for d in dates if d is not None and d not in _RATE_CACHE}
    if not needed:
        return

    # 2. Bulk DB load — single query for all needed dates
    db_rows = ExchangeRate.query.filter(ExchangeRate.rate_date.in_(list(needed))).all()
    for row in db_rows:
        _RATE_CACHE[row.rate_date] = row.usd_to_idr
    needed -= {row.rate_date for row in db_rows}

    if not needed:
        return

    # 3. Optionally fetch remaining from API (only past/today dates).
    # Dashboard/page-load code should call this with fetch_missing=False; missing
    # historical rates can be filled once via /api/exchange-rate/fetch and then
    # reused from the DB forever.
    if fetch_missing:
        today = date.today()
        to_api = sorted(d for d in needed if d <= today)
        fetched_rows = []
        for d in to_api:
            rate = _fetch_rate_from_api(d)
            if rate:
                _RATE_CACHE[d] = rate
                fetched_rows.append(ExchangeRate(rate_date=d, usd_to_idr=rate, source='frankfurter'))
                needed.discard(d)

        if fetched_rows:
            try:
                db.session.bulk_save_objects(fetched_rows)
                db.session.commit()
            except Exception:
                db.session.rollback()

    # 4. Proxy remaining with nearest known rate (no extra HTTP calls)
    if needed:
        fallback = _get_fallback_rate()
        # Load all rates once for proximity search
        all_rates = ExchangeRate.query.order_by(ExchangeRate.rate_date).all()
        for d in needed:
            if all_rates:
                nearest = min(all_rates, key=lambda r: abs((r.rate_date - d).days))
                _RATE_CACHE[d] = nearest.usd_to_idr
            else:
                _RATE_CACHE[d] = fallback


def convert_to_idr(amount, currency, rate_date=None, cache_only=False):
    """Convert amount to IDR. Non-USD/IDR currencies returned as-is.

    When called inside a loop preceded by prefetch_exchange_rates(), pass
    cache_only=True to skip the per-row HTTP fallback path entirely."""
    if not amount:
        return 0.0
    cur = (currency or 'IDR').strip().upper()
    if cur in ('IDR', ''):
        return float(amount)
    if cur == 'USD':
        return float(amount) * get_usd_to_idr(rate_date, cache_only=cache_only)
    return float(amount)


def raw_purchase_amount(s):
    """Return the original purchasing amount before currency conversion."""
    raw = float(s.purchasing_amount or 0)
    if raw == 0 and s.purchasing_price:
        raw = float(s.purchasing_price) * float(s.so_qty or 0)
    return raw


def purchase_amount_idr(s, allow_persist=False):
    """Return cached/persisted purchase amount in IDR for SOData.

    Delivery Completed pages call this for every row. For IDR rows the value is
    already final. For USD rows we persist the converted result in so_data, so
    old rows are not converted again on every page load. Only rows whose cache
    is still empty (typically newly uploaded/updated non-IDR rows) are computed.
    """
    cached = getattr(s, 'purchasing_amount_idr', None)
    if cached is not None:
        return float(cached)

    raw = raw_purchase_amount(s)
    cur = (s.purchasing_currency or 'IDR').strip().upper()
    if cur in ('IDR', ''):
        converted = raw
    else:
        converted = convert_to_idr(raw, s.purchasing_currency, s.so_create_date, cache_only=True)

    if allow_persist:
        s.purchasing_amount_idr = converted
        s.purchasing_amount_idr_cached_at = datetime.utcnow()
    return converted


def ensure_purchase_amount_idr_cache(rows):
    """Persist missing IDR purchase amounts for rows in the current result set."""
    missing = [s for s in rows if getattr(s, 'purchasing_amount_idr', None) is None]
    if not missing:
        return 0

    usd_missing = [s for s in missing if (s.purchasing_currency or '').strip().upper() == 'USD']
    if usd_missing:
        prefetch_exchange_rates({s.so_create_date for s in usd_missing if s.so_create_date}, fetch_missing=False)

    for s in missing:
        purchase_amount_idr(s, allow_persist=True)

    try:
        db.session.commit()
    except Exception:
        db.session.rollback()
    return len(missing)


def _ensure_extra_columns():
    """Online migration: add optional columns that older local DBs may miss."""
    is_sqlite = 'sqlite' in app.config['SQLALCHEMY_DATABASE_URI']

    def existing_columns(table_name):
        try:
            if is_sqlite:
                result = db.session.execute(text(f"PRAGMA table_info({table_name})"))
                return {row[1].lower() for row in result}
            result = db.session.execute(text(
                "SELECT column_name FROM information_schema.columns "
                f"WHERE table_name = '{table_name}'"
            ))
            return {row[0].lower() for row in result}
        except Exception:
            return set()

    migration_plan = {
        'so_data': [
            ('specification',                      'TEXT'),
            ('product_id',                         'VARCHAR(100)'),
            ('purchasing_currency',                'VARCHAR(10)'),
            ('purchasing_amount_idr',              'DOUBLE PRECISION'),
            ('purchasing_amount_idr_cached_at',    'TIMESTAMP'),
            ('pic_name',                           'VARCHAR(100)'),
        ],
        'po_data': [
            ('delivery_plan_date',   'DATE'),
            ('remarks',              'TEXT'),
        ],
    }

    for table_name, columns in migration_plan.items():
        cols = existing_columns(table_name)
        for col_name, col_type in columns:
            if col_name.lower() not in cols:
                try:
                    db.session.execute(text(f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_type}"))
                    db.session.commit()
                    print(f'DB migration: added column {table_name}.{col_name}')
                except Exception as exc:
                    db.session.rollback()
                    print(f'DB migration warning ({table_name}.{col_name}): {exc}')


def _ensure_so_extra_columns():
    """Online migration: add `specification` and `product_id` columns to
    so_data if they don't exist yet.

    Strategy (works on all SQLite versions and PostgreSQL):
    1. Query the actual column list from the DB (PRAGMA for SQLite,
       information_schema for Postgres).
    2. Only issue ALTER TABLE when the column is genuinely absent.
    This avoids the `IF NOT EXISTS` clause that many SQLite builds reject.
    """
    is_sqlite = 'sqlite' in app.config['SQLALCHEMY_DATABASE_URI']
    needed = [
        ('specification',                      'TEXT'),
        ('product_id',                         'VARCHAR(100)'),
        ('purchasing_currency',                'VARCHAR(10)'),
        ('purchasing_amount_idr',              'DOUBLE PRECISION'),
        ('purchasing_amount_idr_cached_at',    'TIMESTAMP'),
    ]
    try:
        if is_sqlite:
            result = db.session.execute(text("PRAGMA table_info(so_data)"))
            existing_cols = {row[1].lower() for row in result}
        else:
            result = db.session.execute(text(
                "SELECT column_name FROM information_schema.columns "
                "WHERE table_name = 'so_data'"
            ))
            existing_cols = {row[0].lower() for row in result}
    except Exception:
        existing_cols = set()

    for col_name, col_type in needed:
        if col_name.lower() not in existing_cols:
            try:
                db.session.execute(
                    text(f"ALTER TABLE so_data ADD COLUMN {col_name} {col_type}")
                )
                db.session.commit()
                print(f'DB migration: added column so_data.{col_name}')
            except Exception as exc:
                db.session.rollback()
                print(f'DB migration warning (so_data.{col_name}): {exc}')


with app.app_context():
    db.create_all()
    _ensure_extra_columns()
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

def is_return_so_status(so_status):
    return 'return' in str(so_status or '').strip().lower()

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
    
    # Get category from ProductIDDB (level 1 only — before first >)
    category_name = ''
    if s.product_id:
        prod = db.session.query(ProductIDDB).filter_by(product_id=str(s.product_id).strip()).first()
        if prod and prod.category_name:
            # Extract level 1 category only (before first >)
            full_category = prod.category_name.strip()
            category_name = full_category.split('>')[0].strip() if '>' in full_category else full_category
    
    return {
        'id': s.id, 'so_number': s.so_number, 'so_item': s.so_item,
        'so_status': s.so_status, 'operation_unit_name': s.operation_unit_name,
        'vendor_name': s.vendor_name, 'customer_po_number': s.customer_po_number,
        'delivery_memo': s.delivery_memo, 'product_name': s.product_name,
        'specification': s.specification, 'product_id': s.product_id,
        'category_name': category_name,
        'svo_po': s.matched_po_number or '',
        'so_qty': s.so_qty, 'sales_price': s.sales_price, 'sales_amount': s.sales_amount,
        'purchasing_price': s.purchasing_price, 'purchasing_amount': s.purchasing_amount,
        'purchasing_currency': s.purchasing_currency,
        'so_create_date': s.so_create_date.isoformat() if s.so_create_date else '',
        'delivery_possible_date': s.delivery_possible_date.isoformat() if s.delivery_possible_date else '',
        'delivery_plan_date': s.delivery_plan_date.isoformat() if s.delivery_plan_date else '',
        'remarks': s.remarks or '',
        'pic_name': s.pic_name or '',
        'aging_days': age_days,
        'aging_label': get_aging_label(age_days)
    }

# ─── Build hidden set from delete requests ────────────────────────────────
def get_hidden_po_hli_keys():
    """Return set of hidden PO HLI keys. Format stored: 'po_number-item_no' or just 'po_number'."""
    rows = db.session.query(DeleteRequest.ref_number).filter_by(ref_type='PO', is_hidden=True).all()
    return {r[0] for r in rows}

def get_hidden_so_items():
    """Return set of SO items/numbers that are hidden from dashboard."""
    rows = db.session.query(DeleteRequest.ref_number).filter_by(ref_type='SO', is_hidden=True).all()
    return {r[0] for r in rows}

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

        po_numbers = get_po_hli_key_set()

        matched_set = build_matched_set()

        po_without_so_count = 0
        for p in POData.query.with_entities(
            POData.po_number, POData.item_no, POData.po_item_type, POData.item_code
        ).all():
            if is_po_hidden(p.po_number, p.item_no, hidden_po):
                continue
            op_unit = get_operation_unit(p.po_item_type, p.item_code)
            if op_unit in EXCLUDED_OP_UNITS:
                continue
            if not po_is_matched(p.po_number, p.item_no, matched_set):
                po_without_so_count += 1

        po_suffix_index = build_po_suffix_index(po_numbers)

        # Single pass over open SO rows — compute total_so_count AND so_without_po_count
        # together instead of two separate full-table scans.
        open_so_rows = so_q(
            open_so_filter()
        ).all()

        total_so_count = 0
        so_without_po_count = 0
        for s in open_so_rows:
            if s.so_item in hidden_so or s.so_number in hidden_so:
                continue
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue
            total_so_count += 1
            # KPI SO without PO HLI excludes:
            # - return statuses
            # - HLI Consumable operation unit
            # - SOs created before 2024
            if (
                not is_return_so_status(s.so_status)
                and s.operation_unit_name not in EXCLUDED_OP_UNITS
                and s.so_create_date and s.so_create_date.year >= 2024
                and not so_has_matching_po_hli(s, po_numbers, po_suffix_index)
            ):
                so_without_po_count += 1

        # All SO-based dashboard aggregates below must use exactly the same
        # canonical Open SO dataset as the KPI above:
        #   - SO create date filter
        #   - open status filter
        #   - excluded operation units removed
        #   - hidden SO removed
        #   - non-countable SO item/customer PO/delivery memo removed
        # This keeps KPI, status distribution, pie chart, vendor, op unit, and
        # aging-derived drilldowns consistent.
        canonical_open_sos = []
        total_open_so_amount = 0.0
        monthly = {}
        vendor_map = {}
        op_unit_map = {}
        status_map = {}
        monthly_by_status = {}
        all_months_set = set()

        for s in open_so_rows:
            if s.so_item in hidden_so or s.so_number in hidden_so:
                continue
            if not so_is_countable(s.so_item, customer_po_number=s.customer_po_number, delivery_memo=s.delivery_memo):
                continue

            canonical_open_sos.append(s)
            amount = float(s.sales_amount or 0)
            total_open_so_amount += amount

            if s.so_create_date:
                month_key = s.so_create_date.strftime('%b %Y')
                if month_key not in monthly:
                    monthly[month_key] = {
                        'month': month_key,
                        'so_count': 0,
                        'amount': 0.0,
                        '_s': s.so_create_date.replace(day=1)
                    }
                monthly[month_key]['so_count'] += 1
                monthly[month_key]['amount'] += round(amount / 1_000_000, 2)
                all_months_set.add((s.so_create_date.replace(day=1), month_key))
            else:
                month_key = None

            if s.vendor_name:
                if s.vendor_name not in vendor_map:
                    vendor_map[s.vendor_name] = {'vendor': s.vendor_name, 'so_count': 0, 'total_amount': 0.0}
                vendor_map[s.vendor_name]['so_count'] += 1
                vendor_map[s.vendor_name]['total_amount'] += amount

            if s.operation_unit_name:
                if s.operation_unit_name not in op_unit_map:
                    op_unit_map[s.operation_unit_name] = {
                        'op_unit': s.operation_unit_name,
                        'so_count': 0,
                        'total_amount': 0.0
                    }
                op_unit_map[s.operation_unit_name]['so_count'] += 1
                op_unit_map[s.operation_unit_name]['total_amount'] += amount

            status_name = s.so_status or 'Unknown'
            if status_name not in status_map:
                status_map[status_name] = {'name': status_name, 'value': 0, 'amount': 0.0}
            status_map[status_name]['value'] += 1
            status_map[status_name]['amount'] += amount

            if status_name not in monthly_by_status:
                monthly_by_status[status_name] = {'monthly': {}, 'total': 0, 'amount': 0.0}
            monthly_by_status[status_name]['total'] += 1
            monthly_by_status[status_name]['amount'] += amount
            if month_key:
                monthly_by_status[status_name]['monthly'][month_key] = (
                    monthly_by_status[status_name]['monthly'].get(month_key, 0) + 1
                )

        monthly_trend = sorted(monthly.values(), key=lambda x: x['_s'])
        for m in monthly_trend:
            del m['_s']

        top_vendors = sorted(
            [{'vendor': v['vendor'], 'so_count': v['so_count'], 'total_amount': round(v['total_amount'], 2)}
             for v in vendor_map.values()],
            key=lambda x: x['total_amount'],
            reverse=True
        )[:5]

        top_op_units = sorted(
            [{'op_unit': v['op_unit'], 'so_count': v['so_count'], 'total_amount': round(v['total_amount'], 2)}
             for v in op_unit_map.values()],
            key=lambda x: x['total_amount'],
            reverse=True
        )[:10]

        total_open_for_pct = total_so_count or 1
        so_status = sorted(
            [{'name': v['name'], 'value': v['value'],
              'percentage': round(v['value'] / total_open_for_pct * 100, 1),
              'amount': round(v['amount'], 2)}
             for v in status_map.values()],
            key=lambda x: x['value'],
            reverse=True
        )

        sorted_months = [mk for _, mk in sorted(all_months_set)]
        so_status_monthly = sorted(
            [{'name': st, 'monthly': d['monthly'], 'total': d['total'],
              'percentage': round(d['total'] / total_open_for_pct * 100, 1),
              'amount': round(d['amount'], 2)}
             for st, d in monthly_by_status.items()],
            key=lambda x: x['total'], reverse=True
        )

        po_date_range = db.session.query(func.min(POData.po_date), func.max(POData.po_date)).first()
        so_date_range = db.session.query(func.min(SOData.so_create_date), func.max(SOData.so_create_date)).first()

        # Last updated: latest successful upload timestamp per source file.
        last_po_upload = db.session.query(func.max(UploadLog.uploaded_at)).filter(UploadLog.file_type == 'PO').scalar()
        last_so_upload = db.session.query(func.max(UploadLog.uploaded_at)).filter(UploadLog.file_type == 'SO').scalar()
        if not last_po_upload:
            last_po_upload = db.session.query(func.max(POData.uploaded_at)).scalar()
        if not last_so_upload:
            last_so_upload = db.session.query(func.max(SOData.uploaded_at)).scalar()
        candidates = [x for x in [last_po_upload, last_so_upload] if x]
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
            'last_updated_po': utc_isoformat(last_po_upload),
            'last_updated_smro': utc_isoformat(last_so_upload),
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
    
    Checks both Customer PO Number (primary) and Delivery Memo (secondary) fields.
    Any PO HLI number found in either field counts as matched.
    """
    matched = set()
    # Only load the four columns we actually need — avoids fetching every field
    for row in db.session.query(
            SOData.customer_po_number, SOData.delivery_memo,
            SOData.so_item, SOData.operation_unit_name).all():
        cust_po, memo, so_item, op_unit = row
        # Skip excluded op units
        if op_unit in EXCLUDED_OP_UNITS:
            continue
        # Skip return items (SO Item starting with 9)
        if is_return_so_item(so_item):
            continue
        # Extract PO references from BOTH Customer PO Number and Delivery Memo
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


@app.route('/api/debug/smro-columns', methods=['POST'])
def debug_smro_columns():
    """Inspect column names of an uploaded SMRO file without saving anything.
    Returns all column names and which ones were detected as spec/pid/so_item."""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl', nrows=3)
        df.columns = [str(c).strip() for c in df.columns]
        all_cols = df.columns.tolist()

        detected = {
            'col_so':      find_column(df, ['SO Number','SO No','SO No.','SO','SO Item','Sales Order Number','No SO','Nomor SO']),
            'col_soitem':  find_column(df, ['SO Item No','Item No','Line','SO Line','SO Item']),
            'col_spec':    find_column(df, ['Specification','Spec','Specifications','Product Specification','Material Description','Material Desc','Short Text']),
            'col_pid':     find_column(df, ['Product ID','Product Id','Product Code','Material','Material No','Material Number','Material Code','SKU','Article','Article Number']),
            'col_prod':    find_column(df, ['Product Name','Item Name','Description','Product']),
            'col_status':  find_column(df, ['SO Status','Status','Order Status']),
            'col_vendor':  find_column(df, ['Vendor Name','Vendor','Supplier']),
            'col_sodate':  find_column(df, ['SO Create Date','Order Date','SO Date','Create Date']),
        }

        missing_critical = [k for k in ('col_spec', 'col_pid') if not detected[k]]

        return jsonify({
            'total_columns': len(all_cols),
            'all_columns': all_cols,
            'detected': detected,
            'missing_critical': missing_critical,
            'diagnosis': (
                'col_spec and col_pid NOT detected — column names in this file do not match any known alias. '
                'Check "all_columns" list and update backend aliases.'
                if missing_critical else
                'col_spec and col_pid both detected — upload should populate Specification and Product ID correctly.'
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
        page = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 10))
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
        # Keep backward compatibility: existing dashboard calls this endpoint
        # without pagination params and expects a plain array. When pagination
        # params are supplied, default/start rows per page is 10.
        if 'page' in request.args or 'per_page' in request.args:
            total = len(result)
            start = (page - 1) * per_page
            end = page * per_page
            return jsonify({
                'data': result[start:end],
                'total': total,
                'page': page,
                'per_page': per_page,
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
                open_so_filter()
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
            if is_return_so_status(s.so_status):
                continue
            if s.operation_unit_name in EXCLUDED_OP_UNITS:
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
      - includes all operation units, including HLI GREEN POWER (CONSUMABLE)
      - excludes hidden SO items
      - includes return SO/statuses
      - excludes internal PO refs
      - SO records without so_create_date are bucketed as '180+' (not silently dropped)
    """
    try:
        today = date.today()
        hidden_so = get_hidden_so_items()
        vendors = {}

        for s in db.session.query(SOData).filter(
            open_so_filter()
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
            SOData.vendor_name == vendor_name
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
            open_so_filter()
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
    """Paginated SO list with filters."""
    try:
        page = int(request.args.get('page', 1))
        per_page = int(request.args.get('per_page', 10))
        op_units = request.args.getlist('op_unit')
        vendors = request.args.getlist('vendor')
        statuses = request.args.getlist('status')
        aging_list = request.args.getlist('aging')
        so_items = request.args.getlist('so_item')
        pics = request.args.getlist('pic')
        margin_filter = request.args.get('margin_filter', 'all')
        sort_order = request.args.get('sort_order', 'newest')  # 'newest' or 'oldest'
        date_year, date_from, date_to = parse_so_date_args()

        q = SOData.query.filter(open_so_filter())
        if op_units: q = q.filter(SOData.operation_unit_name.in_(op_units))
        if vendors: q = q.filter(SOData.vendor_name.in_(vendors))
        if statuses: q = q.filter(SOData.so_status.in_(statuses))
        if so_items: q = q.filter(SOData.so_item.in_(so_items))
        if pics:
            if '(Kosong)' in pics:
                others = [p for p in pics if p != '(Kosong)']
                if others:
                    q = q.filter(db.or_(SOData.pic_name.in_(others), SOData.pic_name.is_(None), SOData.pic_name == ''))
                else:
                    q = q.filter(db.or_(SOData.pic_name.is_(None), SOData.pic_name == ''))
            else:
                q = q.filter(SOData.pic_name.in_(pics))

        # SO Create Date filter
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
        else:
            if date_from:
                q = q.filter(SOData.so_create_date >= date_from)
            if date_to:
                q = q.filter(SOData.so_create_date <= date_to)

        # Apply sort order (deterministic by SO Create Date, then SO Item).
        if sort_order == 'oldest':
            all_sos = q.order_by(SOData.so_create_date.asc(), SOData.so_item.asc()).all()
        else:  # newest
            all_sos = q.order_by(SOData.so_create_date.desc(), SOData.so_item.asc()).all()

        # Keep Open SO table count aligned with dashboard total_so_count KPI:
        # exclude hidden SO rows and internal/HLI-referenced rows.
        hidden_so = get_hidden_so_items()
        all_sos = [
            s for s in all_sos
            if s.so_item not in hidden_so
            and s.so_number not in hidden_so
            and so_is_countable(
                s.so_item,
                customer_po_number=s.customer_po_number,
                delivery_memo=s.delivery_memo
            )
        ]

        if aging_list:
            today = date.today()
            def matches_aging(s):
                age = workdays_since(s.so_create_date, today)
                return get_aging_label(age) in aging_list
            all_sos = [s for s in all_sos if matches_aging(s)]

        if margin_filter in ('positive', 'negative'):
            # Warm cache before filtering loop to avoid per-row HTTP calls.
            usd_sos = [s for s in all_sos if (s.purchasing_currency or '').strip().upper() == 'USD']
            if usd_sos:
                prefetch_exchange_rates({s.so_create_date for s in usd_sos if s.so_create_date})

            def calc_margin(s):
                po_amt = convert_to_idr((s.purchasing_amount or 0) or (s.purchasing_price or 0) * (s.so_qty or 0), s.purchasing_currency, s.so_create_date, cache_only=True)
                return float(s.sales_amount or 0) - po_amt
            if margin_filter == 'negative':
                all_sos = [s for s in all_sos if calc_margin(s) < 0]
            else:
                all_sos = [s for s in all_sos if calc_margin(s) >= 0]

        approval_statuses = {'Approval Apply', 'Approval Complete Step', 'Approval Reject'}
        approval_q = SOData.query.filter(SOData.so_status.in_(list(approval_statuses)))
        if op_units: approval_q = approval_q.filter(SOData.operation_unit_name.in_(op_units))
        if vendors: approval_q = approval_q.filter(SOData.vendor_name.in_(vendors))
        if statuses: approval_q = approval_q.filter(SOData.so_status.in_(statuses))
        if so_items: approval_q = approval_q.filter(SOData.so_item.in_(so_items))
        approval_q = apply_so_create_date_filter(approval_q, date_year, date_from, date_to, is_sqlite)
        if sort_order == 'oldest':
            approval_sos = approval_q.order_by(SOData.so_create_date.asc(), SOData.so_item.asc()).all()
        else:
            approval_sos = approval_q.order_by(SOData.so_create_date.desc(), SOData.so_item.asc()).all()
        approval_sos = [
            s for s in approval_sos
            if s.so_item not in hidden_so
            and s.so_number not in hidden_so
            and so_is_countable(
                s.so_item,
                customer_po_number=s.customer_po_number,
                delivery_memo=s.delivery_memo
            )
        ]

        total = len(all_sos)
        subtotal_amount = sum(float(s.sales_amount or 0) for s in all_sos)
        paged = all_sos[(page-1)*per_page : page*per_page]

        op_units_opts = sorted({s.operation_unit_name for s in all_sos if s.operation_unit_name})
        vendors_opts  = sorted({s.vendor_name for s in all_sos if s.vendor_name})
        statuses_opts = sorted({s.so_status for s in all_sos if s.so_status})
        pics_opts     = sorted({s.pic_name for s in all_sos if s.pic_name})

        return jsonify({
            'data': [so_dict(s) for s in paged],
            'approval_data': [so_dict(s) for s in approval_sos],
            'total': total, 'subtotal_amount': round(subtotal_amount, 2), 'page': page, 'per_page': per_page,
            'filters': {'op_units': list(op_units_opts), 'vendors': list(vendors_opts), 'statuses': list(statuses_opts), 'pics': list(pics_opts)}
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
# EXCHANGE RATE ENDPOINTS
# ═══════════════════════════════════════════════════════════════════

@app.route('/api/exchange-rate', methods=['GET'])
def list_exchange_rates():
    """Return all stored USD->IDR rates, newest first."""
    try:
        rates = ExchangeRate.query.order_by(ExchangeRate.rate_date.desc()).limit(120).all()
        return jsonify([{
            'id': r.id, 'date': r.rate_date.isoformat(),
            'usd_to_idr': r.usd_to_idr, 'source': r.source,
        } for r in rates])
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/exchange-rate', methods=['POST'])
def upsert_exchange_rate():
    """Manually set or update a USD->IDR rate for a specific date."""
    try:
        data = request.json
        d = parse_date(data.get('date'))
        rate = float(data.get('usd_to_idr', 0))
        if not d:
            return jsonify({'error': 'Invalid date'}), 400
        if rate <= 0:
            return jsonify({'error': 'Rate must be > 0'}), 400
        rec = ExchangeRate.query.filter_by(rate_date=d).first()
        if rec:
            rec.usd_to_idr = rate
            rec.source = 'manual'
        else:
            rec = ExchangeRate(rate_date=d, usd_to_idr=rate, source='manual')
            db.session.add(rec)
        db.session.commit()
        # Invalidate cache for this date
        _RATE_CACHE.pop(d, None)
        return jsonify({'success': True, 'date': d.isoformat(), 'usd_to_idr': rate})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500


@app.route('/api/exchange-rate/fetch', methods=['POST'])
def fetch_exchange_rates_bulk():
    """Auto-fetch USD->IDR rates from Frankfurter API for all SO create dates
    that have USD purchasing currency but no rate stored yet.
    Returns count of rates fetched."""
    try:
        # Find distinct SO create dates where purchasing_currency = USD and no rate stored
        is_sqlite = 'sqlite' in app.config['SQLALCHEMY_DATABASE_URI']
        usd_rows = db.session.query(SOData.so_create_date).filter(
            SOData.purchasing_currency == 'USD',
            SOData.so_create_date.isnot(None)
        ).distinct().all()

        dates_needed = {r[0] for r in usd_rows}
        existing_dates = {r[0] for r in db.session.query(ExchangeRate.rate_date).all()}
        to_fetch = sorted(dates_needed - existing_dates)

        fetched = 0
        failed = []
        for d in to_fetch:
            rate = _fetch_rate_from_api(d)
            if rate:
                try:
                    db.session.add(ExchangeRate(rate_date=d, usd_to_idr=rate, source='frankfurter'))
                    db.session.flush()
                    _RATE_CACHE[d] = rate
                    fetched += 1
                except Exception:
                    db.session.rollback()
            else:
                failed.append(d.isoformat())

        db.session.commit()
        return jsonify({
            'dates_needed': len(dates_needed),
            'already_stored': len(existing_dates & dates_needed),
            'fetched': fetched,
            'failed': failed,
            'message': f'{fetched} kurs berhasil di-fetch dari Frankfurter API.'
                       + (f' {len(failed)} tanggal gagal: {", ".join(failed[:5])}' if failed else '')
        })
    except Exception as e:
        db.session.rollback()
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/exchange-rate/preview', methods=['GET'])
def preview_exchange_rate():
    """Preview what rate would be used for a given date (for debugging)."""
    try:
        d = parse_date(request.args.get('date', ''))
        if not d:
            return jsonify({'error': 'Provide ?date=YYYY-MM-DD'}), 400
        rate = get_usd_to_idr(d)
        rec = ExchangeRate.query.filter_by(rate_date=d).first()
        return jsonify({
            'date': d.isoformat(),
            'usd_to_idr': rate,
            'source': rec.source if rec else 'fallback/nearest',
            'stored_exact': rec is not None,
        })
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
        col_spec    = find_column(df, ['Specification','Spec','Specifications','Product Specification','Material Description','Material Desc','Short Text'])
        col_pid     = find_column(df, ['Product ID','Product Id','Product Code','Material','Material No','Material Number','Material Code','SKU','Article','Article Number'])
        col_qty     = find_column(df, ['SO Quantity','SO Qty','Qty','Quantity'])
        col_sunit   = find_column(df, ['Sales Unit','Unit','UOM'])
        col_sprice  = find_column(df, ['Sales Price(Exclude Tax)','Sales Price','Price','Unit Price'])
        col_samt    = find_column(df, ['Sales Amount(Exclude Tax)','Sales Amount','Amount','Total'])
        col_cur     = find_column(df, ['Currency','Curr'])
        col_pprice  = find_column(df, ['Purchasing Price','Purchase Price','PO Price'])
        col_pamt    = find_column(df, ['Purchasing Amount','Purchase Amount','PO Amount'])
        col_pcur    = find_column(df, ['Purchasing Currency','Purchase Currency','PO Currency','Purchasing Curr','Purchase Curr'])
        col_sodate  = find_column(df, ['SO Create Date','Order Date','SO Date','Create Date'])
        col_delposs = find_column(df, ['Delivery Possible Date','Possible Delivery Date','Est Delivery'])
        col_matchpo = find_column(df, ['Matched PO Number','Matched PO','PO HLI','PO HLI Number','Purchasing Order Number','PO Number'])

        # Build lookup of existing SO records by so_item
        existing_so = {}
        for s in SOData.query.all():
            if s.so_item:
                existing_so[s.so_item] = s

        count = 0
        updated = 0
        inserted = 0

        # Track how many rows had non-empty Specification / Product ID values
        # so the user can see whether the upload actually carried that data.
        spec_filled = 0
        pid_filled  = 0

        for _, row in df.iterrows():
            so_val = clean(df_val(row, col_so))
            if not so_val: continue
            so_item_val = clean(df_val(row, col_soitem))

            spec_val = clean(df_val(row, col_spec)) if col_spec else None
            pid_val  = clean(df_val(row, col_pid))  if col_pid  else None
            if spec_val: spec_filled += 1
            if pid_val:  pid_filled  += 1

            new_data = {
                'so_number': so_val,
                'so_item': so_item_val,
                'so_status': clean(df_val(row, col_status)),
                'operation_unit_name': clean(df_val(row, col_opunit)),
                'vendor_name': clean(df_val(row, col_vendor)),
                'customer_po_number': clean(df_val(row, col_custpo)),
                'delivery_memo': clean(df_val(row, col_memo)),
                'product_name': clean(df_val(row, col_prod)),
                'specification': spec_val,
                'product_id': pid_val,
                'so_qty': safe_float(df_val(row, col_qty)),
                'sales_unit': clean(df_val(row, col_sunit)),
                'sales_price': safe_float(df_val(row, col_sprice)),
                'sales_amount': safe_float(df_val(row, col_samt)),
                'currency': clean(df_val(row, col_cur)) or 'IDR',
                'purchasing_price': safe_float(df_val(row, col_pprice)),
                'purchasing_amount': safe_float(df_val(row, col_pamt)),
                'purchasing_currency': clean(df_val(row, col_pcur)) if col_pcur else None,
                'purchasing_amount_idr': None,
                'purchasing_amount_idr_cached_at': None,
                'so_create_date': parse_date(df_val(row, col_sodate)),
                'delivery_possible_date': parse_date(df_val(row, col_delposs)),
                'matched_po_number': clean(df_val(row, col_matchpo)),
                'uploaded_at': datetime.utcnow()
            }

            if so_item_val and so_item_val in existing_so:
                # Update existing record — preserve remarks & delivery_plan_date.
                # Also preserve specification / product_id if either:
                #  (a) the uploaded file doesn't have that column at all, or
                #  (b) the row's value is blank in the file.
                # This protects spec/pid that was previously populated via the
                # SMRO upload itself or via the backfill endpoint.
                existing = existing_so[so_item_val]
                preserved_remarks   = existing.remarks
                preserved_plan_date = existing.delivery_plan_date
                preserved_spec      = existing.specification
                preserved_pid       = existing.product_id
                preserved_amount_idr = existing.purchasing_amount_idr
                preserved_amount_idr_cached_at = existing.purchasing_amount_idr_cached_at
                old_purchase_signature = (
                    float(existing.purchasing_amount or 0),
                    float(existing.purchasing_price or 0),
                    float(existing.so_qty or 0),
                    (existing.purchasing_currency or 'IDR').strip().upper(),
                    existing.so_create_date,
                )
                new_purchase_signature = (
                    float(new_data.get('purchasing_amount') or 0),
                    float(new_data.get('purchasing_price') or 0),
                    float(new_data.get('so_qty') or 0),
                    (new_data.get('purchasing_currency') or 'IDR').strip().upper(),
                    new_data.get('so_create_date'),
                )
                purchase_inputs_changed = old_purchase_signature != new_purchase_signature
                for field, val in new_data.items():
                    setattr(existing, field, val)
                existing.remarks = preserved_remarks
                existing.delivery_plan_date = preserved_plan_date
                if not purchase_inputs_changed:
                    existing.purchasing_amount_idr = preserved_amount_idr
                    existing.purchasing_amount_idr_cached_at = preserved_amount_idr_cached_at
                if not col_spec or spec_val is None:
                    existing.specification = preserved_spec
                if not col_pid or pid_val is None:
                    existing.product_id = preserved_pid
                # Auto-fill pic_name from ProductIDDB + MasterPIC
                _pid_for_pic = existing.product_id
                if _pid_for_pic:
                    existing.pic_name = _lookup_pic(_pid_for_pic)
                updated += 1
            else:
                new_rec = SOData(**new_data)
                # Auto-fill pic_name on insert
                _pid_for_pic = new_data.get('product_id')
                if _pid_for_pic:
                    new_rec.pic_name = _lookup_pic(_pid_for_pic)
                db.session.add(new_rec)
                inserted += 1

            count += 1
            if count % CHUNK_SIZE == 0:
                db.session.flush()

        # ── KEY CHANGE: Do NOT delete records not in this file ──
        # Old records with different SO Items are preserved as-is.

        db.session.add(UploadLog(file_type='SO', filename=file.filename, records_count=count))
        db.session.commit()

        # Diagnostic block — surfaces which Spec / Product ID columns were
        # detected so the user can immediately tell whether the upload
        # carried those values.  Returned even on success.
        diagnostics = {
            'columns_detected': {
                'specification': col_spec,
                'product_id':    col_pid,
            },
            'rows_with_specification': spec_filled,
            'rows_with_product_id':    pid_filled,
            'all_file_columns':        df.columns.tolist(),
        }
        # Build the warning by *accumulating* every condition that applies,
        # so we never silently drop a Product ID warning when Specification
        # also has an issue (or vice versa).
        warnings = []
        if not col_spec and not col_pid:
            warnings.append(
                "File ini tidak mengandung kolom 'Specification' maupun 'Product ID' "
                "(atau alias yang dikenal). Spec/Product ID di DB tidak diubah."
            )
        else:
            if not col_spec:
                warnings.append("Kolom 'Specification' tidak ditemukan di file ini — Specification di DB dipertahankan.")
            elif spec_filled == 0:
                warnings.append(f"Kolom '{col_spec}' terdeteksi tapi semua baris kosong.")
            if not col_pid:
                warnings.append("Kolom 'Product ID' tidak ditemukan di file ini — Product ID di DB dipertahankan.")
            elif pid_filled == 0:
                warnings.append(f"Kolom '{col_pid}' terdeteksi tapi semua baris kosong.")
        if warnings:
            diagnostics['warning'] = ' '.join(warnings)

        return jsonify({
            'message': f'Berhasil: {inserted} SO baru ditambahkan, {updated} SO diperbarui. Data lama yang tidak ada di file ini tetap dipertahankan.',
            'uploaded': count,
            'inserted': inserted,
            'updated': updated,
            'diagnostics': diagnostics,
        })
    except Exception as e:
        db.session.rollback(); import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/upload/smro-backfill-spec', methods=['POST'])
def upload_smro_backfill_spec():
    """Backfill-only upload: reads Specification and Product ID from an SMRO
    Excel file and updates those two fields on existing SO records.

    Matching strategy (in order):
    1. Exact SO Item match  (e.g. '9008123456-10' in file == '9008123456-10' in DB)
    2. SO Number + item-line suffix match  (file SO Item '9008123456-10' →
       SO Number '9008123456', item line '10')
    3. SO Number only match (last resort — updates all lines of that SO)

    Safe to run multiple times.  Only writes when spec or pid is non-null in file.
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        col_sonum  = find_column(df, ['SO Number', 'SO No', 'SO No.', 'SO'])
        col_soitem = find_column(df, ['SO Item', 'SO Item No', 'SO Line', 'Item No', 'Line'])
        col_spec   = find_column(df, ['Specification', 'Spec', 'Specifications', 'Product Specification'])
        col_pid    = find_column(df, ['Product ID', 'Product Id', 'Product Code',
                                      'Material', 'Material No', 'Material Number', 'Material Code', 'SKU'])

        if not col_soitem and not col_sonum:
            return jsonify({'error': f'SO Item / SO Number column not found. Columns: {df.columns.tolist()}'}), 400
        if not col_spec and not col_pid:
            return jsonify({'error': 'Neither Specification nor Product ID column found.'}), 400

        # Ensure columns exist in DB before writing
        _ensure_so_extra_columns()

        # Build lookups from all SO records in DB
        all_so = SOData.query.all()
        # Primary: exact so_item match
        by_soitem = {}
        # Secondary: so_number → list of records (for fallback)
        by_sonum  = {}
        for s in all_so:
            if s.so_item:
                by_soitem[s.so_item] = s
                # Also index by so_number+item_suffix e.g. '9008123456' + '-10'
                parts = s.so_item.rsplit('-', 1)
                if len(parts) == 2:
                    by_soitem[s.so_item] = s  # already done
            if s.so_number:
                by_sonum.setdefault(s.so_number, []).append(s)

        updated = 0
        skipped_no_match = 0
        skipped_no_data  = 0
        flush_counter    = 0

        for _, row in df.iterrows():
            so_item_val = clean(df_val(row, col_soitem)) if col_soitem else None
            so_num_val  = clean(df_val(row, col_sonum))  if col_sonum  else None
            spec_val    = clean(df_val(row, col_spec))   if col_spec   else None
            pid_val     = clean(df_val(row, col_pid))    if col_pid    else None

            if spec_val is None and pid_val is None:
                skipped_no_data += 1
                continue

            # Resolve matching DB records
            matched_recs = []

            if so_item_val:
                # Try exact so_item match first
                rec = by_soitem.get(so_item_val)
                if rec:
                    matched_recs = [rec]
                else:
                    # Try: maybe DB stores so_number only as so_item (older uploads)
                    # Extract so_number from so_item (strip '-NN' suffix)
                    parts = so_item_val.rsplit('-', 1)
                    so_num_from_item = parts[0] if len(parts) == 2 else so_item_val
                    candidates = by_sonum.get(so_num_from_item, [])
                    if len(parts) == 2:
                        # Try to match by item line number suffix
                        item_line = parts[1]
                        line_matched = [
                            c for c in candidates
                            if c.so_item and c.so_item.endswith(f'-{item_line}')
                        ]
                        matched_recs = line_matched or candidates
                    else:
                        matched_recs = candidates

            if not matched_recs and so_num_val:
                matched_recs = by_sonum.get(so_num_val, [])

            if not matched_recs:
                skipped_no_match += 1
                continue

            for rec in matched_recs:
                changed = False
                if spec_val is not None and rec.specification != spec_val:
                    rec.specification = spec_val
                    changed = True
                if pid_val is not None and rec.product_id != pid_val:
                    rec.product_id = pid_val
                    changed = True
                if changed:
                    updated += 1
                    flush_counter += 1
                    if flush_counter % 300 == 0:
                        db.session.flush()

        db.session.commit()
        return jsonify({
            'message': (
                f'Backfill selesai: {updated} SO record diperbarui'
                + (f', {skipped_no_match} baris tidak cocok di DB' if skipped_no_match else '')
                + (f', {skipped_no_data} baris tidak ada data Spec/PID' if skipped_no_data else '')
                + '.'
            ),
            'updated': updated,
            'skipped_no_match': skipped_no_match,
            'skipped_no_data': skipped_no_data,
            'spec_column_detected': col_spec,
            'pid_column_detected': col_pid,
            'soitem_column_detected': col_soitem,
        })
    except Exception as e:
        db.session.rollback()
        import traceback; traceback.print_exc()
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


@app.route('/api/data/po/<int:po_id>', methods=['PUT'])
def update_po(po_id):
    try:
        data = request.json
        po = db.session.get(POData, po_id)
        if not po:
            return jsonify({'error': 'Not found'}), 404
        if 'delivery_plan_date' in data:
            po.delivery_plan_date = parse_date(data['delivery_plan_date'])
        if 'remarks' in data:
            po.remarks = data['remarks']
        db.session.commit()
        return jsonify({'success': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/po/template', methods=['GET'])
def download_po_batch_template():
    """Download Excel template for PO HLI Without SO editable fields."""
    try:
        matched_set = build_matched_set()
        hidden_po = get_hidden_po_numbers()
        wb = Workbook()
        ws = wb.active
        ws.title = "PO Batch Upload"
        headers = ['PO HLI Number', 'Delivery Plan Date', 'Remarks']
        ws.append(headers)

        header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        widths = [35, 25, 60]
        for i, cell in enumerate(ws[1], 1):
            cell.fill = header_fill
            cell.font = Font(bold=True, color="000000")
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]

        ws.append(['example : 4502358819-10', 'example : 2025-12-31', 'example : Waiting for vendor confirmation'])
        grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        red_font = Font(color="FF0000")
        for cell in ws[2]:
            cell.font = red_font
            cell.fill = grey_fill

        seen_keys = set()
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
                po_key = po_hli_key(p.po_number, p.item_no) or p.po_number
                ws.append([
                    po_key,
                    p.delivery_plan_date.isoformat() if p.delivery_plan_date else '',
                    p.remarks or ''
                ])

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"Template_PO_BatchUpload_{datetime.now().strftime('%Y%m%d')}.xlsx")
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/po/batch-upload', methods=['POST'])
def batch_upload_po():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file'}), 400
        file = request.files['file']
        df = pd.read_excel(file, engine='openpyxl', skiprows=[1])
        df.columns = [str(c).strip() for c in df.columns]
        col_po = find_column(df, ['PO HLI Number', 'NO PO HLI', 'PO HLI', 'PO Number-Item No', 'PO Number'])
        col_plan = find_column(df, ['Delivery Plan Date', 'Plan Date'])
        col_rem = find_column(df, ['Remarks', 'Remark'])
        if not col_po:
            return jsonify({'error': f'Column "PO HLI Number" not found. Available: {df.columns.tolist()}'}), 400

        updated = 0
        not_found = 0
        for _, row in df.iterrows():
            ref = clean(df_val(row, col_po))
            if not ref or ref.lower().startswith('example'):
                continue
            ref = ref.replace('example :', '').replace('example:', '').strip()
            po_num, item_no = ref, None
            if '-' in ref:
                po_num, item_no = ref.split('-', 1)
                po_num = po_num.strip()
                item_no = item_no.strip()
            candidates = POData.query.filter_by(po_number=po_num).all()
            if item_no:
                item_variants = _normalize_item_no(item_no)
                candidates = [p for p in candidates if any(str(p.item_no or '').strip() == v for v in item_variants)]
            if not candidates:
                not_found += 1
                continue
            for po in candidates:
                if col_plan:
                    d = parse_date(df_val(row, col_plan))
                    if d or clean(df_val(row, col_plan)) == '':
                        po.delivery_plan_date = d
                if col_rem:
                    po.remarks = clean(df_val(row, col_rem)) or ''
                updated += 1
        db.session.commit()
        return jsonify({'updated': updated, 'not_found': not_found})
    except Exception as e:
        db.session.rollback()
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/so/template', methods=['GET'])
def download_so_batch_template():
    """Download Excel template for SO batch upload.
    If filters are supplied (same params as /api/data/all-so), the filtered SO Items
    are pre-populated in the template starting from row 3, with existing
    delivery_plan_date and remarks already filled in so the user only needs to
    update what changed.  Either column may be left blank on upload.
    """
    try:
        # ── Apply same filters as fetchSOData ──────────────────────────────
        op_units   = request.args.getlist('op_unit')
        vendors    = request.args.getlist('vendor')
        statuses   = request.args.getlist('status')
        aging_list = request.args.getlist('aging')
        so_items   = request.args.getlist('so_item')
        margin_filter = request.args.get('margin_filter', 'all')
        date_year, date_from, date_to = parse_so_date_args()

        q = SOData.query.filter(open_so_filter())
        if op_units:  q = q.filter(SOData.operation_unit_name.in_(op_units))
        if vendors:   q = q.filter(SOData.vendor_name.in_(vendors))
        if statuses:  q = q.filter(SOData.so_status.in_(statuses))
        if so_items:  q = q.filter(SOData.so_item.in_(so_items))
        q = apply_so_create_date_filter(q, date_year, date_from, date_to)
        all_sos = q.order_by(SOData.so_create_date.asc()).all()

        # Aging filter (post-query, same as all-so endpoint)
        if aging_list:
            today = date.today()
            def matches_aging(s):
                return get_aging_label(workdays_since(s.so_create_date, today)) in aging_list
            all_sos = [s for s in all_sos if matches_aging(s)]

        # Margin filter
        if margin_filter in ('positive', 'negative'):
            # Warm cache before filtering loop to avoid per-row HTTP calls.
            usd_sos = [s for s in all_sos if (s.purchasing_currency or '').strip().upper() == 'USD']
            if usd_sos:
                prefetch_exchange_rates({s.so_create_date for s in usd_sos if s.so_create_date})

            def calc_margin(s):
                po_amt = convert_to_idr((s.purchasing_amount or 0) or (s.purchasing_price or 0) * (s.so_qty or 0), s.purchasing_currency, s.so_create_date, cache_only=True)
                return float(s.sales_amount or 0) - po_amt
            if margin_filter == 'negative':
                all_sos = [s for s in all_sos if calc_margin(s) < 0]
            else:
                all_sos = [s for s in all_sos if calc_margin(s) >= 0]

        # ── Build workbook ─────────────────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.title = "SO Batch Upload"

        headers = ['SO Item', 'Delivery Plan Date', 'Remarks']
        ws.append(headers)

        # Row 1: header — yellow, bold, centered
        header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        col_widths   = [35, 25, 50]
        for i, cell in enumerate(ws[1], 1):
            cell.fill = header_fill
            cell.font = Font(bold=True, color="000000")
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(i)].width = col_widths[i - 1]

        # Row 2: example — red font, light grey background
        grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        red_font  = Font(color="FF0000")
        ws.append(['example : 9008988017-10', 'example : 2025-12-31', 'example : Waiting for vendor confirmation'])
        for cell in ws[2]:
            cell.font = red_font
            cell.fill = grey_fill

        # Rows 3+: pre-populate with filtered SO items (if any filter active)
        for s in all_sos:
            if not s.so_item:
                continue
            plan = s.delivery_plan_date.isoformat() if s.delivery_plan_date else ''
            ws.append([s.so_item, plan, s.remarks or ''])

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"Template_SO_BatchUpload_{datetime.now().strftime('%Y%m%d')}.xlsx")
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/so/batch-upload', methods=['POST'])
def batch_upload_so():
    try:
        if 'file' not in request.files: return jsonify({'error': 'No file'}), 400
        file = request.files['file']
        # Row 0 = header, row 1 = example (red row) → skip with skiprows so
        # actual data starts at row index 0 of the resulting DataFrame (Excel row 3+).
        df = pd.read_excel(file, engine='openpyxl', skiprows=[1])
        df.columns = [str(c).strip() for c in df.columns]
        col_so_item = find_column(df, ['SO Item', 'SO Item No', 'SO Item Number'])
        col_plan    = find_column(df, ['Delivery Plan Date', 'Plan Date'])
        col_rem     = find_column(df, ['Remarks', 'Remark'])
        if not col_so_item:
            return jsonify({'error': f'Column "SO Item" not found. Available: {df.columns.tolist()}'}), 400
        updated = 0
        not_found = 0
        for _, row in df.iterrows():
            so_item_val = clean(df_val(row, col_so_item)) if col_so_item else None
            if not so_item_val: continue
            # Lookup by so_item (unique identifier) — NOT so_number
            so = SOData.query.filter_by(so_item=so_item_val).first()
            if so:
                if col_plan:
                    d = parse_date(df_val(row, col_plan))
                    if d: so.delivery_plan_date = d
                if col_rem:
                    r = clean(df_val(row, col_rem))
                    if r: so.remarks = r
                updated += 1
            else:
                not_found += 1
        db.session.commit()
        return jsonify({'updated': updated, 'not_found': not_found})
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
        _style_wb(ws, ['Aging','SO Number','SO Item','Status','Op Unit','Vendor','Product','PIC',
                       'SO Qty','Sales Price','Sales Amount','PO Price','PO Amount',
                       'SO Date','Delivery Possible','Customer PO','Delivery Memo',
                       'Delivery Plan Date','Remarks'], num_cols=[9,10,11,12,13])
        for s in sos:
            age = (today - s.so_create_date).days if s.so_create_date else None
            ws.append([get_aging_label(age), s.so_number, s.so_item, s.so_status,
                s.operation_unit_name, s.vendor_name, s.product_name,
                s.pic_name or '',
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
        headers = ['SO Item', 'Reason']
        ws.append(headers)
        # Header: yellow background, bold, centered
        header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for i, cell in enumerate(ws[1], 1):
            cell.fill = header_fill
            cell.font = Font(bold=True, color="000000")
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(i)].width = 35 if i == 1 else 50
        # Row 2: example row in red font
        ws.append(['9008988017-10', 'Reason why this SO Item should be hidden'])
        example_font = Font(color="FF0000")
        for cell in ws[2]:
            cell.font = example_font
    else:
        ws.title = "Hide PO HLI Template"
        headers = ['NO PO HLI (PO Number-Item No)', 'Reason']
        ws.append(headers)
        # Header: yellow background, bold, centered
        header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for i, cell in enumerate(ws[1], 1):
            cell.fill = header_fill
            cell.font = Font(bold=True, color="000000")
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[get_column_letter(i)].width = 35 if i == 1 else 50
        # Row 2: example row in red font
        ws.append(['4502358819-10', 'Reason why this PO HLI should be hidden'])
        example_font = Font(color="FF0000")
        for cell in ws[2]:
            cell.font = example_font

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

        df = pd.read_excel(file, engine='openpyxl', skiprows=[1])
        df.columns = [str(c).strip() for c in df.columns]

        # Detect column names
        if hide_type == 'SO':
            col_ref = find_column(df, ['SO Item', 'SO Item No', 'SO Number', 'SO No', 'SO Number-Item No'])
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
        q = q.filter(~SOData.operation_unit_name.in_(list(EXCLUDED_OP_UNITS)))

        cache_key = (
            year_filter or 'all',
            date_year or '',
            date_from or '',
            date_to or '',
        )
        db_signature = q.with_entities(
            func.count(SOData.id),
            func.max(SOData.id),
            func.max(SOData.purchasing_amount_idr_cached_at),
        ).one()
        cache_entry = _COMPLETED_SUMMARY_CACHE.get(cache_key)
        now_ts = datetime.utcnow().timestamp()
        if (
            cache_entry
            and cache_entry.get('signature') == tuple(db_signature)
            and now_ts - cache_entry.get('created_at', 0) < _COMPLETED_SUMMARY_CACHE_TTL_SECONDS
        ):
            return jsonify(cache_entry['payload'])

        rows = q.all()

        missing_conversion_count = sum(
            1 for s in rows
            if s.purchasing_amount_idr is None
            and str(s.purchasing_currency or 'IDR').strip().upper() != 'IDR'
            and raw_purchase_amount(s) > 0
        )

        # Persist missing converted purchase amounts once. Subsequent page loads
        # reuse so_data.purchasing_amount_idr instead of converting every row.
        converted_count = ensure_purchase_amount_idr_cache(rows)

        def po_amt_of(s):
            return purchase_amount_idr(s)

        # Pre-compute per-row sales/purchase/margin once, then reuse.
        enriched = []
        for s in rows:
            po_amt = po_amt_of(s)
            sales = float(s.sales_amount or 0)
            # Margin is only calculated if purchase price exists and is not 0
            # If purchase price is 0 or missing, margin should be None (displayed as "-")
            has_purchase_data = (
                (s.purchasing_amount is not None and s.purchasing_amount != 0) or
                (s.purchasing_price is not None and s.purchasing_price != 0)
            )
            # Only calculate margin if we have valid purchase data
            margin = (sales - po_amt) if has_purchase_data else None
            enriched.append((s, po_amt, sales, margin))

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
            if m is not None:
                vendor_map[v]['margin'] += m

        top_vendors = sorted(vendor_map.values(), key=lambda x: x['sales_amount'], reverse=True)[:5]

        # Margin distribution + totals (KPI cards)
        pos = neg = zero = 0
        total_sales = 0.0
        total_purchase = 0.0
        for _s, po_amt, sales, m in enriched:
            total_sales += sales
            total_purchase += po_amt
            if m is not None:
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
            if m is not None:
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

        payload = {
            'total_count': len(rows),
            'total_sales': total_sales,
            'total_purchase': total_purchase,
            'total_margin': (total_sales - total_purchase) if (total_sales > 0 and total_purchase > 0) else None,
            'monthly_trend': monthly_trend,
            'top_vendors': top_vendors,
            'top_items': top_items,
            'worst_margin_vendors': worst_margin_vendors,
            'worst_margin_transactions': worst_margin_transactions,
            'margin_distribution': {
                'positive': pos,
                'negative': neg,
                'zero': zero
            },
            'conversion_status': {
                'checked': True,
                'had_missing_cache': missing_conversion_count > 0,
                'converted_count': converted_count,
                'pending_count': max(missing_conversion_count - converted_count, 0),
                'message': (
                    f'Konversi currency selesai dan disimpan untuk {converted_count} data baru.'
                    if converted_count
                    else 'Tidak ada data currency baru yang perlu dikonversi.'
                )
            }
        }
        _COMPLETED_SUMMARY_CACHE[cache_key] = {
            'signature': tuple(db_signature),
            'created_at': now_ts,
            'payload': payload,
        }
        return jsonify(payload)

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

        # Persist missing converted purchase amounts once. Subsequent popup
        # loads reuse the stored IDR value.
        ensure_purchase_amount_idr_cache(rows)

        def get_po_amt(s):
            return purchase_amount_idr(s)

        result = []
        for s in rows:
            po_amt = get_po_amt(s)
            # Check if purchase data exists (not 0 or None)
            has_purchase_data = (
                (s.purchasing_amount is not None and s.purchasing_amount != 0) or
                (s.purchasing_price is not None and s.purchasing_price != 0)
            )
            # Only calculate margin if we have valid purchase data
            m = (float(s.sales_amount or 0) - po_amt) if has_purchase_data else None
            
            # Skip rows without margin data for positive/negative categories
            if m is None and category in ('positive', 'negative'):
                continue
            if category == 'positive' and (m is None or m <= 0):
                continue
            elif category == 'negative' and (m is None or m >= 0):
                continue
            elif category == 'zero' and (m is None or m != 0):
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




# ─── Product ID Database & Master PIC endpoints ───────────────────────────

def _lookup_pic(product_id_str):
    """Return PIC name for a product_id string, or None if not found."""
    if not product_id_str:
        return None
    pid = str(product_id_str).strip()
    prod = db.session.query(ProductIDDB).filter_by(product_id=pid).first()
    if not prod or not prod.category_id:
        return None
    pic = db.session.query(MasterPIC).filter_by(category_id=str(prod.category_id).strip()).first()
    return pic.pic_name if pic else None


@app.route('/api/upload/product-id', methods=['POST'])
def upload_product_id():
    """Upload Prod_ID Excel from SAP. Upserts product_id → category_id mapping."""
    try:
        file = request.files.get('file')
        if not file:
            return jsonify({'error': 'No file provided'}), 400

        # SAP ZMMR3190 kadang generate file BIFF8 (.xls) meski ekstensinya .xlsx.
        # Deteksi format sebenarnya dari magic bytes, bukan dari nama file.
        raw = file.read()
        file.seek(0)

        # Magic bytes: BIFF8/XLS = D0 CF 11 E0 | XLSX (ZIP) = 50 4B 03 04
        is_xls_format = raw[:4] == b'\xd0\xcf\x11\xe0'
        filename = (file.filename or '').lower()

        if is_xls_format:
            engine = 'xlrd'
        elif filename.endswith('.xls'):
            engine = 'xlrd'
        else:
            engine = 'openpyxl'

        df = pd.read_excel(file, sheet_name=0, engine=engine)
        df.columns = [str(c).strip() for c in df.columns]

        pid_col = next((c for c in df.columns if 'Product ID' in c or c.lower() == 'product id'), None)
        cat_col = next((c for c in df.columns if 'Category ID' in c or c.lower() == 'category id'), None)
        catn_col = next((c for c in df.columns if 'Category Name' in c or c.lower() == 'category name'), None)
        pname_col = next((c for c in df.columns if 'Product Name' in c and 'EN' not in c), None)

        if not pid_col or not cat_col:
            return jsonify({'error': f'Missing required columns. Found: {list(df.columns)[:10]}'}), 400

        added = updated = 0
        pic_cache = {}  # category_id → pic_name

        for _, row in df.iterrows():
            pid = str(row[pid_col]).strip() if pd.notna(row[pid_col]) else None
            cat_id = str(row[cat_col]).strip() if pd.notna(row[cat_col]) else None
            if not pid or pid == 'nan':
                continue
            cat_name = str(row[catn_col]).strip() if catn_col and pd.notna(row[catn_col]) else None
            pname = str(row[pname_col]).strip() if pname_col and pd.notna(row[pname_col]) else None

            existing = db.session.query(ProductIDDB).filter_by(product_id=pid).first()
            if existing:
                existing.category_id = cat_id
                existing.category_name = cat_name
                existing.product_name = pname
                existing.updated_at = datetime.utcnow()
                updated += 1
            else:
                db.session.add(ProductIDDB(
                    product_id=pid, category_id=cat_id,
                    category_name=cat_name, product_name=pname,
                    updated_at=datetime.utcnow()
                ))
                added += 1

        db.session.commit()

        # After upserting ProductIDDB, refresh pic_name on SO rows that have a product_id
        so_rows = db.session.query(SOData).filter(
            SOData.product_id.isnot(None), SOData.product_id != ''
        ).all()
        refreshed = 0
        for s in so_rows:
            cat_id_key = None
            prod = db.session.query(ProductIDDB).filter_by(product_id=str(s.product_id).strip()).first()
            if prod and prod.category_id:
                cat_id_key = str(prod.category_id).strip()
            if cat_id_key:
                if cat_id_key not in pic_cache:
                    pic_obj = db.session.query(MasterPIC).filter_by(category_id=cat_id_key).first()
                    pic_cache[cat_id_key] = pic_obj.pic_name if pic_obj else None
                new_pic = pic_cache[cat_id_key]
                if s.pic_name != new_pic:
                    s.pic_name = new_pic
                    refreshed += 1
        db.session.commit()

        return jsonify({
            'status': 'ok',
            'added': added, 'updated': updated,
            'so_pic_refreshed': refreshed,
            'total_in_db': db.session.query(ProductIDDB).count()
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        db.session.rollback()
        return jsonify({'error': str(e)}), 500


@app.route('/api/upload/master-pic', methods=['POST'])
def upload_master_pic():
    """Upload Master PIC Excel. Upserts category_id → PIC mapping, then refreshes SO pic_name."""
    try:
        file = request.files.get('file')
        if not file:
            return jsonify({'error': 'No file provided'}), 400

        filename = (file.filename or '').lower()
        engine = 'xlrd' if filename.endswith('.xls') else 'openpyxl'
        df = pd.read_excel(file, sheet_name=0, engine=engine)
        df.columns = [str(c).strip() for c in df.columns]

        cat_col = next((c for c in df.columns if 'Category ID' in c or c.lower() == 'category id'), None)
        catn_col = next((c for c in df.columns if 'Category Name' in c or c.lower() == 'category name'), None)
        pic_col = next((c for c in df.columns if c.upper() == 'PIC' or 'PIC' in c), None)

        if not cat_col or not pic_col:
            return jsonify({'error': f'Missing required columns. Found: {list(df.columns)}'}), 400

        added = updated = 0
        for _, row in df.iterrows():
            cat_id = str(row[cat_col]).strip() if pd.notna(row[cat_col]) else None
            if not cat_id or cat_id == 'nan':
                continue
            cat_name = str(row[catn_col]).strip() if catn_col and pd.notna(row[catn_col]) else None
            pic_name = str(row[pic_col]).strip() if pd.notna(row[pic_col]) else None

            existing = db.session.query(MasterPIC).filter_by(category_id=cat_id).first()
            if existing:
                existing.category_name = cat_name
                existing.pic_name = pic_name
                existing.updated_at = datetime.utcnow()
                updated += 1
            else:
                db.session.add(MasterPIC(
                    category_id=cat_id, category_name=cat_name,
                    pic_name=pic_name, updated_at=datetime.utcnow()
                ))
                added += 1
        db.session.commit()

        # Rebuild pic_cache from DB for refreshing SO rows
        pic_map = {m.category_id: m.pic_name for m in db.session.query(MasterPIC).all()}
        prod_map = {p.product_id: p.category_id for p in db.session.query(ProductIDDB).all()}

        so_rows = db.session.query(SOData).filter(
            SOData.product_id.isnot(None), SOData.product_id != ''
        ).all()
        refreshed = 0
        for s in so_rows:
            cat_id_key = prod_map.get(str(s.product_id).strip())
            new_pic = pic_map.get(cat_id_key) if cat_id_key else None
            if s.pic_name != new_pic:
                s.pic_name = new_pic
                refreshed += 1
        db.session.commit()

        return jsonify({
            'status': 'ok',
            'added': added, 'updated': updated,
            'so_pic_refreshed': refreshed,
            'total_categories': db.session.query(MasterPIC).count()
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        db.session.rollback()
        return jsonify({'error': str(e)}), 500


@app.route('/api/master-pic/status', methods=['GET'])
def master_pic_status():
    """Return summary of Master PIC and ProductID database."""
    try:
        total_pid = db.session.query(ProductIDDB).count()
        last_pid = db.session.query(func.max(ProductIDDB.updated_at)).scalar()
        total_pic = db.session.query(MasterPIC).count()
        last_pic = db.session.query(func.max(MasterPIC.updated_at)).scalar()
        return jsonify({
            'product_id_count': total_pid,
            'last_product_id_upload': last_pid.isoformat() if last_pid else None,
            'master_pic_count': total_pic,
            'last_pic_update': last_pic.isoformat() if last_pic else None,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/data/pic-kpi', methods=['GET'])
def get_pic_kpi():
    """Return KPI metrics per PIC for Open SO."""
    try:
        # Get date filter params
        date_from = request.args.get('date_from', '')
        date_to = request.args.get('date_to', '')
        date_year = request.args.get('date_year', '')
        is_sqlite = 'sqlite' in app.config['SQLALCHEMY_DATABASE_URI']
        
        # Base query: Open SO only (exclude Delivery Completed and SO Cancel)
        q = db.session.query(SOData).filter(
            SOData.so_status.notin_(['Delivery Completed', 'SO Cancel'])
        )
        
        # Apply date filters
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
        
        rows = q.all()
        
        # Group by PIC
        pic_map = {}
        for s in rows:
            pic = s.pic_name or 'Unassigned'
            if pic not in pic_map:
                pic_map[pic] = {
                    'pic_name': pic,
                    'so_count': 0,
                    'total_sales': 0.0,
                    'total_purchase': 0.0,
                    'total_margin': 0.0,
                    'positive_margin_count': 0,
                    'negative_margin_count': 0,
                }
            
            pic_map[pic]['so_count'] += 1
            sales = float(s.sales_amount or 0)
            po_price = float(s.purchasing_price or 0)
            qty = float(s.so_qty or 0)
            po_amount = po_price * qty
            margin = sales - po_amount
            
            pic_map[pic]['total_sales'] += sales
            pic_map[pic]['total_purchase'] += po_amount
            pic_map[pic]['total_margin'] += margin
            
            if margin > 0:
                pic_map[pic]['positive_margin_count'] += 1
            elif margin < 0:
                pic_map[pic]['negative_margin_count'] += 1
        
        # Convert to list and sort by SO count descending
        result = sorted(pic_map.values(), key=lambda x: x['so_count'], reverse=True)
        
        return jsonify(result)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/template/master-pic', methods=['GET'])
def download_master_pic_template():
    """Generate and download Master PIC Excel template."""
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = 'Master PIC'
        
        # Header styling
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        header_align = Alignment(horizontal='center', vertical='center')
        
        # Headers
        headers = ['Category ID', 'Category Name', 'PIC']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align
        
        # Column widths
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 25
        
        # Sample data rows
        sample_data = [
            ['CAT001', 'Electronics', 'John Doe'],
            ['CAT002', 'Mechanical Parts', 'Jane Smith'],
            ['CAT003', 'Chemical Supplies', 'Bob Johnson'],
        ]
        
        for row_idx, row_data in enumerate(sample_data, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.alignment = Alignment(horizontal='left', vertical='center')
        
        # Instructions sheet
        ws_inst = wb.create_sheet('Instructions')
        instructions = [
            ['Master PIC Template - Instructions'],
            [''],
            ['Column Descriptions:'],
            ['1. Category ID: Unique identifier for the product category (required)'],
            ['2. Category Name: Name of the product category (optional)'],
            ['3. PIC: Person In Charge name for this category (required)'],
            [''],
            ['How to use:'],
            ['1. Fill in the Category ID and PIC columns (required)'],
            ['2. Category Name is optional but recommended for clarity'],
            ['3. Delete the sample data rows before uploading'],
            ['4. Upload the filled template via Manual Update > Update PIC'],
            [''],
            ['Notes:'],
            ['- Existing categories will be updated with new PIC names'],
            ['- New categories will be added to the database'],
            ['- After upload, SO rows will automatically refresh their PIC assignments'],
        ]
        
        for row_idx, row_data in enumerate(instructions, 1):
            cell = ws_inst.cell(row=row_idx, column=1, value=row_data[0])
            if row_idx == 1:
                cell.font = Font(bold=True, size=14, color='4472C4')
            elif 'Column Descriptions:' in row_data[0] or 'How to use:' in row_data[0] or 'Notes:' in row_data[0]:
                cell.font = Font(bold=True, size=11)
        
        ws_inst.column_dimensions['A'].width = 80
        
        # Save to BytesIO
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Template_Master_PIC.xlsx'
        )
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ═══════════════════════════════════════════════════════════════════════════
#  DELIVERY MONITORING
# ═══════════════════════════════════════════════════════════════════════════

class DeliveryMonitoring(db.Model):
    """
    Stores delivery process tracking data from Search_PO_Details exports.
    Primary key is po_number (PO No.) — one row per PO item.
    On re-upload: upsert by po_number, no duplicates allowed.
    PO Cancel rows are stored but excluded from leadtime calculations.
    """
    __tablename__ = 'delivery_monitoring'
    id              = db.Column(db.Integer, primary_key=True)
    po_number       = db.Column(db.String(100), unique=True, nullable=False, index=True)
    so_number       = db.Column(db.String(100))
    po_status       = db.Column(db.String(100))
    so_status       = db.Column(db.String(100))
    vendor_id       = db.Column(db.String(100))
    vendor_name     = db.Column(db.String(300))
    prod_id         = db.Column(db.String(100))
    prod_name       = db.Column(db.Text)
    op_unit_id      = db.Column(db.String(100))
    op_unit_name    = db.Column(db.Text)
    dlv_type        = db.Column(db.String(100))
    pur_pic         = db.Column(db.String(200))
    sales_pic       = db.Column(db.String(200))
    # Process date columns
    po_create_date      = db.Column(db.DateTime)
    so_erp_create_date  = db.Column(db.DateTime)
    po_rcvd_date        = db.Column(db.DateTime)
    ship_odr_date       = db.Column(db.DateTime)
    ship_compl_date     = db.Column(db.DateTime)
    hub_rcv_date        = db.Column(db.DateTime)
    hub_ship_date       = db.Column(db.DateTime)
    dlv_compl_date      = db.Column(db.DateTime)
    # Extra info
    dlv_due_date        = db.Column(db.DateTime)
    dlv_possible_date   = db.Column(db.DateTime)
    reject_date         = db.Column(db.DateTime)
    uploaded_at     = db.Column(db.DateTime, default=datetime.utcnow)


# Process stages in order, with human labels
DLV_PROCESS_STAGES = [
    ('po_create_date',     'PO Create Date',      'PO Created'),
    ('so_erp_create_date', 'SO(ERP) Create Date', 'SO ERP Created'),
    ('po_rcvd_date',       'PO Rcvd. Date',       'PO Received'),
    ('ship_odr_date',      'Ship. Odr. Date',     'Shipping Order'),
    ('ship_compl_date',    'Ship. Compl. Date',   'Ship Completed'),
    ('hub_rcv_date',       'HUB Rcv. Date',       'HUB Received'),
    ('hub_ship_date',      'HUB Ship. Date',       'HUB Shipped'),
    ('dlv_compl_date',     'Dlv. Compl. Date',    'Delivery Completed'),
]


def _parse_dt(val):
    """Parse a value to date/datetime, return None if not parseable."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (datetime, date)):
        return val
    try:
        import pandas as _pd
        ts = _pd.to_datetime(val, errors='coerce')
        if _pd.isna(ts):
            return None
        return ts.to_pydatetime()
    except Exception:
        return None


def _dt_to_date(val):
    """Convert datetime/date to date object."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    return None


def _calc_stage_leadtimes(row):
    """
    Calculate workday leadtime between consecutive process stages.
    Returns list of dicts: {stage_from, stage_to, label_from, label_to, workdays, pending}
    pending=True means the 'to' date is still not set (process not completed yet).
    """
    if row.po_status and 'cancel' in row.po_status.lower():
        return []

    results = []
    stages = DLV_PROCESS_STAGES
    for i in range(len(stages) - 1):
        field_from, _, label_from = stages[i]
        field_to,   _, label_to   = stages[i + 1]
        dt_from = getattr(row, field_from, None)
        dt_to   = getattr(row, field_to,   None)
        d_from  = _dt_to_date(dt_from) if dt_from else None
        d_to    = _dt_to_date(dt_to)   if dt_to   else None

        if d_from is None:
            # Can't compute; skip
            continue

        if d_to is None:
            # Stage not yet completed — pending, count from d_from to today
            today = date.today()
            wdays = count_workdays(d_from, today)
            results.append({
                'stage_from': field_from,
                'stage_to':   field_to,
                'label_from': label_from,
                'label_to':   label_to,
                'workdays':   wdays,
                'pending':    True,
            })
        else:
            wdays = count_workdays(d_from, d_to)
            results.append({
                'stage_from': field_from,
                'stage_to':   field_to,
                'label_from': label_from,
                'label_to':   label_to,
                'workdays':   wdays,
                'pending':    False,
            })
    return results


def _row_to_dict(row):
    """Convert a DeliveryMonitoring row to a JSON-serialisable dict."""
    stages = _calc_stage_leadtimes(row)
    total_wdays = None
    if row.po_create_date and row.dlv_compl_date:
        total_wdays = count_workdays(
            _dt_to_date(row.po_create_date),
            _dt_to_date(row.dlv_compl_date)
        )

    # Find where the process is currently pending (first pending stage)
    pending_stage = None
    if row.po_status and 'cancel' not in row.po_status.lower():
        for s in stages:
            if s['pending']:
                pending_stage = s['label_to']
                break

    return {
        'id':               row.id,
        'po_number':        row.po_number,
        'so_number':        row.so_number,
        'po_status':        row.po_status,
        'so_status':        row.so_status,
        'vendor_name':      row.vendor_name,
        'prod_id':          row.prod_id,
        'prod_name':        row.prod_name,
        'op_unit_name':     row.op_unit_name,
        'dlv_type':         row.dlv_type,
        'pur_pic':          row.pur_pic,
        'sales_pic':        row.sales_pic,
        'po_create_date':   row.po_create_date.isoformat() if row.po_create_date else None,
        'so_erp_create_date': row.so_erp_create_date.isoformat() if row.so_erp_create_date else None,
        'po_rcvd_date':     row.po_rcvd_date.isoformat() if row.po_rcvd_date else None,
        'ship_odr_date':    row.ship_odr_date.isoformat() if row.ship_odr_date else None,
        'ship_compl_date':  row.ship_compl_date.isoformat() if row.ship_compl_date else None,
        'hub_rcv_date':     row.hub_rcv_date.isoformat() if row.hub_rcv_date else None,
        'hub_ship_date':    row.hub_ship_date.isoformat() if row.hub_ship_date else None,
        'dlv_compl_date':   row.dlv_compl_date.isoformat() if row.dlv_compl_date else None,
        'dlv_due_date':     row.dlv_due_date.isoformat() if row.dlv_due_date else None,
        'stage_leadtimes':  stages,
        'total_workdays':   total_wdays,
        'pending_at':       pending_stage,
        'uploaded_at':      row.uploaded_at.isoformat() if row.uploaded_at else None,
    }


@app.route('/api/delivery-monitoring/upload', methods=['POST'])
def upload_delivery_monitoring():
    """
    Upload Search PO Details Excel file.
    Upsert by po_number — no duplicate PO numbers allowed.
    PO Cancel rows are stored but excluded from leadtime display.
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']

        df = pd.read_excel(file, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]

        # Column mapping (flexible aliases)
        def fc(aliases):
            for a in aliases:
                for c in df.columns:
                    if c.strip().lower() == a.lower():
                        return c
            return None

        col_po      = fc(['PO No.', 'PO No', 'PO Number'])
        col_so      = fc(['SO No.', 'SO No', 'SO Number'])
        col_postatus= fc(['PO Status'])
        col_sostatus= fc(['SO Status'])
        col_vid     = fc(['Vendor ID'])
        col_vnm     = fc(['Vendor Nm.', 'Vendor Name'])
        col_pid     = fc(['Prod. ID', 'Product ID'])
        col_pnm     = fc(['Prod. Nm.', 'Product Name', 'Prod. Nm.(Eng.)'])
        col_opid    = fc(['Op. Unit ID'])
        col_opnm    = fc(['Op. Unit Nm.', 'Op. Unit Name'])
        col_dtype   = fc(['Dlv. Type', 'Delivery Type'])
        col_ppic    = fc(['Pur. PIC', 'Purchase PIC'])
        col_spic    = fc(['Sales PIC', 'Sales PIC.1'])
        # Date columns
        col_pocreate = fc(['PO Create Date'])
        col_soerp    = fc(['SO(ERP) Create Date', 'SO Create Date'])
        col_porcvd   = fc(['PO Rcvd. Date', 'PO Received Date'])
        col_shodr    = fc(['Ship. Odr. Date', 'Shipping Order Date'])
        col_shcompl  = fc(['Ship. Compl. Date', 'Ship. Completed Date'])
        col_hubrcv   = fc(['HUB Rcv. Date', 'HUB Receive Date'])
        col_hubship  = fc(['HUB Ship. Date', 'HUB Ship Date'])
        col_dlvcompl = fc(['Dlv. Compl. Date', 'Delivery Complete Date'])
        col_dlvdue   = fc(['Dlv. Due Date', 'Delivery Due Date'])
        col_dlvposs  = fc(['Dlv. Possible Date', 'Delivery Possible Date'])
        col_reject   = fc(['Reject Date'])

        if not col_po:
            return jsonify({'error': f'Kolom "PO No." tidak ditemukan. Kolom tersedia: {df.columns.tolist()}'}), 400
        if not col_pocreate:
            return jsonify({'error': f'Kolom "PO Create Date" tidak ditemukan. Pastikan file yang diupload adalah Search PO Details.'}), 400

        # Check for in-file duplicates
        file_po_counts = df[col_po].dropna().astype(str).value_counts()
        in_file_dupes = file_po_counts[file_po_counts > 1]
        if len(in_file_dupes) > 0:
            return jsonify({
                'error': f'File mengandung {len(in_file_dupes)} PO duplikat: {", ".join(in_file_dupes.index[:5].tolist())}{"..." if len(in_file_dupes) > 5 else ""}. Harap periksa file dan hapus data duplikat.'
            }), 400

        # Load existing records
        existing = {r.po_number: r for r in DeliveryMonitoring.query.all()}

        count = 0
        skipped = 0
        for _, row in df.iterrows():
            po_num = str(row[col_po]).strip() if col_po and pd.notna(row[col_po]) else None
            if not po_num or po_num.lower() in ('nan', ''):
                skipped += 1
                continue

            def gv(col):
                if col is None: return None
                v = row[col]
                if pd.isna(v) if hasattr(pd, 'isna') else v is None: return None
                return str(v).strip() if not isinstance(v, (int, float)) else v

            new_data = dict(
                po_number        = po_num,
                so_number        = gv(col_so),
                po_status        = gv(col_postatus),
                so_status        = gv(col_sostatus),
                vendor_id        = gv(col_vid),
                vendor_name      = gv(col_vnm),
                prod_id          = gv(col_pid),
                prod_name        = gv(col_pnm),
                op_unit_id       = gv(col_opid),
                op_unit_name     = gv(col_opnm),
                dlv_type         = gv(col_dtype),
                pur_pic          = gv(col_ppic),
                sales_pic        = gv(col_spic),
                po_create_date   = _parse_dt(gv(col_pocreate)),
                so_erp_create_date = _parse_dt(gv(col_soerp)),
                po_rcvd_date     = _parse_dt(gv(col_porcvd)),
                ship_odr_date    = _parse_dt(gv(col_shodr)),
                ship_compl_date  = _parse_dt(gv(col_shcompl)),
                hub_rcv_date     = _parse_dt(gv(col_hubrcv)),
                hub_ship_date    = _parse_dt(gv(col_hubship)),
                dlv_compl_date   = _parse_dt(gv(col_dlvcompl)),
                dlv_due_date     = _parse_dt(gv(col_dlvdue)),
                dlv_possible_date= _parse_dt(gv(col_dlvposs)),
                reject_date      = _parse_dt(gv(col_reject)),
                uploaded_at      = datetime.utcnow(),
            )

            if po_num in existing:
                rec = existing[po_num]
                for k, v in new_data.items():
                    setattr(rec, k, v)
            else:
                rec = DeliveryMonitoring(**new_data)
                db.session.add(rec)
                existing[po_num] = rec

            count += 1
            if count % 200 == 0:
                db.session.flush()

        db.session.add(UploadLog(file_type='DELIVERY', filename=file.filename, records_count=count))
        db.session.commit()
        return jsonify({'message': f'Berhasil upload {count} data delivery monitoring (skip: {skipped})', 'uploaded': count, 'skipped': skipped})
    except Exception as e:
        db.session.rollback()
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/delivery-monitoring/data', methods=['GET'])
def get_delivery_monitoring():
    """
    Return delivery monitoring records with leadtime calculations.
    Excludes PO Cancel from leadtime.
    Supports filters: status, pending_at, search (po/so number).
    """
    try:
        status_filter  = request.args.get('status', '')
        pending_filter = request.args.get('pending_at', '')
        search         = request.args.get('search', '').strip()
        page           = int(request.args.get('page', 1))
        per_page       = int(request.args.get('per_page', 50))

        q = DeliveryMonitoring.query

        if status_filter:
            q = q.filter(DeliveryMonitoring.po_status.ilike(f'%{status_filter}%'))
        if search:
            q = q.filter(
                db.or_(
                    DeliveryMonitoring.po_number.ilike(f'%{search}%'),
                    DeliveryMonitoring.so_number.ilike(f'%{search}%'),
                    DeliveryMonitoring.vendor_name.ilike(f'%{search}%'),
                )
            )

        total = q.count()
        rows  = q.order_by(DeliveryMonitoring.po_create_date.desc()).offset((page-1)*per_page).limit(per_page).all()

        data = [_row_to_dict(r) for r in rows]

        # Apply pending_at filter client-side (after calculation)
        if pending_filter:
            data = [d for d in data if d.get('pending_at') == pending_filter]

        # Summary stats for cards
        all_rows = DeliveryMonitoring.query.all()
        non_cancel = [r for r in all_rows if not (r.po_status and 'cancel' in r.po_status.lower())]
        pending_counts = {}
        for r in non_cancel:
            stages = _calc_stage_leadtimes(r)
            for s in stages:
                if s['pending']:
                    pending_counts[s['label_to']] = pending_counts.get(s['label_to'], 0) + 1
                    break  # only count the first pending stage

        # Average leadtime per stage (across completed stages)
        stage_avg = {}
        for _, _, label_from in DLV_PROCESS_STAGES[:-1]:
            pass
        stage_pairs = [(DLV_PROCESS_STAGES[i][0], DLV_PROCESS_STAGES[i+1][0],
                        DLV_PROCESS_STAGES[i][2], DLV_PROCESS_STAGES[i+1][2])
                       for i in range(len(DLV_PROCESS_STAGES)-1)]
        for f_from, f_to, lbl_from, lbl_to in stage_pairs:
            vals = []
            for r in non_cancel:
                d_from = _dt_to_date(getattr(r, f_from))
                d_to   = _dt_to_date(getattr(r, f_to))
                if d_from and d_to:
                    vals.append(count_workdays(d_from, d_to))
            key = f'{lbl_from} → {lbl_to}'
            stage_avg[key] = round(sum(vals)/len(vals), 1) if vals else None

        return jsonify({
            'data':          data,
            'total':         total,
            'page':          page,
            'per_page':      per_page,
            'pending_counts': pending_counts,
            'stage_avg':     stage_avg,
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/delivery-monitoring/summary', methods=['GET'])
def get_delivery_monitoring_summary():
    """Summary cards for Delivery Monitoring dashboard."""
    try:
        all_rows    = DeliveryMonitoring.query.all()
        total       = len(all_rows)
        cancelled   = sum(1 for r in all_rows if r.po_status and 'cancel' in r.po_status.lower())
        completed   = sum(1 for r in all_rows if r.dlv_compl_date is not None)
        non_cancel  = [r for r in all_rows if not (r.po_status and 'cancel' in r.po_status.lower())]

        pending_counts = {}
        longest_pending = []  # top 10 longest pending per stage

        for r in non_cancel:
            stages = _calc_stage_leadtimes(r)
            for s in stages:
                if s['pending']:
                    lbl = s['label_to']
                    pending_counts[lbl] = pending_counts.get(lbl, 0) + 1
                    longest_pending.append({
                        'po_number':  r.po_number,
                        'so_number':  r.so_number,
                        'po_status':  r.po_status,
                        'vendor_name': r.vendor_name,
                        'pending_at': lbl,
                        'workdays':   s['workdays'],
                    })
                    break

        longest_pending.sort(key=lambda x: x['workdays'] if x['workdays'] else 0, reverse=True)

        # Stage avg leadtimes
        stage_pairs = [(DLV_PROCESS_STAGES[i][0], DLV_PROCESS_STAGES[i+1][0],
                        DLV_PROCESS_STAGES[i][2], DLV_PROCESS_STAGES[i+1][2])
                       for i in range(len(DLV_PROCESS_STAGES)-1)]
        stage_avg = []
        for f_from, f_to, lbl_from, lbl_to in stage_pairs:
            vals = []
            for r in non_cancel:
                d_from = _dt_to_date(getattr(r, f_from))
                d_to   = _dt_to_date(getattr(r, f_to))
                if d_from and d_to:
                    vals.append(count_workdays(d_from, d_to))
            stage_avg.append({
                'stage': f'{lbl_from} → {lbl_to}',
                'label_from': lbl_from,
                'label_to': lbl_to,
                'avg_workdays': round(sum(vals)/len(vals), 1) if vals else None,
                'count': len(vals),
            })

        last_upload = db.session.query(func.max(UploadLog.uploaded_at)).filter(UploadLog.file_type == 'DELIVERY').scalar()

        return jsonify({
            'total':           total,
            'cancelled':       cancelled,
            'completed':       completed,
            'in_progress':     total - cancelled - completed,
            'pending_counts':  pending_counts,
            'longest_pending': longest_pending[:20],
            'stage_avg':       stage_avg,
            'last_updated':    utc_isoformat(last_upload),
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("Backend: http://127.0.0.1:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
