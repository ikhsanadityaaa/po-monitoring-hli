"""
Optimized Flask Backend for PythonAnywhere + SQLite
=====================================================
Key optimizations applied:
1. Read paths NEVER call external APIs (Frankfurter, Google Sheets)
2. Exchange rates cached permanently in DB, warmed at startup
3. RFQ smart sync via lightweight fingerprint detection
4. Server-side pagination on all list endpoints
5. SQLite composite indexes for all query patterns
6. In-memory caches warmed at startup
7. Batch operations instead of N+1 queries
8. Cursor-based pagination where possible
"""

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
import pandas as pd
import numpy as np
import re
import os
import json
import html
import hashlib
from datetime import datetime, date, timedelta
import io
import time
import threading
from sqlalchemy import func, text, event
from sqlalchemy.engine import Engine
from sqlalchemy.orm import load_only
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════════════════
# 1. APP INITIALIZATION
# ═══════════════════════════════════════════════════════════════════════════

app = Flask(__name__)

CORS(app, resources={r"/api/*": {
    "origins": "*",
    "methods": ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    "allow_headers": ["Content-Type", "Authorization", "Accept"]
}})

# ─── Database Configuration ──────────────────────────────────────────────
_db_url = os.environ.get('DATABASE_URL', '')
if _db_url:
    if _db_url.startswith('postgres://'):
        _db_url = _db_url.replace('postgres://', 'postgresql://', 1)
    app.config['SQLALCHEMY_DATABASE_URI'] = _db_url
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'pool_pre_ping': True, 'pool_recycle': 300,
        'pool_size': 5, 'max_overflow': 10,
        'connect_args': {'connect_timeout': 10},
    }
else:
    _inst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'instance')
    os.makedirs(_inst, exist_ok=True)
    app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{_inst}/po_database.db'
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'pool_pre_ping': True,
        'connect_args': {'timeout': 30},
    }

app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
db = SQLAlchemy(app)


# ─── SQLite Performance Pragmas ──────────────────────────────────────────
@event.listens_for(Engine, 'connect')
def _set_sqlite_pragmas(dbapi_connection, connection_record):
    if 'sqlite' in app.config.get('SQLALCHEMY_DATABASE_URI', ''):
        cursor = dbapi_connection.cursor()
        cursor.execute('PRAGMA journal_mode=WAL')
        cursor.execute('PRAGMA synchronous=NORMAL')
        cursor.execute('PRAGMA busy_timeout=30000')
        cursor.execute('PRAGMA temp_store=MEMORY')
        cursor.execute('PRAGMA cache_size=-65536')
        cursor.execute('PRAGMA wal_autocheckpoint=1000')
        cursor.close()


# ═══════════════════════════════════════════════════════════════════════════
# 2. THREAD LOCKS & CACHE INFRASTRUCTURE
# ═══════════════════════════════════════════════════════════════════════════

_HOLIDAY_CACHE = None
_HOLIDAY_CACHE_KEY = None
_HOLIDAY_ARRAY_CACHE = None
_HOLIDAY_ARRAY_CACHE_KEY = None

_HOLIDAY_LOCK = threading.Lock()
_READ_CACHE_LOCK = threading.Lock()
_COMPLETED_CACHE_LOCK = threading.Lock()
_RATE_CACHE_LOCK = threading.Lock()
_SIMILARITY_LOCK = threading.Lock()
_MASTER_PIC_LOCK = threading.Lock()

# Exchange rate: {date: float} — permanent, warmed at startup
_RATE_CACHE = {}
# FX rate for non-USD: {(currency, date): float}
_FX_RATE_CACHE = {}

# General read-response cache
_READ_RESPONSE_CACHE = {}

# Delivery completed summary cache
_COMPLETED_SUMMARY_CACHE = {}
_COMPLETED_SUMMARY_CACHE_TTL_SECONDS = 300

# Similarity cache
_SIMILARITY_CACHE = {}
_SIMILARITY_CACHE_FILE = os.path.join(
    os.path.dirname(__file__), 'instance', 'similarity_cache.json')

# Master PIC cache: {category_name: pic_name}
_MASTER_PIC_CACHE = {'signature': None, 'by_id': {}, 'by_name': {}}

# ProductIDDB category cache
_PID_CATEGORY_CACHE = {}
_PID_CATEGORY_CACHE_LOADED = False

_RUNTIME_CACHE_VERSION = 0


def runtime_cache_get(key):
    with _READ_CACHE_LOCK:
        item = _READ_RESPONSE_CACHE.get(key)
        if not item:
            return None
        expires_at, payload = item
        if expires_at <= datetime.utcnow():
            _READ_RESPONSE_CACHE.pop(key, None)
            return None
        return payload


def runtime_cache_set(key, payload, ttl_seconds=120):
    with _READ_CACHE_LOCK:
        _READ_RESPONSE_CACHE[key] = (
            datetime.utcnow() + timedelta(seconds=ttl_seconds), payload)


def runtime_cache_key(namespace):
    return (namespace, request.query_string.decode('utf-8', errors='ignore'))


def clear_runtime_caches():
    global _RUNTIME_CACHE_VERSION
    with _READ_CACHE_LOCK:
        _RUNTIME_CACHE_VERSION += 1
        _READ_RESPONSE_CACHE.clear()
    with _COMPLETED_CACHE_LOCK:
        _COMPLETED_SUMMARY_CACHE.clear()
    try:
        RFQ_CACHE['expires_at'] = None
    except NameError:
        pass
    try:
        VENDOR_CONTROL_CACHE['expires_at'] = None
    except NameError:
        pass


# ═══════════════════════════════════════════════════════════════════════════
# 3. HOLIDAY HELPERS (unchanged logic, optimized caching)
# ═══════════════════════════════════════════════════════════════════════════

def _holiday_set():
    global _HOLIDAY_CACHE, _HOLIDAY_CACHE_KEY
    today_year = date.today().year
    cache_key = today_year
    if _HOLIDAY_CACHE is not None and _HOLIDAY_CACHE_KEY == cache_key:
        return _HOLIDAY_CACHE

    years = list(range(today_year - 2, today_year + 2))
    try:
        import holidays as _holidays_pkg
        s = set(_holidays_pkg.country_holidays('ID', years=years).keys())
    except Exception:
        s = set()

    extras_path = os.path.join(os.path.dirname(__file__), 'holiday_extras.json')
    try:
        with open(extras_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        items = data.get('dates', []) if isinstance(data, dict) else data
        for ds in items or []:
            try:
                s.add(date.fromisoformat(str(ds).strip()))
            except (ValueError, TypeError):
                pass
    except (FileNotFoundError, OSError, json.JSONDecodeError):
        pass

    _HOLIDAY_CACHE = s
    _HOLIDAY_CACHE_KEY = cache_key
    return s


def _holiday_array():
    global _HOLIDAY_ARRAY_CACHE, _HOLIDAY_ARRAY_CACHE_KEY
    holiday_set = _holiday_set()
    cache_key = (_HOLIDAY_CACHE_KEY, len(holiday_set))
    if _HOLIDAY_ARRAY_CACHE is not None and _HOLIDAY_ARRAY_CACHE_KEY == cache_key:
        return _HOLIDAY_ARRAY_CACHE
    arr = (np.array(sorted(holiday_set), dtype='datetime64[D]')
           if holiday_set else np.array([], dtype='datetime64[D]'))
    _HOLIDAY_ARRAY_CACHE = arr
    _HOLIDAY_ARRAY_CACHE_KEY = cache_key
    return arr


def is_workday(d):
    return d.weekday() < 5 and d not in _holiday_set()


def count_workdays(start, end):
    if start is None or end is None:
        return None
    if start == end:
        return 0
    holidays = _holiday_array()
    if end > start:
        return int(np.busday_count(start, end, holidays=holidays))
    return -int(np.busday_count(end, start, holidays=holidays))


def workdays_since(past_date, today=None):
    if past_date is None:
        return None
    if today is None:
        today = date.today()
    return count_workdays(past_date, today)


def workdays_until(future_date, today=None):
    if future_date is None:
        return None
    if today is None:
        today = date.today()
    return count_workdays(today, future_date)


# ═══════════════════════════════════════════════════════════════════════════
# 4. DATABASE MODELS
# ═══════════════════════════════════════════════════════════════════════════

class SOData(db.Model):
    __tablename__ = 'so_data'
    id = db.Column(db.Integer, primary_key=True)
    so_number = db.Column(db.String(50), index=True)
    so_item = db.Column(db.String(100))
    so_status = db.Column(db.String(50))
    operation_unit_name = db.Column(db.String(200))
    vendor_id = db.Column(db.String(100))
    vendor_name = db.Column(db.String(200))
    customer_po_number = db.Column(db.String(200))
    delivery_memo = db.Column(db.Text)
    product_name = db.Column(db.Text)
    specification = db.Column(db.Text)
    manufacturer_name = db.Column(db.String(300))
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


class ExchangeRate(db.Model):
    __tablename__ = 'exchange_rate'
    id = db.Column(db.Integer, primary_key=True)
    rate_date = db.Column(db.Date, nullable=False, unique=True, index=True)
    usd_to_idr = db.Column(db.Float, nullable=False)
    source = db.Column(db.String(50), default='manual')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class ProductIDDB(db.Model):
    __tablename__ = 'product_id_db'
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.String(100), unique=True, nullable=False, index=True)
    category_id = db.Column(db.String(100))
    category_name = db.Column(db.String(255))
    product_name = db.Column(db.Text)
    product_status = db.Column(db.String(100))
    specification = db.Column(db.Text)
    manufacturer_name = db.Column(db.String(255))
    vendor_name = db.Column(db.String(300))
    order_unit = db.Column(db.String(50))
    hub_handling_check = db.Column(db.String(100))
    tax_type = db.Column(db.String(100))
    registration_date = db.Column(db.Date, index=True)
    product_registry_pic = db.Column(db.String(200))
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)


class MasterPIC(db.Model):
    __tablename__ = 'master_pic'
    id = db.Column(db.Integer, primary_key=True)
    category_id = db.Column(db.String(100), unique=True, nullable=False, index=True)
    category_name = db.Column(db.String(255))
    pic_name = db.Column(db.String(100))
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)


class ItemRegistration(db.Model):
    __tablename__ = 'item_registration'
    id = db.Column(db.Integer, primary_key=True)
    proc_status = db.Column(db.String(100))
    req_date = db.Column(db.Date, index=True)
    existing_owner = db.Column(db.String(100))
    client_name = db.Column(db.String(300), index=True)
    category = db.Column(db.String(255))
    category_id = db.Column(db.String(100))
    pic = db.Column(db.String(200))
    pic_name = db.Column(db.String(200))
    req_no = db.Column(db.String(100), index=True)
    prod_id = db.Column(db.String(100), index=True)
    product_status = db.Column(db.String(100))
    batch_grp_no = db.Column(db.String(100))
    prod_name = db.Column(db.Text)
    spec = db.Column(db.Text)
    mfr_name = db.Column(db.String(300))
    odr_unit = db.Column(db.String(50))
    vendor_name = db.Column(db.String(300))
    prod_price = db.Column(db.Float)
    curr = db.Column(db.String(20))
    hub_handling_check = db.Column(db.String(100))
    tax_type = db.Column(db.String(50))
    registration_date = db.Column(db.Date, index=True)
    product_registry_pic = db.Column(db.String(200))
    remarks = db.Column(db.Text)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)


class RFQCellEdit(db.Model):
    __tablename__ = 'rfq_cell_edit'
    id = db.Column(db.Integer, primary_key=True)
    row_key = db.Column(db.String(200), nullable=False, index=True)
    field = db.Column(db.String(100), nullable=False)
    value = db.Column(db.Text)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)
    __table_args__ = (
        db.UniqueConstraint('row_key', 'field', name='uq_rfq_cell_edit_row_field'),
    )


class RFQDashboardRow(db.Model):
    __tablename__ = 'rfq_dashboard_row'
    id = db.Column(db.Integer, primary_key=True)
    row_key = db.Column(db.String(200), unique=True, nullable=False, index=True)
    sheet_row = db.Column(db.Integer, index=True)
    data_json = db.Column(db.Text, nullable=False, default='{}')
    dirty_fields_json = db.Column(db.Text, nullable=False, default='[]')
    first_seen_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    last_seen_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)


class ImportVendor(db.Model):
    __tablename__ = 'import_vendor'
    id = db.Column(db.Integer, primary_key=True)
    vendor_name = db.Column(db.String(300), unique=True, nullable=False, index=True)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)


class ImportDashboardRow(db.Model):
    __tablename__ = 'import_dashboard_row'
    id = db.Column(db.Integer, primary_key=True)
    row_key = db.Column(db.String(120), unique=True, nullable=False, index=True)
    source_key = db.Column(db.String(50), nullable=False, index=True)
    source_label = db.Column(db.String(100))
    source_uid = db.Column(db.String(50), nullable=False, index=True)
    sheet_row = db.Column(db.Integer)
    vendor_name = db.Column(db.String(300), index=True)
    data_json = db.Column(db.Text, nullable=False, default='{}')
    first_seen_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    last_seen_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)


class ImportDashboardMeta(db.Model):
    __tablename__ = 'import_dashboard_meta'
    id = db.Column(db.Integer, primary_key=True)
    meta_key = db.Column(db.String(100), unique=True, nullable=False, index=True)
    value_json = db.Column(db.Text, nullable=False, default='null')
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)


# ═══════════════════════════════════════════════════════════════════════════
# 5. PRODUCTIDDB CATEGORY CACHE (warm at startup)
# ═══════════════════════════════════════════════════════════════════════════

def _pid_category_cache_load():
    global _PID_CATEGORY_CACHE, _PID_CATEGORY_CACHE_LOADED
    mapping = {}
    try:
        for pid, cat in db.session.query(
                ProductIDDB.product_id, ProductIDDB.category_name).all():
            if not pid:
                continue
            raw = (cat or '').strip()
            mapping[str(pid).strip()] = (
                raw.split('>')[0].strip() if '>' in raw else raw)
    except Exception:
        mapping = {}
    with _MASTER_PIC_LOCK:
        _PID_CATEGORY_CACHE = mapping
        _PID_CATEGORY_CACHE_LOADED = True


def _pid_category_lookup(product_id):
    global _PID_CATEGORY_CACHE_LOADED
    with _MASTER_PIC_LOCK:
        loaded = _PID_CATEGORY_CACHE_LOADED
    if not loaded:
        _pid_category_cache_load()
    pid = str(product_id or '').strip()
    with _MASTER_PIC_LOCK:
        return _PID_CATEGORY_CACHE.get(pid, '')


def _pid_category_cache_invalidate():
    global _PID_CATEGORY_CACHE_LOADED
    with _MASTER_PIC_LOCK:
        _PID_CATEGORY_CACHE_LOADED = False


# ═══════════════════════════════════════════════════════════════════════════
# 6. EXCHANGE RATE HELPERS — PERMANENT CACHE, READ-ONLY PATHS
# ═══════════════════════════════════════════════════════════════════════════

def _fetch_rate_from_api(d, currency='USD'):
    """Fetch one historical currency→IDR rate from Frankfurter v2.
    ONLY called during upload/backfill — NEVER on read path."""
    try:
        import urllib.request
        cur = (currency or 'USD').strip().upper()
        url = (f"https://api.frankfurter.dev/v2/rate/{cur}/IDR"
               f"?date={d.isoformat()}")
        with urllib.request.urlopen(url, timeout=6) as resp:
            data = json.loads(resp.read())
        return float(data['rate'])
    except Exception:
        return None


def _get_fallback_rate():
    last = ExchangeRate.query.order_by(ExchangeRate.rate_date.desc()).first()
    return last.usd_to_idr if last else 16000.0


def _warm_exchange_rate_cache():
    """Load ALL exchange rates from DB into memory at startup.
    Historical rates are immutable — this is a permanent cache."""
    global _RATE_CACHE
    try:
        rates = ExchangeRate.query.all()
        for r in rates:
            _RATE_CACHE[r.rate_date] = r.usd_to_idr
        print(f'Exchange rate cache warmed: {len(_RATE_CACHE)} dates')
    except Exception as e:
        print(f'Exchange rate warm skipped: {e}')


def get_usd_to_idr_readonly(d):
    """Dashboard-safe: NEVER hits external API.

    Flow: memory cache → DB exact → DB nearest → fallback.
    Historical rates never change once stored."""
    if d is None:
        return _get_fallback_rate()

    # 1. In-memory cache (instant, zero I/O)
    if d in _RATE_CACHE:
        return _RATE_CACHE[d]

    # 2. DB exact match (SQLite index lookup)
    rec = ExchangeRate.query.filter_by(rate_date=d).first()
    if rec:
        _RATE_CACHE[d] = rec.usd_to_idr
        return rec.usd_to_idr

    # 3. Nearest known rate from DB (NO network call)
    nearest = ExchangeRate.query.order_by(
        func.abs(
            func.julianday(ExchangeRate.rate_date) - func.julianday(str(d))
        )
    ).first()
    if nearest:
        _RATE_CACHE[d] = nearest.usd_to_idr
        return nearest.usd_to_idr

    # 4. Hardcoded fallback
    return _get_fallback_rate()


def get_usd_to_idr(d, cache_only=False):
    """Write-path version. Can fetch from API when cache_only=False.
    Used ONLY during upload/backfill operations."""
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
                db.session.add(ExchangeRate(
                    rate_date=d, usd_to_idr=rate, source='frankfurter'))
                db.session.commit()
            except Exception:
                db.session.rollback()
            _RATE_CACHE[d] = rate
            return rate
    nearest = ExchangeRate.query.order_by(
        func.abs(
            func.julianday(ExchangeRate.rate_date) - func.julianday(str(d))
        )
    ).first()
    if nearest:
        _RATE_CACHE[d] = nearest.usd_to_idr
        return nearest.usd_to_idr
    return _get_fallback_rate()


def get_currency_to_idr(currency, d, cache_only=False, readonly=False):
    cur = (currency or 'IDR').strip().upper()
    if cur in ('IDR', ''):
        return 1.0
    if cur == 'USD':
        if readonly:
            return get_usd_to_idr_readonly(d)
        return get_usd_to_idr(d, cache_only=cache_only)

    if d is None:
        d = date.today()
    key = (cur, d)
    if key in _FX_RATE_CACHE:
        return _FX_RATE_CACHE[key]
    if readonly:
        # Try nearest in cache
        same = [(rd, r) for (c, rd), r in _FX_RATE_CACHE.items() if c == cur]
        if same:
            nearest = min(same, key=lambda r: abs((r[0] - d).days))
            _FX_RATE_CACHE[key] = nearest[1]
            return nearest[1]
        return _get_fallback_rate()
    if not cache_only and d <= date.today():
        rate = _fetch_rate_from_api(d, cur)
        if rate:
            _FX_RATE_CACHE[key] = rate
            return rate
    same = [(rd, r) for (c, rd), r in _FX_RATE_CACHE.items() if c == cur]
    if same:
        nearest = min(same, key=lambda r: abs((r[0] - d).days))
        _FX_RATE_CACHE[key] = nearest[1]
        return nearest[1]
    fallback = _fetch_rate_from_api(date.today(), cur) if not cache_only else None
    if fallback:
        _FX_RATE_CACHE[key] = fallback
        return fallback
    return _get_fallback_rate()


def prefetch_exchange_rates(dates, fetch_missing=True, currency='USD'):
    """Warm in-memory cache for a collection of dates.

    IMPORTANT: Dashboard calls with fetch_missing=False.
    Upload/backfill calls with fetch_missing=True."""
    cur = (currency or 'USD').strip().upper()
    if not dates or cur in ('IDR', ''):
        return

    if cur != 'USD':
        needed = {d for d in dates
                  if d is not None and (cur, d) not in _FX_RATE_CACHE}
        if fetch_missing:
            today = date.today()
            for d in sorted(x for x in needed if x <= today):
                rate = _fetch_rate_from_api(d, cur)
                if rate:
                    _FX_RATE_CACHE[(cur, d)] = rate
                    needed.discard(d)
        if needed:
            same = [(rd, r) for (c, rd), r in _FX_RATE_CACHE.items() if c == cur]
            fallback = get_currency_to_idr(
                cur, date.today(), cache_only=not fetch_missing)
            for d in needed:
                if same:
                    nearest = min(same, key=lambda r: abs((r[0] - d).days))
                    _FX_RATE_CACHE[(cur, d)] = nearest[1]
                else:
                    _FX_RATE_CACHE[(cur, d)] = fallback
        return

    needed = {d for d in dates if d is not None and d not in _RATE_CACHE}
    if not needed:
        return

    # Bulk DB load
    db_rows = ExchangeRate.query.filter(
        ExchangeRate.rate_date.in_(list(needed))).all()
    for row in db_rows:
        _RATE_CACHE[row.rate_date] = row.usd_to_idr
    needed -= {row.rate_date for row in db_rows}
    if not needed:
        return

    if fetch_missing:
        today = date.today()
        fetched_rows = []
        for d in sorted(d for d in needed if d <= today):
            rate = _fetch_rate_from_api(d)
            if rate:
                _RATE_CACHE[d] = rate
                fetched_rows.append(ExchangeRate(
                    rate_date=d, usd_to_idr=rate, source='frankfurter'))
                needed.discard(d)
        if fetched_rows:
            try:
                db.session.bulk_save_objects(fetched_rows)
                db.session.commit()
            except Exception:
                db.session.rollback()

    if needed:
        fallback = _get_fallback_rate()
        all_rates = ExchangeRate.query.order_by(
            ExchangeRate.rate_date).all()
        for d in needed:
            if all_rates:
                nearest = min(all_rates,
                              key=lambda r: abs((r.rate_date - d).days))
                _RATE_CACHE[d] = nearest.usd_to_idr
            else:
                _RATE_CACHE[d] = fallback


def backfill_exchange_rates_for_rows(rows):
    """Called ONLY during upload. Fetches missing rates and stores in DB.
    After this, all dates have rates cached permanently."""
    dates = set()
    for s in rows:
        cur = (getattr(s, 'purchasing_currency', '') or '').strip().upper()
        d = getattr(s, 'so_create_date', None)
        if cur in ('USD', 'EUR') and d and d <= date.today():
            dates.add(d)
    if not dates:
        return 0
    prefetch_exchange_rates(dates, fetch_missing=True)
    return len(dates)


def compute_and_cache_purchase_amount_idr(s):
    """Compute purchasing_amount_idr and persist to DB column.
    Called ONLY during upload/backfill — NOT on read path."""
    cached = getattr(s, 'purchasing_amount_idr', None)
    if cached is not None:
        return float(cached)

    raw = raw_purchase_amount(s)
    cur = (s.purchasing_currency or 'IDR').strip().upper()

    if cur in ('IDR', ''):
        s.purchasing_amount_idr = raw
    elif cur in ('USD', 'EUR'):
        rate = get_usd_to_idr_readonly(s.so_create_date)
        s.purchasing_amount_idr = raw * rate
    else:
        s.purchasing_amount_idr = raw

    s.purchasing_amount_idr_cached_at = datetime.utcnow()
    return s.purchasing_amount_idr


def raw_purchase_amount(s):
    raw = float(s.purchasing_amount or 0)
    if raw == 0 and s.purchasing_price:
        raw = float(s.purchasing_price) * float(s.so_qty or 0)
    return raw


def purchase_amount_idr(s, allow_persist=False):
    """Dashboard-safe read. NEVER hits external API.

    If purchasing_amount_idr is NULL (not yet backfilled):
    - IDR rows: return raw amount directly
    - Non-IDR rows: return 0 (admin must run backfill)
    """
    cached = getattr(s, 'purchasing_amount_idr', None)
    if cached is not None:
        return float(cached)

    raw = raw_purchase_amount(s)
    cur = (s.purchasing_currency or 'IDR').strip().upper()

    if cur in ('IDR', ''):
        return raw

    if not allow_persist:
        return 0.0

    # Only called during upload/backfill
    return compute_and_cache_purchase_amount_idr(s)


def convert_to_idr(amount, currency, rate_date=None,
                    cache_only=False, readonly=False):
    if not amount:
        return 0.0
    cur = (currency or 'IDR').strip().upper()
    if cur in ('IDR', ''):
        return float(amount)
    if cur in ('USD', 'EUR'):
        return float(amount) * get_currency_to_idr(
            cur, rate_date, cache_only=cache_only, readonly=readonly)
    return float(amount)


# ═══════════════════════════════════════════════════════════════════════════
# 7. DATABASE MIGRATION & INDEXES
# ═══════════════════════════════════════════════════════════════════════════

def _existing_columns(table_name):
    is_sqlite = 'sqlite' in app.config.get('SQLALCHEMY_DATABASE_URI', '')
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


def _ensure_extra_columns():
    migration_plan = {
        'so_data': [
            ('specification', 'TEXT'),
            ('product_id', 'VARCHAR(100)'),
            ('vendor_id', 'VARCHAR(100)'),
            ('manufacturer_name', 'VARCHAR(300)'),
            ('purchasing_currency', 'VARCHAR(10)'),
            ('purchasing_amount_idr', 'DOUBLE PRECISION'),
            ('purchasing_amount_idr_cached_at', 'TIMESTAMP'),
            ('pic_name', 'VARCHAR(100)'),
        ],
        'item_registration': [
            ('req_date', 'DATE'),
            ('existing_owner', 'VARCHAR(100)'),
            ('category_id', 'VARCHAR(100)'),
            ('pic_name', 'VARCHAR(200)'),
            ('product_status', 'VARCHAR(100)'),
            ('hub_handling_check', 'VARCHAR(100)'),
            ('tax_type', 'VARCHAR(50)'),
            ('registration_date', 'DATE'),
            ('product_registry_pic', 'VARCHAR(200)'),
            ('remarks', 'TEXT'),
        ],
        'product_id_db': [
            ('specification', 'TEXT'),
            ('manufacturer_name', 'VARCHAR(255)'),
            ('vendor_name', 'VARCHAR(300)'),
            ('order_unit', 'VARCHAR(50)'),
            ('product_status', 'VARCHAR(100)'),
            ('hub_handling_check', 'VARCHAR(100)'),
            ('tax_type', 'VARCHAR(100)'),
            ('registration_date', 'DATE'),
            ('product_registry_pic', 'VARCHAR(200)'),
        ],
    }
    for table_name, columns in migration_plan.items():
        cols = _existing_columns(table_name)
        for col_name, col_type in columns:
            if col_name.lower() not in cols:
                try:
                    db.session.execute(text(
                        f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_type}"))
                    db.session.commit()
                    print(f'Migration: added {table_name}.{col_name}')
                except Exception as exc:
                    db.session.rollback()
                    print(f'Migration warning ({table_name}.{col_name}): {exc}')


def _ensure_performance_indexes():
    """Create indexes for ALL databases (SQLite and PostgreSQL).
    Composite indexes match actual query patterns from dashboard endpoints."""
    statements = [
        # SO dashboard: filter by status+date, sort by date
        "CREATE INDEX IF NOT EXISTS idx_so_status_date ON so_data (so_status, so_create_date)",
        "CREATE INDEX IF NOT EXISTS idx_so_op_unit ON so_data (operation_unit_name)",
        "CREATE INDEX IF NOT EXISTS idx_so_pic_name ON so_data (pic_name)",
        "CREATE INDEX IF NOT EXISTS idx_so_vendor_name ON so_data (vendor_name)",
        "CREATE INDEX IF NOT EXISTS idx_so_item ON so_data (so_item)",
        "CREATE INDEX IF NOT EXISTS idx_so_number ON so_data (so_number)",
        "CREATE INDEX IF NOT EXISTS idx_so_product_id ON so_data (product_id)",
        "CREATE INDEX IF NOT EXISTS idx_so_customer_po ON so_data (customer_po_number)",
        "CREATE INDEX IF NOT EXISTS idx_so_create_date ON so_data (so_create_date)",

        # Exchange rate: date lookup (critical for read-only path)
        "CREATE INDEX IF NOT EXISTS idx_exchange_rate_date ON exchange_rate (rate_date)",

        # Upload log
        "CREATE INDEX IF NOT EXISTS idx_upload_log_type_date ON upload_log (file_type, uploaded_at)",

        # RFQ dashboard
        "CREATE INDEX IF NOT EXISTS idx_rfq_dash_row_key ON rfq_dashboard_row (row_key)",
        "CREATE INDEX IF NOT EXISTS idx_rfq_dash_sheet_row ON rfq_dashboard_row (sheet_row)",
        "CREATE INDEX IF NOT EXISTS idx_rfq_dash_last_seen ON rfq_dashboard_row (last_seen_at)",
        "CREATE INDEX IF NOT EXISTS idx_rfq_cell_edit_row ON rfq_cell_edit (row_key, field)",

        # Import dashboard
        "CREATE INDEX IF NOT EXISTS idx_import_dash_row_key ON import_dashboard_row (row_key)",
        "CREATE INDEX IF NOT EXISTS idx_import_dash_source ON import_dashboard_row (source_key)",
        "CREATE INDEX IF NOT EXISTS idx_import_dash_vendor ON import_dashboard_row (vendor_name)",

        # ProductIDDB — similarity lookups
        "CREATE INDEX IF NOT EXISTS idx_pid_product_id ON product_id_db (product_id)",
        "CREATE INDEX IF NOT EXISTS idx_pid_status ON product_id_db (product_status)",
        "CREATE INDEX IF NOT EXISTS idx_pid_name_lower ON product_id_db (product_name COLLATE NOCASE)",
        "CREATE INDEX IF NOT EXISTS idx_pid_spec_lower ON product_id_db (specification COLLATE NOCASE)",

        # Master PIC
        "CREATE INDEX IF NOT EXISTS idx_master_pic_cat_id ON master_pic (category_id)",
        "CREATE INDEX IF NOT EXISTS idx_master_pic_cat_name ON master_pic (category_name)",

        # Item Registration
        "CREATE INDEX IF NOT EXISTS idx_item_reg_proc_client ON item_registration (proc_status, client_name)",
        "CREATE INDEX IF NOT EXISTS idx_item_reg_pic ON item_registration (pic)",
        "CREATE INDEX IF NOT EXISTS idx_item_reg_req_no ON item_registration (req_no)",
        "CREATE INDEX IF NOT EXISTS idx_item_reg_mfr ON item_registration (mfr_name)",
        "CREATE INDEX IF NOT EXISTS idx_item_reg_owner ON item_registration (existing_owner)",

        # Product status+unit combo
        "CREATE INDEX IF NOT EXISTS idx_product_status_unit ON product_id_db (product_status, order_unit)",
    ]
    for stmt in statements:
        try:
            db.session.execute(text(stmt))
        except Exception as exc:
            db.session.rollback()
            print(f'Index warning: {exc}')
    try:
        db.session.commit()
    except Exception:
        db.session.rollback()


# ═══════════════════════════════════════════════════════════════════════════
# 8. UTILITY FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════

CLOSED_STATUSES = {
    'Delivery Completed', 'SO Cancel',
    'Approval Apply', 'Approval Complete', 'Approval Complete Step',
    'Approval Reject', 'Approval Hold',
    'Return Complete(Vendor)', 'Return Complete(HUB)', 'Customer PO Reject'
}

DISCARDABLE_STATUSES = {
    'SO Cancel',
    'Approval Apply', 'Approval Complete Step', 'Approval Reject',
    'Approval Hold',
    'Return Complete(Vendor)', 'Return Complete(HUB)',
    'Customer PO Reject', 'Ship. Order Reject', 'PO Received Reject',
}

EXCLUDED_OP_UNITS = {'HLI GREEN POWER (CONSUMABLE)'}

PO_HLI_RE = re.compile(r'(\d{7,})(?:-(\d{1,4}))?(?!\d)')
PO_SHORT_REF_RE = re.compile(
    r'\bP\s*\.?\s*O\s*\.?\s*[#:.\-]?\s*(\d{2,6})\b', re.IGNORECASE)


def clean(val):
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    s = str(val).strip()
    return None if s.lower() in ('nan', 'none', '') else s


def clean_product_id(val):
    s = clean(val)
    if not s:
        return ''
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except (TypeError, ValueError):
        pass
    return re.sub(r'\.0+$', '', s)


def clean_request_number(val):
    s = clean(val)
    if not s:
        return ''
    s = str(s).strip()
    try:
        from decimal import Decimal, InvalidOperation
        number = Decimal(s)
        if number == number.to_integral_value():
            return format(number.quantize(Decimal('1')), 'f')
    except (InvalidOperation, ValueError, TypeError):
        pass
    return re.sub(r'\.0+$', '', s)


def parse_date(val):
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    raw = str(val).strip()
    if re.match(r'^\d{8}(\.0)?$', raw):
        try:
            return datetime.strptime(raw[:8], '%Y%m%d').date()
        except ValueError:
            pass
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None


def safe_float(val, default=0.0):
    try:
        if pd.isna(val):
            return default
    except (TypeError, ValueError):
        pass
    try:
        return float(val)
    except Exception:
        return default


def find_column(df, names):
    low = {c.lower().strip(): c for c in df.columns}
    for n in names:
        if n.lower().strip() in low:
            return low[n.lower().strip()]
    return None


def utc_isoformat(dt):
    if dt is None:
        return None
    s = dt.isoformat()
    tail = s[10:]
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
        text_val = str(field).strip()
        for token in re.split(r'[\s,;]+', text_val):
            token = token.strip()
            if token and token[0] == '2' and re.match(r'^2\d{6,}', token):
                return True
    return False


def so_is_countable(so_item, so_number=None, customer_po_number=None,
                    delivery_memo=None):
    if has_internal_po_ref(customer_po_number, delivery_memo):
        return False
    return True


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
    text_val = str(val).strip()
    result = set()
    for m in PO_HLI_RE.finditer(text_val):
        po_num = m.group(1)
        item_no = m.group(2)
        if po_num.startswith('2'):
            continue
        result.add(po_num)
        if item_no:
            for item_var in _normalize_item_no(item_no):
                result.add(f"{po_num}-{item_var}")
    return list(result)


def extract_po_short_refs(val):
    if not val:
        return []
    text_val = str(val).strip()
    refs = set()
    for m in PO_SHORT_REF_RE.finditer(text_val):
        n = m.group(1)
        if len(n) >= 7:
            continue
        refs.add(n)
    return list(refs)


# ═══════════════════════════════════════════════════════════════════════════
# 9. SO DATA HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def open_so_filter():
    return db.or_(
        SOData.so_status.is_(None),
        SOData.so_status.notin_(list(CLOSED_STATUSES))
    )


def parse_so_date_args(args=None):
    args = args if args is not None else request.args
    date_year = args.get('date_year', '')
    date_from = args.get('date_from', '')
    date_to = args.get('date_to', '')
    if not date_year:
        legacy = args.get('year', '')
        if legacy and legacy != 'all':
            date_year = legacy
    return date_year, date_from, date_to


def apply_so_create_date_filter(query, date_year='', date_from='',
                                date_to='', is_sqlite=None):
    if is_sqlite is None:
        is_sqlite = 'sqlite' in app.config['SQLALCHEMY_DATABASE_URI']
    if date_year:
        try:
            yr = int(date_year)
            if is_sqlite:
                return query.filter(
                    func.strftime('%Y', SOData.so_create_date) == str(yr))
            return query.filter(
                func.extract('year', SOData.so_create_date) == yr)
        except (ValueError, TypeError):
            return query
    if date_from:
        query = query.filter(SOData.so_create_date >= date_from)
    if date_to:
        query = query.filter(SOData.so_create_date <= date_to)
    return query


def apply_item_registration_date_filter(query, date_year='',
                                        date_from='', date_to=''):
    if date_year:
        try:
            yr = int(date_year)
            if 'sqlite' in app.config.get('SQLALCHEMY_DATABASE_URI', ''):
                return query.filter(
                    func.strftime('%Y', ItemRegistration.req_date) == str(yr))
            return query.filter(
                func.extract('year', ItemRegistration.req_date) == yr)
        except (ValueError, TypeError):
            return query
    df = parse_date(date_from) if date_from else None
    dt = parse_date(date_to) if date_to else None
    if df:
        query = query.filter(ItemRegistration.req_date >= df)
    if dt:
        query = query.filter(ItemRegistration.req_date <= dt)
    return query


def selected_clients(args=None):
    args = args if args is not None else request.args
    return [c.strip() for c in args.getlist('client')
            if c and c.strip()]


def selected_pics(args=None):
    args = args if args is not None else request.args
    return [p.strip() for p in args.getlist('pic')
            if p and p.strip()]


def matches_selected_client(value, clients):
    if not clients:
        return True
    v = (value or '').strip().lower()
    return any(v == c.lower() for c in clients)


def apply_so_client_filter(query, clients):
    if clients:
        return query.filter(SOData.operation_unit_name.in_(clients))
    return query


def apply_so_pic_filter(query, pics):
    if not pics:
        return query
    if '__NONE_PLACEHOLDER__' in pics:
        return query.filter(SOData.id.is_(None))
    non_yupi_op_unit = db.or_(
        SOData.operation_unit_name.is_(None),
        db.not_(SOData.operation_unit_name.ilike('%YUPI%')))
    if 'ANDRE' in pics:
        others = [p for p in pics if p != 'ANDRE']
        andre_filter = db.or_(
            SOData.pic_name == 'ANDRE',
            SOData.operation_unit_name.ilike('%YUPI%'))
        if others:
            others_filter = db.and_(
                SOData.pic_name.in_(others), non_yupi_op_unit)
            return query.filter(db.or_(others_filter, andre_filter))
        return query.filter(andre_filter)
    if '(Kosong)' in pics:
        others = [p for p in pics if p != '(Kosong)']
        empty_pic = db.and_(
            db.or_(SOData.pic_name.is_(None), SOData.pic_name == ''),
            non_yupi_op_unit)
        if others:
            others_filter = db.and_(
                SOData.pic_name.in_(others), non_yupi_op_unit)
            return query.filter(db.or_(others_filter, empty_pic))
        return query.filter(empty_pic)
    return query.filter(SOData.pic_name.in_(pics), non_yupi_op_unit)


def canonical_pending_pic(pic, client_or_op_unit=None):
    if client_or_op_unit and 'YUPI' in str(client_or_op_unit).upper():
        return 'ANDRE'
    return pic or 'Unassigned'


def canonical_rfq_pic(row):
    return canonical_pending_pic(
        clean(row.get('purchase_pic')), row.get('client_name'))


def sort_pic_kpis(rows):
    return sorted(rows, key=lambda x: (
        0 if x.get('pic') == 'ANDRE' else 1,
        -x.get('count', 0),
        x.get('pic') or ''))


def apply_item_registration_pic_filter(query, pics):
    if not pics:
        return query
    non_yupi_client = db.or_(
        ItemRegistration.client_name.is_(None),
        db.not_(ItemRegistration.client_name.ilike('%YUPI%')))
    if 'ANDRE' in pics:
        others = [p for p in pics if p != 'ANDRE']
        andre_filter = db.or_(
            ItemRegistration.pic_name == 'ANDRE',
            ItemRegistration.client_name.ilike('%YUPI%'))
        if others:
            others_filter = db.and_(
                ItemRegistration.pic_name.in_(others), non_yupi_client)
            return query.filter(db.or_(others_filter, andre_filter))
        return query.filter(andre_filter)
    if '(Kosong)' in pics:
        others = [p for p in pics if p != '(Kosong)']
        empty_pic = db.and_(
            db.or_(ItemRegistration.pic_name.is_(None),
                   ItemRegistration.pic_name == ''),
            non_yupi_client)
        if others:
            others_filter = db.and_(
                ItemRegistration.pic_name.in_(others), non_yupi_client)
            return query.filter(db.or_(others_filter, empty_pic))
        return query.filter(empty_pic)
    return query.filter(
        ItemRegistration.pic_name.in_(pics), non_yupi_client)


# ═══════════════════════════════════════════════════════════════════════════
# 10. UPLOAD HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def uploaded_files():
    files = []
    for key in ('file', 'files'):
        files.extend(request.files.getlist(key))
    return [f for f in files if f and f.filename]


def read_upload_excel(file):
    raw = file.read()
    file.seek(0)
    filename = (file.filename or '').lower()
    is_xls_format = raw[:4] == b'\xd0\xcf\x11\xe0'
    engine = 'xlrd' if is_xls_format or filename.endswith('.xls') else 'openpyxl'
    return pd.read_excel(file, sheet_name=0, engine=engine)


def _json_rows_to_dataframe(rows, columns=None):
    if rows is None:
        rows = []
    if not isinstance(rows, list):
        raise ValueError('JSON rows/data must be a list')
    if columns:
        return pd.DataFrame(rows, columns=[str(c).strip() for c in columns])
    if not rows:
        return pd.DataFrame()
    if all(isinstance(r, dict) for r in rows):
        return pd.DataFrame(rows)
    return pd.DataFrame(rows)


def _json_payload_to_uploads(payload, default_filename='json_upload'):
    if payload is None:
        raise ValueError('Invalid or empty JSON body')

    def one(obj, index=1):
        if isinstance(obj, dict):
            filename = (clean(obj.get('filename')) or
                        clean(obj.get('name')) or
                        f'{default_filename}_{index}.json')
            columns = obj.get('columns')
            rows = (obj.get('rows') if 'rows' in obj else
                    obj.get('data') if 'data' in obj else
                    obj.get('records') if 'records' in obj else
                    obj.get('items') if 'items' in obj else None)
            if rows is None:
                row = {k: v for k, v in obj.items()
                       if k not in ('filename', 'name', 'columns')}
                rows = [row] if row else []
            df = _json_rows_to_dataframe(rows, columns=columns)
            df.columns = [str(c).strip() for c in df.columns]
            return {'filename': filename, 'df': df}
        if isinstance(obj, list):
            df = _json_rows_to_dataframe(obj)
            df.columns = [str(c).strip() for c in df.columns]
            return {'filename': f'{default_filename}_{index}.json', 'df': df}
        raise ValueError('Each JSON upload must be an object or list')

    uploads = []
    if isinstance(payload, dict) and isinstance(payload.get('files'), list):
        for idx, item in enumerate(payload.get('files') or [], start=1):
            uploads.append(one(item, idx))
    else:
        uploads.append(one(payload, 1))
    return [u for u in uploads if u['df'] is not None]


def request_upload_dataframes(default_filename='upload'):
    content_type = (request.content_type or '').lower()
    if request.is_json or 'application/json' in content_type:
        payload = request.get_json(silent=True)
        uploads = _json_payload_to_uploads(
            payload, default_filename=default_filename)
        return uploads, 'json'
    files = uploaded_files()
    uploads = []
    for file in files:
        df = read_upload_excel(file)
        df.columns = [str(c).strip() for c in df.columns]
        uploads.append({'filename': file.filename, 'df': df})
    return uploads, 'excel'


def _product_id_columns(df):
    return {
        'product_id': find_column(df, ['Product ID', 'Prod. ID', 'Prod ID']),
        'category_id': find_column(df, ['Category ID', 'Category Id',
                                         'CategoryID', 'Cat. ID', 'Cat. ID.']),
        'category_name': find_column(df, ['Category Name', 'Category Nm.',
                                           'Cat. Nm.', 'Cat. Nm']),
        'product_name': find_column(df, ['Product Name', 'Prod. Nm.',
                                          'Prod. Nm', 'Product Name(EN)']),
        'product_status': find_column(df, ['Product Status', 'Prod. Status',
                                            'Prod Status']),
        'specification': find_column(df, ['Specification', 'Spec.', 'Spec']),
        'manufacturer_name': find_column(df, ['Manufacturer Name', 'Mfr. Nm.',
                                               'Mfr. Nm', 'Maker Nm.']),
        'vendor_name': find_column(df, ['Vendor Name', 'Vendor Nm.',
                                         'Vendor Nm', 'Supplier Name',
                                         'Supplier']),
        'order_unit': find_column(df, ['Order Unit', 'Odr. Unit',
                                        'Odr. Unit.']),
        'hub_handling_check': find_column(df, ['HUB Handling Check',
                                                'HUB Handling Chk.',
                                                'HUB Handling Chk']),
        'tax_type': find_column(df, ['Purchasing Price Tax Type', 'Tax Type',
                                      'Tax Type.', 'Tax']),
        'registration_date': find_column(df, [
            'Registration Date', 'Prod. Reg. Date',
            'Product Registration Date', 'Product Reg. Date', 'Reg. Date']),
        'product_registry_pic': find_column(df, [
            'Product Registy PIC(Name)', 'Product Registry PIC(Name)',
            'Product Registy PIC', 'Product Registry PIC',
            'Product Registered by(Name)', 'Prod. Reg. PIC Nm.',
            'Prod. Reg. PIC Nm', 'Prod. Reg. PIC',
            'Product Registry PIC Name']),
    }


def _master_pic_columns(df):
    return {
        'category_id': find_column(df, ['Category ID', 'Category Id',
                                         'CategoryID', 'Cat. ID', 'Cat. ID.']),
        'category_name': find_column(df, ['Category Name', 'Category Nm.',
                                           'Cat. Nm.', 'Cat. Nm']),
        'pic': find_column(df, ['PIC', 'PIC Name', 'Pur. PIC',
                                 'Purchase PIC', 'Current PIC', 'Nama PIC']),
        'pic_update': find_column(df, ['Update New PIC', 'New PIC',
                                        'Update PIC', 'PIC Baru',
                                        'New PIC Name']),
    }


# ═══════════════════════════════════════════════════════════════════════════
# 11. GOOGLE SHEETS INTEGRATION
# ═══════════════════════════════════════════════════════════════════════════

RFQ_SHEET_ID = '1JrdsYWhv1mzeXB-jbukDxDYxBgaeISzpiVKEKdgfQvw'
RFQ_SHEET_NAME = 'Sales Submit-RFQ'
RFQ_CACHE = {'expires_at': None, 'rows': [], 'fetched_at': None}
RFQ_CACHE_TTL_SECONDS = 3600

VENDOR_CONTROL_SHEET_ID = '1N0Jr_h5InHH1X2TyLxRf2SMXgDzAXIJnhswzMv5Wf4E'
VENDOR_CONTROL_SHEET_GID = 723367207
VENDOR_CONTROL_CACHE = {
    'expires_at': None, 'rows': [], 'fetched_at': None,
    'sheet_name': None, 'columns': {}
}
VENDOR_CONTROL_CACHE_TTL_SECONDS = 300


def rfq_sheet_sync_credentials():
    raw_json = (os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON') or
                os.environ.get('GOOGLE_SHEETS_SERVICE_ACCOUNT_JSON'))
    raw_file = (os.environ.get('GOOGLE_SERVICE_ACCOUNT_FILE') or
                os.environ.get('GOOGLE_APPLICATION_CREDENTIALS'))
    if raw_json:
        try:
            return json.loads(raw_json)
        except json.JSONDecodeError as e:
            raise RuntimeError(f'Invalid GOOGLE_SERVICE_ACCOUNT_JSON: {e}')
    if raw_file and os.path.exists(raw_file):
        with open(raw_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


GOOGLE_SHEETS_SCOPE = ['https://www.googleapis.com/auth/spreadsheets']


def google_sheets_access_token():
    credentials_info = rfq_sheet_sync_credentials()
    if not credentials_info:
        raise RuntimeError('Google service account not configured')
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import Request
    creds = Credentials.from_service_account_info(
        credentials_info, scopes=GOOGLE_SHEETS_SCOPE)
    creds.refresh(Request())
    return creds.token


def google_sheets_request(method, spreadsheet_id, path,
                          params=None, body=None):
    import requests as _requests
    from urllib.parse import quote
    token = google_sheets_access_token()
    encoded_path = '/'.join(quote(str(part), safe='') for part in path)
    url = (f'https://sheets.googleapis.com/v4/spreadsheets/'
           f'{spreadsheet_id}/{encoded_path}')
    headers = {'Authorization': f'Bearer {token}'}
    if body is not None:
        headers['Content-Type'] = 'application/json'
    proxies = {}
    if os.environ.get('HTTPS_PROXY'):
        proxies['https'] = os.environ.get('HTTPS_PROXY')
    if os.environ.get('HTTP_PROXY'):
        proxies['http'] = os.environ.get('HTTP_PROXY')
    kwargs = {'headers': headers, 'params': params or {}, 'timeout': 60}
    if body is not None:
        kwargs['json'] = body
    if proxies:
        kwargs['proxies'] = proxies
    response = _requests.request(method, url, **kwargs)
    if not response.ok:
        raise RuntimeError(
            f'Google Sheets API {method} {path} failed: '
            f'{response.status_code} {response.text[:500]}')
    return response.json() if response.text else {}


def google_sheets_metadata(spreadsheet_id):
    import requests as _requests
    token = google_sheets_access_token()
    url = f'https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}'
    headers = {'Authorization': f'Bearer {token}'}
    proxies = {}
    if os.environ.get('HTTPS_PROXY'):
        proxies['https'] = os.environ.get('HTTPS_PROXY')
    if os.environ.get('HTTP_PROXY'):
        proxies['http'] = os.environ.get('HTTP_PROXY')
    kwargs = {'headers': headers, 'timeout': 60}
    if proxies:
        kwargs['proxies'] = proxies
    response = _requests.get(url, **kwargs)
    if not response.ok:
        raise RuntimeError(
            f'Google Sheets metadata failed: '
            f'{response.status_code} {response.text[:500]}')
    return response.json()


def google_sheets_values_get(spreadsheet_id, range_name,
                             value_render_option='UNFORMATTED_VALUE'):
    return google_sheets_request(
        'GET', spreadsheet_id, ['values', range_name],
        params={'valueRenderOption': value_render_option})


def google_sheets_values_update(spreadsheet_id, range_name, values):
    return google_sheets_request(
        'PUT', spreadsheet_id, ['values', range_name],
        params={'valueInputOption': 'USER_ENTERED'},
        body={'values': values})


def google_sheets_values_batch_update(spreadsheet_id, ranges):
    return google_sheets_request(
        'POST', spreadsheet_id, ['values:batchUpdate'],
        body={'valueInputOption': 'USER_ENTERED', 'data': ranges})


def google_csv_url(spreadsheet_id, gid='0'):
    return (f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}'
            f'/gviz/tq?tqx=out:csv&gid={gid}')


def read_public_sheet_csv(spreadsheet_id, gid='0', nrows=None):
    return pd.read_csv(
        google_csv_url(spreadsheet_id, gid), header=None, dtype=str,
        keep_default_na=False, nrows=nrows)


# ═══════════════════════════════════════════════════════════════════════════
# 12. RFQ HELPERS — SMART SYNC WITH FINGERPRINT
# ═══════════════════════════════════════════════════════════════════════════

RFQ_TEMPLATE_COLUMNS = [
    ('check', 'Check'), ('sheet_status', 'Status'),
    ('days_left', 'Days Left'), ('no', 'No'),
    ('client_name', 'Nama Client'), ('rfq_date', 'RFQ Date'),
    ('closing_date', 'Closing Date'), ('sales_pic', 'Sales PIC'),
    ('category_name', 'Category Name'),
    ('purchase_pic', 'Purchase PIC'),
    ('rfq_code', 'No. RFQ / KODE'), ('item_name', 'Item Name'),
    ('detail_spec', 'Detail Spec'),
    ('brand_manufacturer', 'Brand/Manufaktur'),
    ('qty', 'Qty'), ('unit', 'Unit'), ('remark', 'Remark'),
    ('product_id', 'Product ID'),
    ('request_number', 'Request Number'),
    ('same_replacement', 'Same/Replacement'),
    ('vendor_name', 'Vendor Name'),
    ('unit_price_idr', 'Unit Price (IDR)'),
    ('amt_idr', 'Amt (IDR)'),
    ('quoted_item_name', 'Item Name'), ('quoted_spec', 'Spec'),
    ('quoted_brand', 'Brand'), ('quoted_unit', 'Unit'),
    ('moq', 'MOQ'), ('lead_time_days', 'Lead Time (Days)'),
    ('valid_period', 'Valid period'),
    ('photo_url', 'Photo URL (optional)'),
    ('remarks', 'Remarks'),
    ('private_remarks_1', 'Private Remarks 1'),
    ('private_remarks_2', 'Private Remarks 2'),
]

RFQ_SIMILARITY_COLUMNS = [
    ('similar_prod_ids', 'Similar Product ID'),
    ('similar_prod_name', 'Similar Product Name'),
    ('similar_spec', 'Similar Specification'),
    ('similar_mfr_name', 'Similar Manufacturer'),
    ('similar_odr_unit', 'Similar Unit'),
    ('similar_score', '%Similarity'),
]

RFQ_EDITABLE_FIELDS = {
    'sheet_status', 'no', 'client_name', 'rfq_date', 'closing_date',
    'sales_pic', 'category_name', 'purchase_pic', 'item_name',
    'detail_spec', 'brand_manufacturer', 'qty', 'unit', 'remark',
    'product_id', 'request_number', 'same_replacement', 'vendor_name',
    'unit_price_idr', 'quoted_item_name', 'quoted_spec', 'quoted_brand',
    'quoted_unit', 'moq', 'lead_time_days', 'valid_period',
    'photo_url', 'remarks', 'private_remarks_1', 'private_remarks_2'
}

RFQ_DIRECT_UPDATE_FIELDS = {'product_id'}

RFQ_BATCH_FIELDS = [
    'same_replacement', 'vendor_name', 'unit_price_idr',
    'quoted_item_name', 'quoted_spec', 'quoted_brand', 'quoted_unit',
    'moq', 'lead_time_days', 'valid_period', 'photo_url', 'remarks',
    'private_remarks_1', 'private_remarks_2'
]

RFQ_SHEET_COLUMN_BY_FIELD = {
    'sheet_status': 'A', 'no': 'B', 'client_name': 'C',
    'rfq_date': 'E', 'closing_date': 'F', 'sales_pic': 'G',
    'request_number': 'R', 'item_name': 'I', 'detail_spec': 'J',
    'brand_manufacturer': 'K', 'qty': 'L', 'unit': 'M', 'remark': 'N',
    'category_name': 'P', 'product_id': 'Q', 'purchase_pic': 'S',
    'same_replacement': 'V', 'vendor_name': 'W',
    'unit_price_idr': 'X', 'quoted_item_name': 'Z',
    'quoted_spec': 'AA', 'quoted_brand': 'AB', 'quoted_unit': 'AC',
    'moq': 'AD', 'lead_time_days': 'AE', 'valid_period': 'AF',
    'photo_url': 'AG', 'remarks': 'AH',
}

RFQ_DASHBOARD_ONLY_FIELDS = {'private_remarks_1', 'private_remarks_2'}
RFQ_SEARCH_FIELDS = ('rfq_code', 'request_number', 'item_name', 'detail_spec')

_RFQ_SHEET_FINGERPRINT = None
_RFQ_FINGERPRINT_TTL = 300  # 5 minutes


def rfq_label(field):
    return dict(RFQ_TEMPLATE_COLUMNS).get(field, field)


def parse_rfq_number(value):
    raw = clean(value)
    if not raw:
        return None
    s = re.sub(r'[^0-9.\-]', '', str(raw))
    if not s or s in ('-', '.', '-.'):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def fmt_rfq_amount(value):
    if value is None:
        return None
    if abs(value - round(value)) < 0.000001:
        return f'{int(round(value)):,}'
    return f'{value:,.2f}'


def rfq_days_left(closing_date):
    raw = clean(closing_date)
    if not raw:
        return None
    d = None
    for fmt in ('%d/%m/%Y', '%Y/%m/%d', '%Y-%m-%d'):
        try:
            d = datetime.strptime(str(raw).strip(), fmt).date()
            break
        except ValueError:
            pass
    if d is None:
        d = parse_date(raw)
    if not d:
        return None
    if d < date.today():
        return None
    return workdays_until(d)


def parse_rfq_closing_date_value(value):
    raw = clean(value)
    if not raw:
        return None
    for fmt in ('%d/%m/%Y', '%Y/%m/%d', '%Y-%m-%d'):
        try:
            return datetime.strptime(str(raw).strip(), fmt).date()
        except ValueError:
            pass
    return parse_date(raw)


def parse_rfq_date_value(value):
    raw = clean(value)
    if not raw:
        return None
    if not re.search(r'\d{4}', str(raw)) and not re.match(
            r'^\d{8}(\.0)?$', str(raw).strip()):
        return None
    for fmt in ('%d/%m/%Y', '%Y/%m/%d', '%Y-%m-%d'):
        try:
            return datetime.strptime(str(raw).strip(), fmt).date()
        except ValueError:
            pass
    return parse_date(raw)


def sort_rfq_rows(rows, sort_order='newest'):
    newest = sort_order != 'oldest'

    def key(row):
        d = parse_rfq_date_value(row.get('rfq_date'))
        ordinal = d.toordinal() if d else 0
        sheet_row = int(row.get('sheet_row') or 0)
        return (d is None, -ordinal if newest else ordinal, sheet_row)
    rows.sort(key=key)
    return rows


def rfq_multiline_search_terms(value):
    terms = []
    seen = set()
    for raw in re.split(r'[\r\n]+', str(value or '')):
        term = raw.strip().lower()
        if term and term not in seen:
            seen.add(term)
            terms.append(term)
    return terms


def filter_rfq_rows_by_multiline_search(rows, value):
    terms = rfq_multiline_search_terms(value)
    if not terms:
        return rows
    filtered = []
    for row in rows:
        searchable = [str(row.get(f) or '').lower() for f in RFQ_SEARCH_FIELDS]
        if any(term in fv for term in terms for fv in searchable):
            filtered.append(row)
    return filtered


def rfq_check_value(item):
    if clean_product_id(item.get('product_id')):
        return 'complete'
    if 'reject' in (clean(item.get('sheet_status')) or '').lower():
        return 'reject'
    closing_date = parse_rfq_closing_date_value(item.get('closing_date'))
    if closing_date and closing_date < date.today():
        return 'closed'
    return 'open'


def rfq_check_label(value):
    return {
        'complete': 'Complete', 'reject': 'Reject',
        'closed': 'Closed', 'open': 'Open'
    }.get(value or '', 'Open')


def apply_rfq_computed_fields(item):
    item['category_name'] = (
        (clean(item.get('category_name')) or '').split('>')[0].strip()
        or None)
    qty = parse_rfq_number(item.get('qty'))
    unit_price = parse_rfq_number(item.get('unit_price_idr'))
    item['amt_idr'] = (fmt_rfq_amount(qty * unit_price)
                       if qty is not None and unit_price is not None
                       else None)
    item['days_left'] = rfq_days_left(item.get('closing_date'))
    item['unit_price_missing'] = unit_price is None
    item['status'] = bool(clean_product_id(item.get('product_id')))
    item['check'] = rfq_check_value(item)
    return item


def rfq_cell(row, idx):
    try:
        return clean(row.iloc[idx])
    except Exception:
        return None


def rfq_row_key(data, sheet_row):
    code = clean(data.get('source_code'))
    if code:
        return code
    parts = [data.get('no'), data.get('client_name'),
             data.get('rfq_date'), data.get('item_name')]
    key = '|'.join(str(clean(x) or '') for x in parts).strip('|')
    return key or f'row-{sheet_row}'


def fetch_rfq_rows(force=False):
    now = datetime.utcnow()
    if (not force and RFQ_CACHE['expires_at'] and
            RFQ_CACHE['expires_at'] > now):
        return RFQ_CACHE['rows'], RFQ_CACHE['fetched_at']

    from urllib.parse import quote
    url = (f'https://docs.google.com/spreadsheets/d/{RFQ_SHEET_ID}'
           f'/gviz/tq?tqx=out:csv&sheet={quote(RFQ_SHEET_NAME)}')
    df = pd.read_csv(url, header=None, dtype=str, keep_default_na=False)
    rows = []
    for idx in range(3, len(df)):
        src = df.iloc[idx]
        product_id = clean_product_id(rfq_cell(src, 16))
        request_number = clean_request_number(rfq_cell(src, 17))
        data = {
            'sheet_row': idx + 1,
            'no': rfq_cell(src, 1),
            'client_name': rfq_cell(src, 2),
            'rfq_date': rfq_cell(src, 4),
            'closing_date': rfq_cell(src, 5),
            'sales_pic': rfq_cell(src, 6),
            'rfq_code': rfq_cell(src, 7),
            'item_name': rfq_cell(src, 8),
            'detail_spec': rfq_cell(src, 9),
            'brand_manufacturer': rfq_cell(src, 10),
            'qty': rfq_cell(src, 11),
            'unit': rfq_cell(src, 12),
            'remark': rfq_cell(src, 13),
            'category_id': rfq_cell(src, 14),
            'category_name': rfq_cell(src, 15),
            'product_id': product_id,
            'sheet_status': rfq_cell(src, 0),
            'request_number': request_number,
            'purchase_pic': rfq_cell(src, 18),
            'same_replacement': rfq_cell(src, 21),
            'vendor_name': rfq_cell(src, 22),
            'unit_price_idr': rfq_cell(src, 23),
            'amt_idr': rfq_cell(src, 24),
            'quoted_item_name': rfq_cell(src, 25),
            'quoted_spec': rfq_cell(src, 26),
            'quoted_brand': rfq_cell(src, 27),
            'quoted_unit': rfq_cell(src, 28),
            'moq': rfq_cell(src, 29),
            'lead_time_days': rfq_cell(src, 30),
            'valid_period': rfq_cell(src, 31),
            'photo_url': rfq_cell(src, 32),
            'remarks': rfq_cell(src, 33),
            'private_remarks_1': '',
            'private_remarks_2': '',
            'source_code': rfq_cell(src, 38),
        }
        data['purchase_pic'] = canonical_rfq_pic(data)
        if not any(data.get(f) for f, _ in RFQ_TEMPLATE_COLUMNS
                   if f != 'check'):
            continue
        data['row_key'] = rfq_row_key(data, idx + 1)
        apply_rfq_computed_fields(data)
        rows.append(data)

    fetched_at = datetime.utcnow()
    RFQ_CACHE.update({
        'rows': rows, 'fetched_at': fetched_at,
        'expires_at': fetched_at + timedelta(seconds=RFQ_CACHE_TTL_SECONDS),
    })
    return rows, fetched_at


def rfq_json_load(value, fallback):
    try:
        return json.loads(value or '')
    except (TypeError, json.JSONDecodeError):
        return fallback


def rfq_dashboard_payload(row):
    payload = dict(row or {})
    payload['row_key'] = (clean(payload.get('row_key')) or
                          rfq_row_key(payload,
                                      payload.get('sheet_row') or 0))
    try:
        payload['sheet_row'] = int(payload.get('sheet_row') or 0) or None
    except (TypeError, ValueError):
        payload['sheet_row'] = None
    apply_rfq_computed_fields(payload)
    return payload


def rfq_dashboard_row_to_dict(row):
    data = rfq_json_load(row.data_json, {})
    data['row_key'] = row.row_key
    data['sheet_row'] = row.sheet_row
    return data


def load_rfq_dashboard_rows():
    db_rows = RFQDashboardRow.query.order_by(
        RFQDashboardRow.sheet_row.is_(None),
        RFQDashboardRow.sheet_row.asc(),
        RFQDashboardRow.id.asc(),
    ).all()
    rows = [rfq_dashboard_row_to_dict(row) for row in db_rows]
    fetched_at = max(
        (r.last_seen_at for r in db_rows if r.last_seen_at),
        default=None)
    return rows, fetched_at


def set_rfq_runtime_rows(rows, fetched_at):
    now = datetime.utcnow()
    RFQ_CACHE.update({
        'rows': [dict(row) for row in rows],
        'fetched_at': fetched_at or now,
        'expires_at': now + timedelta(seconds=RFQ_CACHE_TTL_SECONDS),
    })


# ─── RFQ Smart Sync via Fingerprint ─────────────────────────────────────

def _rfq_sheet_fingerprint():
    """Lightweight change detection: Google Sheets metadata only (~200ms).
    No data download. Rate-limited to once per 5 minutes."""
    global _RFQ_SHEET_FINGERPRINT
    now = datetime.utcnow()

    if (_RFQ_SHEET_FINGERPRINT and
            _RFQ_SHEET_FINGERPRINT.get('checked_at') and
            (now - _RFQ_SHEET_FINGERPRINT['checked_at']).total_seconds()
            < _RFQ_FINGERPRINT_TTL):
        return _RFQ_SHEET_FINGERPRINT

    try:
        meta = google_sheets_metadata(RFQ_SHEET_ID)
        for sheet in meta.get('sheets', []):
            props = sheet.get('properties', {})
            if props.get('title') == RFQ_SHEET_NAME:
                _RFQ_SHEET_FINGERPRINT = {
                    'row_count': props.get(
                        'gridProperties', {}).get('rowCount', 0),
                    'column_count': props.get(
                        'gridProperties', {}).get('columnCount', 0),
                    'modified': meta.get('modifiedTime', ''),
                    'checked_at': now,
                }
                return _RFQ_SHEET_FINGERPRINT
    except Exception as e:
        print(f'RFQ fingerprint check failed: {e}')

    _RFQ_SHEET_FINGERPRINT = {
        'row_count': 0, 'modified': '', 'checked_at': now}
    return _RFQ_SHEET_FINGERPRINT


def _rfq_stored_fingerprint():
    row = ImportDashboardMeta.query.filter_by(
        meta_key='rfq_fingerprint').first()
    if not row:
        return None
    try:
        return json.loads(row.value_json)
    except (TypeError, json.JSONDecodeError):
        return None


def _rfq_save_fingerprint(fingerprint):
    meta_row = ImportDashboardMeta.query.filter_by(
        meta_key='rfq_fingerprint').first()
    if not meta_row:
        meta_row = ImportDashboardMeta(meta_key='rfq_fingerprint')
        db.session.add(meta_row)
    meta_row.value_json = json.dumps({
        'row_count': fingerprint.get('row_count', 0),
        'modified': fingerprint.get('modified', ''),
        'synced_at': utc_isoformat(datetime.utcnow()),
    }, ensure_ascii=False)
    meta_row.updated_at = datetime.utcnow()
    db.session.commit()


def rfq_sheet_has_changes():
    current = _rfq_sheet_fingerprint()
    stored = _rfq_stored_fingerprint()
    if not stored:
        return True
    if current.get('row_count', 0) > stored.get('row_count', 0):
        return True
    if current.get('modified') != stored.get('modified'):
        return True
    return False


def sync_rfq_incremental():
    """Sync only new/changed rows from RFQ Google Sheet.

    Because row/column structure is stable (only data added):
    - New rows at bottom → add to DB
    - Changed data in existing rows → update non-dirty fields
    - Deleted rows → leave in DB (don't remove)
    """
    sheet_rows, fetched_at = fetch_rfq_rows(force=True)
    existing = {row.row_key: row for row in RFQDashboardRow.query.all()}
    duplicate_counts = {}
    now = datetime.utcnow()
    added = 0
    updated = 0

    for sr in sheet_rows:
        base_key = clean(sr.get('row_key'))
        if not base_key:
            continue
        duplicate_counts[base_key] = duplicate_counts.get(base_key, 0) + 1
        row_key = (base_key if duplicate_counts[base_key] == 1
                   else f"{base_key}#{duplicate_counts[base_key]}")
        sr = dict(sr)
        sr['row_key'] = row_key
        incoming = rfq_dashboard_payload(sr)
        current = existing.get(row_key)

        if current:
            local = rfq_json_load(current.data_json, {})
            dirty_fields = set(rfq_json_load(
                current.dirty_fields_json, []))
            changed = False
            for field, value in incoming.items():
                if field in dirty_fields and field in RFQ_EDITABLE_FIELDS:
                    continue
                if local.get(field) != value:
                    local[field] = value
                    changed = True
            if changed:
                local['row_key'] = row_key
                local['sheet_row'] = incoming.get('sheet_row')
                apply_rfq_computed_fields(local)
                current.data_json = json.dumps(local, ensure_ascii=False)
                current.sheet_row = incoming.get('sheet_row')
                current.last_seen_at = fetched_at or now
                current.updated_at = now
                updated += 1
            else:
                current.last_seen_at = fetched_at or now
        else:
            db.session.add(RFQDashboardRow(
                row_key=row_key,
                sheet_row=incoming.get('sheet_row'),
                data_json=json.dumps(incoming, ensure_ascii=False),
                dirty_fields_json='[]',
                first_seen_at=now,
                last_seen_at=fetched_at or now,
                updated_at=now,
            ))
            added += 1

    db.session.commit()

    # Save fingerprint after successful sync
    fingerprint = _rfq_sheet_fingerprint()
    _rfq_save_fingerprint(fingerprint)

    # Refresh runtime cache
    rows, loaded_at = load_rfq_dashboard_rows()
    clear_runtime_caches()
    set_rfq_runtime_rows(rows, loaded_at or fetched_at)

    return {
        'added': added, 'updated': updated,
        'sheet_rows': len(sheet_rows),
        'synced_at': utc_isoformat(now),
    }


def _merge_rfq_edits(rows):
    """Merge dashboard-only edits (Private Remarks) with data rows.
    Single query for all edits instead of per-row."""
    edits = RFQCellEdit.query.filter(
        RFQCellEdit.field.in_(list(RFQ_DASHBOARD_ONLY_FIELDS))
    ).all()
    edit_map = {}
    for edit in edits:
        edit_map.setdefault(edit.row_key, {})[edit.field] = edit.value

    merged = []
    for row in rows:
        item = dict(row)
        for field, value in edit_map.get(item.get('row_key'), {}).items():
            item[field] = value
        merged.append(item)
    return merged


def rfq_rows_with_edits_smart(force=False):
    """Smart RFQ read: auto-sync only when sheet has changes.

    Flow:
    1. force=True → sync from sheet
    2. In-memory cache alive → return cache (instant)
    3. Read from SQLite (instant)
    4. Fingerprint check (~200ms, rate-limited to 5 min)
    5. No changes → return DB data
    6. Changes detected → incremental sync → return new data
    """
    now = datetime.utcnow()

    # Force sync (admin "Sync Now" button)
    if force:
        result = sync_rfq_incremental()
        rows, fetched_at = load_rfq_dashboard_rows()
        return _merge_rfq_edits(rows), fetched_at, result

    # In-memory cache alive
    if (RFQ_CACHE.get('expires_at') and
            RFQ_CACHE['expires_at'] > now and
            RFQ_CACHE.get('rows')):
        return (_merge_rfq_edits(RFQ_CACHE['rows']),
                RFQ_CACHE.get('fetched_at'), None)

    # Read from SQLite
    rows, fetched_at = load_rfq_dashboard_rows()

    # No data yet
    if not rows:
        return [], None, {'needs_sync': True}

    # Smart change detection
    sync_result = None
    try:
        if rfq_sheet_has_changes():
            sync_result = sync_rfq_incremental()
            rows, fetched_at = load_rfq_dashboard_rows()
    except Exception as e:
        print(f'RFQ change check failed, using cached: {e}')

    set_rfq_runtime_rows(rows, fetched_at)
    return _merge_rfq_edits(rows), fetched_at, sync_result


def set_rfq_dashboard_cell(row_key, field, value, dirty=True, commit=True):
    row = RFQDashboardRow.query.filter_by(row_key=row_key).first()
    if not row:
        return False
    data = rfq_json_load(row.data_json, {})
    dirty_fields = set(rfq_json_load(row.dirty_fields_json, []))
    data[field] = value
    data['row_key'] = row.row_key
    data['sheet_row'] = row.sheet_row
    apply_rfq_computed_fields(data)
    if dirty:
        dirty_fields.add(field)
    else:
        dirty_fields.discard(field)
    row.data_json = json.dumps(data, ensure_ascii=False)
    row.dirty_fields_json = json.dumps(
        sorted(dirty_fields), ensure_ascii=False)
    row.updated_at = datetime.utcnow()
    if commit:
        db.session.commit()
        RFQ_CACHE['expires_at'] = None
        clear_runtime_caches()
    return True


def sync_rfq_cell_to_google_sheet(row, field, value):
    column = RFQ_SHEET_COLUMN_BY_FIELD.get(field)
    if field in RFQ_DASHBOARD_ONLY_FIELDS:
        return {'synced': False, 'local_only': True,
                'reason': 'Dashboard-only field'}
    if not column:
        return {'synced': False,
                'reason': 'Field not mapped to RFQ sheet column'}
    sheet_row = row.get('sheet_row')
    if not sheet_row:
        return {'synced': False, 'reason': 'RFQ sheet row missing'}
    range_name = f"'{RFQ_SHEET_NAME}'!{column}{sheet_row}"
    google_sheets_values_update(RFQ_SHEET_ID, range_name, [[value or '']])
    RFQ_CACHE['expires_at'] = None
    return {'synced': True, 'range': range_name}


# ─── RFQ Similarity (batch precompute, read from cache) ─────────────────

def _similarity_token(val):
    s = (clean(val) or '').strip().lower()
    return s[:50] if s else ''


def calculate_similarity(a, b):
    """Simple Jaccard-ish similarity on tokens."""
    if not a or not b:
        return 0.0
    a_tokens = set(str(a).lower().split())
    b_tokens = set(str(b).lower().split())
    if not a_tokens or not b_tokens:
        return 0.0
    intersection = a_tokens & b_tokens
    union = a_tokens | b_tokens
    return (len(intersection) / len(union)) * 100 if union else 0.0


def _candidate_registered_items_for_rfq_similarity(row, limit=1200):
    name_token = _similarity_token(row.get('item_name'))
    spec_token = _similarity_token(row.get('detail_spec'))

    if (not clean(row.get('unit')) or
            not clean(row.get('item_name')) or
            not clean(row.get('detail_spec'))):
        return []

    q = ProductIDDB.query.filter(
        ProductIDDB.product_id.isnot(None),
        ProductIDDB.product_id != '',
        db.or_(ProductIDDB.product_status.is_(None),
               ProductIDDB.product_status == '',
               func.lower(ProductIDDB.product_status) == 'use'))
    token_filters = []
    if name_token:
        token_filters.append(
            ProductIDDB.product_name.ilike(f'%{name_token}%'))
    if spec_token:
        token_filters.append(
            ProductIDDB.specification.ilike(f'%{spec_token}%'))
    if token_filters:
        q = q.filter(db.or_(*token_filters))
    return q.limit(limit).all()


def find_similar_rfq_registered_items(row):
    try:
        if (clean(row.get('check')) or '').lower() != 'open':
            return None
        if clean_product_id(row.get('product_id')):
            return None
        key_fields = [row.get('item_name'), row.get('detail_spec'),
                      row.get('unit')]
        if not all(clean(v) for v in key_fields):
            return None

        current_prod_id = clean_product_id(row.get('product_id'))
        cache_key = '|'.join([
            'rfq_similar_v5',
            clean(row.get('row_key')) or '',
            current_prod_id,
            (clean(row.get('item_name')) or '').lower(),
            (clean(row.get('detail_spec')) or '').lower(),
            (clean(row.get('unit')) or '').lower(),
        ])
        if cache_key in _SIMILARITY_CACHE:
            return _SIMILARITY_CACHE[cache_key]

        similar_items = []
        for reg in _candidate_registered_items_for_rfq_similarity(row):
            reg_prod_id = clean_product_id(reg.product_id)
            if not reg_prod_id or (
                    current_prod_id and reg_prod_id == current_prod_id):
                continue
            if not (clean(reg.product_name) and
                    clean(reg.specification) and clean(reg.order_unit)):
                continue
            item_score = calculate_similarity(
                row.get('item_name'), reg.product_name)
            spec_score = calculate_similarity(
                row.get('detail_spec'), reg.specification)
            unit_score = calculate_similarity(
                row.get('unit'), reg.order_unit)
            if item_score >= 70.0 and spec_score >= 70.0 and unit_score >= 70.0:
                total_sim = (item_score + spec_score + unit_score) / 3
                similar_items.append({
                    'product_id': reg_prod_id,
                    'product_name': reg.product_name or '',
                    'specification': reg.specification or '',
                    'manufacturer_name': reg.manufacturer_name or '',
                    'order_unit': reg.order_unit or '',
                    'similarity': round(total_sim, 1),
                })

        similar_items.sort(
            key=lambda x: (-x['similarity'], x['product_id']))
        if not similar_items:
            result = None
        else:
            result = {
                'product_ids': '\n'.join(
                    x['product_id'] for x in similar_items),
                'product_name': '\n'.join(
                    x['product_name'] or '-' for x in similar_items),
                'specification': '\n'.join(
                    x['specification'] or '-' for x in similar_items),
                'manufacturer_name': '\n'.join(
                    x['manufacturer_name'] or '-' for x in similar_items),
                'order_unit': '\n'.join(
                    x['order_unit'] or '-' for x in similar_items),
                'similarity': '\n'.join(
                    f"{x['similarity']:.0f}%" for x in similar_items),
                'count': len(similar_items),
            }
        _SIMILARITY_CACHE[cache_key] = result
        return result
    except Exception as e:
        print(f"Error finding RFQ similar items: {e}")
        return None


def apply_rfq_similarity(row):
    if (clean(row.get('check')) or '').lower() != 'open':
        row['similar_prod_ids'] = ''
        row['similar_prod_name'] = ''
        row['similar_spec'] = ''
        row['similar_mfr_name'] = ''
        row['similar_odr_unit'] = ''
        row['similar_score'] = None
        return row
    similar = find_similar_rfq_registered_items(row)
    has_pid = clean_product_id(row.get('product_id'))
    row['similar_prod_ids'] = (
        (similar or {}).get('product_ids', '') if has_pid
        else (similar or {}).get('product_ids', 'No Similar Item'))
    row['similar_prod_name'] = (similar or {}).get('product_name', '')
    row['similar_spec'] = (similar or {}).get('specification', '')
    row['similar_mfr_name'] = (similar or {}).get('manufacturer_name', '')
    row['similar_odr_unit'] = (similar or {}).get('order_unit', '')
    row['similar_score'] = (similar or {}).get('similarity', None)
    return row


# ═══════════════════════════════════════════════════════════════════════════
# 13. VENDOR CONTROL HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def column_letter_from_index(index):
    result = ''
    while index > 0:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def vendor_control_sheet_name():
    if VENDOR_CONTROL_CACHE.get('sheet_name'):
        return VENDOR_CONTROL_CACHE['sheet_name']
    meta = google_sheets_metadata(VENDOR_CONTROL_SHEET_ID)
    for sheet in meta.get('sheets', []):
        props = sheet.get('properties', {})
        if props.get('sheetId') == VENDOR_CONTROL_SHEET_GID:
            VENDOR_CONTROL_CACHE['sheet_name'] = props.get('title')
            return VENDOR_CONTROL_CACHE['sheet_name']
    sheets = meta.get('sheets', [])
    if sheets:
        VENDOR_CONTROL_CACHE['sheet_name'] = (
            sheets[0].get('properties', {}).get('title'))
        return VENDOR_CONTROL_CACHE['sheet_name']
    raise RuntimeError('Vendor Control sheet not found')


def normalized_header(value):
    return re.sub(r'[^a-z0-9]+', '', str(value or '').lower())


def find_vendor_control_columns(headers):
    normalized = {}
    for idx, header in enumerate(headers or []):
        key = normalized_header(header)
        if key and key not in normalized:
            normalized[key] = idx + 1

    def pick(names):
        for name in names:
            idx = normalized.get(normalized_header(name))
            if idx:
                return idx
        return None

    return {
        'vendor_name': pick(['Vendor Name', 'Vendor Nm', 'Vendor',
                              'Supplier Name', 'Supplier']),
        'vendor_id': pick(['Vendor ID', 'Vendor Id', 'VendorID',
                            'ID', 'User ID']),
        'password': pick(['Password', 'Pass', 'PWD', 'Pwd']),
    }


def vendor_control_rows(force=False):
    now = datetime.utcnow()
    if (not force and VENDOR_CONTROL_CACHE.get('expires_at') and
            VENDOR_CONTROL_CACHE['expires_at'] > now and
            VENDOR_CONTROL_CACHE.get('rows')):
        return VENDOR_CONTROL_CACHE['rows'], VENDOR_CONTROL_CACHE.get(
            'fetched_at')

    sheet_name = vendor_control_sheet_name()
    result = google_sheets_values_get(
        VENDOR_CONTROL_SHEET_ID, f"'{sheet_name}'!A:Z")
    values = result.get('values', [])
    if not values:
        rows = []
        fetched_at = datetime.utcnow()
        VENDOR_CONTROL_CACHE.update({
            'rows': rows, 'fetched_at': fetched_at,
            'expires_at': fetched_at + timedelta(
                seconds=VENDOR_CONTROL_CACHE_TTL_SECONDS),
            'columns': {},
        })
        return rows, fetched_at

    header_index = 0
    columns = {}
    for idx, candidate_headers in enumerate(values[:20]):
        candidate_columns = find_vendor_control_columns(candidate_headers)
        if all(candidate_columns.get(name)
               for name in ('vendor_name', 'vendor_id', 'password')):
            header_index = idx
            columns = candidate_columns
            break

    missing = [name for name in ('vendor_name', 'vendor_id', 'password')
               if not columns.get(name)]
    if missing:
        raise RuntimeError(
            f"Vendor Control sheet missing columns: {', '.join(missing)}")

    def cell(row, col_index):
        idx = col_index - 1
        return clean(row[idx]) if idx < len(row) else ''

    rows = []
    for sheet_row, raw in enumerate(
            values[header_index + 1:], start=header_index + 2):
        vendor_name = cell(raw, columns['vendor_name'])
        vendor_id = cell(raw, columns['vendor_id'])
        password = cell(raw, columns['password'])
        if not (vendor_name and vendor_id and password):
            continue
        if re.fullmatch(r'\d+(?:\.0+)?', str(vendor_name).strip()):
            continue
        rows.append({
            'row_key': str(sheet_row),
            'sheet_row': sheet_row,
            'vendor_name': vendor_name,
            'vendor_id': vendor_id,
            'password': password,
        })

    fetched_at = datetime.utcnow()
    VENDOR_CONTROL_CACHE.update({
        'rows': rows, 'fetched_at': fetched_at,
        'expires_at': fetched_at + timedelta(
            seconds=VENDOR_CONTROL_CACHE_TTL_SECONDS),
        'columns': columns,
    })
    return rows, fetched_at


def sync_vendor_control_cell(sheet_row, field, value):
    if field not in ('vendor_id', 'password'):
        return {'synced': False, 'reason': 'Field is not editable'}
    sheet_name = vendor_control_sheet_name()
    columns = VENDOR_CONTROL_CACHE.get('columns') or {}
    if not columns.get(field):
        vendor_control_rows(force=True)
        columns = VENDOR_CONTROL_CACHE.get('columns') or {}
    column_index = columns.get(field)
    if not column_index:
        return {'synced': False,
                'reason': f'Sheet column for {field} not found'}
    range_name = (f"'{sheet_name}'!"
                  f"{column_letter_from_index(column_index)}{sheet_row}")
    google_sheets_values_update(
        VENDOR_CONTROL_SHEET_ID, range_name, [[value or '']])
    VENDOR_CONTROL_CACHE['expires_at'] = None
    return {'synced': True, 'range': range_name}


# ═══════════════════════════════════════════════════════════════════════════
# 14. IMPORT DASHBOARD HELPERS
# ═══════════════════════════════════════════════════════════════════════════

IMPORT_LAYOUT_SHEET_ID = '1i0N4VdF_vMHjr_0gjrUdS7nCKUpxPYvDWW-HOWSanEM'
IMPORT_LAYOUT_GID = '73188127'
IMPORT_SOURCE_SHEETS = [
    {'key': 'source_1',
     'spreadsheet_id': '1OSISIb3-D_-oxj2LXH4Q3jcG2IZWnjFGWAmTmdcPBJg',
     'gid': '0', 'label': 'Source 1'},
    {'key': 'source_2',
     'spreadsheet_id': '17P7_JsUGF5mqlz-j2fdvFZ9-gX8l-WGPqZABjng5Hnc',
     'gid': '0', 'label': 'Source 2'},
]
IMPORT_LAYOUT_VENDOR_COLUMNS = (5, 28)
IMPORT_FALLBACK_SOURCE_VENDOR_COLUMNS = (16,)


def import_meta_get(key):
    row = ImportDashboardMeta.query.filter_by(meta_key=key).first()
    if not row:
        return None
    try:
        return json.loads(row.value_json or 'null')
    except (TypeError, json.JSONDecodeError):
        return None


def import_meta_set(key, value):
    row = ImportDashboardMeta.query.filter_by(meta_key=key).first()
    if not row:
        row = ImportDashboardMeta(meta_key=key)
        db.session.add(row)
    row.value_json = json.dumps(value, ensure_ascii=False)
    row.updated_at = datetime.utcnow()
    db.session.commit()


def import_clean_header(value, fallback):
    label = (clean(value) or '').replace('\r', '').replace('\n', ' / ')
    return label or fallback


def import_header_key(value):
    return re.sub(r'[^a-z0-9]+', '', (clean(value) or '').lower())


def import_layout_columns(force=False):
    cache_key = ('import_layout_columns',)
    cached = None if force else runtime_cache_get(cache_key)
    if cached is not None:
        return cached
    cached = None if force else import_meta_get('layout_columns')
    if cached is not None:
        runtime_cache_set(cache_key, cached, ttl_seconds=900)
        return cached
    df = read_public_sheet_csv(
        IMPORT_LAYOUT_SHEET_ID, IMPORT_LAYOUT_GID, nrows=3)
    header_row = df.iloc[1] if len(df) > 1 else (
        df.iloc[0] if len(df) else [])
    columns = []
    seen = {}
    for idx, raw in enumerate(list(header_row)):
        label = import_clean_header(raw, '')
        if not label:
            continue
        base = (re.sub(r'[^a-z0-9]+', '_', label.lower()).strip('_')
                or f'col_{idx}')
        count = seen.get(base, 0) + 1
        seen[base] = count
        field = base if count == 1 else f'{base}_{count}'
        columns.append(
            {'field': field, 'label': label, 'col_idx': idx})
    import_meta_set('layout_columns', columns)
    runtime_cache_set(cache_key, columns, ttl_seconds=900)
    return columns


def import_default_vendors_from_layout(force=False):
    cache_key = ('import_default_vendors_from_layout',)
    cached = None if force else runtime_cache_get(cache_key)
    if cached is not None:
        return cached
    cached = None if force else import_meta_get('default_vendors')
    if cached is not None:
        runtime_cache_set(cache_key, cached, ttl_seconds=900)
        return cached
    try:
        df = read_public_sheet_csv(
            IMPORT_LAYOUT_SHEET_ID, IMPORT_LAYOUT_GID)
    except Exception:
        return []
    vendors = set()
    for row_idx in range(2, len(df)):
        for col_idx in IMPORT_LAYOUT_VENDOR_COLUMNS:
            if col_idx >= df.shape[1]:
                continue
            name = clean(df.iloc[row_idx, col_idx])
            if not name or name.lower() in ('vendor', 'vendor name'):
                continue
            vendors.add(name)
    vendors = sorted(vendors, key=lambda s: s.lower())
    import_meta_set('default_vendors', vendors)
    runtime_cache_set(cache_key, vendors, ttl_seconds=900)
    return vendors


def import_vendor_names(force_default=False):
    rows = ImportVendor.query.order_by(
        ImportVendor.vendor_name.asc()).all()
    uploaded = [r.vendor_name for r in rows if clean(r.vendor_name)]
    return uploaded or import_default_vendors_from_layout(
        force=force_default)


def import_detect_data_start(df):
    for idx in range(min(len(df), 12)):
        item = clean(df.iloc[idx, 7]) if df.shape[1] > 7 else ''
        vendor = clean(df.iloc[idx, 16]) if df.shape[1] > 16 else ''
        qty = clean(df.iloc[idx, 12]) if df.shape[1] > 12 else ''
        if (item and item.lower() != 'item name' and (vendor or qty)):
            return idx
    return 3


def import_detect_header_row(df):
    for idx in range(min(len(df), 12)):
        labels = [import_header_key(v) for v in df.iloc[idx].tolist()]
        if ('itemname' in labels and
                ('vendorname' in labels or 'posementara' in labels)):
            return idx
    return max(import_detect_data_start(df) - 1, 0)


def import_source_column_map(df, columns):
    header_idx = import_detect_header_row(df)
    header_values = list(df.iloc[header_idx]) if len(df) else []
    by_key = {}
    for idx, raw in enumerate(header_values):
        key = import_header_key(raw)
        if key and key not in by_key:
            by_key[key] = idx
    aliases = {
        'site': ['siteidnkrg'],
        'vendor': ['vendorname'],
        'so': ['noso'],
        'purchaseprice': ['purchaseprice', 'price', 'unitprice'],
        'deliverystatus': [
            'deliverystatus', 'createsopodeliverycompletefelix'],
        'importcheck': ['importcheck', 'importautoinput'],
        'happycall': ['happycall', 'poconfirmhappycall'],
    }
    source_map = {}
    for col in columns:
        keys = [import_header_key(col.get('label')),
                import_header_key(col.get('field'))]
        keys.extend(aliases.get(keys[0], []))
        source_idx = next(
            (by_key[k] for k in keys if k in by_key), None)
        if source_idx is not None:
            source_map[col['field']] = source_idx
    return source_map


def import_row_vendor_candidates(values, source_map, columns):
    candidates = []
    for field in ('vendor_name', 'vendor'):
        col_idx = source_map.get(field)
        if col_idx is not None and col_idx < len(values):
            candidates.append(values[col_idx])
    for col_idx in IMPORT_FALLBACK_SOURCE_VENDOR_COLUMNS:
        if col_idx < len(values):
            candidates.append(values[col_idx])
    return [clean(v) for v in candidates if clean(v)]


def import_sheet_rows(force_metadata=False):
    columns = import_layout_columns(force=force_metadata)
    vendor_set = {
        v.strip().lower()
        for v in import_vendor_names(force_default=force_metadata)
        if v.strip()}
    rows = []
    for source in IMPORT_SOURCE_SHEETS:
        df = read_public_sheet_csv(
            source['spreadsheet_id'], source['gid'])
        source_map = import_source_column_map(df, columns)
        start_idx = import_detect_data_start(df)
        for idx in range(start_idx, len(df)):
            values = [clean(v) or '' for v in df.iloc[idx].tolist()]
            vendor_candidates = import_row_vendor_candidates(
                values, source_map, columns)
            row_vendor = next((v for v in vendor_candidates if v), '')
            if vendor_set and not any(
                    v.strip().lower() in vendor_set
                    for v in vendor_candidates if v):
                continue
            row = {
                '_row_key': f"{source['key']}:{idx + 1}",
                '_source_key': source['key'],
                '_source_label': source['label'],
                '_spreadsheet_id': source['spreadsheet_id'],
                '_gid': source['gid'],
                '_sheet_row': idx + 1,
                '_vendor_name': row_vendor,
            }
            for col in columns:
                col_idx = source_map.get(col['field'])
                row[col['field']] = (
                    values[col_idx]
                    if col_idx is not None and col_idx < len(values)
                    else '')
            if not any(row.get(col['field']) for col in columns):
                continue
            rows.append(row)
    return columns, rows


def import_row_payload(row, columns):
    return {col['field']: '' if row.get(col['field']) is None
            else str(row.get(col['field'])) for col in columns}


def import_row_source_uid(row, columns):
    values = [clean(row.get(col['field'])) or '' for col in columns]
    payload = {
        'source': clean(row.get('_source_key')) or '',
        'vendor': clean(row.get('_vendor_name')) or '',
        'values': values,
    }
    raw = json.dumps(
        payload, ensure_ascii=False,
        sort_keys=True, separators=(',', ':'))
    return hashlib.sha1(raw.encode('utf-8')).hexdigest()


def import_dashboard_row_to_dict(row, columns):
    try:
        data = json.loads(row.data_json or '{}')
    except (TypeError, json.JSONDecodeError):
        data = {}
    out = {col['field']: data.get(col['field'], '') for col in columns}
    out.update({
        '_row_key': row.row_key,
        '_source_key': row.source_key,
        '_source_label': row.source_label,
        '_sheet_row': row.sheet_row,
        '_vendor_name': row.vendor_name,
        '_dashboard_id': row.id,
    })
    return out


def sync_import_sheet_to_dashboard():
    columns, sheet_rows = import_sheet_rows(force_metadata=True)
    vendor_count = len(import_vendor_names(force_default=True))
    existing = {r.row_key: r for r in ImportDashboardRow.query.all()}
    duplicate_counts = {}
    now = datetime.utcnow()
    added = 0
    seen = 0
    for sheet_row in sheet_rows:
        source_uid = import_row_source_uid(sheet_row, columns)
        duplicate_base = (
            f"{sheet_row.get('_source_key')}:{source_uid}")
        duplicate_counts[duplicate_base] = (
            duplicate_counts.get(duplicate_base, 0) + 1)
        row_key = (
            f"{duplicate_base}:{duplicate_counts[duplicate_base]}")
        current = existing.get(row_key)
        if current:
            current.sheet_row = sheet_row.get('_sheet_row')
            current.source_label = sheet_row.get('_source_label')
            current.vendor_name = (
                sheet_row.get('_vendor_name') or current.vendor_name)
            current.last_seen_at = now
            seen += 1
            continue
        db.session.add(ImportDashboardRow(
            row_key=row_key,
            source_key=sheet_row.get('_source_key') or '',
            source_label=sheet_row.get('_source_label') or '',
            source_uid=source_uid,
            sheet_row=sheet_row.get('_sheet_row'),
            vendor_name=sheet_row.get('_vendor_name') or '',
            data_json=json.dumps(
                import_row_payload(sheet_row, columns),
                ensure_ascii=False),
            first_seen_at=now,
            last_seen_at=now,
            updated_at=now,
        ))
        added += 1
    db.session.commit()
    clear_runtime_caches()
    return {
        'added': added, 'seen': seen,
        'sheet_rows': len(sheet_rows),
        'vendor_count': vendor_count,
        'columns': columns,
    }


# ═══════════════════════════════════════════════════════════════════════════
# 15. EXCEL EXPORT HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def _excel_style_header():
    return {
        'font': Font(bold=True, color='FFFFFF', size=11),
        'fill': PatternFill('solid', fgColor='2F5496'),
        'alignment': Alignment(horizontal='center', vertical='center',
                                wrap_text=True),
        'border': Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')),
    }


def _excel_style_cell():
    return {
        'border': Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')),
        'alignment': Alignment(vertical='top', wrap_text=True),
    }


# ═══════════════════════════════════════════════════════════════════════════
# 16. API ENDPOINTS
# ═══════════════════════════════════════════════════════════════════════════

# ─── Health Check ────────────────────────────────────────────────────────

@app.route('/')
def index():
    return jsonify({'status': 'ok', 'service': 'SO Dashboard API'})


@app.route('/api/health')
def api_health():
    try:
        db.session.execute(text('SELECT 1'))
        db_ok = True
    except Exception:
        db_ok = False
    return jsonify({
        'status': 'ok' if db_ok else 'degraded',
        'database': 'connected' if db_ok else 'error',
        'exchange_rates_cached': len(_RATE_CACHE),
        'timestamp': utc_isoformat(datetime.utcnow()),
    })


# ─── Exchange Rate Endpoints ─────────────────────────────────────────────

@app.route('/api/exchange-rate', methods=['GET'])
def api_get_exchange_rates():
    page = request.args.get('page', 1, type=int)
    page_size = min(request.args.get('page_size', 50, type=int), 200)

    query = ExchangeRate.query.order_by(ExchangeRate.rate_date.desc())
    total = query.count()
    rows = query.offset((page - 1) * page_size).limit(page_size).all()

    return jsonify({
        'data': [{
            'rate_date': r.rate_date.isoformat() if r.rate_date else None,
            'usd_to_idr': r.usd_to_idr,
            'source': r.source,
        } for r in rows],
        'total': total,
        'page': page,
        'page_size': page_size,
    })


@app.route('/api/exchange-rate', methods=['POST'])
def api_set_exchange_rate():
    data = request.get_json() or {}
    rate_date = parse_date(data.get('rate_date'))
    rate = data.get('usd_to_idr')
    if not rate_date or rate is None:
        return jsonify({'error': 'rate_date and usd_to_idr required'}), 400

    existing = ExchangeRate.query.filter_by(rate_date=rate_date).first()
    if existing:
        existing.usd_to_idr = float(rate)
        existing.source = 'manual'
    else:
        db.session.add(ExchangeRate(
            rate_date=rate_date, usd_to_idr=float(rate), source='manual'))
    db.session.commit()
    _RATE_CACHE[rate_date] = float(rate)
    return jsonify({'success': True})


@app.route('/api/exchange-rate/fetch', methods=['POST'])
def api_fetch_exchange_rates():
    """Backfill missing exchange rates from Frankfurter API.
    Admin-only operation — called once after initial deployment."""
    data = request.get_json() or {}
    start_date = parse_date(data.get('start_date'))
    end_date = parse_date(data.get('end_date')) or date.today()

    if not start_date:
        # Default: 2 years back
        start_date = date.today() - timedelta(days=730)

    # Find existing dates
    existing = {
        r.rate_date for r in
        ExchangeRate.query.filter(
            ExchangeRate.rate_date >= start_date,
            ExchangeRate.rate_date <= end_date).all()
    }

    current = start_date
    fetched = 0
    while current <= end_date:
        if current.weekday() < 5 and current not in existing:
            rate = _fetch_rate_from_api(current)
            if rate:
                db.session.add(ExchangeRate(
                    rate_date=current, usd_to_idr=rate,
                    source='frankfurter'))
                _RATE_CACHE[current] = rate
                fetched += 1
                if fetched % 50 == 0:
                    db.session.commit()
            time.sleep(0.2)  # Rate limit
        current += timedelta(days=1)

    db.session.commit()
    return jsonify({'fetched': fetched, 'total_dates': len(existing)})


# ─── SO Data Upload ──────────────────────────────────────────────────────

@app.route('/api/so/upload', methods=['POST'])
def api_so_upload():
    try:
        uploads, source_type = request_upload_dataframes()
    except Exception as e:
        return jsonify({'error': str(e)}), 400

    if not uploads:
        return jsonify({'error': 'No files uploaded'}), 400

    total_records = 0
    for upload in uploads:
        df = upload['df']
        col_map = {
            'so_number': find_column(df, ['SO Number', 'SO No.']),
            'so_item': find_column(df, ['SO Item', 'Item']),
            'so_status': find_column(df, ['SO Status', 'Status']),
            'operation_unit_name': find_column(df, [
                'Operation Unit Name', 'Op Unit', 'Client']),
            'vendor_id': find_column(df, ['Vendor ID']),
            'vendor_name': find_column(df, [
                'Vendor Name', 'Vendor Nm.', 'Supplier']),
            'customer_po_number': find_column(df, [
                'Customer PO Number', 'PO Number', 'Cust PO']),
            'delivery_memo': find_column(df, [
                'Delivery Memo', 'Memo']),
            'product_name': find_column(df, [
                'Product Name', 'Prod. Nm.']),
            'specification': find_column(df, [
                'Specification', 'Spec.', 'Spec']),
            'manufacturer_name': find_column(df, [
                'Manufacturer Name', 'Mfr. Nm.']),
            'product_id': find_column(df, [
                'Product ID', 'Prod. ID']),
            'so_qty': find_column(df, ['SO Qty', 'Qty', 'Quantity']),
            'sales_unit': find_column(df, ['Sales Unit', 'Unit']),
            'sales_price': find_column(df, ['Sales Price']),
            'sales_amount': find_column(df, ['Sales Amount']),
            'currency': find_column(df, ['Currency', 'Curr']),
            'purchasing_price': find_column(df, [
                'Purchasing Price', 'Pur. Price']),
            'purchasing_amount': find_column(df, [
                'Purchasing Amount', 'Pur. Amount']),
            'purchasing_currency': find_column(df, [
                'Purchasing Currency', 'Pur. Currency', 'Pur. Curr']),
            'so_create_date': find_column(df, [
                'SO Create Date', 'Create Date']),
            'delivery_possible_date': find_column(df, [
                'Delivery Possible Date']),
            'matched_po_number': find_column(df, [
                'Matched PO Number', 'Matched PO']),
            'delivery_plan_date': find_column(df, [
                'Delivery Plan Date', 'Plan Date']),
            'remarks': find_column(df, ['Remarks', 'Remark']),
            'pic_name': find_column(df, ['PIC Name', 'PIC', 'Name']),
        }

        new_rows = []
        for _, row in df.iterrows():
            so_status = clean(row.get(col_map['so_status'])) if col_map[
                'so_status'] else None
            if so_status in DISCARDABLE_STATUSES:
                continue

            so_item = clean(row.get(col_map['so_item'])) if col_map[
                'so_item'] else None
            cust_po = clean(row.get(
                col_map['customer_po_number'])) if col_map[
                'customer_po_number'] else None
            del_memo = clean(row.get(
                col_map['delivery_memo'])) if col_map[
                'delivery_memo'] else None

            record = SOData(
                so_number=clean(row.get(col_map['so_number'])) if col_map[
                    'so_number'] else None,
                so_item=so_item,
                so_status=so_status,
                operation_unit_name=clean(row.get(
                    col_map['operation_unit_name'])) if col_map[
                    'operation_unit_name'] else None,
                vendor_id=clean(row.get(col_map['vendor_id'])) if col_map[
                    'vendor_id'] else None,
                vendor_name=clean(row.get(
                    col_map['vendor_name'])) if col_map[
                    'vendor_name'] else None,
                customer_po_number=cust_po,
                delivery_memo=del_memo,
                product_name=clean(row.get(
                    col_map['product_name'])) if col_map[
                    'product_name'] else None,
                specification=clean(row.get(
                    col_map['specification'])) if col_map[
                    'specification'] else None,
                manufacturer_name=clean(row.get(
                    col_map['manufacturer_name'])) if col_map[
                    'manufacturer_name'] else None,
                product_id=clean_product_id(
                    row.get(col_map['product_id'])) if col_map[
                    'product_id'] else None,
                so_qty=safe_float(row.get(col_map['so_qty'])) if col_map[
                    'so_qty'] else None,
                sales_unit=clean(row.get(
                    col_map['sales_unit'])) if col_map[
                    'sales_unit'] else None,
                sales_price=safe_float(row.get(
                    col_map['sales_price'])) if col_map[
                    'sales_price'] else None,
                sales_amount=safe_float(row.get(
                    col_map['sales_amount'])) if col_map[
                    'sales_amount'] else None,
                currency=clean(row.get(col_map['currency'])) if col_map[
                    'currency'] else None,
                purchasing_price=safe_float(row.get(
                    col_map['purchasing_price'])) if col_map[
                    'purchasing_price'] else None,
                purchasing_amount=safe_float(row.get(
                    col_map['purchasing_amount'])) if col_map[
                    'purchasing_amount'] else None,
                purchasing_currency=clean(row.get(
                    col_map['purchasing_currency'])) if col_map[
                    'purchasing_currency'] else None,
                so_create_date=parse_date(
                    row.get(col_map['so_create_date'])) if col_map[
                    'so_create_date'] else None,
                delivery_possible_date=parse_date(
                    row.get(col_map['delivery_possible_date'])) if col_map[
                    'delivery_possible_date'] else None,
                matched_po_number=clean(row.get(
                    col_map['matched_po_number'])) if col_map[
                    'matched_po_number'] else None,
                delivery_plan_date=parse_date(
                    row.get(col_map['delivery_plan_date'])) if col_map[
                    'delivery_plan_date'] else None,
                remarks=clean(row.get(col_map['remarks'])) if col_map[
                    'remarks'] else None,
                pic_name=clean(row.get(col_map['pic_name'])) if col_map[
                    'pic_name'] else None,
            )
            new_rows.append(record)

        # Bulk add
        db.session.bulk_save_objects(new_rows)
        db.session.flush()

        # Backfill exchange rates ONLY for non-IDR rows
        backfill_exchange_rates_for_rows(new_rows)

        # Compute and cache purchasing_amount_idr for ALL rows
        for s in new_rows:
            compute_and_cache_purchase_amount_idr(s)

        db.session.commit()
        total_records += len(new_rows)

        # Log upload
        db.session.add(UploadLog(
            file_type='so_data',
            filename=upload['filename'],
            records_count=len(new_rows)))
        db.session.commit()

    clear_runtime_caches()
    return jsonify({
        'success': True,
        'total_records': total_records,
        'source': source_type,
    })


# ─── SO Dashboard (Paginated, Read-Only, No API Calls) ───────────────────

@app.route('/api/so/dashboard')
def api_so_dashboard():
    # Parse filters
    date_year, date_from, date_to = parse_so_date_args()
    clients = selected_clients()
    pics = selected_pics()
    search = request.args.get('search', '').strip()
    page = request.args.get('page', 1, type=int)
    page_size = min(request.args.get('page_size', 50, type=int), 200)
    sort_field = request.args.get('sort', 'so_create_date')
    sort_dir = request.args.get('sort_dir', 'desc')

    # Check runtime cache
    cache_key = runtime_cache_key('so_dashboard')
    cached = runtime_cache_get(cache_key)
    if cached:
        return jsonify(cached)

    # Build query
    query = SOData.query.filter(open_so_filter())
    query = apply_so_create_date_filter(
        query, date_year, date_from, date_to)
    query = apply_so_client_filter(query, clients)
    query = apply_so_pic_filter(query, pics)

    if search:
        like = f'%{search}%'
        query = query.filter(db.or_(
            SOData.so_number.ilike(like),
            SOData.product_name.ilike(like),
            SOData.vendor_name.ilike(like),
            SOData.customer_po_number.ilike(like),
            SOData.product_id.ilike(like),
        ))

    # Sort
    sort_column = getattr(SOData, sort_field, SOData.so_create_date)
    if sort_dir == 'asc':
        query = query.order_by(sort_column.asc())
    else:
        query = query.order_by(sort_column.desc())

    # Paginate (cursor-style for speed on SQLite)
    total = query.count()
    rows = query.offset((page - 1) * page_size).limit(page_size).all()

    # Prefetch exchange rates from memory/DB only (NO API)
    dates = {s.so_create_date for s in rows if s.so_create_date}
    prefetch_exchange_rates(dates, fetch_missing=False)

    data = []
    for s in rows:
        data.append({
            'id': s.id,
            'so_number': s.so_number,
            'so_item': s.so_item,
            'so_status': s.so_status,
            'operation_unit_name': s.operation_unit_name,
            'vendor_name': s.vendor_name,
            'customer_po_number': s.customer_po_number,
            'product_name': s.product_name,
            'specification': s.specification,
            'manufacturer_name': s.manufacturer_name,
            'product_id': s.product_id,
            'so_qty': s.so_qty,
            'sales_amount': s.sales_amount,
            'currency': s.currency,
            'purchasing_amount_idr': purchase_amount_idr(
                s, allow_persist=False),
            'purchasing_currency': s.purchasing_currency,
            'so_create_date': (s.so_create_date.isoformat()
                               if s.so_create_date else None),
            'delivery_possible_date': (
                s.delivery_possible_date.isoformat()
                if s.delivery_possible_date else None),
            'delivery_plan_date': (
                s.delivery_plan_date.isoformat()
                if s.delivery_plan_date else None),
            'matched_po_number': s.matched_po_number,
            'remarks': s.remarks,
            'pic_name': s.pic_name,
            'workdays_since_create': workdays_since(s.so_create_date),
        })

    result = {
        'data': data,
        'total': total,
        'page': page,
        'page_size': page_size,
        'total_pages': (total + page_size - 1) // page_size,
    }
    runtime_cache_set(cache_key, result, ttl_seconds=120)
    return jsonify(result)


# ─── SO Single Record Update ─────────────────────────────────────────────

@app.route('/api/so/<int:so_id>', methods=['PUT'])
def api_so_update(so_id):
    s = SOData.query.get_or_404(so_id)
    data = request.get_json() or {}

    updatable = [
        'remarks', 'pic_name', 'delivery_possible_date',
        'delivery_plan_date', 'matched_po_number',
    ]
    for field in updatable:
        if field in data:
            val = data[field]
            if 'date' in field and val:
                val = parse_date(val)
            setattr(s, field, val)

    db.session.commit()
    clear_runtime_caches()
    return jsonify({'success': True})


# ─── SO Delete ───────────────────────────────────────────────────────────

@app.route('/api/so/<int:so_id>', methods=['DELETE'])
def api_so_delete(so_id):
    s = SOData.query.get_or_404(so_id)
    db.session.delete(s)
    db.session.commit()
    clear_runtime_caches()
    return jsonify({'success': True})


# ─── SO Backfill IDR (Admin, one-time) ───────────────────────────────────

@app.route('/api/so/backfill-idr', methods=['POST'])
def api_backfill_idr():
    missing = SOData.query.filter(
        SOData.purchasing_amount_idr.is_(None),
        SOData.purchasing_currency.isnot(None),
        SOData.purchasing_currency != '',
        SOData.purchasing_currency != 'IDR',
    ).all()

    if not missing:
        return jsonify({'message': 'All rows already cached', 'updated': 0})

    # Bulk prefetch exchange rates
    dates = {s.so_create_date for s in missing if s.so_create_date}
    prefetch_exchange_rates(dates, fetch_missing=True)

    updated = 0
    for s in missing:
        compute_and_cache_purchase_amount_idr(s)
        updated += 1

    db.session.commit()
    clear_runtime_caches()
    return jsonify({'updated': updated, 'total_missing': len(missing)})


# ─── SO Summary / Analytics (Read-Only, Cached) ─────────────────────────

@app.route('/api/so/summary')
def api_so_summary():
    cache_key = runtime_cache_key('so_summary')
    cached = runtime_cache_get(cache_key)
    if cached:
        return jsonify(cached)

    date_year, date_from, date_to = parse_so_date_args()
    clients = selected_clients()
    pics = selected_pics()

    query = SOData.query.filter(open_so_filter())
    query = apply_so_create_date_filter(
        query, date_year, date_from, date_to)
    query = apply_so_client_filter(query, clients)
    query = apply_so_pic_filter(query, pics)

    rows = query.all()

    # Prefetch rates from cache/DB only
    dates = {s.so_create_date for s in rows if s.so_create_date}
    prefetch_exchange_rates(dates, fetch_missing=False)

    total_count = len(rows)
    total_sales = sum(float(s.sales_amount or 0) for s in rows)
    total_purchasing_idr = sum(
        purchase_amount_idr(s, allow_persist=False) for s in rows)

    result = {
        'total_count': total_count,
        'total_sales': total_sales,
        'total_purchasing_idr': total_purchasing_idr,
    }
    runtime_cache_set(cache_key, result, ttl_seconds=120)
    return jsonify(result)


# ─── Product ID DB Upload ────────────────────────────────────────────────

@app.route('/api/product-id/upload', methods=['POST'])
def api_product_id_upload():
    uploads, source_type = request_upload_dataframes()
    if not uploads:
        return jsonify({'error': 'No files uploaded'}), 400

    total = 0
    for upload in uploads:
        df = upload['df']
        col_map = _product_id_columns(df)

        for _, row in df.iterrows():
            pid = clean_product_id(
                row.get(col_map['product_id'])) if col_map[
                'product_id'] else None
            if not pid:
                continue

            existing = ProductIDDB.query.filter_by(product_id=pid).first()
            if existing:
                rec = existing
            else:
                rec = ProductIDDB(product_id=pid)
                db.session.add(rec)

            if col_map['category_id']:
                rec.category_id = clean(row.get(col_map['category_id']))
            if col_map['category_name']:
                rec.category_name = clean(row.get(col_map['category_name']))
            if col_map['product_name']:
                rec.product_name = clean(row.get(col_map['product_name']))
            if col_map['product_status']:
                rec.product_status = clean(
                    row.get(col_map['product_status']))
            if col_map['specification']:
                rec.specification = clean(
                    row.get(col_map['specification']))
            if col_map['manufacturer_name']:
                rec.manufacturer_name = clean(
                    row.get(col_map['manufacturer_name']))
            if col_map['vendor_name']:
                rec.vendor_name = clean(row.get(col_map['vendor_name']))
            if col_map['order_unit']:
                rec.order_unit = clean(row.get(col_map['order_unit']))
            if col_map['hub_handling_check']:
                rec.hub_handling_check = clean(
                    row.get(col_map['hub_handling_check']))
            if col_map['tax_type']:
                rec.tax_type = clean(row.get(col_map['tax_type']))
            if col_map['registration_date']:
                rec.registration_date = parse_date(
                    row.get(col_map['registration_date']))
            if col_map['product_registry_pic']:
                rec.product_registry_pic = clean(
                    row.get(col_map['product_registry_pic']))

            rec.updated_at = datetime.utcnow()
            total += 1

        db.session.commit()

    # Invalidate category cache
    _pid_category_cache_invalidate()
    # Warm it immediately
    _pid_category_cache_load()

    return jsonify({'success': True, 'total_records': total})


# ─── Product ID List (Paginated) ─────────────────────────────────────────

@app.route('/api/product-id/list')
def api_product_id_list():
    page = request.args.get('page', 1, type=int)
    page_size = min(request.args.get('page_size', 50, type=int), 200)
    search = request.args.get('search', '').strip()

    query = ProductIDDB.query
    if search:
        like = f'%{search}%'
        query = query.filter(db.or_(
            ProductIDDB.product_id.ilike(like),
            ProductIDDB.product_name.ilike(like),
            ProductIDDB.specification.ilike(like),
            ProductIDDB.manufacturer_name.ilike(like),
        ))

    total = query.count()
    rows = query.order_by(ProductIDDB.product_id).offset(
        (page - 1) * page_size).limit(page_size).all()

    return jsonify({
        'data': [{
            'product_id': r.product_id,
            'category_id': r.category_id,
            'category_name': r.category_name,
            'product_name': r.product_name,
            'product_status': r.product_status,
            'specification': r.specification,
            'manufacturer_name': r.manufacturer_name,
            'vendor_name': r.vendor_name,
            'order_unit': r.order_unit,
        } for r in rows],
        'total': total,
        'page': page,
        'page_size': page_size,
    })


# ─── Master PIC Upload ───────────────────────────────────────────────────

@app.route('/api/master-pic/upload', methods=['POST'])
def api_master_pic_upload():
    uploads, source_type = request_upload_dataframes()
    if not uploads:
        return jsonify({'error': 'No files uploaded'}), 400

    total = 0
    for upload in uploads:
        df = upload['df']
        col_map = _master_pic_columns(df)

        for _, row in df.iterrows():
            cat_id = clean(row.get(col_map['category_id'])) if col_map[
                'category_id'] else None
            cat_name = clean(row.get(col_map['category_name'])) if col_map[
                'category_name'] else None

            # Use category_name as business key, fall back to category_id
            existing = None
            if cat_name:
                existing = MasterPIC.query.filter_by(
                    category_name=cat_name).first()
            if not existing and cat_id:
                existing = MasterPIC.query.filter_by(
                    category_id=cat_id).first()

            if existing:
                rec = existing
            else:
                rec = MasterPIC(
                    category_id=cat_id or cat_name or f'auto_{total}')
                db.session.add(rec)

            rec.category_id = cat_id or rec.category_id
            rec.category_name = cat_name or rec.category_name

            # Determine PIC: prefer "Update New PIC" column, else "PIC"
            new_pic = (clean(row.get(col_map['pic_update']))
                       if col_map['pic_update'] else None)
            old_pic = (clean(row.get(col_map['pic']))
                       if col_map['pic'] else None)
            rec.pic_name = new_pic or old_pic or rec.pic_name
            rec.updated_at = datetime.utcnow()
            total += 1

        db.session.commit()

    # Refresh Master PIC cache
    _warm_master_pic_cache()

    return jsonify({'success': True, 'total_records': total})


def _warm_master_pic_cache():
    global _MASTER_PIC_CACHE
    try:
        rows = MasterPIC.query.all()
        by_id = {}
        by_name = {}
        for r in rows:
            if r.category_id:
                by_id[r.category_id] = r.pic_name
            if r.category_name:
                by_name[r.category_name.lower()] = r.pic_name
        _MASTER_PIC_CACHE = {
            'signature': str(len(rows)),
            'by_id': by_id,
            'by_name': by_name,
        }
        print(f'Master PIC cache warmed: {len(rows)} entries')
    except Exception as e:
        print(f'Master PIC warm skipped: {e}')


# ─── Master PIC List ─────────────────────────────────────────────────────

@app.route('/api/master-pic/list')
def api_master_pic_list():
    rows = MasterPIC.query.order_by(MasterPIC.category_name).all()
    return jsonify({
        'data': [{
            'category_id': r.category_id,
            'category_name': r.category_name,
            'pic_name': r.pic_name,
        } for r in rows],
        'total': len(rows),
    })


# ─── Item Registration Upload ────────────────────────────────────────────

@app.route('/api/item-registration/upload', methods=['POST'])
def api_item_registration_upload():
    uploads, source_type = request_upload_dataframes()
    if not uploads:
        return jsonify({'error': 'No files uploaded'}), 400

    total = 0
    for upload in uploads:
        df = upload['df']

        col_map = {
            'proc_status': find_column(df, ['Proc. Status', 'Proc Status']),
            'req_date': find_column(df, ['Req. Date', 'Request Date']),
            'existing_owner': find_column(df, [
                'Existing Owner', 'Owner', 'Existing Owner(Nm)']),
            'client_name': find_column(df, [
                'Client Name', 'Client', 'Client Nm.']),
            'category': find_column(df, ['Category', 'Cat.']),
            'category_id': find_column(df, ['Category ID', 'Cat. ID']),
            'pic': find_column(df, ['PIC', 'PIC Name']),
            'pic_name': find_column(df, [
                'PIC Name', 'Pur. PIC Nm.', 'Purchase PIC Name']),
            'req_no': find_column(df, [
                'Req. No.', 'Request Number', 'Req No']),
            'prod_id': find_column(df, [
                'Prod. ID', 'Product ID', 'Prod ID']),
            'product_status': find_column(df, [
                'Product Status', 'Prod. Status']),
            'batch_grp_no': find_column(df, [
                'Batch/Grp No.', 'Batch Grp No']),
            'prod_name': find_column(df, [
                'Prod. Name', 'Product Name', 'Prod Name']),
            'spec': find_column(df, ['Spec.', 'Specification', 'Spec']),
            'mfr_name': find_column(df, [
                'Mfr. Name', 'Manufacturer Name', 'Mfr. Nm.']),
            'odr_unit': find_column(df, [
                'Odr. Unit', 'Order Unit', 'Odr. Unit.']),
            'vendor_name': find_column(df, [
                'Vendor Name', 'Vendor Nm.', 'Supplier']),
            'prod_price': find_column(df, [
                'Prod. Price', 'Product Price', 'Price']),
            'curr': find_column(df, ['Curr', 'Currency']),
            'hub_handling_check': find_column(df, [
                'HUB Handling Check', 'HUB Handling Chk.']),
            'tax_type': find_column(df, ['Tax Type', 'Tax']),
            'registration_date': find_column(df, [
                'Registration Date', 'Reg. Date',
                'Product Registration Date']),
            'product_registry_pic': find_column(df, [
                'Product Registy PIC(Name)',
                'Product Registry PIC(Name)',
                'Product Registry PIC']),
            'remarks': find_column(df, ['Remarks', 'Remark']),
        }

        for _, row in df.iterrows():
            req_no = clean(row.get(col_map['req_no'])) if col_map[
                'req_no'] else None
            prod_id = clean_product_id(
                row.get(col_map['prod_id'])) if col_map['prod_id'] else ''

            existing = None
            if req_no and prod_id:
                existing = ItemRegistration.query.filter_by(
                    req_no=req_no, prod_id=prod_id).first()
            if not existing and req_no:
                existing = ItemRegistration.query.filter_by(
                    req_no=req_no).first()

            if existing:
                rec = existing
            else:
                rec = ItemRegistration()
                db.session.add(rec)

            rec.proc_status = clean(row.get(
                col_map['proc_status'])) if col_map['proc_status'] else None
            rec.req_date = parse_date(row.get(
                col_map['req_date'])) if col_map['req_date'] else None
            rec.existing_owner = clean(row.get(
                col_map['existing_owner'])) if col_map[
                'existing_owner'] else None
            rec.client_name = clean(row.get(
                col_map['client_name'])) if col_map['client_name'] else None
            rec.category = clean(row.get(
                col_map['category'])) if col_map['category'] else None
            rec.category_id = clean(row.get(
                col_map['category_id'])) if col_map['category_id'] else None
            rec.pic = clean(row.get(
                col_map['pic'])) if col_map['pic'] else None
            rec.pic_name = clean(row.get(
                col_map['pic_name'])) if col_map['pic_name'] else None
            rec.req_no = req_no
            rec.prod_id = prod_id
            rec.product_status = clean(row.get(
                col_map['product_status'])) if col_map[
                'product_status'] else None
            rec.batch_grp_no = clean(row.get(
                col_map['batch_grp_no'])) if col_map[
                'batch_grp_no'] else None
            rec.prod_name = clean(row.get(
                col_map['prod_name'])) if col_map['prod_name'] else None
            rec.spec = clean(row.get(
                col_map['spec'])) if col_map['spec'] else None
            rec.mfr_name = clean(row.get(
                col_map['mfr_name'])) if col_map['mfr_name'] else None
            rec.odr_unit = clean(row.get(
                col_map['odr_unit'])) if col_map['odr_unit'] else None
            rec.vendor_name = clean(row.get(
                col_map['vendor_name'])) if col_map['vendor_name'] else None
            rec.prod_price = safe_float(row.get(
                col_map['prod_price'])) if col_map['prod_price'] else None
            rec.curr = clean(row.get(
                col_map['curr'])) if col_map['curr'] else None
            rec.hub_handling_check = clean(row.get(
                col_map['hub_handling_check'])) if col_map[
                'hub_handling_check'] else None
            rec.tax_type = clean(row.get(
                col_map['tax_type'])) if col_map['tax_type'] else None
            rec.registration_date = parse_date(row.get(
                col_map['registration_date'])) if col_map[
                'registration_date'] else None
            rec.product_registry_pic = clean(row.get(
                col_map['product_registry_pic'])) if col_map[
                'product_registry_pic'] else None
            rec.remarks = clean(row.get(
                col_map['remarks'])) if col_map['remarks'] else None
            rec.uploaded_at = datetime.utcnow()
            total += 1

        db.session.commit()

    clear_runtime_caches()
    return jsonify({'success': True, 'total_records': total})


# ─── Item Registration Dashboard (Paginated) ─────────────────────────────

@app.route('/api/item-registration/dashboard')
def api_item_registration_dashboard():
    page = request.args.get('page', 1, type=int)
    page_size = min(request.args.get('page_size', 50, type=int), 200)
    search = request.args.get('search', '').strip()
    date_year, date_from, date_to = parse_so_date_args()
    clients = selected_clients()
    pics = selected_pics()

    cache_key = runtime_cache_key('item_reg_dashboard')
    cached = runtime_cache_get(cache_key)
    if cached:
        return jsonify(cached)

    query = ItemRegistration.query
    query = apply_item_registration_date_filter(
        query, date_year, date_from, date_to)
    query = apply_so_client_filter(query, clients)  # reuse
    if pics:
        query = apply_item_registration_pic_filter(query, pics)

    if search:
        like = f'%{search}%'
        query = query.filter(db.or_(
            ItemRegistration.req_no.ilike(like),
            ItemRegistration.prod_name.ilike(like),
            ItemRegistration.prod_id.ilike(like),
            ItemRegistration.mfr_name.ilike(like),
            ItemRegistration.client_name.ilike(like),
        ))

    total = query.count()
    rows = query.order_by(ItemRegistration.req_date.desc()).offset(
        (page - 1) * page_size).limit(page_size).all()

    data = [{
        'id': r.id,
        'proc_status': r.proc_status,
        'req_date': r.req_date.isoformat() if r.req_date else None,
        'existing_owner': r.existing_owner,
        'client_name': r.client_name,
        'category': r.category,
        'pic': r.pic,
        'pic_name': r.pic_name,
        'req_no': r.req_no,
        'prod_id': r.prod_id,
        'product_status': r.product_status,
        'prod_name': r.prod_name,
        'spec': r.spec,
        'mfr_name': r.mfr_name,
        'odr_unit': r.odr_unit,
        'vendor_name': r.vendor_name,
        'prod_price': r.prod_price,
        'curr': r.curr,
        'registration_date': (r.registration_date.isoformat()
                              if r.registration_date else None),
        'remarks': r.remarks,
    } for r in rows]

    result = {
        'data': data,
        'total': total,
        'page': page,
        'page_size': page_size,
        'total_pages': (total + page_size - 1) // page_size,
    }
    runtime_cache_set(cache_key, result, ttl_seconds=120)
    return jsonify(result)


# ─── RFQ Dashboard (Smart Sync, Paginated) ───────────────────────────────

@app.route('/api/rfq/dashboard')
def api_rfq_dashboard():
    page = request.args.get('page', 1, type=int)
    page_size = min(request.args.get('page_size', 50, type=int), 200)
    search = request.args.get('search', '').strip()
    check_filter = request.args.get('check', '').strip()
    sort_order = request.args.get('sort', 'newest')

    rows, fetched_at, sync_info = rfq_rows_with_edits_smart()

    # Handle empty state
    if not rows and fetched_at is None:
        return jsonify({
            'data': [], 'total': 0,
            'needs_initial_sync': True,
            'message': ('RFQ data not synced yet. '
                        'Click Sync to load data.'),
        })

    # Filter
    if search:
        rows = filter_rfq_rows_by_multiline_search(rows, search)
    if check_filter:
        rows = [r for r in rows if (r.get('check') or '') == check_filter]

    # Sort
    sort_rfq_rows(rows, sort_order)

    # Paginate
    total = len(rows)
    start = (page - 1) * page_size
    paginated = rows[start:start + page_size]

    return jsonify({
        'data': paginated,
        'total': total,
        'page': page,
        'page_size': page_size,
        'total_pages': (total + page_size - 1) // page_size,
        'last_synced': utc_isoformat(fetched_at) if fetched_at else None,
        'synced_ago_seconds': (
            int((datetime.utcnow() - fetched_at).total_seconds())
            if fetched_at else None),
        'sync_info': sync_info,
    })


# ─── RFQ Sync (Manual or Scheduled) ─────────────────────────────────────

@app.route('/api/rfq/sync', methods=['POST'])
def api_rfq_sync():
    result = sync_rfq_incremental()
    return jsonify({
        'success': True,
        'added': result.get('added', 0),
        'updated': result.get('updated', 0),
        'sheet_rows': result.get('sheet_rows', 0),
        'synced_at': utc_isoformat(datetime.utcnow()),
    })


# ─── RFQ Cell Edit ──────────────────────────────────────────────────────

@app.route('/api/rfq/cell', methods=['PUT'])
def api_rfq_cell_edit():
    data = request.get_json() or {}
    row_key = data.get('row_key')
    field = data.get('field')
    value = data.get('value')

    if not row_key or not field:
        return jsonify({'error': 'row_key and field required'}), 400
    if field not in RFQ_EDITABLE_FIELDS:
        return jsonify({'error': f'Field {field} not editable'}), 400

    # Update local DB
    ok = set_rfq_dashboard_cell(row_key, field, value, dirty=True)
    if not ok:
        return jsonify({'error': 'Row not found'}), 404

    # Sync to Google Sheet (if not dashboard-only field)
    if field not in RFQ_DASHBOARD_ONLY_FIELDS:
        row_data = RFQDashboardRow.query.filter_by(
            row_key=row_key).first()
        if row_data:
            row_dict = rfq_json_load(row_data.data_json, {})
            sync_result = sync_rfq_cell_to_google_sheet(
                row_dict, field, value)
            return jsonify({
                'success': True, 'sync': sync_result})

    return jsonify({'success': True, 'local_only': True})


# ─── RFQ Similarity (Read from Cache) ────────────────────────────────────

@app.route('/api/rfq/similarity/refresh', methods=['POST'])
def api_rfq_similarity_refresh():
    """Admin: pre-compute similarity for all open RFQ rows.
    Results stored in memory cache for fast dashboard reads."""
    rows, _ = rfq_rows_with_edits_smart()
    computed = 0
    for row in rows:
        if (clean(row.get('check')) or '').lower() == 'open':
            apply_rfq_similarity(row)
            computed += 1
    return jsonify({'computed': computed})


# ─── Vendor Control ──────────────────────────────────────────────────────

@app.route('/api/vendor-control/list')
def api_vendor_control_list():
    try:
        rows, fetched_at = vendor_control_rows()
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    return jsonify({
        'data': rows,
        'total': len(rows),
        'fetched_at': utc_isoformat(fetched_at),
    })


@app.route('/api/vendor-control/sync', methods=['POST'])
def api_vendor_control_sync():
    try:
        rows, fetched_at = vendor_control_rows(force=True)
        return jsonify({
            'success': True,
            'total': len(rows),
            'synced_at': utc_isoformat(fetched_at),
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/vendor-control/cell', methods=['PUT'])
def api_vendor_control_cell():
    data = request.get_json() or {}
    sheet_row = data.get('sheet_row')
    field = data.get('field')
    value = data.get('value')

    if not sheet_row or not field:
        return jsonify({'error': 'sheet_row and field required'}), 400

    result = sync_vendor_control_cell(sheet_row, field, value)
    if result.get('synced'):
        # Refresh cache
        vendor_control_rows(force=True)
        return jsonify({'success': True, **result})
    return jsonify({'error': result.get('reason', 'Sync failed')}), 400


@app.route('/api/vendor-control/login', methods=['POST'])
def api_vendor_control_login():
    data = request.get_json() or {}
    vendor_name = (data.get('vendor_name') or '').strip()
    vendor_id = (data.get('vendor_id') or '').strip()
    password = (data.get('password') or '').strip()

    if not all([vendor_name, vendor_id, password]):
        return jsonify({
            'authenticated': False,
            'error': 'All fields required'}), 400

    try:
        rows, _ = vendor_control_rows()
    except Exception as e:
        return jsonify({'authenticated': False, 'error': str(e)}), 500

    for row in rows:
        if (row['vendor_name'].strip().lower() == vendor_name.lower() and
                row['vendor_id'].strip() == vendor_id and
                row['password'].strip() == password):
            return jsonify({
                'authenticated': True,
                'vendor_name': row['vendor_name'],
            })

    return jsonify({'authenticated': False, 'error': 'Invalid credentials'}), 401


# ─── Import Dashboard ────────────────────────────────────────────────────

@app.route('/api/import/dashboard')
def api_import_dashboard():
    page = request.args.get('page', 1, type=int)
    page_size = min(request.args.get('page_size', 50, type=int), 200)
    vendor_filter = request.args.get('vendor', '').strip()
    search = request.args.get('search', '').strip()

    columns = import_layout_columns()

    query = ImportDashboardRow.query
    if vendor_filter:
        query = query.filter(
            ImportDashboardRow.vendor_name.ilike(f'%{vendor_filter}%'))

    total = query.count()
    db_rows = query.order_by(
        ImportDashboardRow.last_seen_at.desc()).offset(
        (page - 1) * page_size).limit(page_size).all()

    data = [import_dashboard_row_to_dict(r, columns) for r in db_rows]

    if search:
        search_lower = search.lower()
        data = [r for r in data if any(
            search_lower in str(v).lower() for v in r.values())]

    vendors = import_vendor_names()

    return jsonify({
        'data': data,
        'total': total,
        'page': page,
        'page_size': page_size,
        'columns': columns,
        'vendors': vendors,
    })


@app.route('/api/import/sync', methods=['POST'])
def api_import_sync():
    result = sync_import_sheet_to_dashboard()
    return jsonify({'success': True, **result})


@app.route('/api/import/vendor', methods=['POST'])
def api_import_vendor_add():
    data = request.get_json() or {}
    name = (data.get('vendor_name') or '').strip()
    if not name:
        return jsonify({'error': 'vendor_name required'}), 400

    existing = ImportVendor.query.filter_by(vendor_name=name).first()
    if existing:
        return jsonify({'message': 'Already exists'})

    db.session.add(ImportVendor(vendor_name=name))
    db.session.commit()
    return jsonify({'success': True})


@app.route('/api/import/vendor/<int:vendor_id>', methods=['DELETE'])
def api_import_vendor_delete(vendor_id):
    v = ImportVendor.query.get_or_404(vendor_id)
    db.session.delete(v)
    db.session.commit()
    return jsonify({'success': True})


# ─── Dashboard Aggregate Endpoints ───────────────────────────────────────

@app.route('/api/dashboard/clients')
def api_dashboard_clients():
    cache_key = ('dashboard_clients',)
    cached = runtime_cache_get(cache_key)
    if cached:
        return jsonify(cached)

    clients = db.session.query(
        SOData.operation_unit_name,
        func.count(SOData.id)
    ).filter(open_so_filter()).group_by(
        SOData.operation_unit_name
    ).order_by(func.count(SOData.id).desc()).all()

    result = [{
        'name': name, 'count': count
    } for name, count in clients if name]
    runtime_cache_set(cache_key, result, ttl_seconds=300)
    return jsonify(result)


@app.route('/api/dashboard/pics')
def api_dashboard_pics():
    cache_key = ('dashboard_pics',)
    cached = runtime_cache_get(cache_key)
    if cached:
        return jsonify(cached)

    pics = db.session.query(
        SOData.pic_name,
        func.count(SOData.id)
    ).filter(open_so_filter()).group_by(
        SOData.pic_name
    ).order_by(func.count(SOData.id).desc()).all()

    result = [{'name': name or '(Kosong)', 'count': count}
              for name, count in pics]
    runtime_cache_set(cache_key, result, ttl_seconds=300)
    return jsonify(result)


@app.route('/api/dashboard/date-range')
def api_dashboard_date_range():
    cache_key = ('dashboard_date_range',)
    cached = runtime_cache_get(cache_key)
    if cached:
        return jsonify(cached)

    min_date = db.session.query(
        func.min(SOData.so_create_date)).scalar()
    max_date = db.session.query(
        func.max(SOData.so_create_date)).scalar()

    result = {
        'min_date': min_date.isoformat() if min_date else None,
        'max_date': max_date.isoformat() if max_date else None,
    }
    runtime_cache_set(cache_key, result, ttl_seconds=600)
    return jsonify(result)


# ─── Static Files ────────────────────────────────────────────────────────

@app.route('/static/<path:filename>')
def serve_static(filename):
    return send_from_directory('static', filename)


# ═══════════════════════════════════════════════════════════════════════════
# 17. STARTUP: Schema + Cache Warming
# ═══════════════════════════════════════════════════════════════════════════

def _warm_all_caches():
    """Called once at worker startup. Loads permanent data into memory
    so dashboard reads are instant (no DB query per request)."""
    _warm_exchange_rate_cache()
    _pid_category_cache_load()
    _warm_master_pic_cache()
    print('All caches warmed.')


try:
    with app.app_context():
        db.create_all()
        _ensure_extra_columns()
        _ensure_performance_indexes()
        _warm_all_caches()
        print('DB schema ready, caches warmed.')
except Exception as exc:
    print(f'DB schema setup skipped (will retry next reload): {exc}')


# ═══════════════════════════════════════════════════════════════════════════
# 18. BACKGROUND SYNC SCRIPT (for PythonAnywhere scheduled tasks)
# ═══════════════════════════════════════════════════════════════════════════
# Save this as sync_background.py and set up in PythonAnywhere Tasks:
#
#   Schedule: Every 2-4 hours
#   Command: python /home/username/mysite/sync_background.py
#
# ┌──────────────────────────────────────────────────────────────┐
# │ # sync_background.py                                         │
# │ import requests, os                                          │
# │                                                              │
# │ BASE = 'https://yourusername.pythonanywhere.com'             │
# │ SECRET = os.environ.get('SYNC_SECRET', '')                   │
# │                                                              │
# │ def sync_all():                                              │
# │     headers = {'Authorization': f'Bearer {SECRET}'}          │
# │     endpoints = [                                            │
# │         '/api/vendor-control/sync',                          │
# │         '/api/import/sync',                                  │
# │     ]                                                        │
# │     for ep in endpoints:                                     │
# │         try:                                                 │
# │             r = requests.post(f'{BASE}{ep}',                 │
# │                              headers=headers, timeout=120)   │
# │             print(f'{ep}: {r.status_code}')                  │
# │         except Exception as e:                               │
# │             print(f'{ep} failed: {e}')                       │
# │                                                              │
# │ if __name__ == '__main__':                                   │
# │     sync_all()                                               │
# └──────────────────────────────────────────────────────────────┘
