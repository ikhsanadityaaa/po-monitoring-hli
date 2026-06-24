import React, { useState, useEffect, useLayoutEffect, useMemo, useCallback, useRef } from 'react';
import { createPortal } from 'react-dom';
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, AreaChart, Area, ComposedChart
} from 'recharts';
import {
  Upload, Download, AlertCircle, CheckCircle, XCircle,
  Package, TrendingUp, TrendingDown, Award, Calendar, ChevronLeft,
  ChevronRight, Moon, Sun, FileText, BarChart3, FileSpreadsheet,
  Filter, X, ChevronDown, ChevronUp, Building2, Search, Loader2,
  EyeOff, Eye, Trash2, RotateCcw, Plus, Coins, Wallet, Mail, Minus,
  Clock, Wrench, Check, Link as LinkIcon, Pin, PinOff, Ship, FolderOpen, Pencil, Printer,
  DollarSign
} from 'lucide-react';
import axios from 'axios';
import { format, parseISO } from 'date-fns';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { useLocation, useNavigate } from 'react-router-dom';

const BACKEND = import.meta.env.VITE_API_URL || 'http://127.0.0.1:5001';
const api = axios.create({ baseURL: BACKEND, timeout: 600000 });
const DASHBOARD_SUMMARY_CACHE_PREFIX = 'po-monitoring:dashboard-summary:';
const DASHBOARD_STATS_CACHE_KEY = 'po-monitoring:dashboard-stats';
const DASHBOARD_AGING_CACHE_KEY = 'po-monitoring:dashboard-aging';
const DASHBOARD_PENDING_TOTAL_CACHE_KEY = 'po-monitoring:dashboard-pending-total';
const PIC_DB_STATUS_CACHE_KEY = 'po-monitoring:pic-db-status';
const DASHBOARD_CACHE_KEY_ALL = '__all__';
// Cache TTL: keeps Dashboard instant on reload, while uploads still clear it explicitly.
// localStorage is used instead of sessionStorage so refresh / browser reopen does not
// force PythonAnywhere to recalculate the heavy KPI + Delivery Completed summary.
const DASHBOARD_CACHE_TTL_MS = 30 * 60 * 1000;

const storageGet = (key) => {
  if (typeof window === 'undefined') return null;
  try {
    return window.localStorage.getItem(key) || window.sessionStorage.getItem(key);
  } catch {
    return null;
  }
};

const storageSet = (key, value) => {
  if (typeof window === 'undefined') return;
  try {
    window.localStorage.setItem(key, value);
  } catch {
    try { window.sessionStorage.setItem(key, value); } catch {}
  }
};

const storageRemoveWhere = (predicate) => {
  if (typeof window === 'undefined') return;
  for (const store of [window.localStorage, window.sessionStorage]) {
    try {
      Object.keys(store)
        .filter(predicate)
        .forEach((key) => store.removeItem(key));
    } catch {}
  }
};

// ── Filter persistence across page refresh ─────────────────────────────────
// Saves all per-page filter state to localStorage so a refresh restores the
// exact same view the user was looking at. Mirrors Google Sheets behavior:
// refresh never loses context.
const FILTER_STATE_KEY_PREFIX = 'po-monitoring:filter-state:';

const loadFilterState = (pageKey) => {
  if (typeof window === 'undefined') return null;
  try {
    const raw = window.localStorage.getItem(`${FILTER_STATE_KEY_PREFIX}${pageKey}`);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return null;
    return parsed;
  } catch {
    return null;
  }
};

const saveFilterState = (pageKey, state) => {
  if (typeof window === 'undefined') return;
  try {
    window.localStorage.setItem(`${FILTER_STATE_KEY_PREFIX}${pageKey}`, JSON.stringify(state));
  } catch {
    /* storage full or disabled — non-fatal */
  }
};

// ── Offline-first edit queue ───────────────────────────────────────────────
// When a cell PUT fails (network down, server 500, etc.), we store the
// pending update here. On reconnect we replay them in order. Mirrors Google
// Sheets: edits never disappear, they sync when the connection is back.
const OFFLINE_QUEUE_KEY = 'po-monitoring:offline-queue';

const loadOfflineQueue = () => {
  if (typeof window === 'undefined') return [];
  try {
    const raw = window.localStorage.getItem(OFFLINE_QUEUE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

const saveOfflineQueue = (queue) => {
  if (typeof window === 'undefined') return;
  try {
    window.localStorage.setItem(OFFLINE_QUEUE_KEY, JSON.stringify(queue));
  } catch {
    /* storage full — drop oldest entries */
    try {
      const trimmed = queue.slice(-50);
      window.localStorage.setItem(OFFLINE_QUEUE_KEY, JSON.stringify(trimmed));
    } catch {}
  }
};

const enqueueOfflineUpdate = (kind, payload) => {
  const queue = loadOfflineQueue();
  queue.push({ kind, payload, queuedAt: Date.now() });
  saveOfflineQueue(queue);
};

const removeFromOfflineQueue = (index) => {
  const queue = loadOfflineQueue();
  queue.splice(index, 1);
  saveOfflineQueue(queue);
};

const normalizeDashboardCacheQuery = (qs = '') => {
  const raw = String(qs || '').replace(/^\?/, '');
  return raw || DASHBOARD_CACHE_KEY_ALL;
};

const dashboardStatsCacheKey = (qs = '') => `${DASHBOARD_STATS_CACHE_KEY}:${normalizeDashboardCacheQuery(qs)}`;
const dashboardAgingCacheKey = (qs = '') => `${DASHBOARD_AGING_CACHE_KEY}:${normalizeDashboardCacheQuery(qs)}`;
const dashboardPendingCacheKey = (qs = '') => `${DASHBOARD_PENDING_TOTAL_CACHE_KEY}:${normalizeDashboardCacheQuery(qs)}`;

const readCachePayload = (key) => {
  try {
    const raw = storageGet(key);
    if (!raw) return null;
    const { at, data } = JSON.parse(raw);
    if (!at || Date.now() - at > DASHBOARD_CACHE_TTL_MS) return null;
    return data ?? null;
  } catch {
    return null;
  }
};

const writeCachePayload = (key, data) => {
  try { storageSet(key, JSON.stringify({ at: Date.now(), data })); } catch {}
};

const readDashboardSummaryCache = (url) => readCachePayload(`${DASHBOARD_SUMMARY_CACHE_PREFIX}${url}`);
const writeDashboardSummaryCache = (url, data) => writeCachePayload(`${DASHBOARD_SUMMARY_CACHE_PREFIX}${url}`, data);

const clearDashboardSummaryCache = () => {
  storageRemoveWhere((key) => key.startsWith(DASHBOARD_SUMMARY_CACHE_PREFIX));
};

// ─── Stats & aging persistent cache ───────────────────────────────────────
const readStatsCache = (key) => readCachePayload(key);
const writeStatsCache = (key, data) => writeCachePayload(key, data);

// ─── PIC DB status persistent cache ───────────────────────────────────────
// Keeps "Prod ID" / "PIC" timestamps visible in the header even if a fetch
// fails (e.g. transient CORS / cold-start error from the backend), instead
// of silently falling back to '-' forever for that session.
const readPicDbStatusCache = () => readCachePayload(PIC_DB_STATUS_CACHE_KEY);
const writePicDbStatusCache = (data) => writeCachePayload(PIC_DB_STATUS_CACHE_KEY, data);

const clearStatsCache = () => {
  storageRemoveWhere((key) => (
    key === DASHBOARD_STATS_CACHE_KEY ||
    key === DASHBOARD_AGING_CACHE_KEY ||
    key === DASHBOARD_PENDING_TOTAL_CACHE_KEY ||
    key.startsWith(`${DASHBOARD_STATS_CACHE_KEY}:`) ||
    key.startsWith(`${DASHBOARD_AGING_CACHE_KEY}:`) ||
    key.startsWith(`${DASHBOARD_PENDING_TOTAL_CACHE_KEY}:`)
  ));
};

const PIE_COLORS = ['#2563EB','#14B8A6','#22C55E','#EF4444','#06B6D4',
                    '#84CC16','#EC4899','#0EA5E9','#F43F5E','#94A3B8'];

const AGING_LABELS = ['0-30','30-90','90-180','180+'];
const AGING_COLORS = { '0-30':'#10B981','30-90':'#0EA5E9','90-180':'#F43F5E','180+':'#EF4444' };

const PAGE_PATHS = {
  dashboard: '/',
  'all-so': '/Pending_Delivery',
  'item-registration': '/Item_Registration',
  rfq: '/RFQ',
  import: '/Import',
  'vendor-control': '/Vendor_Control',
  'all-registered-items': '/Registered_Items',
};

const PATH_PAGES = Object.fromEntries(
  Object.entries(PAGE_PATHS).map(([page, path]) => [path.toLowerCase(), page])
);

const localISODate = (d) => {
  const dt = new Date(d);
  dt.setMinutes(dt.getMinutes() - dt.getTimezoneOffset());
  return dt.toISOString().slice(0, 10);
};

const getDateFilterBounds = (filter) => {
  if (!filter || filter.mode === 'all') return {};
  if (filter.mode === 'range') return { date_from: filter.start || '', date_to: filter.end || '' };

  const now = new Date();
  const start = new Date(now);
  const end = new Date(now);

  if (filter.mode === 'today') {
    return { date_from: localISODate(now), date_to: localISODate(now) };
  }
  if (filter.mode === 'week') {
    const day = start.getDay() || 7;
    start.setDate(start.getDate() - day + 1);
    end.setTime(start.getTime());
    end.setDate(end.getDate() + 6);
  } else if (filter.mode === 'month') {
    start.setDate(1);
    end.setMonth(start.getMonth() + 1, 0);
  } else if (filter.mode === 'year') {
    start.setMonth(0, 1);
    end.setMonth(11, 31);
  } else {
    return {};
  }

  return { date_from: localISODate(start), date_to: localISODate(end) };
};

// ─── Excluded from PO without SO calculation ──────────────────────────
const EXCLUDED_OP_UNITS = new Set(['ACM ENERGY SOLUTIONS (CONSUMABLE)']);

const renderPctLabel = ({ cx, cy, midAngle, innerRadius, outerRadius, percent }) => {
  if (percent < 0.04) return null;
  const RAD = Math.PI / 180;
  const r = innerRadius + (outerRadius - innerRadius) * 0.58;
  const x = cx + r * Math.cos(-midAngle * RAD);
  const y = cy + r * Math.sin(-midAngle * RAD);
  return (
    <text x={x} y={y} fill="white" textAnchor="middle" dominantBaseline="central"
      fontSize={11} fontWeight="bold" style={{textShadow:'0 1px 2px rgba(0,0,0,0.4)'}}>
      {`${(percent*100).toFixed(0)}%`}
    </text>
  );
};

const fmtNum  = (v) => new Intl.NumberFormat('id-ID').format(v || 0);
const fmtCur  = (v) => `IDR ${new Intl.NumberFormat('id-ID', {maximumFractionDigits:0}).format(v || 0)}`;
const fmtCurShort = (v) => {
  const n = parseFloat(v) || 0;
  if (n >= 1e12) return `IDR ${(n/1e12).toFixed(1)}T`;
  if (n >= 1e9)  return `IDR ${(n/1e9).toFixed(1)}B`;
  if (n >= 1e6)  return `IDR ${(n/1e6).toFixed(1)}M`;
  if (n >= 1e3)  return `IDR ${(n/1e3).toFixed(1)}K`;
  return `IDR ${n.toLocaleString('id-ID')}`;
};
const fmtDate = (d) => { try { return d ? format(parseISO(d),'dd MMM yyyy') : '-'; } catch { return d||'-'; } };
const fmtDateTime = (d) => {
  if (!d) return '-';
  try {
    return new Date(d).toLocaleString('en-GB', { day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' });
  } catch {
    return d || '-';
  }
};
// Backend stores last_copy_at as a plain 'YYYY-MM-DD HH:MM' string already in
// WIB (Asia/Jakarta) local time. Do NOT pass it through new Date(...) — browsers
// would interpret a timezone-less string as the browser's own local time, which
// silently shifts the displayed hour for any user outside WIB. Format it as text
// directly so "Last Copy" always shows the real WIB timestamp the backend wrote.
const fmtWibDateTime = (raw) => {
  if (!raw) return '-';
  const m = String(raw).trim().match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
  if (!m) return String(raw);
  const [, y, mo, d, h, mi] = m;
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const monthLabel = months[parseInt(mo, 10) - 1] || mo;
  return `${d} ${monthLabel} ${y} ${h}:${mi} WIB`;
};
const sanitizeFilename = (name) => String(name || 'Export')
  .replace(/[\\/:*?"<>|]+/g, '_')
  .replace(/\s+/g, '_')
  .slice(0, 160);
const downloadStyledExcel = async ({ columns, rows, filename, sheetName = 'Detail' }) => {
  const XLSXStyleModule = await import('xlsx-js-style');
  const XLSXStyle = XLSXStyleModule.default || XLSXStyleModule;
  const JSZipModule = await import('jszip');
  const JSZip = JSZipModule.default || JSZipModule;
  const headers = columns.map(c => c.header);
  const body = (rows || []).map(row => columns.map(col => {
    const raw = typeof col.value === 'function' ? col.value(row) : row?.[col.key];
    return raw == null || raw === '' ? '' : raw;
  }));
  const ws = XLSXStyle.utils.aoa_to_sheet([headers, ...body]);
  const range = XLSXStyle.utils.decode_range(ws['!ref'] || 'A1:A1');
  const border = {
    top: { style: 'thin', color: { rgb: 'D9E2EF' } },
    bottom: { style: 'thin', color: { rgb: 'D9E2EF' } },
    left: { style: 'thin', color: { rgb: 'D9E2EF' } },
    right: { style: 'thin', color: { rgb: 'D9E2EF' } },
  };
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const addr = XLSXStyle.utils.encode_cell({ r: 0, c });
    if (ws[addr]) {
      ws[addr].s = {
        fill: { patternType: 'solid', fgColor: { rgb: '1D4ED8' } },
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border,
      };
    }
  }
  for (let r = 1; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = XLSXStyle.utils.encode_cell({ r, c });
      if (ws[addr]) {
        ws[addr].s = {
          alignment: { vertical: 'center', wrapText: true },
          border,
        };
      }
    }
  }
  ws['!cols'] = columns.map(col => ({ wch: col.width || 16 }));
  ws['!rows'] = [{ hpt: 24 }];
  ws['!autofilter'] = { ref: XLSXStyle.utils.encode_range(range) };
  ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
  const wb = XLSXStyle.utils.book_new();
  XLSXStyle.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  const xlsxArray = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  const zip = await JSZip.loadAsync(xlsxArray);
  const sheetPath = 'xl/worksheets/sheet1.xml';
  const sheetFile = zip.file(sheetPath);
  if (sheetFile) {
    const freezeSheetView = '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>';
    const xml = await sheetFile.async('string');
    zip.file(sheetPath, xml.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, freezeSheetView));
  }
  const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, `${sanitizeFilename(filename)}.xlsx`);
};
const MARGIN_DETAIL_COLUMNS = [
  { header: 'SO Item', value: r => r.so_item || '', width: 18 },
  { header: 'Product', value: r => r.product || '', width: 34 },
  { header: 'Vendor', value: r => r.vendor || '', width: 28 },
  { header: 'Op Unit', value: r => r.operation_unit_name || '', width: 30 },
  { header: 'Sales', value: r => r.sales_amount ?? '', width: 18 },
  { header: 'Purchase', value: r => r.purchase_amount ?? '', width: 18 },
  { header: 'Margin', value: r => r.margin ?? '', width: 18 },
  { header: '% Margin', value: r => r.margin_pct == null ? '' : `${r.margin_pct}%`, width: 12 },
  { header: 'Date', value: r => r.date || '', width: 16 },
  { header: 'Status', value: r => r.so_status || '', width: 20 },
];
const fmtUpdateShort = (d) => {
  if (!d) return '-';
  try {
    return new Date(d).toLocaleString('en-GB', { day:'2-digit', month:'short', hour:'2-digit', minute:'2-digit' });
  } catch {
    return d || '-';
  }
};

const workingDaysUntilToday = (dateValue) => {
  if (!dateValue) return null;
  const start = new Date(dateValue);
  if (Number.isNaN(start.getTime())) return null;

  const today = new Date();
  const cur = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  if (cur > end) return 0;

  let days = 0;
  while (cur <= end) {
    const day = cur.getDay();
    if (day !== 0 && day !== 6) days += 1;
    cur.setDate(cur.getDate() + 1);
  }
  return days;
};

// ─── Import table: Google Drive link chips, status dropdown & checklist columns ────────────
const GDRIVE_URL_RE = /https?:\/\/(?:drive|docs)\.google\.com\/\S+/i;

const extractGDriveUrl = (value) => {
  const text = String(value || '');
  const m = text.match(GDRIVE_URL_RE);
  return m ? m[0].replace(/[),.;]+$/, '') : '';
};

const gDriveChipLabel = (url) => {
  if (/\/folders\//.test(url)) return 'Folder Drive';
  if (/spreadsheets/.test(url)) return 'Spreadsheet';
  if (/\/document\//.test(url)) return 'Dokumen';
  if (/presentation/.test(url)) return 'Slide';
  if (/\/file\/d\//.test(url)) return 'File Drive';
  return 'Buka Drive';
};

const IMPORT_STATUS_OPTIONS = ['ON PROCESS', 'ON DELIVERY', 'DELIVERED', 'CANCELED'];
const IMPORT_CHECKLIST_TRUE = new Set(['true', '1', 'yes', 'ya', 'y', 'checked', 'done', 'ok', '✓', '✅']);
const IMPORT_CHECKLIST_FALSE = new Set(['false', '0', 'no', 'tidak', 'n', 'unchecked', '❌']);
const IMPORT_CHECKLIST_VALUES = new Set([...IMPORT_CHECKLIST_TRUE, ...IMPORT_CHECKLIST_FALSE]);
// sap_input is a checkbox toggle but NOT part of the hide-checklist group
// (always visible). non_ski is a regular editable text cell (not checkbox).
const IMPORT_CHECKLIST_FIELDS = new Set(['bl_awb', 'invoice', 'pl', 'hc', 'msds', 'coa', 'coo']);
// sap_input renders as a checkmark toggle but is NOT hidden behind the
// "Show Checklist" button — it's always visible.
const IMPORT_CHECKBOX_ALWAYS_VISIBLE = new Set(['sap_input']);
const IMPORT_FORMULA_FIELDS = new Set(['days_left', 'site', 'vendor', 'arrival_check', 'purchase_amount', 'lt_days']);

const isImportChecklistColumn = (col) => Boolean(col?.checkbox) || IMPORT_CHECKLIST_FIELDS.has(col?.field);
const isImportHideableChecklistColumn = (col) => isImportChecklistColumn(col) && !IMPORT_CHECKBOX_ALWAYS_VISIBLE.has(col?.field);
const isImportFormulaColumn = (col) => Boolean(col?.formula) || IMPORT_FORMULA_FIELDS.has(col?.field);
const isImportHyperlinkColumn = (col) => Boolean(col?.hyperlink) || col?.field === 'soft_copy_doc';

const importCheckboxChecked = (value) => IMPORT_CHECKLIST_TRUE.has(String(value ?? '').trim().toLowerCase());

const importStatusClass = (status, darkMode = false) => {
  const s = String(status || '').trim().toUpperCase();
  if (s === 'NEW') return darkMode ? 'bg-blue-950/55 text-blue-100 border-blue-600' : 'bg-blue-50 text-blue-700 border-blue-200';
  if (s === 'DELIVERED') return darkMode ? 'bg-green-900/45 text-green-100 border-green-700' : 'bg-green-50 text-green-700 border-green-200';
  // SWAP per user request: ON DELIVERY → yellow, ON PROCESS → blue.
  if (s === 'ON DELIVERY') return darkMode ? 'bg-amber-900/45 text-amber-100 border-amber-700' : 'bg-amber-50 text-amber-700 border-amber-200';
  if (s === 'ON PROCESS') return darkMode ? 'bg-blue-900/45 text-blue-100 border-blue-700' : 'bg-blue-50 text-blue-700 border-blue-200';
  if (s === 'CANCELED') return darkMode ? 'bg-red-900/45 text-red-100 border-red-700' : 'bg-red-50 text-red-700 border-red-200';
  return darkMode ? 'bg-gray-800 text-gray-100 border-gray-600' : 'bg-white text-gray-700 border-gray-200';
};

const importStatusOptionStyle = (status) => {
  const s = String(status || '').trim().toUpperCase();
  if (s === 'NEW') return { backgroundColor: '#DBEAFE', color: '#1D4ED8', fontWeight: '700' };
  if (s === 'DELIVERED') return { backgroundColor: '#DCFCE7', color: '#166534' };
  // SWAP per user request: ON DELIVERY → yellow, ON PROCESS → blue.
  if (s === 'ON DELIVERY') return { backgroundColor: '#FEF3C7', color: '#92400E' };
  if (s === 'ON PROCESS') return { backgroundColor: '#DBEAFE', color: '#1D4ED8' };
  if (s === 'CANCELED') return { backgroundColor: '#FEE2E2', color: '#B91C1C' };
  return {};
};

const importArrivalClass = (value, darkMode = false) => {
  const s = String(value || '').toLowerCase();
  if (s.includes('delay')) return darkMode ? 'bg-red-900/40 text-red-100 border-red-700' : 'bg-red-50 text-red-700 border-red-200';
  if (s.includes('on schedule')) return darkMode ? 'bg-green-900/40 text-green-100 border-green-700' : 'bg-green-50 text-green-700 border-green-200';
  return darkMode ? 'bg-gray-800 text-gray-200 border-gray-700' : 'bg-slate-50 text-slate-700 border-slate-200';
};

const importDisplayValue = (value) => {
  if (value === null || value === undefined || value === '') return '-';
  return String(value);
};


const DownloadToast = ({ message, onClose }) => {
  return (
    <div className="fixed top-5 right-5 z-[200] flex items-center gap-3 px-5 py-3 rounded-xl shadow-2xl text-white bg-blue-700 max-w-sm animate-slide-in">
      <Loader2 className="w-5 h-5 flex-shrink-0 animate-spin"/>
      <span className="text-sm font-medium">{message}</span>
    </div>
  );
};

const Toast = ({ message, type, onClose }) => {
  useEffect(() => { const t = setTimeout(onClose, 3000); return () => clearTimeout(t); }, [onClose]);
  const bg = type === 'success' ? 'bg-green-600' : type === 'error' ? 'bg-red-600' : 'bg-blue-600';
  return (
    <div className={`fixed top-5 right-5 z-[100] flex items-center gap-3 px-5 py-3 rounded-xl shadow-2xl text-white ${bg} max-w-sm`}>
      {type === 'success' ? <CheckCircle className="w-5 h-5 flex-shrink-0" /> : <AlertCircle className="w-5 h-5 flex-shrink-0" />}
      <span className="text-sm font-medium">{message}</span>
      <button onClick={onClose} className="ml-2 hover:opacity-70"><X className="w-4 h-4" /></button>
    </div>
  );
};

// ─── Download Button with press animation ─────────────────────────────────
const DownloadButton = ({ onClick, className, children, disabled }) => {
  const handleClick = () => {
    onClick && onClick();
  };
  return (
    <button
      onClick={handleClick}
      disabled={disabled}
      className={className}
    >
      {children}
    </button>
  );
};

const SOModal = ({ title, data, onClose, darkMode, onUpdateCell }) => {
  const [dlPage, setDlPage] = useState(1);
  const [editing, setEditing] = useState(null);
  const [editValue, setEditValue] = useState('');
  const PER = 50;
  const safeData = Array.isArray(data) ? data.filter(row => row && typeof row === 'object') : [];
  const rowKey = (row) => row?.id || row?.so_item || row?.so_number || '';
  const pages = Math.ceil((safeData.length || 0) / PER);
  const rows = safeData.slice((dlPage-1)*PER, dlPage*PER);

  // Determine if SO Item column exists in data (show SO Number only when SO Item is absent)
  const hasSoItem = safeData.some(s => s.so_item);
  const columns = [
    { header: 'SO Item', value: s => s.so_item || '', width: 18 },
    ...(!hasSoItem ? [{ header: 'SO Number', value: s => s.so_number || '', width: 18 }] : []),
    { header: 'Status', value: s => s.so_status || '', width: 24 },
    { header: 'PIC', value: s => s.pic_name || '', width: 16 },
    { header: 'Op Unit', value: s => s.operation_unit_name || '', width: 30 },
    { header: 'Vendor', value: s => s.vendor_name || '', width: 24 },
    { header: 'Product', value: s => s.product_name || '', width: 34 },
    { header: 'Qty', value: s => s.so_qty ?? '', width: 10 },
    { header: 'Sales Amount', value: s => s.sales_amount ?? '', width: 18 },
    { header: 'Cust PO', value: s => s.customer_po_number || '', width: 18 },
    { header: 'Delivery Memo', value: s => s.delivery_memo || '', width: 28 },
    { header: 'SO Date', value: s => s.so_create_date || '', width: 16 },
    { header: 'Plan Date', value: s => s.delivery_plan_date || '', width: 16 },
    { header: 'Remarks', value: s => s.remarks || '', width: 46 },
  ];

  const downloadExcel = () => {
    downloadStyledExcel({ columns, rows: safeData, filename: title, sheetName: 'Detail' });
  };
  const startEdit = (row, field) => {
    const key = rowKey(row);
    if (!key || !onUpdateCell) {
      console.warn('SO detail row is missing id/SO Item; cannot edit from modal', row);
      return;
    }
    setEditing({ id: key, field });
    setEditValue(row[field] || '');
  };
  const saveEdit = async () => {
    if (!editing || !onUpdateCell) return;
    await onUpdateCell(editing.id, editing.field, editValue);
    setEditing(null);
  };
  return (
    <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm" onClick={onClose}>
      <div role="dialog" aria-modal="true" aria-label={title} className={`rounded-2xl overflow-hidden shadow-2xl w-full max-w-6xl max-h-[85vh] flex flex-col ${darkMode?'bg-gray-800 text-white':'bg-white'}`} onClick={e=>e.stopPropagation()}>
        <div className={`flex justify-between items-center px-6 py-4 border-b ${darkMode?'border-gray-700':'border-gray-100'}`}>
          <h3 className="font-bold text-lg">{title} <span className={`text-sm font-normal ml-2 ${darkMode?'text-gray-400':'text-gray-500'}`}>({fmtNum(safeData.length)} records)</span></h3>
          <div className="flex gap-2">
            <button onClick={downloadExcel} className="flex items-center gap-1 px-3 py-1.5 bg-green-600 hover:bg-green-700 text-white rounded-lg text-sm"><FileSpreadsheet className="w-4 h-4"/>Excel</button>
            <button onClick={onClose} className={`p-1.5 rounded-lg ${darkMode?'hover:bg-gray-700':'hover:bg-gray-100'}`}><X className="w-5 h-5"/></button>
          </div>
        </div>
        <div className="overflow-auto flex-1 rounded-b-2xl">
          <table className="w-full text-sm">
            <thead className={`sticky top-0 ${darkMode?'bg-gray-700':'bg-blue-50'}`}>
              <tr>{columns.map(({ header: h })=>(
                <th key={h} className={`px-3 py-2 text-center font-bold whitespace-nowrap ${darkMode?'text-gray-200':'text-gray-700'}`}>{h}</th>
              ))}</tr>
            </thead>
            <tbody className={`divide-y ${darkMode?'divide-gray-700':'divide-gray-100'}`}>
              {rows.map((s,i)=>(
                <tr key={i} className={darkMode?'hover:bg-gray-700':'hover:bg-blue-50'}>
                  <td className="px-3 py-2 text-blue-600 font-medium whitespace-nowrap">{s.so_item||'-'}</td>
                  {!hasSoItem && <td className="px-3 py-2 whitespace-nowrap">{s.so_number}</td>}
                  <td className="px-3 py-2 whitespace-nowrap"><span className={`px-2 py-0.5 rounded-full text-xs font-medium ${s.so_status==='Delivery Completed'?'bg-green-100 text-green-700':s.so_status==='SO Cancel'?'bg-red-100 text-red-700':'bg-blue-100 text-blue-700'}`}>{s.so_status||'-'}</span></td>
                  <td className="px-3 py-2 whitespace-nowrap text-center font-semibold text-slate-600">{s.pic_name||'-'}</td>
                  <td className="px-3 py-2 whitespace-nowrap min-w-[180px]">{s.operation_unit_name}</td>
                  <td className="px-3 py-2 whitespace-nowrap max-w-[140px] truncate">{s.vendor_name}</td>
                  <td className="px-3 py-2 max-w-[160px] truncate">{s.product_name}</td>
                  <td className="px-3 py-2 text-right">{fmtNum(s.so_qty)}</td>
                  <td className="px-3 py-2 text-center font-bold text-slate-700 whitespace-nowrap">{fmtCur(s.sales_amount)}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{s.customer_po_number||'-'}</td>
                  <td className="px-3 py-2 max-w-[160px] truncate">{s.delivery_memo||'-'}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{s.so_create_date||'-'}</td>
                  <td className="px-3 py-2 whitespace-nowrap text-blue-600">
                    {editing?.id === rowKey(s) && editing.field === 'delivery_plan_date' ? (
                      <input
                        type="date"
                        value={editValue || ''}
                        onChange={e=>setEditValue(e.target.value)}
                        onBlur={saveEdit}
                        onKeyDown={e=>{ if(e.key==='Enter') saveEdit(); if(e.key==='Escape') setEditing(null); }}
                        className={`w-36 px-2 py-1 rounded text-xs border ${darkMode?'bg-gray-700 border-gray-600 text-white':'bg-white border-gray-300 text-gray-900'}`}
                        autoFocus
                      />
                    ) : (
                      <button type="button" onClick={()=>startEdit(s, 'delivery_plan_date')} className="text-blue-600 hover:underline text-xs whitespace-nowrap">
                        {s.delivery_plan_date||'Set'}
                      </button>
                    )}
                  </td>
                  <td className="px-3 py-2 min-w-[560px] max-w-[560px] truncate">
                    {editing?.id === rowKey(s) && editing.field === 'remarks' ? (
                      <input
                        type="text"
                        value={editValue || ''}
                        onChange={e=>setEditValue(e.target.value)}
                        onBlur={saveEdit}
                        onKeyDown={e=>{ if(e.key==='Enter') saveEdit(); if(e.key==='Escape') setEditing(null); }}
                        className={`w-full px-2 py-1 rounded text-xs border ${darkMode?'bg-gray-700 border-gray-600 text-white':'bg-white border-gray-300 text-gray-900'}`}
                        autoFocus
                      />
                    ) : (
                      <button type="button" onClick={()=>startEdit(s, 'remarks')} className="block max-w-full truncate text-left text-blue-600 hover:underline text-xs">
                        {s.remarks||'Add'}
                      </button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {pages > 1 && (
          <div className={`flex justify-between items-center px-6 py-3 border-t ${darkMode?'border-gray-700':'border-gray-100'}`}>
            <span className={`text-sm ${darkMode?'text-gray-400':'text-gray-600'}`}>{(dlPage-1)*PER+1}–{Math.min(dlPage*PER,safeData.length)} / {fmtNum(safeData.length)}</span>
            <div className="flex gap-2">
              <button disabled={dlPage===1} onClick={()=>setDlPage(p=>p-1)} className={`p-1.5 rounded ${dlPage===1?'opacity-40':'hover:bg-gray-200'}`}><ChevronLeft className="w-4 h-4"/></button>
              <span className="px-3 py-1 bg-blue-100 rounded text-sm text-blue-700">{dlPage}/{pages}</span>
              <button disabled={dlPage===pages} onClick={()=>setDlPage(p=>p+1)} className={`p-1.5 rounded ${dlPage===pages?'opacity-40':'hover:bg-gray-200'}`}><ChevronRight className="w-4 h-4"/></button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

// ─── Reusable floating-dropdown hook ──────────────────────────────────────
// Returns a ref (attach to the trigger button) and a `menuPos` object to
// pass as `style` on the floating menu. The menu uses `position: fixed` so
// it escapes any `overflow-hidden` ancestor (e.g. the rounded card wrapper
// around every table). The position is recomputed on scroll/resize and the
// menu flips above the trigger when there isn't enough space below.
//
// Usage:
//   const { triggerRef, menuPos } = useFloatingDropdown(open);
//   <button ref={triggerRef} onClick={...}>Open</button>
//   {open && (
//     <div style={menuPos.style} className="fixed z-[180] ...">
//       ...menu content...
//     </div>
//   )}
const useFloatingDropdown = (open, minWidth = 320, maxWidth = 520, menuHeight = 300) => {
  const triggerRef = useRef(null);
  const [menuPos, setMenuPos] = useState({ top: 0, left: 0, width: minWidth, style: {} });

  useEffect(() => {
    if (!open) return undefined;
    const compute = () => {
      const anchor = triggerRef.current;
      if (!anchor || typeof window === 'undefined') return;
      const rect = anchor.getBoundingClientRect();
      const viewportW = window.innerWidth || 1024;
      const viewportH = window.innerHeight || 768;
      const width = Math.min(Math.max(rect.width, minWidth), Math.min(maxWidth, viewportW - 32));
      const left = Math.min(Math.max(rect.left, 16), viewportW - width - 16);
      const spaceBelow = viewportH - rect.bottom - 12;
      const spaceAbove = rect.top - 12;
      const preferAbove = spaceBelow < menuHeight && spaceAbove > spaceBelow;
      const top = preferAbove
        ? Math.max(8, rect.top - Math.min(menuHeight, spaceAbove) - 10)
        : Math.min(rect.bottom + 6, viewportH - 120);
      setMenuPos({
        top,
        left,
        width,
        style: { position: 'fixed', top: `${top}px`, left: `${left}px`, width: `${width}px`, maxWidth: 'calc(100vw - 32px)', zIndex: 180 },
      });
    };
    compute();
    window.addEventListener('resize', compute);
    window.addEventListener('scroll', compute, true);
    return () => {
      window.removeEventListener('resize', compute);
      window.removeEventListener('scroll', compute, true);
    };
  }, [open, minWidth, maxWidth, menuHeight]);

  return { triggerRef, menuPos };
};

// ─── MultiSelect dropdown — Excel-style (all checked by default) ─────────
const MultiSelect = ({ label, options, selected, onChange, darkMode, txt2, hideLabel = false }) => {
  const [open, setOpen] = useState(false);
  const [draftSelected, setDraftSelected] = useState([]);
  const [draftNone, setDraftNone] = useState(false);
  const [searchText, setSearchText] = useState('');
  const [menuPos, setMenuPos] = useState({ top: 0, left: 0, width: 320, maxHeight: 240 });
  const ref = useRef(null);
  const menuRef = useRef(null);

  const safeOptions = useMemo(() => {
    const seen = new Set();
    return (Array.isArray(options) ? options : [])
      .map(v => String(v ?? '').trim())
      .filter(v => {
        if (!v || seen.has(v)) return false;
        seen.add(v);
        return true;
      });
  }, [options]);

  const noSelection = selected === '__NONE__';
  const safeSelected = Array.isArray(selected) ? selected : [];
  const noneSelected = !noSelection && safeSelected.length === 0;
  const currentSelected = open ? draftSelected : safeSelected;
  const currentNone = open ? draftNone : noSelection;
  const currentAll = !currentNone && currentSelected.length === 0;
  const someSelected = !currentNone && currentSelected.length > 0 && currentSelected.length < safeOptions.length;

  const updateMenuPosition = useCallback(() => {
    const anchor = ref.current;
    if (!anchor || typeof window === 'undefined') return;
    const rect = anchor.getBoundingClientRect();
    const viewportW = window.innerWidth || 1024;
    const viewportH = window.innerHeight || 768;
    const width = Math.min(Math.max(rect.width, 320), Math.min(520, viewportW - 32));
    const left = Math.min(Math.max(rect.left, 16), viewportW - width - 16);
    const spaceBelow = viewportH - rect.bottom - 12;
    const spaceAbove = rect.top - 12;
    const preferAbove = spaceBelow < 260 && spaceAbove > spaceBelow;
    const maxHeight = Math.max(168, Math.min(300, (preferAbove ? spaceAbove : spaceBelow) - 20));
    const top = preferAbove
      ? Math.max(8, rect.top - maxHeight - 10)
      : Math.min(rect.bottom + 6, viewportH - 120);
    setMenuPos({ top, left, width, maxHeight });
  }, []);

  const closeDropdown = useCallback(() => {
    setOpen(false);
    setDraftSelected([]);
    setDraftNone(false);
    setSearchText('');
  }, []);

  useEffect(() => {
    const handler = (e) => {
      const target = e.target;
      if (ref.current?.contains(target) || menuRef.current?.contains(target)) return;
      closeDropdown();
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [closeDropdown]);

  useEffect(() => {
    if (!open) return undefined;
    setDraftSelected(safeSelected);
    setDraftNone(noSelection);
    updateMenuPosition();
    const reposition = () => updateMenuPosition();
    window.addEventListener('resize', reposition);
    window.addEventListener('scroll', reposition, true);
    return () => {
      window.removeEventListener('resize', reposition);
      window.removeEventListener('scroll', reposition, true);
    };
  }, [open, selected, noSelection, updateMenuPosition]);

  const filteredOptions = searchText.trim()
    ? safeOptions.filter(opt => String(opt).toLowerCase().includes(searchText.trim().toLowerCase()))
    : safeOptions;

  const applySelection = () => {
    if (searchText.trim()) {
      const next = filteredOptions.length === 0
        ? '__NONE__'
        : filteredOptions.length === safeOptions.length
        ? []
        : filteredOptions;
      onChange(next);
      closeDropdown();
      return;
    }
    onChange(currentNone ? '__NONE__' : currentSelected);
    closeDropdown();
  };

  const resetSelection = () => {
    setDraftSelected([]);
    setDraftNone(true);
  };

  const toggleAll = () => {
    if (currentAll) {
      setDraftSelected([]);
      setDraftNone(true);
    } else {
      setDraftSelected([]);
      setDraftNone(false);
    }
  };

  const toggle = (val) => {
    if (currentAll) {
      const next = safeOptions.filter(x => x !== val);
      if (next.length === 0) {
        setDraftSelected([]);
        setDraftNone(true);
        return;
      }
      setDraftSelected(next);
      setDraftNone(false);
      return;
    }

    if (currentNone) {
      setDraftSelected([val]);
      setDraftNone(false);
      return;
    }

    if (currentSelected.includes(val)) {
      const next = currentSelected.filter(x => x !== val);
      if (next.length === 0) {
        setDraftSelected([]);
        setDraftNone(true);
        return;
      }
      setDraftSelected(next);
    } else {
      const next = [...currentSelected, val];
      const normalized = next.length === safeOptions.length ? [] : next;
      setDraftSelected(normalized);
      setDraftNone(false);
    }
  };

  const isChecked = (val) => {
    if (currentNone) return false;
    if (currentAll) return true;
    return currentSelected.includes(val);
  };

  const displayLabel = currentNone
    ? `0 selected`
    : noneSelected
    ? `All ${label}`
    : safeSelected.length === 1
    ? String(safeSelected[0])
    : `${safeSelected.length} selected`;
  const hasActiveFilter = noSelection || !noneSelected;

  const dropdown = open ? (
    <div
      ref={menuRef}
      className={`fixed z-[180] rounded-lg shadow-2xl border overflow-hidden ${darkMode?'bg-gray-700 border-gray-600':'bg-white border-gray-200'}`}
      style={{ top: menuPos.top, left: menuPos.left, width: menuPos.width, maxWidth: 'calc(100vw - 32px)' }}
    >
      <div className={`px-2 pt-2 pb-1 border-b ${darkMode?'border-gray-600':'border-gray-100'}`}>
        <input
          type="text"
          value={searchText}
          onChange={e => setSearchText(e.target.value)}
          placeholder={`Search ${label}...`}
          autoFocus
          className={`w-full px-2 py-1.5 rounded text-xs border ${darkMode?'bg-gray-600 border-gray-500 text-white placeholder-gray-400':'bg-gray-50 border-gray-200 text-gray-800 placeholder-gray-400'}`}
          onClick={e => e.stopPropagation()}
          onKeyDown={e => { if (e.key === 'Escape') closeDropdown(); if (e.key === 'Enter') applySelection(); }}
        />
      </div>
      <div className="overflow-auto" style={{ maxHeight: menuPos.maxHeight }}>
        <label style={{cursor:'pointer'}} className={`flex items-center gap-2 px-3 py-2 text-xs font-semibold border-b
          ${darkMode?'border-gray-600 hover:bg-gray-600 text-white':'border-gray-100 hover:bg-blue-50 text-gray-700'}`}>
          <input type="checkbox"
            checked={currentAll}
            ref={el => { if (el) el.indeterminate = someSelected; }}
            onChange={toggleAll}
            className="accent-blue-600" style={{cursor:'pointer'}}/>
          <span>(Select All)</span>
        </label>
        {filteredOptions.map(opt => (
          <label key={opt} style={{cursor:'pointer'}} className={`flex items-center gap-2 px-3 py-2 text-xs
            ${darkMode?'hover:bg-gray-600 text-white':'hover:bg-blue-50 text-gray-700'}`}>
            <input type="checkbox" checked={isChecked(opt)} onChange={()=>toggle(opt)}
              className="accent-blue-600" style={{cursor:'pointer'}}/>
            <span className="min-w-0 break-words leading-snug" title={opt}>{opt}</span>
          </label>
        ))}
        {filteredOptions.length === 0 && <div className={`px-3 py-2 text-xs ${darkMode?'text-gray-400':'text-gray-500'}`}>{searchText.trim() ? 'No matching options' : 'No options'}</div>}
      </div>
      <div className={`flex gap-2 px-3 py-2 border-t shadow-inner ${darkMode ? 'bg-gray-800 border-gray-600' : 'bg-gray-50 border-gray-200'}`}>
        <button
          type="button"
          onClick={applySelection}
          className="flex-1 px-3 py-2 rounded-lg text-xs font-bold bg-blue-600 text-white hover:bg-blue-700 shadow-sm"
        >
          Apply
        </button>
        <button
          type="button"
          onClick={resetSelection}
          className={`flex-1 px-3 py-2 rounded-lg text-xs font-semibold ${darkMode ? 'bg-gray-600 text-gray-100 hover:bg-gray-500' : 'bg-white border border-gray-300 text-gray-700 hover:bg-gray-100'}`}
        >
          Clear All
        </button>
        {searchText.trim() && (
          <button
            type="button"
            onClick={() => setSearchText('')}
            title="Clear search"
            className={`px-3 py-2 rounded-lg text-xs font-semibold flex items-center justify-center gap-1 ${darkMode ? 'bg-gray-600 text-gray-100 hover:bg-gray-500' : 'bg-white border border-gray-300 text-gray-500 hover:bg-gray-100'}`}
          >
            <X className="w-3.5 h-3.5" />
          </button>
        )}
      </div>
    </div>
  ) : null;

  return (
    <div className="relative w-full min-w-0" ref={ref}>
      {!hideLabel && <label className={`block text-xs font-medium mb-1 ${txt2}`}>{label}</label>}
      <button onClick={(e)=>{ e.stopPropagation(); setOpen(o=>!o); }} style={{cursor:'pointer'}}
        className={`w-full h-10 px-3 py-2 rounded-lg text-sm border text-left flex justify-between items-center transition-colors
          ${darkMode
            ? hasActiveFilter
              ? 'bg-amber-900/30 border-amber-500 text-amber-100 hover:bg-amber-900/40'
              : 'bg-gray-600 border-gray-500 text-white hover:bg-gray-500'
            : hasActiveFilter
              ? 'bg-amber-50 border-amber-300 text-amber-800 hover:bg-amber-100'
              : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'}`}>
        <span className={`truncate ${hasActiveFilter ? 'font-semibold' : ''}`}>{displayLabel}</span>
        <ChevronDown className={`w-4 h-4 flex-shrink-0 ml-1 transition-transform ${open ? 'rotate-180' : ''}`}/>
      </button>
      {typeof document !== 'undefined' ? createPortal(dropdown, document.body) : dropdown}
    </div>
  );
};

// ─── Search Input for SO / PO numbers ─────────────────────────────────────
const SearchInput = ({ placeholder, onSearch, darkMode, txt2, label }) => {
  const [open, setOpen] = useState(false);
  const [value, setValue] = useState('');
  const ref = useRef(null);
  // Floating dropdown — escapes `overflow-hidden` parents so the search
  // panel is never clipped by the table card border.
  const float = useFloatingDropdown(open, 256, 320, 260);

  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target) && !float.triggerRef.current?.contains(e.target)) setOpen(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const handleSearch = () => {
    const numbers = value.split('\n').map(s=>s.trim()).filter(Boolean);
    onSearch(numbers);
    setOpen(false);
  };

  const handleClear = () => {
    setValue('');
    onSearch([]);
    setOpen(false);
  };

  return (
    <div className="relative w-full" ref={ref}>
      <button
        ref={float.triggerRef}
        onClick={() => setOpen(o => !o)}
        title={`Search ${label}`}
        className={`w-full h-10 flex items-center justify-between gap-1.5 px-3 py-2 rounded-lg text-sm border font-medium transition-all
          ${darkMode ? 'bg-gray-600 border-gray-500 text-white hover:bg-gray-500' : 'bg-white border-gray-300 text-gray-700 hover:bg-blue-50 hover:border-blue-400'}`}
      >
        <span className="flex items-center gap-1.5 min-w-0">
          <Search className="w-4 h-4 flex-shrink-0"/>
          <span className="truncate">Search {label}</span>
        </span>
        <ChevronDown className="w-3.5 h-3.5 opacity-60 flex-shrink-0"/>
      </button>
      {open && (
        <div
          style={float.menuPos.style}
          className={`rounded-xl shadow-2xl border p-3 ${darkMode?'bg-gray-800 border-gray-700':'bg-white border-gray-200'}`}
        >
          <p className={`text-xs font-semibold mb-1.5 ${darkMode?'text-gray-300':'text-gray-600'}`}>
            Enter {label} (one per line):
          </p>
          <textarea
            value={value}
            onChange={e => setValue(e.target.value)}
            placeholder={placeholder}
            rows={4}
            className={`w-full px-2 py-1.5 rounded-lg text-xs border resize-none font-mono
              ${darkMode?'bg-gray-700 border-gray-600 text-white placeholder-gray-500':'bg-gray-50 border-gray-300 text-gray-800 placeholder-gray-400'}`}
            autoFocus
          />
          <div className="flex gap-2 mt-2">
            <button onClick={handleSearch}
              className="flex-1 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold">
              Search
            </button>
            <button onClick={handleClear}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium ${darkMode?'bg-gray-600 text-gray-200 hover:bg-gray-500':'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>
              Clear
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

// ─── Multiline search dropdown ────────────────────────────────────────────
const RFQMultiSearch = ({
  value,
  onChange,
  onSearch,
  darkMode,
  txt2,
  label = 'Search',
  description = 'Enter one Request Number, Item Name, or Spec per line. Results match any entered value.',
  placeholder = 'REQ-0001\nBearing SKF\nStainless steel 304',
}) => {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  // Floating dropdown — escapes `overflow-hidden` parents.
  const float = useFloatingDropdown(open, 360, 440, 320);

  useEffect(() => {
    const handler = (event) => {
      if (ref.current && !ref.current.contains(event.target) && !float.triggerRef.current?.contains(event.target)) setOpen(false);
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const searchValues = String(value || '')
    .split(/\r?\n/)
    .map(item => item.trim())
    .filter(Boolean);

  const applySearch = () => {
    const normalized = searchValues.join('\n');
    onChange(normalized);
    onSearch(normalized);
    setOpen(false);
  };

  const clearSearch = () => {
    onChange('');
    onSearch('');
    setOpen(false);
  };

  return (
    <div className="relative w-full min-w-0" ref={ref}>
      <label className={`block text-xs font-semibold mb-1 ${txt2}`}>{label}</label>
      <button
        ref={float.triggerRef}
        type="button"
        onClick={() => setOpen(current => !current)}
        className={`w-full h-10 px-3 py-2 rounded-xl text-sm border text-left flex items-center justify-between gap-2 transition-colors ${
          darkMode
            ? searchValues.length
              ? 'bg-amber-900/30 border-amber-500 text-amber-100 hover:bg-amber-900/40'
              : 'bg-gray-700 border-gray-600 text-white hover:bg-gray-600'
            : searchValues.length
              ? 'bg-amber-50 border-amber-300 text-amber-800 hover:bg-amber-100'
              : 'bg-white border-gray-200 text-gray-700 hover:bg-gray-50'
        }`}
      >
        <span className="flex min-w-0 items-center gap-2">
          <Search className="w-4 h-4 flex-shrink-0" />
          <span className="truncate">
            {searchValues.length ? `${searchValues.length} value${searchValues.length > 1 ? 's' : ''}` : label}
          </span>
        </span>
        <ChevronDown className="w-4 h-4 flex-shrink-0" />
      </button>

      {open && (
        <div
          style={float.menuPos.style}
          className={`rounded-xl border p-3 shadow-2xl ${darkMode ? 'bg-gray-800 border-gray-600' : 'bg-white border-gray-200'}`}
        >
          <p className={`mb-2 text-xs leading-relaxed ${txt2}`}>
            {description}
          </p>
          <textarea
            value={value}
            onChange={event => onChange(event.target.value)}
            onKeyDown={event => {
              if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
                event.preventDefault();
                applySearch();
              }
              if (event.key === 'Escape') setOpen(false);
            }}
            placeholder={'REQ-0001\nBearing SKF\nStainless steel 304'}
            className={`h-52 w-full overflow-y-auto resize-y rounded-lg border px-3 py-2 font-mono text-sm leading-6 ${darkMode ? 'bg-gray-700 border-gray-600 text-white placeholder:text-gray-400' : 'bg-gray-50 border-gray-300 text-gray-800 placeholder:text-gray-400'}`}
            autoFocus
          />
          <div className={`mt-1 text-[11px] ${txt2}`}>
            {searchValues.length} search value{searchValues.length === 1 ? '' : 's'}
          </div>
          <div className="mt-3 flex gap-2">
            <button
              type="button"
              onClick={applySearch}
              className="flex-1 rounded-lg bg-blue-600 px-3 py-2 text-xs font-bold text-white hover:bg-blue-700"
            >
              Search
            </button>
            <button
              type="button"
              onClick={clearSearch}
              className={`flex-1 rounded-lg px-3 py-2 text-xs font-semibold ${darkMode ? 'bg-gray-600 text-gray-100 hover:bg-gray-500' : 'border border-gray-300 bg-white text-gray-700 hover:bg-gray-100'}`}
            >
              Clear Search
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

const FilterPanel = ({ children, darkMode, className = '' }) => (
  <div className={`relative z-[70] overflow-visible mx-5 my-3 rounded-xl border p-3 ${darkMode ? 'border-gray-700 bg-gray-800/70' : 'border-gray-100 bg-[#f6f6f4]'} ${className}`}>
    {children}
  </div>
);

const PagePagination = ({ darkMode, txt2, page, totalPages, total, perPage, onPageChange, onPerPageChange }) => {
  const from = total ? (page - 1) * perPage + 1 : 0;
  const to = Math.min(page * perPage, total);
  return (
    <div className={`px-5 py-3 border-t ${darkMode ? 'border-gray-700' : 'border-gray-100'} flex flex-wrap justify-between items-center gap-3`}>
      <div className="flex items-center gap-3">
        <span className={`text-sm ${txt2}`}>Showing {from}-{to} of {fmtNum(total)}</span>
        <label className={`flex items-center gap-1 text-xs ${txt2}`}>Rows
          <select
            className={`px-2 py-1 rounded-lg text-xs border ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-white border-gray-200'}`}
            value={perPage}
            onChange={e => onPerPageChange(Number(e.target.value))}
          >
            <option value={10}>10</option><option value={25}>25</option><option value={50}>50</option><option value={100}>100</option><option value={500}>500</option>
          </select>
        </label>
      </div>
      <div className="flex gap-1 items-center">
        <button disabled={page === 1} onClick={() => onPageChange(page - 1)} className={`p-1.5 rounded ${page === 1 ? 'opacity-40' : 'hover:bg-blue-100'}`}><ChevronLeft className="w-4 h-4" /></button>
        <span className={`px-3 py-1 rounded text-sm font-semibold ${darkMode ? 'bg-gray-700 text-white' : 'bg-blue-100 text-blue-700'}`}>{page}/{totalPages}</span>
        <button disabled={page === totalPages} onClick={() => onPageChange(page + 1)} className={`p-1.5 rounded ${page === totalPages ? 'opacity-40' : 'hover:bg-blue-100'}`}><ChevronRight className="w-4 h-4" /></button>
      </div>
    </div>
  );
};

const FloatingTableScrollbar = ({ targetRef, darkMode }) => {
  const barRef = useRef(null);
  const syncingRef = useRef(false);
  const [state, setState] = useState({ visible: false, left: 0, width: 0, scrollWidth: 0 });

  useEffect(() => {
    const target = targetRef.current;
    if (!target) return undefined;

    const update = () => {
      const rect = target.getBoundingClientRect();
      const hasOverflow = target.scrollWidth > target.clientWidth + 2;
      const tableVisible = rect.bottom > 96 && rect.top < window.innerHeight - 42;
      const nativeBottomScrollbarVisible = rect.bottom <= window.innerHeight - 12;
      setState({
        visible: hasOverflow && tableVisible && !nativeBottomScrollbarVisible,
        left: Math.max(8, rect.left),
        width: Math.min(rect.width, window.innerWidth - Math.max(8, rect.left) - 8),
        scrollWidth: target.scrollWidth,
      });
      if (barRef.current && barRef.current.scrollLeft !== target.scrollLeft) {
        barRef.current.scrollLeft = target.scrollLeft;
      }
    };

    const syncFromTarget = () => {
      if (syncingRef.current) return;
      syncingRef.current = true;
      if (barRef.current) barRef.current.scrollLeft = target.scrollLeft;
      syncingRef.current = false;
      update();
    };

    update();
    target.addEventListener('scroll', syncFromTarget, { passive: true });
    window.addEventListener('scroll', update, { passive: true });
    window.addEventListener('resize', update);
    const resizeObserver = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(update) : null;
    resizeObserver?.observe(target);

    return () => {
      target.removeEventListener('scroll', syncFromTarget);
      window.removeEventListener('scroll', update);
      window.removeEventListener('resize', update);
      resizeObserver?.disconnect();
    };
  }, [targetRef]);

  if (!state.visible) return null;

  return (
    <div
      ref={barRef}
      className={`floating-table-scrollbar ${darkMode ? 'floating-table-scrollbar-dark' : ''}`}
      style={{ left: state.left, width: state.width }}
      onScroll={(e) => {
        const target = targetRef.current;
        if (!target || syncingRef.current) return;
        syncingRef.current = true;
        target.scrollLeft = e.currentTarget.scrollLeft;
        syncingRef.current = false;
      }}
    >
      <div style={{ width: state.scrollWidth, height: 1 }} />
    </div>
  );
};

const DataTableScroll = ({ children, className = '', darkMode }) => {
  const ref = useRef(null);

  useEffect(() => {
    const frame = ref.current;
    if (!frame || typeof document === 'undefined') return undefined;

    let raf = null;
    let mirror = null;
    let mirrorTable = null;
    let lastHtml = '';

    const removeMirror = () => {
      if (mirror) {
        mirror.remove();
        mirror = null;
        mirrorTable = null;
        lastHtml = '';
      }
    };

    const ensureMirror = () => {
      if (mirror) return;
      mirror = document.createElement('div');
      mirror.className = `window-sticky-table-header ${darkMode ? 'window-sticky-table-header-dark' : ''}`;
      mirror.style.display = 'none';
      mirror.style.position = 'fixed';
      mirror.style.top = '0px';
      mirror.style.overflow = 'hidden';
      mirror.style.pointerEvents = 'none';
      mirror.style.zIndex = '130';
      mirror.style.boxSizing = 'border-box';

      mirrorTable = document.createElement('table');
      mirrorTable.className = 'window-sticky-table-header-table';
      mirrorTable.style.borderCollapse = 'collapse';
      mirrorTable.style.tableLayout = 'fixed';
      mirrorTable.style.margin = '0';
      mirrorTable.style.transformOrigin = 'top left';
      mirror.appendChild(mirrorTable);
      document.body.appendChild(mirror);
    };

    const syncMirrorHeader = (table, thead, frameRect) => {
      ensureMirror();
      if (!mirrorTable) return;

      // Strategy: position every mirror <th> absolutely inside the mirror
      // container (which is a fixed overlay matching the frame's viewport
      // rectangle). Each th's left offset is read directly from
      // getBoundingClientRect() on the ORIGINAL th — which already accounts
      // for the frame's horizontal scroll position because the real table is
      // rendered inside the scrollable frame. We then subtract frameRect.left
      // so the offset is relative to the mirror container.
      //
      // Pinned (sticky) columns: their real th.getBoundingClientRect().left
      // is already "stuck" at their pinned position — so we just read it the
      // same way, and the clone naturally stays at the correct left edge.
      //
      // This approach requires no translateX on the table, no counter-hacks,
      // and works correctly for both pinned and non-pinned columns.

      const tableStyles = window.getComputedStyle(table);
      mirrorTable.className = `${table.className || ''} window-sticky-table-header-table`;
      mirrorTable.style.fontSize = tableStyles.fontSize;
      mirrorTable.style.lineHeight = tableStyles.lineHeight;
      mirrorTable.style.letterSpacing = tableStyles.letterSpacing;
      // Mirror table must fill the container and use absolute positioning for
      // each th — no table layout engine interference.
      mirrorTable.style.width = '100%';
      mirrorTable.style.transform = '';
      mirrorTable.style.position = 'relative';

      const html = thead.outerHTML;
      if (html !== lastHtml) {
        mirrorTable.innerHTML = html;
        lastHtml = html;
      }

      const originalThs = Array.from(thead.querySelectorAll('th'));
      const mirrorThs = Array.from(mirrorTable.querySelectorAll('th'));
      const headerHeight = thead.getBoundingClientRect().height || 0;

      // Make the mirror thead the positioning context for absolute th clones.
      const mirrorThead = mirrorTable.querySelector('thead');
      const mirrorTr = mirrorTable.querySelector('tr');
      if (mirrorThead) {
        mirrorThead.style.position = 'relative';
        mirrorThead.style.display = 'block';
        mirrorThead.style.height = `${headerHeight}px`;
        mirrorThead.style.width = '100%';
      }
      if (mirrorTr) {
        mirrorTr.style.position = 'relative';
        mirrorTr.style.display = 'block';
        mirrorTr.style.height = `${headerHeight}px`;
      }

      originalThs.forEach((th, idx) => {
        const clone = mirrorThs[idx];
        if (!clone) return;
        const rect = th.getBoundingClientRect();
        const width = Math.max(1, rect.width);
        const thStyles = window.getComputedStyle(th);

        // Position absolutely inside the mirror container.
        // rect.left - frameRect.left gives offset from left edge of the mirror.
        // For sticky (pinned) columns the browser already holds rect.left at
        // the stuck position, so no special case needed.
        clone.style.position = 'absolute';
        clone.style.top = '0';
        clone.style.left = `${Math.round(rect.left - frameRect.left)}px`;
        clone.style.width = `${width}px`;
        clone.style.minWidth = `${width}px`;
        clone.style.maxWidth = `${width}px`;
        clone.style.height = `${headerHeight}px`;
        clone.style.boxSizing = 'border-box';
        clone.style.backgroundClip = 'padding-box';
        clone.style.transform = '';

        // BACKGROUND — must be FULLY OPAQUE for every clone. The original th
        // often has a transparent computed background-color (it relies on
        // <thead className="bg-slate-50"> inheritance, which doesn't show up
        // in getComputedStyle for the th itself — background is NOT
        // inherited). When the clone ends up transparent, the next column's
        // text scrolls underneath and shows through, producing the overlap
        // bug the user reported.
        //
        // Solution:
        //  • For sticky (pinned) ths → always use solid header bg color.
        //  • For non-pinned ths → use the original bg if it's opaque,
        //    otherwise fall back to the same solid header color.
        const solidHeaderBg = darkMode ? '#374151' : '#e2e8f0';
        const origBg = thStyles.backgroundColor;
        const isTransparentBg = !origBg
          || origBg === 'transparent'
          || origBg === 'rgba(0, 0, 0, 0)'
          || origBg === 'rgba(0,0,0,0)';
        const isSticky = thStyles.position === 'sticky';
        clone.style.background = (isSticky || isTransparentBg) ? solidHeaderBg : origBg;

        if (isSticky) {
          clone.style.zIndex = '45';
          clone.style.boxShadow = thStyles.boxShadow || '10px 0 14px -14px rgba(15, 23, 42, 0.55)';
        } else {
          clone.style.zIndex = '1';
          clone.style.boxShadow = '';
        }

        clone.style.fontSize = thStyles.fontSize;
        clone.style.fontWeight = thStyles.fontWeight;
        clone.style.lineHeight = thStyles.lineHeight;
        clone.style.paddingTop = thStyles.paddingTop;
        clone.style.paddingRight = thStyles.paddingRight;
        clone.style.paddingBottom = thStyles.paddingBottom;
        clone.style.paddingLeft = thStyles.paddingLeft;
        clone.style.textAlign = thStyles.textAlign;
        clone.style.verticalAlign = thStyles.verticalAlign;
        clone.style.overflow = 'hidden';
      });
    };

    const applyHeaderLock = () => {
      raf = null;
      const table = frame.querySelector('table');
      const thead = frame.querySelector('thead');
      if (!table || !thead) {
        removeMirror();
        return;
      }

      const frameRect = frame.getBoundingClientRect();
      const tableRect = table.getBoundingClientRect();
      const headerHeight = thead.getBoundingClientRect().height || 0;
      const topOffset = 0;

      const shouldStick = (
        frameRect.top < topOffset &&
        tableRect.bottom > topOffset + headerHeight &&
        frameRect.right > 0 &&
        frameRect.left < window.innerWidth
      );

      if (!shouldStick) {
        if (mirror) mirror.style.display = 'none';
        return;
      }

      const left = Math.max(frameRect.left, 0);
      const right = Math.min(frameRect.right, window.innerWidth);
      const width = Math.max(0, right - left);

      // Pass frameRect so syncMirrorHeader can compute each th's offset
      // relative to the mirror container's left edge.
      // NOTE: frameRect here reflects the *visible* left edge of the frame
      // as clipped to the viewport — same as mirror.style.left below.
      const visibleFrameRect = { ...frameRect, left };
      syncMirrorHeader(table, thead, visibleFrameRect);
      if (!mirror || !mirrorTable) return;

      mirror.style.display = 'block';
      mirror.style.left = `${left}px`;
      mirror.style.width = `${width}px`;
      mirror.style.height = `${headerHeight}px`;
      mirror.style.position = 'fixed';
      mirror.style.overflow = 'hidden';
      // Mirror container must be 'relative' so absolute-positioned th clones
      // are anchored to its top-left corner.
      mirrorTable.style.position = 'relative';
      mirrorTable.style.height = `${headerHeight}px`;
      mirror.style.background = darkMode ? '#111827' : '#ffffff';
      mirror.classList.toggle('window-sticky-table-header-dark', !!darkMode);
    };

    const schedule = () => {
      if (raf != null) return;
      raf = window.requestAnimationFrame(applyHeaderLock);
    };

    schedule();
    window.addEventListener('scroll', schedule, { passive: true });
    window.addEventListener('resize', schedule);
    frame.addEventListener('scroll', schedule, { passive: true });
    const resizeObserver = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(schedule) : null;
    resizeObserver?.observe(frame);

    return () => {
      if (raf != null) window.cancelAnimationFrame(raf);
      window.removeEventListener('scroll', schedule);
      window.removeEventListener('resize', schedule);
      frame.removeEventListener('scroll', schedule);
      resizeObserver?.disconnect();
      removeMirror();
    };
  }, [children, darkMode]);

  return (
    <>
      <div
        ref={ref}
        className={`data-table-scroll data-table-scroll-frame ${className}`}
        style={{ overflowX: 'auto', overflowY: 'visible' }}
      >
        {children}
      </div>
      <FloatingTableScrollbar targetRef={ref} darkMode={darkMode} />
    </>
  );
};

// ─── SO Status Pie ─────────────────────────────────────────────────────────
const StatusPie = ({ data, darkMode }) => {
  const [etcHover, setEtcHover] = useState(false);
  const [etcPos, setEtcPos] = useState({x:0, y:0});
  const sorted = [...(data||[])].sort((a,b) => b.value - a.value);
  const top5 = sorted.slice(0, 5);
  const rest = sorted.slice(5);
  const etcValue = rest.reduce((s, d) => s + d.value, 0);
  const pieData = etcValue > 0
    ? [...top5, { name: `Etc (${rest.length} others)`, value: etcValue, isEtc: true, etcItems: rest }]
    : top5;
  return (
    <div style={{position:'relative'}}>
      <ResponsiveContainer width="100%" height={300}>
        <PieChart>
          <Pie data={pieData} cx="50%" cy="42%" innerRadius={52} outerRadius={88} isAnimationActive={false}
            paddingAngle={0} dataKey="value" labelLine={false} label={renderPctLabel}>
            {pieData.map((d,i)=>(
              <Cell key={i} fill={d.isEtc ? '#9CA3AF' : PIE_COLORS[i % PIE_COLORS.length]}/>
            ))}
          </Pie>
          <Tooltip contentStyle={{backgroundColor:darkMode?'#1F2937':'#fff',borderRadius:'8px'}}
            formatter={(v,n,p)=> p.payload.isEtc
              ? [fmtNum(v), `${p.payload.etcItems?.map(x=>x.name).join(', ')}`]
              : [fmtNum(v), n]}/>
          <Legend iconSize={8} layout="horizontal" align="center" verticalAlign="bottom"
            formatter={(v, entry) => {
              if (entry.payload?.isEtc) {
                return (
                  <span className="text-xs" style={{cursor:'help', color: darkMode?'#D1D5DB':'#374151'}}
                    onMouseEnter={e=>{setEtcHover(true);setEtcPos({x:e.clientX,y:e.clientY});}}
                    onMouseLeave={()=>setEtcHover(false)}>
                    {v}
                  </span>
                );
              }
              return <span className="text-xs" style={{color: darkMode?'#D1D5DB':'#374151'}}>{v}</span>;
            }}/>
        </PieChart>
      </ResponsiveContainer>
      {etcHover && rest.length > 0 && (
        <div className="fixed z-[200] bg-gray-900 text-white text-xs rounded-lg px-3 py-2 shadow-xl pointer-events-none max-w-xs"
          style={{left: etcPos.x + 12, top: etcPos.y - 10}}>
          <div className="font-bold mb-1">Etc ({rest.length} statuses):</div>
          {rest.map((r,i)=>(
            <div key={i} className="flex justify-between gap-3">
              <span>{r.name}</span><span className="font-semibold">{fmtNum(r.value)}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};


// ─── Date Range Filter ────────────────────────────────────────────────────
const DateRangeFilter = ({ darkMode, txt, txt2, card, onFilter, value, label = 'Filter SO Create Date', compact = false }) => {
  const [mode, setMode] = useState(value?.mode || 'all'); // all | today | week | month | year | range
  const [startDate, setStartDate] = useState(value?.start || '');
  const [endDate, setEndDate] = useState(value?.end || '');
  const [rangeOpen, setRangeOpen] = useState(false);
  // Floating dropdown for the Custom Date Range picker — escapes
  // `overflow-hidden` parents so the date inputs are never clipped.
  const rangeFloat = useFloatingDropdown(mode === 'range' && rangeOpen, 360, 440, 120);

  // Keep internal state in sync when the controlled `value` changes externally
  // (e.g. user changes filter on another page that shares the same global state).
  useEffect(() => {
    if (!value) return;
    setMode(value.mode || 'all');
    if (value.start !== undefined) setStartDate(value.start || '');
    if (value.end   !== undefined) setEndDate(value.end || '');
  }, [value?.mode, value?.start, value?.end]);

  useEffect(() => {
    const next =
      mode === 'all'
        ? { mode: 'all' }
        : mode === 'range'
        ? { mode: 'range', start: startDate, end: endDate }
        : { mode };

    if (mode === 'range') return;

    const current =
      !value || value.mode === 'all'
        ? { mode: 'all' }
        : value.mode === 'range'
        ? { mode: 'range', start: value.start || '', end: value.end || '' }
        : { mode: value.mode };

    // Only notify the parent when the user actually changed the filter.
    // Without this guard, mounting the filter emits a new `{ mode: 'all' }`
    // object, which re-fetches Delivery Completed, hides the page behind the
    // loading state, remounts the filter, and creates a fast blank/loading loop.
    if (JSON.stringify(next) !== JSON.stringify(current)) {
      onFilter(next);
    }
  }, [mode, startDate, endDate, value, onFilter]);

  const reset = () => {
    setMode('all');
    setStartDate(''); setEndDate('');
    setRangeOpen(false);
    onFilter({ mode: 'all' });
  };

  const applyRange = () => {
    if (!startDate || !endDate) return;
    onFilter({ mode: 'range', start: startDate, end: endDate });
    setRangeOpen(false);
  };

  const formatRangeLabel = (start, end) => {
    if (!start || !end) return '';
    try {
      return `${format(parseISO(start), 'dd/MM/yyyy')} - ${format(parseISO(end), 'dd/MM/yyyy')}`;
    } catch {
      return `${start} - ${end}`;
    }
  };

  const appliedRangeLabel = value?.mode === 'range' ? formatRangeLabel(value.start, value.end) : '';

  return (
    <div data-tour="date-filter" className={`relative flex min-h-[64px] min-w-0 flex-1 flex-col items-start gap-2 px-5 py-3 rounded-xl ${card} shadow ${compact ? 'mb-0' : 'mb-4'}`}>
      <div className="flex items-center gap-3">
        <Calendar className="w-4 h-4 text-blue-500 flex-shrink-0"/>
        <span className={`text-sm font-semibold ${txt} flex-shrink-0`}>{label}:</span>
      </div>
      {/* Mode selector */}
      <div className="relative flex w-full flex-wrap items-start gap-1.5">
        {[
          ['all','All'], ['today','Today'], ['week','This Week'],
          ['month','This Month'], ['year','This Year'], ['range','Custom Date Range']
        ].map(([m, lbl]) => {
          const isRange = m === 'range';
          return (
            <div key={m} className={isRange ? 'flex flex-col items-start gap-0.5' : ''}>
              <button
                ref={isRange ? rangeFloat.triggerRef : undefined}
                type="button"
                onClick={() => {
                  setMode(m);
                  setRangeOpen(isRange ? true : false);
                }}
                className={`px-3 py-1 rounded-full text-xs font-semibold transition-all
                  ${mode === m ? 'bg-blue-600 text-white shadow' : darkMode ? 'bg-gray-700 text-gray-300 hover:bg-gray-600' : 'bg-gray-100 text-gray-600 hover:bg-blue-100'}`}
              >
                {lbl}
              </button>
              {isRange && appliedRangeLabel && (
                <span className={`pl-3 whitespace-nowrap text-[11px] font-semibold leading-tight ${txt2}`}>{appliedRangeLabel}</span>
              )}
            </div>
          );
        })}
        {mode === 'range' && rangeOpen && (
          <div
            style={rangeFloat.menuPos.style}
            className={`flex items-center gap-2 rounded-xl border p-3 shadow-xl ${darkMode ? 'bg-gray-800 border-gray-600' : 'bg-white border-gray-200'}`}
          >
            <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)}
              className={`px-3 py-1.5 rounded-lg text-sm border ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-white border-gray-300'}`}/>
            <span className={`text-xs ${txt2}`}>to</span>
            <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)}
              className={`px-3 py-1.5 rounded-lg text-sm border ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-white border-gray-300'}`}/>
            <button
              type="button"
              disabled={!startDate || !endDate}
              onClick={applyRange}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${!startDate || !endDate ? 'opacity-50 cursor-not-allowed' : ''} ${darkMode ? 'bg-blue-600 text-white hover:bg-blue-500' : 'bg-blue-600 text-white hover:bg-blue-700'}`}
            >
              Set
            </button>
          </div>
        )}
        {mode !== 'all' && (
          <button onClick={reset} className={`px-3 py-1 rounded-lg text-xs font-medium ${darkMode ? 'bg-gray-600 text-gray-200 hover:bg-gray-500' : 'bg-gray-200 text-gray-600 hover:bg-gray-300'}`}>
            Reset
          </button>
        )}
      </div>
    </div>
  );
};

// ═══════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════
const App = () => {
  const location = useLocation();
  const navigate = useNavigate();
  const [darkMode, setDarkMode] = useState(false);
  const activePage = PATH_PAGES[location.pathname.toLowerCase()] || 'dashboard';
  const setActivePage = useCallback((page) => {
    navigate(PAGE_PATHS[page] || '/', { replace: false });
  }, [navigate]);
  const openPage = useCallback((event, page, reset = () => {}) => {
    event.preventDefault();
    reset();
    setActivePage(page);
    window.scrollTo({ top: 0 });
  }, [setActivePage]);
  // ── Filter persistence: load saved state on first mount ──────────────────
  // Each page's filters are stored to localStorage so a refresh restores the
  // exact same view the user had. Loaded once at mount via lazy useState
  // initializers, then auto-saved on every change.
  const savedImportFilters = loadFilterState('import') || {};
  const savedRfqFilters = loadFilterState('rfq') || {};
  const savedItemRegFilters = loadFilterState('item-registration') || {};
  const savedRegisteredItemsFilters = loadFilterState('all-registered-items') || {};
  const savedVendorControlFilters = loadFilterState('vendor-control') || {};
  const savedSoFilters = loadFilterState('all-so') || {};
  const savedDashboardFilters = loadFilterState('dashboard') || {};
  const [showUploadDropdown, setShowUploadDropdown] = useState(false);
  const [sidebarExpanded, setSidebarExpanded] = useState(false);
  const uploadDropdownRef = useRef(null);
  const dashboardRequestSeq = useRef(0);
  const [frozenColumns, setFrozenColumns] = useState({});
  // Widths (in px) of every column for each table that currently has at least
  // one pinned column. Used to compute the cumulative `left` offset for
  // stacked pinned columns (col 2 at left=0, col 5 at left=width(col 2), …).
  // Measured imperatively from the live DOM so it stays correct regardless of
  // cell content, font metrics, or column min-width settings.
  const [frozenColumnWidths, setFrozenColumnWidths] = useState({});

  const [stats, setStats] = useState(() => readStatsCache(dashboardStatsCacheKey()));
  const [summaryPendingTotal, setSummaryPendingTotal] = useState(() => {
    const cachedStats = readStatsCache(dashboardStatsCacheKey());
    const cachedPending = readStatsCache(dashboardPendingCacheKey());
    const n = Number(cachedPending?.total ?? cachedStats?.total_so_count);
    return Number.isFinite(n) ? n : null;
  });
  const [agingData, setAgingData] = useState(() => readStatsCache(dashboardAgingCacheKey()) || []);
  const [allSOData, setAllSOData] = useState([]);
  const [approvalSOData, setApprovalSOData] = useState([]);
  const [picAggregations, setPicAggregations] = useState([]); // PIC aggregations from backend (all filtered data)
  const [soTotal, setSoTotal] = useState(0);
  const [soSubtotalAmount, setSoSubtotalAmount] = useState(0);
  const [soFilterOptions, setSoFilterOptions] = useState({ op_units: [], vendors: [], manufacturers: [], statuses: [], pics: [] });

  // SO filters — load from localStorage so filters persist across refresh
  // (savedSoFilters is already declared above from loadFilterState('all-so'))
  const [soFilters, setSoFilters] = useState(() => savedSoFilters.filters || { op_units: [], vendors: [], manufacturers: [], statuses: [], aging: [], pics: [] });
  const [soSearchNums, setSoSearchNums] = useState(() => savedSoFilters.searchNums || []); // search SO Item
  const [soMarginFilter, setSoMarginFilter] = useState(() => savedSoFilters.marginFilter || 'all'); // 'all' | 'positive' | 'negative'
  const [soSortOrder, setSoSortOrder] = useState(() => savedSoFilters.sortOrder || 'oldest'); // 'oldest' | 'newest'
  const [soPage, setSoPage] = useState(() => savedSoFilters.page || 1);
  const [soPerPage, setSoPerPage] = useState(() => savedSoFilters.perPage || 10);
  const [pendingPicHighlight, setPendingPicHighlight] = useState('');

  // SO Approval Status filters (same as Open SO except Vendor Name)
  const [approvalFilters, setApprovalFilters] = useState({ op_units: [], statuses: [], aging: [] });
  const [approvalSearchNums, setApprovalSearchNums] = useState([]);
  const [approvalPage, setApprovalPage] = useState(1);
  const [approvalPerPage, setApprovalPerPage] = useState(10);

  // Item Registration
  const [itemRegData, setItemRegData] = useState([]);
  const [itemRegTotal, setItemRegTotal] = useState(0);
  const [itemRegPage, setItemRegPage] = useState(() => savedItemRegFilters.page || 1);
  const [itemRegPerPage, setItemRegPerPage] = useState(() => savedItemRegFilters.perPage || 10);
  const [itemRegSearch, setItemRegSearch] = useState(() => savedItemRegFilters.search || []);
  const [itemRegAppliedSearch, setItemRegAppliedSearch] = useState(() => savedItemRegFilters.appliedSearch || []);
  const [itemRegLastUpdated, setItemRegLastUpdated] = useState(null);
  const [itemRegFilters, setItemRegFilters] = useState(() => savedItemRegFilters.filters || { clients: [], categories: [], pics: [], proc_statuses: [], mfr_names: [] });
  const [itemRegOptions, setItemRegOptions] = useState({ clients: [], categories: [], pics: [], proc_statuses: [], mfr_names: [] });
  const [itemRegMissingPicKpis, setItemRegMissingPicKpis] = useState([]);
  const [itemRegPicHighlight, setItemRegPicHighlight] = useState('');

  // RFQ
  const [rfqData, setRfqData] = useState([]);
  const [rfqTotal, setRfqTotal] = useState(0);
  const [rfqPage, setRfqPage] = useState(() => savedRfqFilters.page || 1);
  const [rfqPerPage, setRfqPerPage] = useState(() => savedRfqFilters.perPage || 10);
  const [rfqSearch, setRfqSearch] = useState(() => savedRfqFilters.search || '');
  const [rfqAppliedSearch, setRfqAppliedSearch] = useState(() => savedRfqFilters.appliedSearch || '');
  const [rfqColumns, setRfqColumns] = useState([]);
  const [rfqSimilarityColumns, setRfqSimilarityColumns] = useState([]);
  const [rfqShowSimilarity, setRfqShowSimilarity] = useState(false);
  const [rfqEditableFields, setRfqEditableFields] = useState([]);
  const [rfqPicKpis, setRfqPicKpis] = useState([]);
  const [rfqPicFilter, setRfqPicFilter] = useState(() => savedRfqFilters.picFilter || '');
  const [rfqFilters, setRfqFilters] = useState(() => savedRfqFilters.filters || { checks: [], clients: [], rfq_numbers: [], brands: [], purchase_pics: [], vendors: [] });
  const [rfqOptions, setRfqOptions] = useState({ checks: [], clients: [], rfq_numbers: [], brands: [], purchase_pics: [], vendors: [] });
  const [rfqSelectedCell, setRfqSelectedCell] = useState(null);
  const [rfqFillRange, setRfqFillRange] = useState(null);
  // Multi-select state for Shift+click in the RFQ table (same pattern as Import).
  const [rfqSelectedCells, setRfqSelectedCells] = useState(null);
  const [rfqSelectionAnchor, setRfqSelectionAnchor] = useState(null);
  const [rfqSimilarAction, setRfqSimilarAction] = useState(null);
  const [rfqLastUpdated, setRfqLastUpdated] = useState(null);

  // Import
  const [importData, setImportData] = useState([]);
  const [importColumns, setImportColumns] = useState([]);
  const [importTotal, setImportTotal] = useState(0);
  const [importPage, setImportPage] = useState(() => savedImportFilters.page || 1);
  const [importPerPage, setImportPerPage] = useState(() => savedImportFilters.perPage || 10);
  const [importSearch, setImportSearch] = useState(() => savedImportFilters.search || '');
  const [importAppliedSearch, setImportAppliedSearch] = useState(() => savedImportFilters.appliedSearch || '');
  const [importVendorCount, setImportVendorCount] = useState(0);
  const [importLastCopyAt, setImportLastCopyAt] = useState('');
  const [importEditingCell, setImportEditingCell] = useState(null);
  const [importEditValue, setImportEditValue] = useState('');
  const [showImportChecklist, setShowImportChecklist] = useState(false);
  // Show/Hide Detail — toggles the per-item block (SO through PURCHASE AMOUNT).
  // When hidden, the table becomes much narrower (1 line per row) since the
  // long spec / remark / item name columns disappear. Useful for quick overview.
  const [showImportDetail, setShowImportDetail] = useState(true);
  const [importSelectedCell, setImportSelectedCell] = useState(null);
  const [importFillRange, setImportFillRange] = useState(null);
  // Multi-select state for Shift+click (Excel-like). Stores a Set of
  // "{rowKey}|{field}" strings so we can select arbitrary rectangles of
  // cells. Anchor stores the first cell clicked so Shift+click extends from
  // there. Same-column multi-select (clicking multiple rows in one column)
  // is supported by checking if the Shift-clicked cell shares the field.
  const [importSelectedCells, setImportSelectedCells] = useState(null); // Set or null
  const [importSelectionAnchor, setImportSelectionAnchor] = useState(null); // {rowIndex, field}
  const [importVendorMenuOpen, setImportVendorMenuOpen] = useState(false);
  // Floating dropdown for the Vendor Import menu. Uses `position: fixed` so
  // the menu escapes the table card's `overflow-hidden` and is never clipped
  // or covered by the table/filter below.
  const importVendorDropdown = useFloatingDropdown(importVendorMenuOpen, 224, 280, 200);
  const [importFilters, setImportFilters] = useState(() => {
    // Coerce daysLeft to array — previously it was a string, and old
    // localStorage entries may still contain a string. Guard against
    // TypeError when calling .map() on a string.
    const rawDaysLeft = savedImportFilters.daysLeft;
    const daysLeftArr = Array.isArray(rawDaysLeft)
      ? rawDaysLeft
      : (typeof rawDaysLeft === 'string' && rawDaysLeft ? [rawDaysLeft] : []);
    return {
      yupi_po: savedImportFilters.yupi_po || [],
      vendors: savedImportFilters.vendors || [],
      statuses: savedImportFilters.statuses || [],
      daysLeft: daysLeftArr,
    };
  });
  const [importOptions, setImportOptions] = useState(() => ({
    yupi_po: [],
    vendors: [],
    statuses: ['NEW', ...IMPORT_STATUS_OPTIONS],
  }));
  const [importReqDlvSort, setImportReqDlvSort] = useState(() => savedImportFilters.reqDlvSort || 'oldest');
  const [importYupiPoSort, setImportYupiPoSort] = useState(() => savedImportFilters.yupiPoSort || '');
  // Import page KPIs (returned by /api/import/data). Computed backend-side
  // across ALL filtered rows so they don't change with pagination.
  const [importKpis, setImportKpis] = useState({
    total_po: 0,
    this_week_arrival: 0,
    this_week_no_sap: 0,
    sales_amount: 0,
    po_amount_idr: 0,
    gross_margin: 0,
  });
  const [rfqEditedRowKeys, setRfqEditedRowKeys] = useState(new Set());
  const rfqDashboardOnlyFields = new Set(['private_remarks_1', 'private_remarks_2']);

  // All Registered Items
  const [registeredItemsData, setRegisteredItemsData] = useState([]);
  const [registeredItemsTotal, setRegisteredItemsTotal] = useState(0);
  const [registeredItemsPage, setRegisteredItemsPage] = useState(() => savedRegisteredItemsFilters.page || 1);
  const [registeredItemsPerPage, setRegisteredItemsPerPage] = useState(() => savedRegisteredItemsFilters.perPage || 10);
  const [registeredItemsSearch, setRegisteredItemsSearch] = useState(() => savedRegisteredItemsFilters.search || '');
  const [registeredItemsAppliedSearch, setRegisteredItemsAppliedSearch] = useState(() => savedRegisteredItemsFilters.appliedSearch || '');
  const [registeredItemsProdIds, setRegisteredItemsProdIds] = useState(() => savedRegisteredItemsFilters.prodIds || []);
  const [registeredItemsAppliedProdIds, setRegisteredItemsAppliedProdIds] = useState(() => savedRegisteredItemsFilters.appliedProdIds || []);
  const [registeredItemsPicFilter, setRegisteredItemsPicFilter] = useState(() => savedRegisteredItemsFilters.picFilter || '');
  const [registeredItemsAppliedPicFilter, setRegisteredItemsAppliedPicFilter] = useState(() => savedRegisteredItemsFilters.appliedPicFilter || '');
  // NOTE: vendor_name state removed — source Excel has no Vendor column for
  // product master data, so the filter would always be empty.
  const [registeredItemsFilters, setRegisteredItemsFilters] = useState(() => ({ mfr_names: savedRegisteredItemsFilters.mfr_names || [] }));
  const [registeredItemsOptions, setRegisteredItemsOptions] = useState({ mfr_names: [], pic_options: [] });

  // Vendor Control
  const [vendorControlData, setVendorControlData] = useState([]);
  const [vendorControlTotal, setVendorControlTotal] = useState(0);
  const [vendorControlPage, setVendorControlPage] = useState(1);
  const [vendorControlPerPage, setVendorControlPerPage] = useState(10);
  const [vendorControlSearch, setVendorControlSearch] = useState('');
  const [vendorControlAppliedSearch, setVendorControlAppliedSearch] = useState('');
  const [vendorControlSelectedVendors, setVendorControlSelectedVendors] = useState([]);
  const [vendorControlAppliedVendors, setVendorControlAppliedVendors] = useState([]);
  const [vendorControlSuggestions, setVendorControlSuggestions] = useState([]);
  const [vendorControlSuggestOpen, setVendorControlSuggestOpen] = useState(false);
  // Floating dropdown for the vendor search suggestion list — escapes
  // `overflow-hidden` parents so suggestions are never clipped.
  const vendorControlSuggestFloat = useFloatingDropdown(vendorControlSuggestOpen && vendorControlSuggestions.length > 0, 360, 520, 280);
  const [vendorControlLastUpdated, setVendorControlLastUpdated] = useState(null);
  const [vendorPasswordVisible, setVendorPasswordVisible] = useState({});

  const [pageLoading, setPageLoading] = useState(() => activePage === 'dashboard' && stats === null);
  const [initialPageLoading, setInitialPageLoading] = useState(() => activePage === 'dashboard' && stats === null);
  const setLoading = setPageLoading;
  const [uploadProgress, setUploadProgress] = useState(null);
  const [toasts, setToasts] = useState([]);
  const [modal, setModal] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState('');
  const [downloadToast, setDownloadToast] = useState(null);
  const [completedData, setCompletedData] = useState(null);
  const [completedYear, setCompletedYear] = useState('all');
  const [dashboardMarginData, setDashboardMarginData] = useState(null);
  const [vendorPurchaseType, setVendorPurchaseType] = useState('all');
  const [completedLoading, setCompletedLoading] = useState(false);
  const [completedLoaded, setCompletedLoaded] = useState(false);
  const [marginDetailModal, setMarginDetailModal] = useState(null); // {category, data}
  const [picDbStatus, setPicDbStatus] = useState(() => readPicDbStatusCache()); // {product_id_count, master_pic_count, last_product_id_upload, last_pic_update}
  const [picUploadMsg, setPicUploadMsg] = useState(''); // feedback message for PIC uploads

  // Dynamic color palette for PIC badges — each unique name gets a consistent color
  const PIC_COLORS = [
    { bg: 'bg-indigo-100',  text: 'text-indigo-700'  },
    { bg: 'bg-emerald-100', text: 'text-emerald-700' },
    { bg: 'bg-amber-100',   text: 'text-amber-700'   },
    { bg: 'bg-cyan-100',    text: 'text-cyan-700'    },
    { bg: 'bg-blue-100',  text: 'text-blue-700'  },
    { bg: 'bg-slate-100',   text: 'text-slate-700'   },
    { bg: 'bg-teal-100',    text: 'text-teal-700'    },
    { bg: 'bg-pink-100',    text: 'text-pink-700'    },
    { bg: 'bg-lime-100',    text: 'text-lime-700'    },
    { bg: 'bg-fuchsia-100', text: 'text-fuchsia-700' },
    { bg: 'bg-blue-100',  text: 'text-blue-700'  },
    { bg: 'bg-slate-100',   text: 'text-slate-700'   },
  ];
  const picColorMap = useRef({});
  const picColorCounter = useRef(0);
  const getPicColor = (name) => {
    if (!name) return null;
    if (!picColorMap.current[name]) {
      picColorMap.current[name] = PIC_COLORS[picColorCounter.current % PIC_COLORS.length];
      picColorCounter.current += 1;
    }
    return picColorMap.current[name];
  };
  // ── Pinned (frozen) column logic ───────────────────────────────────────────
  //
  // Design (mirrors Excel "freeze panes"):
  //   • `frozenColumns[tableKey]` is a SORTED array of pinned column indices
  //     (1-based, matching :nth-child). Multiple columns can be pinned at once.
  //   • Each pinned column gets `position: sticky` with a cumulative `left`:
  //       - the leftmost pinned column has left = 0
  //       - the next pinned column has left = sum of widths of pinned columns
  //         to its left
  //     so stacked pinned columns sit side-by-side at the left edge instead of
  //     overlapping.
  //   • `position: sticky !important` is required because the global rule
  //     `.data-table-scroll thead th { position: relative !important; }`
  //     would otherwise override sticky on header cells (that was the root
  //     cause of the header "moving" but body cells sticking — they desynced).
  //   • Only the pinned column is frozen; all other columns scroll normally
  //     because they don't match the :nth-child selector.
  //   • The mirror sticky header (DataTableScroll) reads
  //     getBoundingClientRect() on the real <th>, which already reflects the
  //     stuck position — so the mirror automatically tracks pinned columns
  //     correctly once the underlying sticky positioning works.

  const toggleFrozenColumn = useCallback((tableKey, colIndex) => {
    setFrozenColumns(prev => {
      const current = Array.isArray(prev[tableKey]) ? prev[tableKey] : [];
      const exists = current.includes(colIndex);
      const next = exists
        ? current.filter(i => i !== colIndex)
        : [...current, colIndex].sort((a, b) => a - b);
      return { ...prev, [tableKey]: next };
    });
  }, []);

  // Measure the width of every <th> in each table that has pinned columns.
  // Re-runs on resize, on pin/unpin, and a few times shortly after pinning to
  // catch layout shifts caused by content reflow / image loading / fonts.
  useLayoutEffect(() => {
    if (typeof document === 'undefined') return undefined;
    const measure = () => {
      setFrozenColumnWidths(prev => {
        const next = { ...prev };
        let changed = false;
        Object.keys(frozenColumns).forEach(tableKey => {
          const indices = frozenColumns[tableKey];
          if (!Array.isArray(indices) || indices.length === 0) {
            if (next[tableKey]) { delete next[tableKey]; changed = true; }
            return;
          }
          const table = document.querySelector(`table.freeze-table-${CSS.escape(tableKey)}`);
          if (!table) return;
          const ths = table.querySelectorAll('thead th');
          if (!ths || ths.length === 0) return;
          const widths = {};
          ths.forEach((th, idx) => {
            const w = th.getBoundingClientRect().width;
            if (Number.isFinite(w) && w > 0) widths[idx + 1] = w;
          });
          const prevWidths = prev[tableKey] || {};
          const prevKeys = Object.keys(prevWidths);
          const newKeys = Object.keys(widths);
          const same = prevKeys.length === newKeys.length &&
            newKeys.every(k => Math.abs((prevWidths[k] || 0) - widths[k]) < 0.5);
          if (!same) {
            next[tableKey] = widths;
            changed = true;
          }
        });
        return changed ? next : prev;
      });
    };

    measure();
    // Retry shortly after to catch delayed layout shifts (font load, images,
    // table reflow when sticky kicks in, etc.).
    const timeouts = [60, 200, 450, 900].map(t => setTimeout(measure, t));
    window.addEventListener('resize', measure, { passive: true });

    // ResizeObserver on each pinned table — catches width changes from
    // anything other than window resize (e.g. sidebar toggle, parent flex
    // changes, content edits).
    const tables = [];
    const observers = [];
    Object.keys(frozenColumns).forEach(tableKey => {
      const t = document.querySelector(`table.freeze-table-${CSS.escape(tableKey)}`);
      if (!t) return;
      tables.push(t);
      if (typeof ResizeObserver !== 'undefined') {
        const ro = new ResizeObserver(() => measure());
        ro.observe(t);
        observers.push(ro);
      }
    });

    return () => {
      timeouts.forEach(clearTimeout);
      window.removeEventListener('resize', measure);
      observers.forEach(o => o.disconnect());
    };
  }, [frozenColumns]);

  const renderFreezeHeader = (tableKey, colIndex, label) => {
    const indices = Array.isArray(frozenColumns[tableKey]) ? frozenColumns[tableKey] : [];
    const active = indices.includes(colIndex);
    return (
      <div className="freeze-header group relative flex min-h-8 w-full min-w-0 items-center justify-center">
        <span className="freeze-header-label max-w-full text-center leading-tight">{label}</span>
        <button
          type="button"
          aria-label={active ? `Unfreeze ${label}` : `Freeze ${label}`}
          title={active ? `Unfreeze ${label}` : `Freeze ${label}`}
          onClick={(e) => { e.stopPropagation(); toggleFrozenColumn(tableKey, colIndex); }}
          // When pinned (active), the pin button is ALWAYS visible (opacity-100)
          // so the user can see which columns are frozen and click to unpin.
          // When NOT pinned, the button only appears on hover/focus to keep
          // the header tidy.
          className={`absolute right-0 top-1/2 inline-flex h-5 w-5 -translate-y-1/2 items-center justify-center rounded-md border shadow-sm transition-all ${active ? 'opacity-100 border-amber-300 bg-amber-100 text-amber-700' : `opacity-0 group-hover:opacity-100 group-focus-within:opacity-100 ${darkMode ? 'border-gray-600 bg-gray-700/90 text-gray-300 hover:bg-gray-600' : 'border-slate-200 bg-white/95 text-slate-500 hover:bg-slate-100'}`}`}
        >
          {active ? <PinOff className="h-3 w-3" /> : <Pin className="h-3 w-3" />}
        </button>
      </div>
    );
  };

  const frozenColumnCss = useMemo(() => {
    return Object.entries(frozenColumns).map(([tableKey, indices]) => {
      if (!Array.isArray(indices) || indices.length === 0) return '';
      const widths = frozenColumnWidths[tableKey] || {};
      // Walk pinned columns in DOM order; each one's `left` is the sum of
      // widths of pinned columns to its left. This produces the stacked,
      // non-overlapping "frozen pane" effect.
      let cumulativeLeft = 0;
      const rules = [];
      for (const idx of indices) {
        const left = cumulativeLeft;
        const w = widths[idx];
        if (Number.isFinite(w) && w > 0) cumulativeLeft += w;
        // IMPORTANT: We use [data-col-index="${idx}"] instead of :nth-child(${idx}).
        // The :nth-child selector BREAKS when a table has rowspan (merged cells):
        // in rows covered by a rowspan, the td count shifts, so :nth-child(1) in
        // a gap row points to a DIFFERENT column than in the first row. This caused
        // the "Item Yupi tertahan" bug — Item Yupi values in rowspan gap rows were
        // getting the pinned CSS treatment.
        //
        // z-index: 60 for thead th — MUST be higher than the global
        // `.data-table-scroll thead th { z-index: 56 }` rule, otherwise
        // non-pinned th elements (z-index 56) scroll OVER the pinned th,
        // making it look transparent/overlapping.
        rules.push(`
          /* Base sticky rule — applies to EVERY freeze-table-N table
             regardless of light/dark wrapper. Uses [data-col-index] so it
             works correctly even with rowspan merged cells. */
          .freeze-table-${tableKey} th[data-col-index="${idx}"],
          .freeze-table-${tableKey} td[data-col-index="${idx}"] {
            position: sticky !important;
            left: ${left}px;
            z-index: 25;
            box-shadow: 10px 0 14px -14px rgba(15, 23, 42, 0.55);
            background-clip: padding-box;
            /* Use inherit for tbody td so the pinned cell picks up the
               zebra-stripe color from its parent tr. This preserves the
               alternating row colors across the frozen pane. The thead th
               gets a solid color below. */
            background-color: inherit !important;
          }
          .freeze-table-${tableKey} thead th[data-col-index="${idx}"] {
            z-index: 60;
            background-color: #e2e8f0 !important;
          }
          .data-table-page-dark .freeze-table-${tableKey} thead th[data-col-index="${idx}"] {
            background-color: #374151 !important;
          }
          .freeze-table-${tableKey} tfoot td[data-col-index="${idx}"] {
            background-color: #f1f5f9 !important;
          }
          .data-table-page-dark .freeze-table-${tableKey} tfoot td[data-col-index="${idx}"] {
            background-color: #111827 !important;
          }
          /* Row hover should still tint the pinned cell so users see which
             row they're hovering across the frozen pane. */
          .freeze-table-${tableKey} tbody tr:hover td[data-col-index="${idx}"] {
            background-color: #f8fafc !important;
          }
          .data-table-page-dark .freeze-table-${tableKey} tbody tr:hover td[data-col-index="${idx}"] {
            background-color: #283548 !important;
          }
        `);
      }
      return rules.join('\n');
    }).join('\n');
  }, [frozenColumns, frozenColumnWidths]);
  // ── Global SO Create Date filter (shared across Dashboard / All SO / Delivery Completed)
  const [globalDateFilter, setGlobalDateFilter] = useState(() => savedDashboardFilters.dateFilter || { mode: 'all' });
  const [globalClientFilter, setGlobalClientFilter] = useState(() => savedDashboardFilters.clients || []);
  const [globalPicFilter, setGlobalPicFilter] = useState(() => savedDashboardFilters.pics || []);
  const [dashboardFilterOptions, setDashboardFilterOptions] = useState(() => {
    const cachedStats = readStatsCache(dashboardStatsCacheKey());
    return cachedStats?.filters || { clients: [], pics: [] };
  });
  // Aliases kept so existing references continue to compile.
  const dashDateFilter      = globalDateFilter;
  const setDashDateFilter   = setGlobalDateFilter;
  const soDateFilter        = globalDateFilter;
  const setSODateFilter     = setGlobalDateFilter;
  const completedDateFilter = globalDateFilter;
  const setCompletedDateFilter = setGlobalDateFilter;

  // Click-outside handlers

  useEffect(() => {
    const handler = (e) => { if (uploadDropdownRef.current && !uploadDropdownRef.current.contains(e.target)) setShowUploadDropdown(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  // ── Auto-persist every page's filter state to localStorage ───────────────
  // On any filter change, debounce-save so refresh restores the same view.
  // Mirrors Google Sheets: refresh never loses your place.
  useEffect(() => {
    saveFilterState('import', {
      page: importPage, perPage: importPerPage,
      search: importSearch, appliedSearch: importAppliedSearch,
      yupi_po: importFilters.yupi_po, vendors: importFilters.vendors,
      statuses: importFilters.statuses, daysLeft: importFilters.daysLeft,
      reqDlvSort: importReqDlvSort, yupiPoSort: importYupiPoSort,
    });
  }, [importPage, importPerPage, importSearch, importAppliedSearch, importFilters, importReqDlvSort, importYupiPoSort]);

  useEffect(() => {
    saveFilterState('rfq', {
      page: rfqPage, perPage: rfqPerPage,
      search: rfqSearch, appliedSearch: rfqAppliedSearch,
      picFilter: rfqPicFilter, filters: rfqFilters,
    });
  }, [rfqPage, rfqPerPage, rfqSearch, rfqAppliedSearch, rfqPicFilter, rfqFilters]);

  useEffect(() => {
    saveFilterState('item-registration', {
      page: itemRegPage, perPage: itemRegPerPage,
      search: itemRegSearch, appliedSearch: itemRegAppliedSearch,
      filters: itemRegFilters,
    });
  }, [itemRegPage, itemRegPerPage, itemRegSearch, itemRegAppliedSearch, itemRegFilters]);

  useEffect(() => {
    saveFilterState('all-registered-items', {
      page: registeredItemsPage, perPage: registeredItemsPerPage,
      search: registeredItemsSearch, appliedSearch: registeredItemsAppliedSearch,
      prodIds: registeredItemsProdIds, appliedProdIds: registeredItemsAppliedProdIds,
      picFilter: registeredItemsPicFilter, appliedPicFilter: registeredItemsAppliedPicFilter,
      mfr_names: registeredItemsFilters.mfr_names,
    });
  }, [registeredItemsPage, registeredItemsPerPage, registeredItemsSearch, registeredItemsAppliedSearch, registeredItemsProdIds, registeredItemsAppliedProdIds, registeredItemsPicFilter, registeredItemsAppliedPicFilter, registeredItemsFilters]);

  useEffect(() => {
    saveFilterState('vendor-control', {
      page: vendorControlPage, perPage: vendorControlPerPage,
      search: vendorControlSearch, appliedSearch: vendorControlAppliedSearch,
      appliedVendors: vendorControlAppliedVendors,
    });
  }, [vendorControlPage, vendorControlPerPage, vendorControlSearch, vendorControlAppliedSearch, vendorControlAppliedVendors]);

  useEffect(() => {
    saveFilterState('all-so', {
      page: soPage, perPage: soPerPage,
      searchNums: soSearchNums, filters: soFilters, marginFilter: soMarginFilter, sortOrder: soSortOrder,
    });
  }, [soPage, soPerPage, soSearchNums, soFilters, soMarginFilter, soSortOrder]);

  useEffect(() => {
    saveFilterState('dashboard', {
      dateFilter: globalDateFilter, clients: globalClientFilter, pics: globalPicFilter,
    });
  }, [globalDateFilter, globalClientFilter, globalPicFilter]);

  // ── Auto-populate data-col-index for tables that don't set it inline ──────
  // The pending-delivery table renders 27 <td> per row manually (not via
  // .map with index), so we can't easily add data-col-index in JSX. This
  // effect walks every freeze-table row and stamps each th/td with its
  // 1-based column index. Runs on mount + whenever table content changes
  // (via ResizeObserver). This is a safety net — tables that already set
  // data-col-index inline are not affected (we only set if missing).
  useEffect(() => {
    if (typeof document === 'undefined') return;
    const stamp = () => {
      document.querySelectorAll('table[class*="freeze-table-"]').forEach(table => {
        // Stamp thead
        table.querySelectorAll('thead tr').forEach(tr => {
          tr.querySelectorAll('th').forEach((th, i) => {
            if (!th.hasAttribute('data-col-index')) th.setAttribute('data-col-index', String(i + 1));
          });
        });
        // Stamp tbody — but rowspan complicates this. For each tr, count
        // columns by position (1-based), skipping cells covered by a
        // rowspan from a previous row.
        const colOffsets = []; // colOffsets[rowIndex] = array of column indices still covered by rowspan
        table.querySelectorAll('tbody tr').forEach((tr, rowIdx) => {
          let colIdx = 1;
          let tdIdx = 0;
          const tds = tr.querySelectorAll('td');
          const offset = colOffsets[rowIdx] || [];
          tds.forEach(td => {
            // Skip columns covered by rowspan from above
            while (offset.includes(colIdx)) colIdx++;
            if (!td.hasAttribute('data-col-index')) td.setAttribute('data-col-index', String(colIdx));
            const rs = parseInt(td.getAttribute('rowspan') || '1', 10);
            if (rs > 1) {
              for (let r = 1; r < rs; r++) {
                if (!colOffsets[rowIdx + r]) colOffsets[rowIdx + r] = [];
                colOffsets[rowIdx + r].push(colIdx);
              }
            }
            colIdx++;
            tdIdx++;
          });
        });
        // Stamp tfoot
        table.querySelectorAll('tfoot tr').forEach(tr => {
          tr.querySelectorAll('td').forEach((td, i) => {
            if (!td.hasAttribute('data-col-index')) td.setAttribute('data-col-index', String(i + 1));
          });
        });
      });
    };
    stamp();
    // Re-stamp periodically to catch re-renders. Use a MutationObserver for
    // immediate response + interval as fallback.
    const observer = typeof MutationObserver !== 'undefined' ? new MutationObserver(stamp) : null;
    if (observer) {
      document.querySelectorAll('table[class*="freeze-table-"]').forEach(t => observer.observe(t, { childList: true, subtree: true }));
    }
    const interval = setInterval(stamp, 500);
    return () => {
      if (observer) observer.disconnect();
      clearInterval(interval);
    };
  }, [activePage]);

  // ── Toast helpers (declared FIRST so the offline-queue logic below can
  // reference them without hitting a temporal-dead-zone error). ──────────────
  const addToast = useCallback((message, type='success') => {
    const id = Date.now(); setToasts(t => [...t, { id, message, type }]);
  }, []);
  const removeToast = useCallback((id) => setToasts(t => t.filter(x => x.id !== id)), []);

  // ── Offline queue replay ─────────────────────────────────────────────────
  // When the browser fires `online` (or the app mounts with an existing
  // queue and the connection is up), drain the queue and replay every
  // pending edit in order. On success the entry is removed; on failure it
  // stays in the queue for the next attempt.
  const replayOfflineQueue = useCallback(async () => {
    const queue = loadOfflineQueue();
    if (!queue.length) return;
    let remaining = [...queue];
    let changed = false;
    for (let i = 0; i < queue.length; i++) {
      const item = queue[i];
      try {
        if (item.kind === 'import-cell') {
          await api.put('/api/import/cell', item.payload);
        } else if (item.kind === 'import-cells') {
          await api.put('/api/import/cells', item.payload);
        } else if (item.kind === 'rfq-cell') {
          await api.put(`/api/rfq/${encodeURIComponent(item.payload.row_key)}`, { field: item.payload.field, value: item.payload.value });
        } else if (item.kind === 'rfq-cells') {
          await api.put('/api/rfq/batch-cells', item.payload);
        } else {
          // Unknown kind — drop it so we don't loop forever.
          remaining = remaining.filter((_, j) => j !== i);
          changed = true;
          continue;
        }
        // Success: remove this entry from remaining.
        remaining = remaining.filter(x => x !== item);
        changed = true;
      } catch (e) {
        // Stop on first failure — connection probably still bad. Leave
        // this entry and everything after it in the queue for the next
        // `online` event.
        break;
      }
    }
    if (changed) {
      saveOfflineQueue(remaining);
      if (remaining.length === 0) {
        addToast(`Offline edits synced successfully`, 'success');
      } else {
        addToast(`${remaining.length} edit(s) still pending — will retry`, 'warning');
      }
    }
  }, [addToast]);

  // Replay on mount (if online) + whenever `online` event fires.
  useEffect(() => {
    if (typeof window === 'undefined') return;
    if (typeof navigator !== 'undefined' && navigator.onLine) {
      // Slight delay so React state (rfqData, importData) has time to mount.
      const t = setTimeout(replayOfflineQueue, 1500);
      return () => clearTimeout(t);
    }
    const onOnline = () => {
      addToast(`Connection restored — syncing offline edits...`, 'success');
      replayOfflineQueue();
    };
    window.addEventListener('online', onOnline);
    return () => window.removeEventListener('online', onOnline);
  }, [replayOfflineQueue, addToast]);

  // Also expose the pending count so the user can see it (optional badge).
  const [offlineQueueCount, setOfflineQueueCount] = useState(0);
  useEffect(() => {
    if (typeof window === 'undefined') return;
    const update = () => setOfflineQueueCount(loadOfflineQueue().length);
    update();
    window.addEventListener('online', update);
    window.addEventListener('storage', update);
    return () => {
      window.removeEventListener('online', update);
      window.removeEventListener('storage', update);
    };
  }, []);

  function appendMultiParam(params, key, value) {
    if (value === '__NONE__') {
      params.append(key, '__NONE_PLACEHOLDER__');
      return;
    }
    (Array.isArray(value) ? value : []).forEach(v => params.append(key, v));
  }

  const fetchDashboard = useCallback(async (dateFilter) => {
    const requestId = dashboardRequestSeq.current + 1;
    dashboardRequestSeq.current = requestId;
    const isCurrent = () => dashboardRequestSeq.current === requestId;
    const f = dateFilter || globalDateFilter;
    const params = new URLSearchParams();
    Object.entries(dateFilterParams(f)).forEach(([key, value]) => { if (value) params.append(key, value); });
    appendMultiParam(params, 'client', globalClientFilter);
    appendMultiParam(params, 'pic', globalPicFilter);
    const qs = params.toString() ? `?${params}` : '';

    const completedParams = new URLSearchParams();
    params.forEach((value, key) => completedParams.append(key, value));
    // Dashboard uses a lightweight SQL-aggregated payload. Drilldown details stay lazy.
    completedParams.set('mode', 'dashboard');

    const summaryUrl = (summaryParams) => {
      const summaryQs = summaryParams.toString();
      return summaryQs ? `/api/completed/summary?${summaryQs}` : '/api/completed/summary';
    };
    const completedUrl = summaryUrl(completedParams);
    const cachedCompleted = readDashboardSummaryCache(completedUrl);
    const hasSummaryCache = Boolean(cachedCompleted);

    const statsCacheKey = dashboardStatsCacheKey(qs);
    const agingCacheKey = dashboardAgingCacheKey(qs);
    const pendingCacheKey = dashboardPendingCacheKey(qs);
    const cachedStats = readStatsCache(statsCacheKey);
    const cachedAging = readStatsCache(agingCacheKey);
    const cachedPending = readStatsCache(pendingCacheKey);
    const pendingNumber = Number(cachedPending?.total ?? cachedStats?.total_so_count);
    const hasStatsCache = Boolean(cachedStats && Number.isFinite(pendingNumber));
    const hasAgingCache = Array.isArray(cachedAging);

    if (hasStatsCache && isCurrent()) {
      setStats(cachedStats);
      setSummaryPendingTotal(pendingNumber);
      setDashboardFilterOptions(cachedStats?.filters || { clients: [], pics: [] });
      setLoading(false);
      setInitialPageLoading(false);
    }
    if (hasAgingCache && isCurrent()) {
      setAgingData(cachedAging);
    }

    setCompletedLoading(!hasSummaryCache);
    setCompletedLoaded(hasSummaryCache);
    setCompletedData(hasSummaryCache ? cachedCompleted : null);
    setDashboardMarginData(hasSummaryCache ? cachedCompleted : null);

    // If the light Dashboard and completed chart are already cached, do not touch PythonAnywhere.
    if (hasStatsCache && hasAgingCache && hasSummaryCache) {
      setLoading(false);
      setInitialPageLoading(false);
      setCompletedLoading(false);
      return;
    }

    // Ping is intentionally DB-free in the backend now, so it only wakes the worker.
    api.get('/api/ping').catch(() => {});

    if (!hasStatsCache) {
      setLoading(true);
      setInitialPageLoading(true);
      // Retry with short backoff before giving up. /api/dashboard/stats is the
      // heaviest endpoint on the dashboard (several SQL aggregate queries), so
      // it's the most likely one to time out on a cold/slow PythonAnywhere
      // worker — which is why "SO"/"Reg" in the header can go blank while the
      // lighter /api/master-pic/status (Prod ID) still succeeds.
      let sRes = null;
      let lastErr = null;
      for (let attempt = 0; attempt <= 2; attempt++) {
        try {
          sRes = await api.get(`/api/dashboard/stats${qs}`);
          lastErr = null;
          break;
        } catch (e) {
          lastErr = e;
          if (attempt < 2) await new Promise(r => setTimeout(r, 1500 * (attempt + 1)));
        }
      }
      if (!isCurrent()) return;
      if (sRes) {
        const nextStats = sRes.data || {};
        const nextPending = { total: Number(nextStats.total_so_count) || 0 };
        setStats(nextStats);
        setSummaryPendingTotal(nextPending.total);
        setDashboardFilterOptions(nextStats?.filters || { clients: [], pics: [] });
        writeStatsCache(statsCacheKey, nextStats);
        writeStatsCache(pendingCacheKey, nextPending);
      } else {
        // All retries failed: keep whatever stats we already have (e.g. from
        // an older cache entry that just expired) instead of blanking the
        // "Updates" timestamps in the header to '-'.
        addToast(`Error: ${lastErr?.response?.data?.error || lastErr?.message}`, 'error');
        setCompletedLoading(false);
        setLoading(false);
        setInitialPageLoading(false);
        return;
      }
      if (isCurrent()) {
        setLoading(false);
        setInitialPageLoading(false);
      }
    } else {
      setLoading(false);
      setInitialPageLoading(false);
    }

    if (!hasAgingCache) {
      api.get(`/api/data/aging${qs}`)
        .then((aRes) => {
          if (!isCurrent()) return;
          const nextAging = Array.isArray(aRes.data) ? aRes.data : [];
          setAgingData(nextAging);
          writeStatsCache(agingCacheKey, nextAging);
        })
        .catch((e) => {
          if (isCurrent()) addToast(`Error memuat aging: ${e.response?.data?.error || e.message}`, 'error');
        });
    }

    if (hasSummaryCache) {
      setCompletedLoading(false);
      return;
    }

    // Delay the heavier completed chart until the Dashboard has painted.
    // requestIdleCallback keeps page switches/table loads from being blocked by summary processing.
    await new Promise((resolve) => {
      if (typeof window !== 'undefined' && 'requestIdleCallback' in window) {
        const timeout = window.setTimeout(resolve, 1800);
        window.requestIdleCallback(() => {
          window.clearTimeout(timeout);
          resolve();
        }, { timeout: 1800 });
      } else {
        window.setTimeout(resolve, 1200);
      }
    });
    if (!isCurrent() || activePage !== 'dashboard') {
      setCompletedLoading(false);
      return;
    }

    try {
      const res = await api.get(completedUrl);
      if (!isCurrent()) return;
      writeDashboardSummaryCache(completedUrl, res.data);
      setCompletedData(res.data);
      setDashboardMarginData(res.data);
      setCompletedLoaded(true);
    } catch (e) {
      if (isCurrent()) {
        setCompletedLoaded(true);
        addToast(`Error memuat summary: ${e.response?.data?.error || e.message}`, 'error');
      }
    } finally {
      if (isCurrent()) setCompletedLoading(false);
    }
  }, [addToast, activePage, globalDateFilter, globalClientFilter, globalPicFilter]);

  // Helper: filter array of objects by date field using a DateRangeFilter config
  const applyDateFilter = useCallback((arr, dateField, filter) => {
    if (!filter || filter.mode === 'all') return arr;
    const bounds = getDateFilterBounds(filter);
    return arr.filter(item => {
      const d = item[dateField];
      if (!d) return false;
      const iso = d.slice(0, 10);
      if (bounds.date_from && iso < bounds.date_from) return false;
      if (bounds.date_to && iso > bounds.date_to) return false;
      return true;
    });
  }, []);

  // Helper: build date query params for backend
  const dateFilterParams = (filter) => {
    if (!filter || filter.mode === 'all') return {};
    if (filter.mode === 'range') return { date_from: filter.start || '', date_to: filter.end || '' };
    return getDateFilterBounds(filter);
  };

  const appendDateQuery = (url, filter = globalDateFilter) => {
    const params = new URLSearchParams(dateFilterParams(filter));
    appendMultiParam(params, 'client', globalClientFilter);
    appendMultiParam(params, 'pic', globalPicFilter);
    const qs = params.toString();
    if (!qs) return url;
    return `${url}${url.includes('?') ? '&' : '?'}${qs}`;
  };

  const openNegativeVendorDetail = async (vendor) => {
    try {
      const params = new URLSearchParams({ category: 'negative' });
      Object.entries(dateFilterParams(completedDateFilter)).forEach(([key, value]) => { if (value) params.append(key, value); });
      appendMultiParam(params, 'client', globalClientFilter);
      appendMultiParam(params, 'pic', globalPicFilter);
      const res = await api.get(`/api/completed/margin-detail?${params}`);
      const rows = (Array.isArray(res.data) ? res.data : []).filter(row => String(row.vendor || '-') === String(vendor || '-'));
      setMarginDetailModal({ category: `Vendor: ${vendor || '-'}`, data: rows });
    } catch(e) {
      addToast(`Failed to load vendor margin detail: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  // Helper: resolve filter array
  const resolveFilter = (val) => {
    if (val === '__NONE__') return ['__NONE_PLACEHOLDER__']; // backend will return 0 rows
    if (!Array.isArray(val) || val.length === 0) return []; // empty = no filter = all
    return val;
  };
  const filterValues = (val) => Array.isArray(val) ? val : [];

  const fetchSOData = useCallback(async (filters, page, perPage, searchNums, marginFilter, dateFilter, sortOrder = soSortOrder, kpiPic = pendingPicHighlight) => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage, sort_order: sortOrder });
      resolveFilter(filters.op_units).forEach(v => params.append('op_unit', v));
      resolveFilter(filters.vendors).forEach(v => params.append('vendor', v));
      resolveFilter(filters.manufacturers).forEach(v => params.append('manufacturer', v));
      resolveFilter(filters.statuses).forEach(v => params.append('status', v));
      resolveFilter(filters.pics).forEach(v => params.append('pic', v));
      if (kpiPic) params.append('kpi_pic', kpiPic);
      filterValues(filters.aging).forEach(a => params.append('aging', a));
      (searchNums || []).forEach(n => params.append('so_item', n));
      appendMultiParam(params, 'client', globalClientFilter);
      appendMultiParam(params, 'global_pic', globalPicFilter);
      if (marginFilter && marginFilter !== 'all') params.append('margin_filter', marginFilter);
      Object.entries(dateFilterParams(dateFilter)).forEach(([key, value]) => { if (value) params.append(key, value); });
      const res = await api.get(`/api/data/all-so?${params}`);
      setAllSOData(Array.isArray(res.data.data) ? res.data.data : []);
      setApprovalSOData(Array.isArray(res.data.approval_data) ? res.data.approval_data : []);
      setPicAggregations(Array.isArray(res.data.pic_aggregations) ? res.data.pic_aggregations : []);
      setSoTotal(res.data.total || 0);
      setSoSubtotalAmount(Number(res.data.subtotal_amount) || 0);
      setSoFilterOptions(res.data.filters || { op_units: [], vendors: [], manufacturers: [], statuses: [], pics: [] });
    } catch (e) {
      addToast(`Failed to load SO: ${e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast, soSortOrder, pendingPicHighlight, globalClientFilter, globalPicFilter]);

  const fetchItemRegistration = useCallback(async (page = itemRegPage, perPage = itemRegPerPage, search = itemRegAppliedSearch, filters = itemRegFilters, kpiPic = itemRegPicHighlight, dateFilter = globalDateFilter) => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      Object.entries(dateFilterParams(dateFilter)).forEach(([key, value]) => { if (value) params.append(key, value); });
      appendMultiParam(params, 'client', globalClientFilter);
      appendMultiParam(params, 'global_pic', globalPicFilter);
      if (Array.isArray(search)) search.forEach(v => params.append('req_no', v));
      else if (search) params.append('search', search);
      resolveFilter(filters.clients).forEach(v => params.append('item_client', v));
      resolveFilter(filters.categories).forEach(v => params.append('category', v));
      resolveFilter(filters.pics).forEach(v => params.append('pic', v));
      resolveFilter(filters.proc_statuses).forEach(v => params.append('proc_status', v));
      resolveFilter(filters.mfr_names).forEach(v => params.append('mfr_name', v));
      if (kpiPic) params.append('kpi_pic', kpiPic);
      const res = await api.get(`/api/item-registration/data?${params}`);
      const rows = Array.isArray(res.data.data) ? res.data.data : [];
      setItemRegData(rows);
      setItemRegTotal(res.data.total || 0);
      setItemRegLastUpdated(res.data.last_updated || null);
      setItemRegMissingPicKpis(Array.isArray(res.data.missing_prod_id_by_pic) ? res.data.missing_prod_id_by_pic : []);
      setItemRegOptions({
        clients: res.data.client_options || [],
        categories: res.data.category_options || [],
        pics: res.data.pic_options || [],
        proc_statuses: res.data.proc_status_options || [],
        mfr_names: res.data.mfr_name_options || []
      });
    } catch (e) {
      addToast(`Failed to load Item Registration: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast, itemRegPage, itemRegPerPage, itemRegAppliedSearch, itemRegFilters, itemRegPicHighlight, globalDateFilter, globalClientFilter, globalPicFilter]);

  const fetchRegisteredItems = useCallback(async (
    page = registeredItemsPage,
    perPage = registeredItemsPerPage,
    search = registeredItemsAppliedSearch,
    prodIds = registeredItemsAppliedProdIds,
    filters = registeredItemsFilters,
    picFilter = registeredItemsAppliedPicFilter
  ) => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      if (search) params.append('search', search);
      (prodIds || []).forEach(v => params.append('prod_id', v));
      resolveFilter(filters.mfr_names).forEach(v => params.append('mfr_name', v));
      // vendor_name param intentionally omitted — source has no Vendor column.
      if (picFilter) params.append('pic_name', picFilter);
      const res = await api.get(`/api/all-registered-items?${params}`);
      setRegisteredItemsData(Array.isArray(res.data.data) ? res.data.data : []);
      setRegisteredItemsTotal(res.data.total || 0);
      setRegisteredItemsOptions(res.data.filters || { mfr_names: [], pic_options: [] });
    } catch (e) {
      addToast(`Failed to load All Registered Items: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [
    addToast,
    registeredItemsPage,
    registeredItemsPerPage,
    registeredItemsAppliedSearch,
    registeredItemsAppliedProdIds,
    registeredItemsFilters,
    registeredItemsAppliedPicFilter
  ]);

  const fetchRFQData = useCallback(async (page = rfqPage, perPage = rfqPerPage, search = rfqAppliedSearch, refresh = false, filters = rfqFilters, pic = rfqPicFilter, showSimilarity = rfqShowSimilarity) => {
    setRfqEditedRowKeys(new Set());
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      if (search) params.append('search', search);
      if (refresh) params.append('refresh', '1');
      if (pic) params.append('pic', pic);
      if (showSimilarity) params.append('similarity', '1');
      resolveFilter(filters.checks).forEach(v => params.append('check', v));
      resolveFilter(filters.clients).forEach(v => params.append('client_name', v));
      resolveFilter(filters.rfq_numbers).forEach(v => params.append('rfq_no', v));
      resolveFilter(filters.brands).forEach(v => params.append('brand_manufacturer', v));
      resolveFilter(filters.purchase_pics)
        .filter(v => String(v || '').trim().toLowerCase() !== 'unassigned')
        .forEach(v => params.append('purchase_pic', v));
      resolveFilter(filters.vendors).forEach(v => params.append('vendor_name', v));
      const res = await api.get(`/api/rfq/data?${params}`);
      setRfqData(Array.isArray(res.data.data) ? res.data.data : []);
      setRfqTotal(res.data.total || 0);
      setRfqColumns(Array.isArray(res.data.columns) ? res.data.columns : []);
      setRfqSimilarityColumns(Array.isArray(res.data.similarity_columns) ? res.data.similarity_columns : []);
      setRfqEditableFields(Array.isArray(res.data.editable_fields) ? res.data.editable_fields : []);
      setRfqPicKpis(Array.isArray(res.data.pic_kpis) ? res.data.pic_kpis : []);
      const nextOptions = res.data.filters || { checks: [], clients: [], rfq_numbers: [], brands: [], purchase_pics: [], vendors: [] };
      setRfqOptions({
        ...nextOptions,
        purchase_pics: (nextOptions.purchase_pics || []).filter(v => String(v || '').trim().toLowerCase() !== 'unassigned')
      });
      setRfqLastUpdated(res.data.last_updated || null);
    } catch (e) {
      addToast(`Failed to load RFQ: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast, rfqPage, rfqPerPage, rfqAppliedSearch, rfqFilters, rfqPicFilter, rfqShowSimilarity]);

  const fetchImportData = useCallback(async (page = importPage, perPage = importPerPage, search = importAppliedSearch, refresh = false, filters = importFilters, reqDlvSort = importReqDlvSort, yupiPoSort = importYupiPoSort) => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      if (search) params.append('search', search);
      if (refresh) params.append('refresh', '1');
      if (reqDlvSort) params.append('req_dlv_sort', reqDlvSort);
      if (yupiPoSort) params.append('yupi_po_sort', yupiPoSort);
      resolveFilter(filters?.yupi_po).forEach(v => params.append('yupi_po', v));
      resolveFilter(filters?.vendors).forEach(v => params.append('vendor_name', v));
      resolveFilter(filters?.statuses).forEach(v => params.append('status', v));
      resolveFilter(filters?.daysLeft).forEach(v => params.append('days_left', v));
      const res = await api.get(`/api/import/data?${params}`);
      setImportData(Array.isArray(res.data.data) ? res.data.data : []);
      setImportColumns(Array.isArray(res.data.columns) ? res.data.columns : []);
      setImportTotal(res.data.total || 0);
      setImportVendorCount(res.data.vendor_count || 0);
      setImportLastCopyAt(res.data.last_copy_at || '');
      // Store KPIs returned by backend (computed across all filtered rows).
      if (res.data.kpis && typeof res.data.kpis === 'object') {
        setImportKpis({
          total_po: Number(res.data.kpis.total_po) || 0,
          this_week_arrival: Number(res.data.kpis.this_week_arrival) || 0,
          this_week_no_sap: Number(res.data.kpis.this_week_no_sap) || 0,
          sales_amount: Number(res.data.kpis.sales_amount) || 0,
          po_amount_idr: Number(res.data.kpis.po_amount_idr) || 0,
          gross_margin: Number(res.data.kpis.gross_margin) || 0,
        });
      }
      // Preserve existing status options list (server may not return it).
      setImportOptions(prev => ({
        yupi_po: res.data.filters?.yupi_po || [],
        vendors: res.data.filters?.vendors || [],
        statuses: prev.statuses?.length ? prev.statuses : ['NEW', ...IMPORT_STATUS_OPTIONS],
      }));
      if (refresh && res.data.sync) {
        const added = Number(res.data.sync.added || 0);
        const seen = Number(res.data.sync.seen || 0);
        const sheetRows = Number(res.data.sync.sheet_rows || 0);
        const vendorFilterCount = Number(res.data.sync.vendor_filter_count || 0);
        const vendorSource = res.data.sync.vendor_filter_source === 'existing_import_rows' ? 'current Import vendors' : 'Vendor Import';
        const msg = added
          ? `Copied ${added} new Import rows and skipped ${seen} existing rows`
          : sheetRows
            ? `No new rows. ${seen} existing Import rows already in dashboard (${vendorFilterCount} ${vendorSource})`
            : 'No Import rows found in source sheets';
        addToast(msg, 'success');
      }
    } catch (e) {
      addToast(`Failed to load Import data: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast, importPage, importPerPage, importAppliedSearch, importFilters, importReqDlvSort, importYupiPoSort]);

  const updateImportCell = async (rowKey, field, value) => {
    setImportEditingCell(null);
    const previousRows = importData;
    // Group-level fields: when editing po_send_date or status on one row,
    // sync the same value to ALL rows in the same group (same YUPI PO +
    // Req Dlv Date). This ensures one PO has one send date and one status,
    // preventing the "2 different statuses for same PO" bug.
    const GROUP_SYNC_FIELDS = new Set(['po_send_date', 'status']);
    const editedRow = importData.find(r => r._row_key === rowKey);
    const isGroupSync = GROUP_SYNC_FIELDS.has(field) && editedRow;
    // Find all rows in the same group (same YUPI PO + Req Dlv Date).
    const groupRowKeys = isGroupSync
      ? importData
          .filter(r => String(r.yupi_po ?? '').trim() === String(editedRow.yupi_po ?? '').trim()
                     && String(r.req_dlv_date ?? '').trim() === String(editedRow.req_dlv_date ?? '').trim()
                     && String(r.yupi_po ?? '').trim() !== '')
          .map(r => r._row_key)
      : [rowKey];
    // OPTIMISTIC: update UI immediately for ALL rows in the group.
    setImportData(prev => prev.map(row => groupRowKeys.includes(row._row_key) ? { ...row, [field]: value } : row));
    try {
      // Send batch update to backend so all rows in the group get persisted.
      if (isGroupSync && groupRowKeys.length > 1) {
        const res = await api.put('/api/import/cells', {
          updates: groupRowKeys.map(rk => ({ row_key: rk, field, value })),
        });
        if (Array.isArray(res.data?.rows)) {
          const byKey = new Map(res.data.rows.map(r => [r._row_key, r]));
          setImportData(prev => prev.map(row => byKey.has(row._row_key) ? { ...row, ...byKey.get(row._row_key) } : row));
        }
      } else {
        const res = await api.put('/api/import/cell', { row_key: rowKey, field, value });
        if (res.data?.row) {
          setImportData(prev => prev.map(row => row._row_key === rowKey ? { ...row, ...res.data.row } : row));
        }
      }
      return true;
    } catch (e) {
      // Network/HTTP failure → queue for replay when online. DO NOT revert
      // the local edit (user input is preserved).
      if (typeof navigator === 'undefined' || !navigator.onLine || e.message?.includes('Network') || e.code === 'ERR_NETWORK' || e.response == null) {
        if (isGroupSync && groupRowKeys.length > 1) {
          enqueueOfflineUpdate('import-cells', { updates: groupRowKeys.map(rk => ({ row_key: rk, field, value })) });
        } else {
          enqueueOfflineUpdate('import-cell', { row_key: rowKey, field, value });
        }
        addToast(`Saved offline — will sync when connection returns`, 'warning');
      } else {
        addToast(`Failed to update Import data: ${e.response?.data?.error || e.message}`, 'error');
      }
      return false;
    }
  };

  const updateImportCellsBatch = async (updates) => {
    const safeUpdates = Array.isArray(updates) ? updates.filter(u => u?.row_key && u?.field) : [];
    if (!safeUpdates.length) return false;
    const previousRows = importData;
    setImportEditingCell(null);
    // OPTIMISTIC: apply all updates to local state immediately.
    setImportData(prev => {
      const next = prev.map(row => ({ ...row }));
      for (const update of safeUpdates) {
        const idx = next.findIndex(row => row._row_key === update.row_key);
        if (idx >= 0) next[idx][update.field] = update.value;
      }
      return next;
    });
    try {
      const res = await api.put('/api/import/cells', { updates: safeUpdates });
      const updatedRows = Array.isArray(res.data?.rows) ? res.data.rows : [];
      if (updatedRows.length) {
        const byKey = new Map(updatedRows.map(row => [row._row_key, row]));
        setImportData(prev => prev.map(row => byKey.has(row._row_key) ? { ...row, ...byKey.get(row._row_key) } : row));
      }
      return true;
    } catch (e) {
      // Network/HTTP failure → queue for replay. Keep local edits intact.
      if (typeof navigator === 'undefined' || !navigator.onLine || e.message?.includes('Network') || e.code === 'ERR_NETWORK' || e.response == null) {
        enqueueOfflineUpdate('import-cells', { updates: safeUpdates });
        addToast(`Saved offline — will sync when connection returns`, 'warning');
      } else {
        addToast(`Failed to update Import data: ${e.response?.data?.error || e.message}`, 'error');
      }
      return false;
    }
  };

  const fetchVendorControl = useCallback(async (page = vendorControlPage, perPage = vendorControlPerPage, search = vendorControlAppliedSearch, refresh = false, vendors = vendorControlAppliedVendors) => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      if (search) params.append('search', search);
      (vendors || []).forEach(v => params.append('vendor', v));
      if (refresh) params.append('refresh', '1');
      const res = await api.get(`/api/vendor-control/data?${params}`);
      setVendorControlData(Array.isArray(res.data.data) ? res.data.data : []);
      setVendorControlTotal(res.data.total || 0);
      setVendorControlSuggestions(Array.isArray(res.data.suggestions) ? res.data.suggestions : []);
      setVendorControlLastUpdated(res.data.last_updated || null);
    } catch (e) {
      addToast(`Failed to load Vendor Control: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast, vendorControlPage, vendorControlPerPage, vendorControlAppliedSearch, vendorControlAppliedVendors]);

  const updateVendorControlCell = async (rowKey, field, value) => {
    setEditingCell(null);
    const previousRows = vendorControlData;
    setVendorControlData(prev => prev.map(row => row.row_key === rowKey ? { ...row, [field]: value } : row));
    try {
      const res = await api.put(`/api/vendor-control/${encodeURIComponent(rowKey)}`, { field, value });
      if (res.data?.sheet_sync && res.data.sheet_sync.synced === false) {
        addToast(`Vendor updated locally. Sheet sync not active: ${res.data.sheet_sync.reason}`, 'warning');
      }
    } catch (e) {
      setVendorControlData(previousRows);
      addToast(`Failed to update Vendor Control: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  const openVendorLogin = (row) => {
    if (!row?.row_key) return;
    window.open(`${BACKEND}/api/vendor-control/login/${encodeURIComponent(row.row_key)}`, '_blank', 'noopener,noreferrer');
  };

  useEffect(() => { fetchPicDbStatus(); }, []);
  // Always fetch dashboard stats on mount (regardless of active page) so the
  // "Updates:" timestamps in the header are populated and consistent across
  // ALL pages — not just the Summary page.
  useEffect(() => {
    // One-time fetch on mount if stats cache is empty/expired.
    const cached = readStatsCache(dashboardStatsCacheKey());
    if (!cached) fetchDashboard(globalDateFilter);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useEffect(() => {
    if (activePage === 'dashboard') {
      fetchDashboard(globalDateFilter);
      return;
    }
    // Invalidate any Dashboard request that is still waiting for summary so it
    // cannot keep the global loading overlay alive after the user changes page.
    dashboardRequestSeq.current += 1;
    setCompletedLoading(false);
    setLoading(false);
    setInitialPageLoading(false);
  }, [activePage, globalDateFilter, globalClientFilter, globalPicFilter, fetchDashboard]);
  useEffect(() => {
    if (activePage === 'all-so') {
      fetchSOData(soFilters, soPage, soPerPage, soSearchNums, soMarginFilter, soDateFilter, soSortOrder);
    }
  }, [activePage, soSortOrder, soPage, soPerPage, soFilters, soSearchNums, soMarginFilter, soDateFilter, globalClientFilter, globalPicFilter, fetchSOData]);

  useEffect(() => {
    if (activePage === 'item-registration') {
      fetchItemRegistration(itemRegPage, itemRegPerPage, itemRegAppliedSearch, itemRegFilters, itemRegPicHighlight, globalDateFilter);
    }
  }, [activePage, itemRegPage, itemRegPerPage, itemRegAppliedSearch, itemRegFilters, itemRegPicHighlight, globalDateFilter, globalClientFilter, globalPicFilter, fetchItemRegistration]);

  useEffect(() => {
    if (activePage === 'rfq') {
      fetchRFQData(rfqPage, rfqPerPage, rfqAppliedSearch, true, rfqFilters, rfqPicFilter, rfqShowSimilarity);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activePage]);

  useEffect(() => {
    if (activePage === 'import') {
      fetchImportData(importPage, importPerPage, importAppliedSearch);
    }
  }, [activePage]);

  useEffect(() => {
    if (activePage === 'vendor-control') {
      fetchVendorControl(vendorControlPage, vendorControlPerPage, vendorControlAppliedSearch, false, vendorControlAppliedVendors);
    }
  }, [activePage, vendorControlPage, vendorControlPerPage, vendorControlAppliedSearch, vendorControlAppliedVendors, fetchVendorControl]);

  useEffect(() => {
    const q = vendorControlSearch.trim();
    if (activePage !== 'vendor-control' || q.length < 2) {
      setVendorControlSuggestions([]);
      setVendorControlSuggestOpen(false);
      return;
    }
    let cancelled = false;
    const timer = setTimeout(async () => {
      try {
        const params = new URLSearchParams({ page: 1, per_page: 1, search: q });
        const res = await api.get(`/api/vendor-control/data?${params}`);
        if (cancelled) return;
        const suggestions = Array.isArray(res.data.suggestions) ? res.data.suggestions : [];
        setVendorControlSuggestions(suggestions);
        setVendorControlSuggestOpen(suggestions.length > 0);
      } catch {
        if (!cancelled) {
          setVendorControlSuggestions([]);
          setVendorControlSuggestOpen(false);
        }
      }
    }, 220);
    return () => { cancelled = true; clearTimeout(timer); };
  }, [activePage, vendorControlSearch]);

  useEffect(() => {
    if (activePage === 'all-registered-items') {
      fetchRegisteredItems(
        registeredItemsPage,
        registeredItemsPerPage,
        registeredItemsAppliedSearch,
        registeredItemsAppliedProdIds,
        registeredItemsFilters,
        registeredItemsAppliedPicFilter
      );
    }
  }, [
    activePage,
    registeredItemsPage,
    registeredItemsPerPage,
    registeredItemsAppliedSearch,
    registeredItemsAppliedProdIds,
    registeredItemsAppliedPicFilter,
    fetchRegisteredItems
  ]);

  const handleUpload = async (e, type) => {
    const files = Array.from(e.target.files || []); if (!files.length) return;
    e.target.value = '';
    const label = files.length > 1 ? `SO - Search Client Odr (${files.length} files)` : 'SO - Search Client Odr';
    const endpoint = '/api/upload/smro';

    const fd = new FormData(); files.forEach(file => fd.append('file', file));
    setUploadProgress({ label, pct: 0 });
    try {
      const res = await api.post(endpoint, fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        onUploadProgress: (ev) => setUploadProgress({ label, pct: Math.round(ev.loaded*100/(ev.total||ev.loaded)) })
      });
      setUploadProgress(null);
      // Combine success message and SO Specification/Product ID diagnostics
      // into a single toast so the success and diagnostic don't get the same
      // Date.now() ID and trample each other.
      const diag = res.data.diagnostics;
      let toastMsg = `✅ ${res.data.message}`;
      let toastKind = 'success';
      if (type === 'scor' && diag) {
        const det = diag.columns_detected || {};
        const detail = `Specification col=${det.specification||'(none)'}; Product ID col=${det.product_id||'(none)'}. ` +
                       `Filled rows: spec=${diag.rows_with_specification||0}, pid=${diag.rows_with_product_id||0}.`;
        toastMsg += diag.warning ? `\n⚠️ ${diag.warning} ${detail}` : `\nℹ️ ${detail}`;
        if (diag.warning) toastKind = 'warning';
      }
      addToast(toastMsg, toastKind);
      clearDashboardSummaryCache();
      clearStatsCache();
      if (activePage === 'dashboard') fetchDashboard();
      if (activePage === 'all-so') fetchSOData(soFilters, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
      setSoPage(1);
    } catch (e) {
      setUploadProgress(null);
      addToast(`❌ Failed to upload ${label}: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  const fetchPicDbStatus = async (retryCount = 0) => {
    try {
      const res = await api.get('/api/master-pic/status');
      setPicDbStatus(res.data);
      writePicDbStatusCache(res.data);
    } catch (err) {
      // Transient failures (cold-start, CORS hiccup on the free-tier backend
      // waking up, etc.) shouldn't blank out the header. Retry once after a
      // short delay; if it still fails, just keep whatever we last had
      // (from cache or a previous successful fetch) instead of resetting it.
      if (retryCount < 2) {
        setTimeout(() => fetchPicDbStatus(retryCount + 1), 1500 * (retryCount + 1));
      }
    }
  };

  const handleUploadProductID = async (e) => {
    const files = Array.from(e.target.files || []); if (!files.length) return;
    e.target.value = '';
    const fd = new FormData(); files.forEach(file => fd.append('file', file));
    const label = files.length > 1 ? `Product ID (${files.length} files)` : 'Product ID';
    setPicUploadMsg('⏳ Uploading Product ID database...');
    setUploadProgress({ label, pct: 0 });
    try {
      const res = await api.post('/api/upload/product-id', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        onUploadProgress: (ev) => setUploadProgress({ label, pct: Math.round(ev.loaded*100/(ev.total||ev.loaded)) })
      });
      const d = res.data;
      setUploadProgress(null);
      setPicUploadMsg(`✅ Prod ID (${d.files || files.length} file): +${d.added} added, ${d.updated} updated (total: ${d.total_in_db}). SO PIC refreshed: ${d.so_pic_refreshed} rows.`);
      clearDashboardSummaryCache();
      clearStatsCache();
      fetchPicDbStatus();
      if (activePage === 'all-so') fetchSOData(soFilters, soPage, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
    } catch (err) {
      setUploadProgress(null);
      const msg = err?.response?.data?.error || err.message;
      setPicUploadMsg(`❌ Error: ${msg}`);
      addToast(`❌ ${msg}`, 'error');
    }
  };

  const handleUpdatePIC = async (e) => {
    const files = Array.from(e.target.files || []); if (!files.length) return;
    e.target.value = '';
    const fd = new FormData(); files.forEach(file => fd.append('file', file));
    const label = files.length > 1 ? `Master PIC (${files.length} files)` : 'Master PIC';
    setPicUploadMsg('⏳ Updating Master PIC...');
    setUploadProgress({ label, pct: 0 });
    try {
      const res = await api.post('/api/upload/master-pic', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        onUploadProgress: (ev) => setUploadProgress({ label, pct: Math.round(ev.loaded*100/(ev.total||ev.loaded)) })
      });
      const d = res.data;
      setUploadProgress(null);
      setPicUploadMsg(`✅ Master PIC (${d.files || files.length} file): +${d.added} added, ${d.updated} updated${d.unchanged ? `, ${d.unchanged} unchanged` : ''} (total category names: ${d.total_categories}). SO rows updated: ${d.so_pic_refreshed}.`);
      clearDashboardSummaryCache();
      clearStatsCache();
      fetchPicDbStatus();
      if (activePage === 'all-so') fetchSOData(soFilters, soPage, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
    } catch (err) {
      setUploadProgress(null);
      const msg = err?.response?.data?.error || err.message;
      setPicUploadMsg(`❌ Error: ${msg}`);
      addToast(`❌ ${msg}`, 'error');
    }
  };

  const handleUploadItemRegistration = async (e) => {
    const files = Array.from(e.target.files || []); if (!files.length) return;
    e.target.value = '';
    const fd = new FormData(); files.forEach(file => fd.append('file', file));
    const label = files.length > 1 ? `Item Registration (${files.length} files)` : 'Item Registration';
    setUploadProgress({ label, pct: 0 });
    try {
      const res = await api.post('/api/upload/item-registration', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        onUploadProgress: (ev) => setUploadProgress({ label, pct: Math.round(ev.loaded*100/(ev.total||ev.loaded)) })
      });
      setUploadProgress(null);
      addToast(`✅ ${res.data.message || 'Item Registration uploaded successfully'}`, 'success');
      clearDashboardSummaryCache();
      clearStatsCache();
      setItemRegPage(1);
      if (activePage === 'dashboard') fetchDashboard();
      if (activePage === 'item-registration') fetchItemRegistration(1, itemRegPerPage, itemRegAppliedSearch, itemRegFilters);
    } catch (err) {
      setUploadProgress(null);
      addToast(`❌ Failed to upload Item Registration: ${err?.response?.data?.error || err.message}`, 'error');
    }
  };

  const handleBatchUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    e.target.value = '';
    const fd = new FormData(); fd.append('file', file);
    setUploadProgress({ label: 'Batch Update', pct: 0 });
    try {
      const res = await api.post('/api/data/so/batch-upload', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        onUploadProgress: (ev) => setUploadProgress({ label: 'Batch Update', pct: Math.round(ev.loaded*100/(ev.total||ev.loaded)) })
      });
      setUploadProgress(null);
      addToast(`✅ Batch update: ${res.data.updated} records updated`, 'success');
      clearDashboardSummaryCache();
      clearStatsCache();
      fetchSOData(soFilters, soPage, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
    } catch (e) {
      setUploadProgress(null);
      addToast(`❌ Failed to batch upload: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  const handleItemRegistrationBatchUpload = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    e.target.value = '';
    const fd = new FormData(); fd.append('file', file);
    setUploadProgress({ label: 'Item Registration Batch', pct: 0 });
    try {
      const res = await api.post('/api/item-registration/batch-upload', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        onUploadProgress: (ev) => setUploadProgress({ label: 'Item Registration Batch', pct: Math.round(ev.loaded*100/(ev.total||ev.loaded)) })
      });
      setUploadProgress(null);
      addToast(`Item Registration batch: ${res.data.updated} records updated${res.data.not_found ? `, ${res.data.not_found} Req. No not found` : ''}`, 'success');
      clearDashboardSummaryCache();
      clearStatsCache();
      fetchItemRegistration(itemRegPage, itemRegPerPage, itemRegAppliedSearch, itemRegFilters);
    } catch (e) {
      setUploadProgress(null);
      addToast(`Failed to upload Item Registration batch: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  // Helper: trigger browser download from a blob: URL, then revoke it.
  // Defined before downloadBlob so the closure reference is always ready.
  const triggerDownload = (objectUrl, filename) => {
    const link = document.createElement('a');
    link.href = objectUrl;
    link.setAttribute('download', filename);
    // Some browsers ignore `download` on blob: URLs without a type hint —
    // adding `type` via the Blob constructor (above) is the primary fix,
    // but we also set `target` and `rel` as a belt-and-suspenders measure.
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    link.remove();
    // Revoke the object URL after a short delay so the download has time
    // to start. Revoking immediately can abort the download on some browsers.
    setTimeout(() => { try { window.URL.revokeObjectURL(objectUrl); } catch {} }, 1000);
  };

  const downloadBlob = async (url, filename, label) => {
    const toastId = Date.now();
    setDownloadToast({ id: toastId, message: `Downloading ${label || filename}...` });
    try {
      const res = await api.get(url, { responseType: 'blob' });
      // res.data is ALREADY a Blob (because responseType: 'blob'). Wrapping
      // it in `new Blob([res.data])` without a type loses the MIME type,
      // which causes some browsers to:
      //   - open the file in a new tab instead of downloading it
      //   - download with a .bin extension or no extension
      //   - silently fail the download on mobile
      // Use res.data directly, and preserve the Content-Type from the
      // response headers so the browser knows it's an xlsx/pdf/csv/etc.
      const contentType = res.headers['content-type'] || 'application/octet-stream';
      const blob = res.data instanceof Blob ? res.data : new Blob([res.data], { type: contentType });
      // If the blob somehow lost its type, recreate it with the correct type.
      let objectUrl;
      if (!blob.type && contentType) {
        objectUrl = window.URL.createObjectURL(new Blob([blob], { type: contentType }));
      } else {
        objectUrl = window.URL.createObjectURL(blob);
      }
      triggerDownload(objectUrl, filename);
      setDownloadToast(null);
      addToast(`✅ File "${filename}" downloaded successfully`, 'success');
    } catch (e) {
      setDownloadToast(null);
      // When the backend returns an error (e.g. 500), the response body is
      // JSON but received as a Blob because of responseType: 'blob'. Try to
      // read the actual error message from the blob so the user sees what
      // went wrong instead of a generic "Failed to download file".
      let errMsg = e?.message || 'Unknown error';
      try {
        const errBlob = e?.response?.data;
        if (errBlob instanceof Blob) {
          const text = await errBlob.text();
          try {
            const j = JSON.parse(text);
            errMsg = j.error || j.message || text;
          } catch {
            errMsg = text.slice(0, 200);
          }
        } else if (typeof errBlob === 'string') {
          errMsg = errBlob;
        } else if (errBlob && typeof errBlob === 'object') {
          errMsg = errBlob.error || errBlob.message || JSON.stringify(errBlob);
        }
      } catch {}
      const status = e?.response?.status ? ` (HTTP ${e.response.status})` : '';
      addToast(`❌ Failed to download ${label || filename}${status}: ${errMsg}`, 'error');
    }
  };

  const downloadMasterPICTemplate = () => {
    downloadBlob('/api/template/master-pic', `Master_PIC_Update_Template_${new Date().toISOString().slice(0,10)}.xlsx`, 'Master PIC Update Template');
  };

  const downloadItemRegistrationTemplate = () => {
    const p = new URLSearchParams();
    Object.entries(dateFilterParams(globalDateFilter)).forEach(([key, value]) => { if (value) p.append(key, value); });
    appendMultiParam(p, 'client', globalClientFilter);
    appendMultiParam(p, 'global_pic', globalPicFilter);
    (itemRegAppliedSearch || []).forEach(v => p.append('req_no', v));
    resolveFilter(itemRegFilters.clients).forEach(v => p.append('item_client', v));
    resolveFilter(itemRegFilters.categories).forEach(v => p.append('category', v));
    resolveFilter(itemRegFilters.pics).forEach(v => p.append('pic', v));
    resolveFilter(itemRegFilters.proc_statuses).forEach(v => p.append('proc_status', v));
    resolveFilter(itemRegFilters.mfr_names).forEach(v => p.append('mfr_name', v));
    if (itemRegPicHighlight) p.append('kpi_pic', itemRegPicHighlight);
    downloadBlob(`/api/item-registration/template?${p}`, `Template_ItemRegistration_BatchUpload_${new Date().toISOString().slice(0,10)}.xlsx`, 'Item Registration Batch Upload Template');
  };

  const cleanupImportDuplicates = async () => {
    if (!window.confirm('Bersihkan baris duplikat di Import? Baris dengan identitas bisnis yang sama (PO YUPI + Item Yupi, atau PO Sementara + Item Yupi) akan digabung menjadi satu, baris berlebih akan dihapus. Aksi ini tidak bisa dibatalkan.')) return;
    try {
      const res = await api.post('/api/import/cleanup', {});
      addToast(res.data?.message || `${res.data?.deleted || 0} baris duplikat dihapus`, 'success');
      fetchImportData(1, importPerPage, importAppliedSearch, false, importFilters, importReqDlvSort, importYupiPoSort);
      setImportPage(1);
    } catch (e) {
      addToast(`Gagal membersihkan duplikat: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  const downloadImportExcel = () => {
    const p = new URLSearchParams();
    if (importAppliedSearch) p.append('search', importAppliedSearch);
    resolveFilter(importFilters?.yupi_po).forEach(v => p.append('yupi_po', v));
    resolveFilter(importFilters?.vendors).forEach(v => p.append('vendor_name', v));
    resolveFilter(importFilters?.statuses).forEach(v => p.append('status', v));
    resolveFilter(importFilters?.daysLeft).forEach(v => p.append('days_left', v));
    downloadBlob(`/api/import/export?${p}`, `Import_Dashboard_${new Date().toISOString().slice(0,10)}.xlsx`, 'Import Dashboard Excel');
  };

  const downloadItemRegistrationExcel = () => {
    const p = new URLSearchParams();
    Object.entries(dateFilterParams(globalDateFilter)).forEach(([key, value]) => { if (value) p.append(key, value); });
    appendMultiParam(p, 'client', globalClientFilter);
    appendMultiParam(p, 'global_pic', globalPicFilter);
    (itemRegAppliedSearch || []).forEach(v => p.append('req_no', v));
    resolveFilter(itemRegFilters.clients).forEach(v => p.append('item_client', v));
    resolveFilter(itemRegFilters.categories).forEach(v => p.append('category', v));
    resolveFilter(itemRegFilters.pics).forEach(v => p.append('pic', v));
    resolveFilter(itemRegFilters.proc_statuses).forEach(v => p.append('proc_status', v));
    resolveFilter(itemRegFilters.mfr_names).forEach(v => p.append('mfr_name', v));
    if (itemRegPicHighlight) p.append('kpi_pic', itemRegPicHighlight);
    downloadBlob(`/api/export/item-registration?${p}`, `Item_Registration_${new Date().toISOString().slice(0,10)}.xlsx`, 'Item Registration Excel');
  };

  const fetchCompletedData = useCallback(async (year='all', dateFilter=null) => {
    setCompletedLoading(true);
    try {
      const params = new URLSearchParams({ year });
      Object.entries(dateFilterParams(dateFilter)).forEach(([key, value]) => { if (value) params.append(key, value); });
      appendMultiParam(params, 'client', globalClientFilter);
      appendMultiParam(params, 'pic', globalPicFilter);
      const res = await api.get(`/api/completed/summary?${params}`);
      setCompletedData(res.data);
      setCompletedLoaded(true);
    } catch(e) { addToast(`❌ Failed to load completed data: ${e.message}`, 'error'); }
    finally { setCompletedLoading(false); }
  }, []);

  const downloadSOExcel = () => {
    const p = new URLSearchParams();
    resolveFilter(soFilters.op_units).forEach(v => p.append('op_unit', v));
    resolveFilter(soFilters.vendors).forEach(v => p.append('vendor', v));
    resolveFilter(soFilters.manufacturers).forEach(v => p.append('manufacturer', v));
    resolveFilter(soFilters.statuses).forEach(v => p.append('status', v));
    filterValues(soFilters.aging).forEach(v => p.append('aging', v));
    (soSearchNums||[]).forEach(n => p.append('so_item', n));
    appendMultiParam(p, 'client', globalClientFilter);
    appendMultiParam(p, 'global_pic', globalPicFilter);
    if (pendingPicHighlight) p.append('kpi_pic', pendingPicHighlight);
    if (soSortOrder) p.append('sort_order', soSortOrder);
    if (soMarginFilter && soMarginFilter !== 'all') p.append('margin_filter', soMarginFilter);
    Object.entries(dateFilterParams(soDateFilter)).forEach(([key, value]) => { if (value) p.append(key, value); });
    downloadBlob(`/api/export/all-so?${p}`, `SO_List_${new Date().toISOString().slice(0,10)}.xlsx`, 'SO List');
  };
  const downloadApprovalSOExcel = () => {
    const rows = sortedApprovalSOData.map((so) => {
      const poAmount = (Number(so.purchasing_price) || 0) * (Number(so.so_qty) || 0);
      // Margin valid only when purchase is positive (not empty/zero/negative)
      const purchaseValid = poAmount > 0;
      const margin = purchaseValid ? (Number(so.sales_amount) || 0) - poAmount : null;
      return {
        'SO Item': so.so_item || '',
        'Item Name': so.product_name || '',
        'Status': so.so_status || '',
        'Operation Unit': so.operation_unit_name || '',
        'Vendor': so.vendor_name || '',
        'Qty': Number(so.so_qty) || 0,
        'Sales Amount': Number(so.sales_amount) || 0,
        'PO Price': Number(so.purchasing_price) || 0,
        'PO Amount': poAmount,
        'Margin': margin,
        'SO Create Date': so.so_create_date || '',
        'Possible Delivery': so.delivery_possible_date || '',
        'Plan Date': so.delivery_plan_date || '',
        'Remarks': so.remarks || ''
      };
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'SO Approval Status');
    saveAs(new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],
      {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),
      `SO_Approval_Status_${new Date().toISOString().slice(0,10)}.xlsx`);
    addToast('✅ SO Approval Status Excel downloaded successfully', 'success');
  };
  const downloadSOTemplate = () => {
    const p = new URLSearchParams();
    resolveFilter(soFilters.op_units).forEach(v => p.append('op_unit', v));
    resolveFilter(soFilters.vendors).forEach(v => p.append('vendor', v));
    resolveFilter(soFilters.manufacturers).forEach(v => p.append('manufacturer', v));
    resolveFilter(soFilters.statuses).forEach(v => p.append('status', v));
    filterValues(soFilters.aging).forEach(v => p.append('aging', v));
    (soSearchNums||[]).forEach(n => p.append('so_item', n));
    appendMultiParam(p, 'client', globalClientFilter);
    appendMultiParam(p, 'global_pic', globalPicFilter);
    if (pendingPicHighlight) p.append('kpi_pic', pendingPicHighlight);
    if (soMarginFilter && soMarginFilter !== 'all') p.append('margin_filter', soMarginFilter);
    Object.entries(dateFilterParams(soDateFilter)).forEach(([key, value]) => { if (value) p.append(key, value); });
    downloadBlob(`/api/data/so/template?${p}`, `Template_SO_BatchUpload_${new Date().toISOString().slice(0,10)}.xlsx`, 'SO Batch Upload Template');
  };

  const updateSOCell = async (soId, field, value) => {
    setEditingCell(null);
    try {
      const isNumericId = typeof soId === 'number' || /^\d+$/.test(String(soId));
      const endpoint = isNumericId ? `/api/data/so/${soId}` : `/api/data/so/by-item/${encodeURIComponent(soId)}`;
      await api.put(endpoint, { [field]: value });
      const matches = (s) => (isNumericId ? String(s.id) === String(soId) : String(s.so_item) === String(soId));
      setAllSOData(prev => prev.map(s => matches(s) ? { ...s, [field]: value } : s));
      setModal(prev => prev ? { ...prev, data: (prev.data || []).map(s => matches(s) ? { ...s, [field]: value } : s) } : prev);
    } catch (e) { addToast(`❌ Failed to update: ${e.message}`, 'error'); }
  };

  const updateItemRegistrationCell = async (itemId, field, value) => {
    setEditingCell(null);
    try {
      await api.put(`/api/item-registration/${itemId}`, { [field]: value });
      setItemRegData(prev => prev.map(row => row.id === itemId ? { ...row, [field]: value } : row));
    } catch (e) { addToast(`Failed to update Item Registration: ${e.response?.data?.error || e.message}`, 'error'); }
  };

  const applyRFQLocalUpdates = (updates) => {
    const parseNumber = (v) => {
      const s = String(v ?? '').replace(/[^0-9.-]/g, '');
      const n = Number(s);
      return Number.isFinite(n) ? n : null;
    };
    const formatAmount = (n) => Number.isFinite(n) ? Math.round(n).toLocaleString('en-US') : null;
    setRfqData(prev => prev.map(row => {
      const rowUpdates = updates.filter(item => item.row_key === row.row_key);
      if (!rowUpdates.length) return row;
      const next = { ...row };
      rowUpdates.forEach(({ field, value }) => {
        next[field] = value;
        if (field === 'product_id') {
          next.product_id = value;
          next.status = Boolean(value);
          next.check = value ? 'complete' : next.check;
          next.similar_prod_ids = '';
          next.similar_prod_name = '';
          next.similar_spec = '';
          next.similar_mfr_name = '';
          next.similar_odr_unit = '';
          next.similar_score = null;
        }
      });
      if (rowUpdates.some(({ field }) => field === 'unit_price_idr' || field === 'qty')) {
        const qty = parseNumber(next.qty);
        const unitPrice = parseNumber(next.unit_price_idr);
        next.amt_idr = qty != null && unitPrice != null ? formatAmount(qty * unitPrice) : null;
        next.unit_price_missing = unitPrice == null;
      }
      return next;
    }).filter(row => !rfqPicFilter || rfqEditedRowKeys.has(row.row_key) || (row.purchase_pic === rfqPicFilter && row.check === 'open' && row.unit_price_missing && !row.product_id)));
  };

  const updateRFQCell = async (rowKey, field, value, options = {}) => {
    const quiet = Boolean(options.quiet);
    if (!quiet) setEditingCell(null);
    if (field === 'product_id') setRfqSimilarAction(null);
    setRfqEditedRowKeys(prev => { const s = new Set(prev); s.add(rowKey); return s; });
    const previousRows = rfqData;
    applyRFQLocalUpdates([{ row_key: rowKey, field, value }]);
    try {
      const res = await api.put(`/api/rfq/${encodeURIComponent(rowKey)}`, { field, value });
      if (!quiet && res.data?.sheet_sync && res.data.sheet_sync.synced === false && !res.data.sheet_sync.local_only && !rfqDashboardOnlyFields.has(field)) {
        addToast(`RFQ updated locally. Sheet sync not active: ${res.data.sheet_sync.reason}`, 'warning');
      }
      const parseNumber = (v) => {
        const s = String(v ?? '').replace(/[^0-9.-]/g, '');
        const n = Number(s);
        return Number.isFinite(n) ? n : null;
      };
      const formatAmount = (n) => Number.isFinite(n) ? Math.round(n).toLocaleString('en-US') : null;
      if (field === 'unit_price_idr') {
        const current = rfqData.find(row => row.row_key === rowKey);
        const oldMissing = current?.unit_price_missing;
        const newMissing = parseNumber(value) == null;
        const pic = current?.purchase_pic;
        if (pic && current?.check === 'open' && !current?.product_id && oldMissing !== newMissing) {
          setRfqPicKpis(prev => {
            const next = prev.map(row => row.pic === pic ? { ...row, count: Math.max(0, (Number(row.count) || 0) + (newMissing ? 1 : -1)) } : row)
              .filter(row => Number(row.count) > 0);
            if (newMissing && !next.some(row => row.pic === pic)) next.push({ pic, count: 1 });
            return next.sort((a, b) => (Number(b.count) || 0) - (Number(a.count) || 0) || String(a.pic).localeCompare(String(b.pic)));
          });
        }
      }
      if (field === 'product_id') {
        const current = rfqData.find(row => row.row_key === rowKey);
        if (value && current?.check === 'open' && current?.unit_price_missing && !current?.product_id && current?.purchase_pic) {
          setRfqPicKpis(prev => prev.map(kpi => kpi.pic === current.purchase_pic ? { ...kpi, count: Math.max(0, (Number(kpi.count) || 0) - 1) } : kpi).filter(kpi => Number(kpi.count) > 0));
        }
      }
      setRfqData(prev => prev.map(row => {
        if (row.row_key !== rowKey) return row;
        const next = { ...row, [field]: value };
        if (field === 'product_id') {
          next.product_id = value;
          next.status = Boolean(value);
          next.check = value ? 'complete' : next.check;
          next.similar_prod_ids = '';
          next.similar_prod_name = '';
          next.similar_spec = '';
          next.similar_mfr_name = '';
          next.similar_odr_unit = '';
          next.similar_score = null;
        }
        if (field === 'unit_price_idr' || field === 'qty') {
          const qty = parseNumber(next.qty);
          const unitPrice = parseNumber(next.unit_price_idr);
          next.amt_idr = qty != null && unitPrice != null ? formatAmount(qty * unitPrice) : null;
          if (field === 'unit_price_idr') next.unit_price_missing = unitPrice == null;
        }
        return next;
      }).filter(row => !rfqPicFilter || rfqEditedRowKeys.has(row.row_key) || (row.purchase_pic === rfqPicFilter && row.check === 'open' && row.unit_price_missing && !row.product_id)));
      return true;
    } catch (e) {
      // Network/HTTP failure → queue for replay. KEEP the optimistic local
      // edit (don't revert) so user input is preserved, Google-Sheets-style.
      if (typeof navigator === 'undefined' || !navigator.onLine || e.message?.includes('Network') || e.code === 'ERR_NETWORK' || e.response == null) {
        enqueueOfflineUpdate('rfq-cell', { row_key: rowKey, field, value, options });
        if (!quiet) addToast(`Saved offline — will sync when connection returns`, 'warning');
      } else {
        if (!quiet) addToast(`Failed to update RFQ: ${e.response?.data?.error || e.message}`, 'error');
      }
      return false;
    }
  };

  const updateRFQCellsBatch = async (updates) => {
    const cleanUpdates = (updates || []).filter(item => item?.row_key && item?.field);
    if (!cleanUpdates.length) return false;
    setRfqEditedRowKeys(prev => { const s = new Set(prev); cleanUpdates.forEach(u => s.add(u.row_key)); return s; });
    const previousRows = rfqData;
    applyRFQLocalUpdates(cleanUpdates);
    try {
      const res = await api.put('/api/rfq/batch-cells', { updates: cleanUpdates });
      const onlyDashboardOnly = cleanUpdates.every(item => rfqDashboardOnlyFields.has(item.field));
      if (res.data?.sheet_sync && res.data.sheet_sync.synced === false && !res.data.sheet_sync.local_only && !onlyDashboardOnly) {
        addToast(`RFQ batch updated locally. Sheet sync not active: ${res.data.sheet_sync.reason}`, 'warning');
      }
      if (res.data?.skipped?.length) {
        addToast(`RFQ batch skipped ${res.data.skipped.length} cells`, 'warning');
      }
      return true;
    } catch (e) {
      // Offline / network failure → queue. Keep local edits intact.
      if (typeof navigator === 'undefined' || !navigator.onLine || e.message?.includes('Network') || e.code === 'ERR_NETWORK' || e.response == null) {
        enqueueOfflineUpdate('rfq-cells', { updates: cleanUpdates });
        addToast(`Saved offline — will sync when connection returns`, 'warning');
      } else {
        addToast(`Failed to update RFQ batch: ${e.response?.data?.error || e.message}`, 'error');
      }
      return false;
    }
  };

  const openModal = async (title, endpointOrData) => {
    if (Array.isArray(endpointOrData)) { setModal({ title, data: endpointOrData }); return; }
    try {
      const res = await api.get(endpointOrData);
      const detailRows = Array.isArray(res.data) ? res.data.filter(row => row && typeof row === 'object') : [];
      setModal({ title, data: detailRows });
    } catch (e) { addToast(`❌ Failed to load details: ${e.message}`, 'error'); }
  };

  const toggleAgingFilter = (label) => {
    setSoFilters(f => {
      const aging = f.aging.includes(label) ? f.aging.filter(a=>a!==label) : [...f.aging, label];
      return {...f, aging};
    });
  };

  const openPendingDeliveryWithAging = (label) => {
    const aging = soFilters.aging.includes(label) ? soFilters.aging.filter(a => a !== label) : [...soFilters.aging, label];
    const next = { ...soFilters, aging };
    setSoFilters(next);
    setSoPage(1);
    setActivePage('all-so');
    window.scrollTo({ top: 0 });
  };

  const sortedApprovalSOData = [...approvalSOData]
    .filter((so) => {
      if (approvalFilters.op_units.length && !approvalFilters.op_units.includes(so.operation_unit_name)) return false;
      if (approvalFilters.statuses.length && !approvalFilters.statuses.includes(so.so_status)) return false;
      if (approvalSearchNums.length && !approvalSearchNums.includes(so.so_item)) return false;
      if (approvalFilters.aging.length) {
        const age = so.so_create_date ? Math.floor((new Date() - new Date(so.so_create_date)) / (1000 * 60 * 60 * 24)) : null;
        const label = age === null ? 'Unknown' : age <= 30 ? '0-30 days' : age <= 60 ? '31-60 days' : age <= 90 ? '61-90 days' : '>90 days';
        if (!approvalFilters.aging.includes(label)) return false;
      }
      return true;
    })
    .sort((a, b) => {
      const da = a.so_create_date ? new Date(a.so_create_date).getTime() : 0;
      const db = b.so_create_date ? new Date(b.so_create_date).getTime() : 0;
      return da - db;
    });
  const approvalSOTotalAmount = sortedApprovalSOData.reduce((sum, so) => sum + Number(so.sales_amount || 0), 0);
  const approvalTotalPages = Math.max(1, Math.ceil(sortedApprovalSOData.length / approvalPerPage));
  const approvalRows = sortedApprovalSOData.slice((approvalPage-1)*approvalPerPage, approvalPage*approvalPerPage);

  const sortedSOData = [...allSOData].sort((a, b) => {
    const da = a.so_create_date ? new Date(a.so_create_date).getTime() : 0;
    const db = b.so_create_date ? new Date(b.so_create_date).getTime() : 0;
    return soSortOrder === 'newest' ? db - da : da - db;
  });
  const soTotalPages = Math.max(1, Math.ceil(soTotal / soPerPage));
  const itemRegTotalPages = Math.max(1, Math.ceil(itemRegTotal / itemRegPerPage));

  const card  = darkMode ? 'bg-gray-800 border border-gray-700 shadow-sm' : 'bg-white border border-gray-200/80 shadow-[0_8px_24px_rgba(15,23,42,0.05)]';
  const txt   = darkMode ? 'text-white' : 'text-[#1f2937]';
  const txt2  = darkMode ? 'text-gray-400' : 'text-[#55585d]';
  const tblHd = darkMode ? 'bg-gray-800' : 'bg-slate-200';
  const tblDv = darkMode ? 'divide-gray-700' : 'divide-gray-100';
  const trHov = darkMode ? 'hover:bg-gray-700' : 'hover:bg-[#f7f7f5]';
  const kpiValue = darkMode ? 'text-gray-100' : 'text-[#334155]';
  const neutralIcon = darkMode ? 'bg-gray-700 text-gray-200' : 'bg-slate-100 text-slate-600';

  // ══════════════════════════════════════════════════════════════
  // RENDER COMPLETED TRANSACTIONS PAGE
  // ══════════════════════════════════════════════════════════════
  const renderCompleted = () => {
    const d = completedData;
    const CPIE = ['#10B981','#EF4444','#9CA3AF'];
    const fmtM = (v) => v >= 1e9 ? `${(v/1e9).toFixed(1)}B` : v >= 1e6 ? `${(v/1e6).toFixed(1)}M` : v >= 1e3 ? `${(v/1e3).toFixed(0)}K` : String(Math.round(v));
    const mc = (m) => m > 0 ? 'text-green-600' : m < 0 ? 'text-red-600' : 'text-gray-400';
    const mcBg = (m) => m < 0 ? (darkMode?'bg-red-900/20':'bg-red-50') : (darkMode?'bg-gray-700':'bg-gray-50');
    const monthLabel = (m) => { try { const [y,mo] = m.split('-'); return format(new Date(parseInt(y), parseInt(mo)-1, 1), 'MMM yy'); } catch { return m; } };
    const renderPriceChange = (value, pct, trend) => {
      const isUp = trend === 'up';
      const isDown = trend === 'down';
      const color = isUp ? 'text-green-600' : isDown ? 'text-red-600' : txt2;
      const icon = isUp ? <TrendingUp className="w-4 h-4"/> : isDown ? <TrendingDown className="w-4 h-4"/> : <Minus className="w-4 h-4"/>;
      if (value == null) {
        return <div className={`text-right font-semibold ${txt2}`}>No data</div>;
      }
      return (
        <div className={`text-right font-bold ${color}`}>
          <div className="flex items-center justify-end gap-1">
            {icon}
            <span>{fmtCurShort(Math.abs(value))}</span>
          </div>
          <div className="text-[11px] font-semibold">
            {pct == null ? '0%' : `${pct > 0 ? '+' : ''}${pct}%`}
          </div>
        </div>
      );
    };

    if (completedLoading || !completedLoaded) return (
      <div className="flex items-center justify-center h-64">
        <div className="flex flex-col items-center gap-3">
          <div className="w-12 h-12 border-4 border-blue-600 border-t-transparent rounded-full animate-spin"/>
          <p className={`text-sm ${txt2}`}>Loading completed transactions...</p>
          <p className={`text-xs ${txt2} text-center max-w-sm`}>
            Preparing cached USD to IDR values. Existing historical data will be reused, only new non-IDR rows are converted.
          </p>
        </div>
      </div>
    );

    if (!d) return (
      <div className={`flex flex-col items-center justify-center h-64 rounded-2xl ${card}`}>
        <Coins className="w-16 h-16 text-gray-300 mb-4"/>
        <p className={`text-lg font-semibold ${txt}`}>No completed data yet</p>
        <p className={`text-sm ${txt2} mt-1`}>Upload SO data to see completed transactions</p>
      </div>
    );

    const totalCompleted = d.margin_distribution.positive + d.margin_distribution.negative + d.margin_distribution.zero;
    const marginPieData = [
      { name: 'Positive', value: d.margin_distribution.positive },
      { name: 'Negative', value: d.margin_distribution.negative },
      { name: 'Zero / No Data', value: d.margin_distribution.zero },
    ].filter(x => x.value > 0);

    return (
      <div className="space-y-6">

        {/* ── Date Range Filter ───────────────────────────────── */}
        <DateRangeFilter
          darkMode={darkMode} txt={txt} txt2={txt2} card={card}
          value={globalDateFilter}
          label="Filter SO Create Date"
          onFilter={(f) => {
            setGlobalDateFilter(f);
            setCompletedYear('all'); fetchCompletedData('all', f);
          }}
        />

        {/* ── KPI Cards ───────────────────────────────────────── */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          {[
            { label:'Completed Transactions', value: fmtNum(d.total_count),
              icon:<CheckCircle className="w-6 h-6"/>, bg: neutralIcon, color:kpiValue },
            { label:'Total Sales Amount', value: fmtCurShort(d.total_sales), sub: fmtCur(d.total_sales),
              icon:<Wallet className="w-6 h-6"/>, bg: neutralIcon, color:kpiValue },
            { label:'Total Purchase Amount', value: fmtCurShort(d.total_purchase), sub: fmtCur(d.total_purchase),
              icon:<Coins className="w-6 h-6"/>, bg: neutralIcon, color:kpiValue },
            { label:'Total Margin', value: fmtCurShort(d.total_margin), sub: fmtCur(d.total_margin),
              icon: d.total_margin>=0?<TrendingUp className="w-6 h-6"/>:<TrendingDown className="w-6 h-6"/>,
              bg: neutralIcon, color:kpiValue },
          ].map((k,i)=>(
            <div key={i} className={`p-5 rounded-2xl transition-all ${card}`}>
              <div className="flex justify-between items-start">
                <div>
                  <p className={`text-sm font-medium ${txt2}`}>{k.label}</p>
                  <h3 className={`text-2xl font-bold mt-1 ${k.color}`}>{k.value}</h3>
                  {k.sub && <p className={`text-xs mt-0.5 ${txt2}`}>{k.sub}</p>}
                </div>
                <div className={`p-3 ${k.bg} rounded-xl`}>{k.icon}</div>
              </div>
            </div>
          ))}
        </div>

        {/* ── Monthly Trend: Combined Amount + Transaction Count ── */}
        <div className={`p-5 rounded-2xl shadow ${card}`}>
          <h3 className={`text-base font-bold mb-1 ${txt}`}>Monthly Trend — Delivery Completed</h3>
          <p className={`text-xs mb-4 ${txt2}`}>Sales Amount, Purchase Amount (bar) & Transaction Count (line)</p>
          <ResponsiveContainer width="100%" height={320}>
            <ComposedChart data={(d.monthly_trend||[]).map(m => ({
              ...m,
              monthLabel: (() => { try { const [y,mo] = m.month.split('-'); return format(new Date(parseInt(y), parseInt(mo)-1, 1), 'MMM yy'); } catch { return m.month; } })()
            }))} barGap={2}>
              <defs>
                <linearGradient id="cgSales" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#1D4ED8" stopOpacity={0.92}/><stop offset="95%" stopColor="#1D4ED8" stopOpacity={0.58}/>
                </linearGradient>
                <linearGradient id="cgPurchase" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#93C5FD" stopOpacity={0.9}/><stop offset="95%" stopColor="#BFDBFE" stopOpacity={0.62}/>
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke={darkMode?'#374151':'#F3F4F6'}/>
              <XAxis dataKey="monthLabel" stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={10}/>
              <YAxis yAxisId="amt" stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={10} tickFormatter={fmtM}/>
              <YAxis yAxisId="cnt" orientation="right" stroke="#64748B" fontSize={10}/>
              <Tooltip
                formatter={(v, n) => n === 'Transactions' ? [fmtNum(v), n] : [fmtCur(v), n]}
                contentStyle={{background:darkMode?'#1F2937':'#fff',border:'none',borderRadius:8,fontSize:12}}
                labelStyle={{color:darkMode?'#F3F4F6':'#111827'}}
                itemStyle={{color:darkMode?'#F3F4F6':'#111827'}}/>
              <Legend wrapperStyle={{fontSize:12,color:darkMode?'#F3F4F6':'#111827'}} iconType="rect"/>
              <Bar yAxisId="amt" dataKey="sales_amount" name="Sales Amount" fill="url(#cgSales)" radius={[4,4,0,0]} isAnimationActive={false}/>
              <Bar yAxisId="amt" dataKey="purchase_amount" name="Purchase Amount" fill="url(#cgPurchase)" radius={[4,4,0,0]} isAnimationActive={false}/>
              <Line yAxisId="cnt" type="monotone" dataKey="count" name="Transactions" stroke="#64748B" strokeWidth={3} dot={{r:3,fill:'#64748B'}} activeDot={{r:5, fill:'#64748B'}} z={10} isAnimationActive={false}/>
            </ComposedChart>
          </ResponsiveContainer>
        </div>

        {/* ── Top 5 Vendors  +  Margin Pie ────────────────────── */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">

          {/* Top 5 Vendors */}
          <div className={`p-5 rounded-2xl shadow ${card}`}>
            <h3 className={`text-base font-bold mb-4 ${txt} flex items-center gap-2`}>
              <Award className="w-5 h-5 text-blue-500"/> Top 5 Vendors — Completed Transactions
            </h3>
            <div className="space-y-4">
              {(d.top_vendors||[]).map((v,i)=>{
                const maxAmt = d.top_vendors[0]?.sales_amount || 1;
                const pct = Math.round(v.sales_amount / maxAmt * 100);
                const rankColors = ['bg-blue-600 text-white','bg-slate-200 text-slate-700','bg-teal-100 text-teal-700','bg-blue-100 text-blue-700','bg-slate-100 text-slate-600'];
                return (
                  <div key={i}>
                    <div className="flex items-center justify-between mb-1">
                      <div className="flex items-center gap-2 min-w-0">
                        <span className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold flex-shrink-0 ${rankColors[i]||'bg-gray-100 text-gray-700'}`}>{i+1}</span>
                        <span className={`text-sm font-semibold truncate ${txt}`} title={v.vendor}>{v.vendor}</span>
                      </div>
                      <span className="text-xs font-bold text-blue-600 ml-2 flex-shrink-0">{fmtCurShort(v.sales_amount)}</span>
                    </div>
                    <div className={`w-full h-2 rounded-full ${darkMode?'bg-gray-700':'bg-gray-100'}`}>
                      <div className="h-2 rounded-full bg-gradient-to-r from-blue-600 to-blue-400" style={{width:`${pct}%`}}/>
                    </div>
                    <div className="flex justify-between text-xs mt-0.5">
                      <span className={txt2}>{fmtNum(v.count)} transactions · Purchase {fmtCurShort(v.purchase_amount)}</span>
                      <span className={mc(v.margin)}>Margin: {fmtCurShort(v.margin)}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Margin Pie */}
          <div className={`p-5 rounded-2xl shadow ${card}`}>
            <h3 className={`text-base font-bold mb-3 ${txt} flex items-center gap-2`}>
              <BarChart3 className="w-5 h-5 text-green-500"/> Margin Distribution — PO Count
            </h3>
            <div className="flex flex-col gap-4">
              {/* Top boxes (Positive / Negative / Zero) */}
              <div className="grid grid-cols-3 gap-2">
                {[
                  { cat:'positive', label:'Positive', color:'text-green-600', bg: darkMode?'bg-green-900/30 border-green-700':'bg-green-50 border-green-200', count: d.margin_distribution.positive },
                  { cat:'negative', label:'Negative', color:'text-red-600', bg: darkMode?'bg-red-900/30 border-red-700':'bg-red-50 border-red-200', count: d.margin_distribution.negative },
                  { cat:'zero', label:'Zero / No Data', color:'text-gray-500', bg: darkMode?'bg-gray-700 border-gray-600':'bg-gray-50 border-gray-200', count: d.margin_distribution.zero },
                ].map(({cat, label, color, bg, count}) => (
                  <button key={cat} onClick={async () => {
                    try {
                      const params = new URLSearchParams({category: cat});
                      Object.entries(dateFilterParams(completedDateFilter)).forEach(([key, value]) => { if (value) params.append(key, value); });
                      appendMultiParam(params, 'client', globalClientFilter);
                      appendMultiParam(params, 'pic', globalPicFilter);
                      const res = await api.get(`/api/completed/margin-detail?${params}`);
                      setMarginDetailModal({ category: label, data: Array.isArray(res.data) ? res.data : [] });
                    } catch(e) { addToast(`Failed to load detail: ${e.message}`, 'error'); }
                  }}
                    className={`text-left p-3 rounded-xl border cursor-pointer transition-all hover:shadow-md ${bg}`}>
                    <p className={`text-xs font-bold ${color} mb-1`}>{label}</p>
                    <p className={`text-lg font-bold ${color}`}>{fmtNum(count)} <span className="text-xs font-normal">PO</span></p>
                    <p className={`text-xs ${txt2} mt-0.5`}>{totalCompleted ? Math.round(count/totalCompleted*100) : 0}% of total</p>
                    <p className={`text-xs text-blue-500 font-semibold mt-1`}>Click for details →</p>
                  </button>
                ))}
              </div>
              {/* Pie chart */}
              <div className="w-full">
                <ResponsiveContainer width="100%" height={220}>
                  <PieChart>
                    <Pie data={marginPieData} cx="50%" cy="50%" innerRadius={55} outerRadius={88} isAnimationActive={false}
                      dataKey="value" labelLine={false} label={renderPctLabel}>
                      {marginPieData.map((_,i)=><Cell key={i} fill={CPIE[i]}/>)}
                    </Pie>
                    <Tooltip formatter={(v,n)=>[`${fmtNum(v)} PO (${totalCompleted ? Math.round(v/totalCompleted*100) : 0}%)`, n]}
                      contentStyle={{background:darkMode?'#1F2937':'#fff',border:'none',borderRadius:8,fontSize:12}}/>
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        </div>

        {/* ── Price Tracking ───────────────────────────────────── */}
        <div className={`p-5 rounded-2xl shadow ${card}`}>
          <div className="flex flex-wrap items-start justify-between gap-2 mb-4">
            <div className="min-w-0 pr-3">
              <h3 className={`text-base font-bold ${txt} flex items-center gap-2`}>
                <Package className="w-5 h-5 text-slate-600"/> Price Tracking by Month — Top 20 Items
              </h3>
              <p className={`text-xs mt-1 ${txt2}`}>Monthly unit price trend from completed delivery transactions.</p>
            </div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className={tblHd}>
                  {['#','Item / Product','Specification','Product ID','Order Freq','Total Purchase Amount','Previous Price','Last Price','Trend','MoM Price Change','YoY Price Change'].map(h=>(
                    <th key={h} className={`px-3 py-2 text-center font-bold ${darkMode?'text-blue-300':'text-blue-700'}`}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className={`divide-y ${tblDv}`}>
                {(d.price_tracking||d.top_items||[]).map((item,i)=>{
                  const isUp = item.trend === 'up';
                  const isDown = item.trend === 'down';
                  const trendData = (item.monthly_prices||[]).map(p => ({ ...p, monthLabel: monthLabel(p.month) }));
                  return (
                    <tr key={i} className={trHov}>
                      <td className={`px-3 py-2 font-bold ${txt2}`}>{i+1}</td>
                      <td className={`px-3 py-2 font-medium ${txt} max-w-xs truncate`} title={item.item}>{item.item}</td>
                      <td className={`px-3 py-2 ${txt2} max-w-xs truncate`} title={item.specification||''}>{item.specification||'—'}</td>
                      <td className={`px-3 py-2 ${txt2} font-mono whitespace-nowrap`}>{item.product_id||'—'}</td>
                      <td className={`px-3 py-2 ${txt2}`}>{fmtNum(item.count)}</td>
                      <td className={`px-3 py-2 font-semibold whitespace-nowrap ${kpiValue}`}>{fmtCurShort(item.purchase_amount)}</td>
                      <td className={`px-3 py-2 ${txt2} whitespace-nowrap`}>{item.previous_price ? fmtCurShort(item.previous_price) : '—'}</td>
                      <td className="px-3 py-2 text-blue-600 font-semibold whitespace-nowrap">{item.current_price ? fmtCurShort(item.current_price) : '—'}</td>
                      <td className="px-3 py-2 min-w-[130px]">
                        <div className="w-28 h-10 mx-auto">
                          {trendData.length > 1 ? (
                            <ResponsiveContainer width="100%" height="100%">
                              <LineChart data={trendData}>
                                <XAxis dataKey="monthLabel" hide />
                                <Tooltip
                                  formatter={(v)=>[fmtCurShort(v), 'Price']}
                                  contentStyle={{background:darkMode?'#1F2937':'#fff',border:'none',borderRadius:8,fontSize:12}}
                                />
                                <Line type="monotone" dataKey="price" stroke={isDown?'#EF4444':'#10B981'} strokeWidth={2.5} dot={false} isAnimationActive={false}/>
                              </LineChart>
                            </ResponsiveContainer>
                          ) : <div className={`h-full flex items-center justify-center ${txt2}`}>—</div>}
                        </div>
                      </td>
                      <td className="px-3 py-2 min-w-[120px]">{renderPriceChange(item.change_value, item.change_pct, item.trend)}</td>
                      <td className="px-3 py-2 min-w-[120px]">{renderPriceChange(item.yoy_change_value, item.yoy_change_pct, item.yoy_trend)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        {/* ── Worst Margin Vendors  +  Worst 10 Transactions ── */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">

          {/* Vendors ranked by most negative margin */}
          <div className={`p-5 rounded-2xl shadow ${card}`}>
            <h3 className={`text-base font-bold mb-4 ${txt} flex items-center gap-2`}>
              <TrendingDown className="w-5 h-5 text-red-500"/> Vendors — Largest Negative Margin
            </h3>
            {(d.worst_margin_vendors||[]).length === 0
              ? <div className="flex flex-col items-center justify-center py-10">
                  <CheckCircle className="w-10 h-10 text-green-400 mb-2"/>
                  <p className={`text-sm font-semibold ${txt2}`}>No vendors with negative margin 🎉</p>
                </div>
              : <div className="space-y-2 overflow-y-auto max-h-80">
                  {d.worst_margin_vendors.map((v,i)=>{
                    const maxAbs = Math.abs(d.worst_margin_vendors[0]?.margin || 1);
                    const barPct = Math.round(Math.abs(v.margin) / maxAbs * 100);
                    return (
                      <button type="button" key={i} onClick={() => openNegativeVendorDetail(v.vendor)} className={`block w-full text-left p-3 rounded-xl ${mcBg(v.margin)} transition-all hover:shadow-md hover:-translate-y-0.5`}>
                        <div className="flex items-center justify-between mb-1">
                          <div className="flex items-center gap-2 min-w-0">
                            <span className={`text-xs font-bold flex-shrink-0 ${darkMode?'text-gray-400':'text-gray-500'}`}>#{i+1}</span>
                            <span className={`text-sm font-semibold truncate ${txt}`} title={v.vendor}>{v.vendor}</span>
                          </div>
                          <span className="text-sm font-bold text-red-600 ml-2 flex-shrink-0">{fmtCurShort(v.margin)}</span>
                        </div>
                        <div className={`w-full h-1.5 rounded-full ${darkMode?'bg-gray-600':'bg-red-100'} mb-1`}>
                          <div className="h-1.5 rounded-full bg-red-500" style={{width:`${barPct}%`}}/>
                        </div>
                        <div className="flex justify-between text-xs">
                          <span className={txt2}>{fmtNum(v.count)} negative txn{v.count!==1?'s':''}</span>
                          <span className={txt2}>Sales: {fmtCurShort(v.total_sales||0)} · PO: {fmtCurShort(v.total_purchase||0)}</span>
                        </div>
                      </button>
                    );
                  })}
                </div>
            }
          </div>

          {/* Top 30 worst margin transactions (scrollable, table size unchanged) */}
          <div className={`p-5 rounded-2xl shadow ${card}`}>
            <h3 className={`text-base font-bold mb-4 ${txt} flex items-center gap-2`}>
              <TrendingDown className="w-5 h-5 text-red-500"/> Top 30 Transactions — Largest Negative Margin
            </h3>
            {(d.worst_margin_transactions||[]).length === 0
              ? <div className="flex flex-col items-center justify-center py-10">
                  <CheckCircle className="w-10 h-10 text-green-400 mb-2"/>
                  <p className={`text-sm font-semibold ${txt2}`}>No negative margin transactions 🎉</p>
                </div>
              : <div className="overflow-x-auto overflow-y-auto max-h-80">
                  <table className="w-full text-xs">
                    <thead className={`sticky top-0 z-10 ${darkMode?'bg-gray-800':'bg-white'}`}>
                      <tr className={tblHd}>
                        {['#','SO Item','Product','Vendor','Sales','Purchase','Margin','%','Txns','Last Date'].map(h=>(
                          <th key={h} className={`px-2 py-2 text-center font-bold ${darkMode?'text-blue-300':'text-blue-700'}`}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className={`divide-y ${tblDv}`}>
                      {d.worst_margin_transactions.map((t,i)=>(
                        <tr key={i} className={`${trHov} ${i===0?darkMode?'bg-red-900/20':'bg-red-50':''}`}>
                          <td className={`px-2 py-2 font-bold text-red-600`}>{i+1}</td>
                          <td className="px-2 py-2">
                            <p className="font-semibold text-blue-600 whitespace-nowrap">{t.so_item||'-'}</p>
                          </td>
                          <td className="px-2 py-2">
                            <p className={`truncate max-w-[120px] ${txt}`} title={t.product}>{t.product}</p>
                          </td>
                          <td className={`px-2 py-2 ${txt} truncate max-w-[90px]`} title={t.vendor}>{t.vendor}</td>
                          <td className="px-2 py-2 text-blue-600 whitespace-nowrap">{fmtCurShort(t.sales_amount)}</td>
                          <td className={`px-2 py-2 whitespace-nowrap ${kpiValue}`}>{fmtCurShort(t.purchase_amount)}</td>
                          <td className="px-2 py-2 font-bold text-red-600 whitespace-nowrap">{fmtCurShort(t.margin)}</td>
                          <td className="px-2 py-2 font-semibold text-red-500 whitespace-nowrap">{t.margin_pct != null ? `${t.margin_pct}%` : '—'}</td>
                          <td className={`px-2 py-2 text-center ${txt2}`}>{fmtNum(t.count||1)}</td>
                          <td className={`px-2 py-2 ${txt2} whitespace-nowrap`}>{t.date ? fmtDate(t.date) : '-'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
            }
          </div>
        </div>

      </div>
    );
  };

  const fmtDateRange = (range) => {
    if (!range?.min) return 'No data available';
    return `${fmtDate(range.min)} — ${fmtDate(range.max)}`;
  };

  // ══════════════════════════════════════════════════════════════
  // RENDER DASHBOARD
  // ══════════════════════════════════════════════════════════════
  const globalSlicerPages = new Set(['dashboard', 'all-so', 'item-registration']);
  const renderGlobalSlicer = () => {
    if (!globalSlicerPages.has(activePage)) return null;
    const showDateFilter = true;
    const slicerClientOptions = activePage === 'item-registration' ? (itemRegOptions.clients || []) : (dashboardFilterOptions.clients || []);
    const slicerPicOptions = activePage === 'item-registration' ? (itemRegOptions.pics || []) : (dashboardFilterOptions.pics || []);
    return (
      <div className={showDateFilter ? "mb-5 flex flex-col gap-3 xl:flex-row xl:items-start" : "mb-5 flex justify-end"}>
        {showDateFilter && (
          <DateRangeFilter
            darkMode={darkMode}
            txt={txt}
            txt2={txt2}
            card={card}
            value={globalDateFilter}
            label="Filter SO Create Date"
            compact
            onFilter={(f) => { setGlobalDateFilter(f); if (activePage === 'item-registration') setItemRegPage(1); }}
          />
        )}
        <div className={`flex min-h-[64px] w-full flex-col gap-3 px-5 py-3 rounded-xl sm:flex-row sm:flex-wrap sm:items-end ${card} shadow xl:flex-nowrap xl:w-auto xl:shrink-0 xl:min-w-[440px] xl:max-w-[560px]`}>
          <div className="min-w-0 flex-1">
            <label className={`mb-1 block text-xs font-semibold ${txt}`}>Client Nm.</label>
            <MultiSelect label="Client Nm." options={slicerClientOptions} selected={globalClientFilter} onChange={setGlobalClientFilter} darkMode={darkMode} txt2={txt2} hideLabel />
          </div>
          <div className="min-w-0 flex-1">
            <label className={`mb-1 block text-xs font-semibold ${txt}`}>PIC Name</label>
            <MultiSelect label="PIC Name" options={slicerPicOptions} selected={globalPicFilter} onChange={setGlobalPicFilter} darkMode={darkMode} txt2={txt2} hideLabel />
          </div>
          <button
            type="button"
            onClick={() => { setGlobalDateFilter({ mode: 'all' }); setGlobalClientFilter([]); setGlobalPicFilter([]); }}
            className={`h-10 w-full shrink-0 px-4 rounded-lg text-sm font-medium shadow-sm flex items-center justify-center whitespace-nowrap sm:w-20 ${darkMode ? 'bg-gray-500 text-gray-100 hover:bg-gray-400' : 'bg-gray-400 text-white hover:bg-gray-500'}`}
          >
            Clear
          </button>
        </div>
      </div>
    );
  };

  const renderDashboardOverview = () => {
    const d = completedData || {};
    const marginD = dashboardMarginData || {};
    const summaryLoading = completedLoading || !completedLoaded;
    
    // Create month buckets from the global SO Create Date filter. KPI remains
    // all-data when filter is All; monthly chart/table defaults to current year.
    const currentYear = new Date().getFullYear();
    const monthlyDataMap = {};
    (marginD.monthly_trend || []).forEach(m => {
      if (m.month) monthlyDataMap[m.month] = m;
    });

    const parseLocalISO = (iso) => {
      const [year, month, day] = String(iso || '').split('-').map(Number);
      return Number.isFinite(year) && Number.isFinite(month) && Number.isFinite(day)
        ? new Date(year, month - 1, day)
        : null;
    };
    const monthWindow = (() => {
      const bounds = dateFilterParams(globalDateFilter);
      const fallbackStart = new Date(currentYear, 0, 1);
      const fallbackEnd = new Date(currentYear, 11, 1);
      let start = parseLocalISO(bounds.date_from) || parseLocalISO(bounds.date_to) || fallbackStart;
      let end = parseLocalISO(bounds.date_to) || parseLocalISO(bounds.date_from) || fallbackEnd;
      if (start > end) [start, end] = [end, start];
      start = new Date(start.getFullYear(), start.getMonth(), 1);
      end = new Date(end.getFullYear(), end.getMonth(), 1);
      const months = [];
      const cursor = new Date(start);
      while (cursor <= end && months.length < 36) {
        months.push({ year: cursor.getFullYear(), monthIndex: cursor.getMonth() });
        cursor.setMonth(cursor.getMonth() + 1);
      }
      const sameYear = start.getFullYear() === end.getFullYear();
      const label = sameYear
        ? String(start.getFullYear())
        : `${format(start, 'MMM yyyy')} - ${format(end, 'MMM yyyy')}`;
      return { months, sameYear, label };
    })();
    
    const monthlyCompleted = monthWindow.months.map(({ year, monthIndex }) => {
      const monthKey = `${year}-${String(monthIndex + 1).padStart(2, '0')}`;
      const monthName = format(new Date(year, monthIndex, 1), monthWindow.sameYear ? 'MMMM' : 'MMM yy');
      const existing = monthlyDataMap[monthKey];
      
      if (existing) {
        const margin = (existing.sales_amount || 0) - (existing.purchase_amount || 0);
        return {
          ...existing,
          monthLabel: monthName,
          margin,
          margin_pct: existing.sales_amount ? margin / existing.sales_amount * 100 : null
        };
      }
      
      return {
        month: monthKey,
        monthLabel: monthName,
        sales_amount: null,
        purchase_amount: null,
        margin: null,
        margin_pct: null,
        count: null
      };
    });
    const marginPct = d.total_sales ? ((d.total_margin || 0) / d.total_sales * 100) : null;
    const fmtM = (v) => v >= 1e9 ? `${(v/1e9).toFixed(1)}B` : v >= 1e6 ? `${(v/1e6).toFixed(1)}M` : v >= 1e3 ? `${(v/1e3).toFixed(0)}K` : String(Math.round(v || 0));
    const purchaseYoyYears = (marginD.purchase_yoy_years && marginD.purchase_yoy_years.length)
      ? marginD.purchase_yoy_years
      : [currentYear - 1, currentYear - 2];
    const purchaseYoyData = (marginD.purchase_yoy_trend && marginD.purchase_yoy_trend.length)
      ? marginD.purchase_yoy_trend
      : Array.from({ length: 12 }, (_, i) => ({
          month: i + 1,
          month_label: format(new Date(currentYear, i, 1), 'MMMM'),
          ...purchaseYoyYears.reduce((acc, year) => ({ ...acc, [`purchase_${year}`]: 0 }), {})
        }));
    const activePurchaseYoyYears = purchaseYoyYears.filter(year =>
      purchaseYoyData.some(row => Number(row[`purchase_${year}`] || 0) > 0)
    );
    const completedTrendData = monthlyCompleted.map((m) => {
      const monthNum = Number(String(m.month || '').slice(5, 7));
      const yoyRow = purchaseYoyData.find(row => Number(row.month) === monthNum) || {};
      const yoyValues = purchaseYoyYears.reduce((acc, year) => {
        const value = Number(yoyRow[`purchase_${year}`] || 0);
        acc[`purchase_${year}`] = value > 0 ? value : null;
        return acc;
      }, {});
      return { ...m, ...yoyValues };
    });
    const purchaseYoyColors = ['#14B8A6', '#94A3B8'];
    const completedTrendLegendPayload = [
      { value: 'PO Amount', type: 'rect', color: '#93C5FD', dataKey: 'purchase_amount' },
      { value: 'Sales Amount', type: 'rect', color: '#2563EB', dataKey: 'sales_amount' },
      ...activePurchaseYoyYears.map((year, i) => ({
        value: `PO ${year}`,
        type: 'line',
        color: purchaseYoyColors[i % purchaseYoyColors.length],
        dataKey: `purchase_${year}`,
      })),
    ];
    const sumRows = (rows, key) => (rows || []).reduce((sum, row) => sum + (Number(row?.[key]) || 0), 0);
    const vendorRows = vendorPurchaseType === 'local'
      ? (d.top_vendors_local || [])
      : vendorPurchaseType === 'import'
      ? (d.top_vendors_import || [])
      : (d.top_vendors || []);
    const barList = (rows, labelKey, valueKey, label, color) => (
      <ResponsiveContainer width="100%" height={260}>
        <BarChart data={(rows || []).map(r => ({...r, shortLabel: String(r[labelKey] || '-').slice(0, 34)}))} layout="vertical" margin={{top: 8, right: 18, left: 24, bottom: 8}}>
          <CartesianGrid strokeDasharray="3 3" horizontal={true} stroke={darkMode?'#374151':'#E5E7EB'}/>
          <XAxis type="number" stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={12} tickFormatter={fmtM}/>
          <YAxis type="category" dataKey="shortLabel" width={220} stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={12} tick={{fontSize: 12, textAnchor: 'end'}} tickMargin={8}/>
          <Tooltip formatter={(v)=>[fmtCur(v), label]} labelFormatter={(_, payload)=>payload?.[0]?.payload?.[labelKey] || '-'} contentStyle={{background:darkMode?'#1F2937':'#fff',border:'none',borderRadius:8,fontSize:12}}/>
          <Bar dataKey={valueKey} name={label} fill={color} radius={[0,6,6,0]} isAnimationActive={false}/>
        </BarChart>
      </ResponsiveContainer>
    );

    return (
      <div className="space-y-5">
        <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-5 gap-4">
          {[
            { label:'Total PO', value: summaryLoading ? '...' : fmtNum(d.total_count || 0), sub: summaryLoading ? 'Loading completed records' : 'Delivery Complete records', icon:<FileText className="w-5 h-5"/> },
            { label:'PO Amount', value: summaryLoading ? '...' : fmtCurShort(d.total_purchase || 0), sub: summaryLoading ? 'Loading completed amount' : fmtCur(d.total_purchase || 0), icon:<Coins className="w-5 h-5"/> },
            { label:'Sales Amount', value: summaryLoading ? '...' : fmtCurShort(d.total_sales), sub: summaryLoading ? 'Loading completed amount' : fmtCur(d.total_sales), icon:<Wallet className="w-5 h-5"/> },
            { label:'Margin', value: summaryLoading ? '...' : fmtCurShort(d.total_margin), sub: summaryLoading ? 'Loading margin' : (marginPct == null ? 'Avg margin -' : `Avg margin ${marginPct.toFixed(1)}%`), icon:<TrendingUp className="w-5 h-5"/> },
            { label:'Total Pending Delivery', value: fmtNum(summaryPendingTotal ?? stats?.total_so_count), sub: 'Pending delivery records', icon:<Clock className="w-5 h-5"/>, goPending:true },
          ].map((k,i)=>{
            const Wrapper = k.goPending ? 'button' : 'div';
            return <Wrapper key={i} type={k.goPending ? 'button' : undefined} onClick={k.goPending ? () => { setActivePage('all-so'); setSoPage(1); window.scrollTo({top:0}); } : undefined} className={`p-5 rounded-2xl text-left ${card} ${k.goPending ? 'cursor-pointer hover:border-blue-300' : ''}`}><div className="flex items-start justify-between gap-3"><div className="min-w-0"><p className={`text-sm font-medium ${txt2}`}>{k.label}</p><h3 className={`text-2xl font-bold mt-1 ${kpiValue}`}>{k.value}</h3><p className={`text-xs mt-1 ${txt2}`}>{k.sub}</p></div><div className={`p-2.5 rounded-xl ${neutralIcon}`}>{k.icon}</div></div></Wrapper>;
          })}
        </div>
        {summaryLoading && (
          <div className={`flex items-center gap-2 rounded-xl px-4 py-3 text-sm font-semibold ${card} ${txt2}`}>
            <Loader2 className="h-4 w-4 animate-spin text-blue-600" />
            Loading delivery completed summary...
          </div>
        )}
        <div className={`p-5 rounded-2xl ${card}`}>
          <h3 className={`text-base font-bold mb-1 ${txt}`}>Monthly Trend Delivery Complete</h3>
          <p className={`text-xs mb-4 ${txt2}`}>{monthWindow.label} sales and purchase amount, with comparison purchase amount lines.</p>
          <ResponsiveContainer width="100%" height={320}>
            <ComposedChart data={completedTrendData} barGap={2} margin={{ top: 8, right: 20, left: 0, bottom: 4 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={darkMode?'#374151':'#E5E7EB'}/>
              <XAxis dataKey="monthLabel" stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={10}/>
              <YAxis stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={10} tickFormatter={fmtM}/>
              <Tooltip formatter={(v,n)=>[fmtCur(v), n]} contentStyle={{background:darkMode?'#1F2937':'#fff',border:'none',borderRadius:8,fontSize:12}}/>
              <Legend wrapperStyle={{fontSize:12}} payload={completedTrendLegendPayload}/>
              <Bar dataKey="purchase_amount" name="PO Amount" fill="#93C5FD" radius={[4,4,0,0]} isAnimationActive={false}/>
              <Bar dataKey="sales_amount" name="Sales Amount" fill="#2563EB" radius={[4,4,0,0]} isAnimationActive={false}/>
              {activePurchaseYoyYears.map((year, i) => (
                <Line
                  key={year}
                  type="monotone"
                  dataKey={`purchase_${year}`}
                  name={`PO ${year}`}
                  stroke={purchaseYoyColors[i % purchaseYoyColors.length]}
                  strokeOpacity={0.34}
                  strokeWidth={1.5}
                  dot={{ r: 2, fill: purchaseYoyColors[i % purchaseYoyColors.length], opacity: 0.38 }}
                  activeDot={{ r: 4, opacity: 0.55 }}
                  isAnimationActive={false}
                />
              ))}
            </ComposedChart>
          </ResponsiveContainer>
        </div>
        <div className={`p-3 rounded-2xl ${card}`}>
          <div className="mb-1.5 flex flex-wrap items-center justify-between gap-2">
            <h3 className={`text-sm font-bold ${txt}`}>Gross Margin by Month</h3>
            <span className={`text-xs font-semibold ${txt2}`}>{monthWindow.label}</span>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-[11px] leading-tight">
              <thead className={tblHd}>
                <tr>
                  <th className={`px-2 py-0.5 text-left font-bold ${txt2}`}>Metric</th>
                  {monthlyCompleted.map((m,i)=><th key={i} className={`px-1.5 py-0.5 text-center font-bold ${txt2}`}>{m.monthLabel}</th>)}
                </tr>
              </thead>
              <tbody className={`divide-y ${tblDv}`}>
                <tr className={trHov}><td className={`px-2 py-0.5 font-semibold ${txt}`}>Sales Amount</td>{monthlyCompleted.map((m,i)=><td key={i} className={`px-1.5 py-0.5 text-center ${kpiValue}`}>{m.sales_amount != null ? fmtCurShort(m.sales_amount) : '-'}</td>)}</tr>
                <tr className={trHov}><td className={`px-2 py-0.5 font-semibold ${txt}`}>PO Amount</td>{monthlyCompleted.map((m,i)=><td key={i} className={`px-1.5 py-0.5 text-center ${kpiValue}`}>{m.purchase_amount != null ? fmtCurShort(m.purchase_amount) : '-'}</td>)}</tr>
                <tr className={trHov}><td className={`px-2 py-0.5 font-semibold ${txt}`}>Gross Margin</td>{monthlyCompleted.map((m,i)=><td key={i} className={`px-1.5 py-0.5 text-center font-semibold ${m.margin != null ? (m.margin < 0 ? 'text-red-600' : 'text-green-600') : txt2}`}>{m.margin != null ? fmtCurShort(m.margin) : '-'}</td>)}</tr>
                <tr className={trHov}><td className={`px-2 py-0.5 font-semibold ${txt}`}>Gross Margin %</td>{monthlyCompleted.map((m,i)=><td key={i} className={`px-1.5 py-0.5 text-center font-semibold ${m.margin_pct != null ? (m.margin < 0 ? 'text-red-600' : 'text-green-600') : txt2}`}>{m.margin_pct != null ? `${m.margin_pct.toFixed(1)}%` : '-'}</td>)}</tr>
              </tbody>
            </table>
          </div>
        </div>
        <div className="grid grid-cols-1 xl:grid-cols-2 gap-5">
          <div className={`p-5 rounded-2xl ${card}`}>
            <div className="mb-4 flex flex-wrap items-start justify-between gap-3">
              <div>
                <h3 className={`text-base font-bold ${txt}`}>Top 5 Vendor PO Amount</h3>
                <p className={`text-xs ${txt2}`}>Total: {fmtCurShort(sumRows(vendorRows, 'purchase_amount'))}</p>
              </div>
              <select
                value={vendorPurchaseType}
                onChange={e => setVendorPurchaseType(e.target.value)}
                className={`rounded-lg border px-2 py-1 text-xs font-semibold ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-white border-gray-200 text-slate-700'}`}
              >
                <option value="all">All</option>
                <option value="local">Local</option>
                <option value="import">Import</option>
              </select>
            </div>
            {barList(vendorRows, 'vendor', 'purchase_amount', 'PO Amount', '#2563EB')}
          </div>
          <div className={`p-5 rounded-2xl ${card}`}>
            <h3 className={`text-base font-bold ${txt}`}>Top 5 Client Sales Amount</h3>
            <p className={`text-xs mb-4 ${txt2}`}>Total: {fmtCurShort(sumRows(d.top_clients, 'sales_amount'))}</p>
            {barList(d.top_clients, 'client', 'sales_amount', 'Sales Amount', '#14B8A6')}
          </div>
        </div>
        {renderPendingDeliverySummary()}
      </div>
    );
  };

  const renderVendorControl = () => {
    const totalPages = Math.max(1, Math.ceil(vendorControlTotal / vendorControlPerPage));
    const applyVendorFilters = (vendors) => {
      const unique = [...new Set((vendors || []).map(v => String(v || '').trim()).filter(Boolean))];
      setVendorControlSelectedVendors(unique);
      setVendorControlAppliedVendors(unique);
      setVendorControlAppliedSearch('');
      setVendorControlSearch('');
      setVendorControlSuggestions([]);
      setVendorControlSuggestOpen(false);
      setVendorControlPage(1);
      fetchVendorControl(1, vendorControlPerPage, '', false, unique);
    };
    const addVendorFilter = (vendor) => {
      const value = String(vendor || '').trim();
      if (!value) return;
      applyVendorFilters([...vendorControlSelectedVendors, value]);
    };
    const removeVendorFilter = (vendor) => {
      applyVendorFilters(vendorControlSelectedVendors.filter(v => v !== vendor));
    };
    const handleSearch = () => {
      const q = vendorControlSearch.trim();
      if (q) {
        addVendorFilter(q);
        return;
      }
      setVendorControlAppliedVendors(vendorControlSelectedVendors);
      setVendorControlAppliedSearch('');
      setVendorControlPage(1);
      fetchVendorControl(1, vendorControlPerPage, '', false, vendorControlSelectedVendors);
    };
    const handleClear = () => {
      setVendorControlSearch('');
      setVendorControlAppliedSearch('');
      setVendorControlSelectedVendors([]);
      setVendorControlAppliedVendors([]);
      setVendorControlSuggestions([]);
      setVendorControlSuggestOpen(false);
      setVendorControlPage(1);
      fetchVendorControl(1, vendorControlPerPage, '', false, []);
    };
    const renderEditableVendorCell = (row, field, isPassword = false) => {
      const editing = editingCell?.id === row.row_key && editingCell.field === `vendor_${field}`;
      const visible = Boolean(vendorPasswordVisible[row.row_key]);
      if (editing) {
        return (
          <input
            type={isPassword && !visible ? 'password' : 'text'}
            value={editValue || ''}
            onChange={e => setEditValue(e.target.value)}
            onBlur={() => updateVendorControlCell(row.row_key, field, editValue)}
            onKeyDown={e => {
              if (e.key === 'Enter') updateVendorControlCell(row.row_key, field, editValue);
              if (e.key === 'Escape') setEditingCell(null);
            }}
            className={`w-full h-9 px-2 rounded-lg border text-sm ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-white border-blue-200 text-gray-900'}`}
            autoFocus
          />
        );
      }
      const display = isPassword && !visible ? '********' : row[field];
      return (
        <div className="flex items-center gap-2 min-w-0">
          <button
            type="button"
            onClick={() => { setEditingCell({ id: row.row_key, field: `vendor_${field}` }); setEditValue(row[field] || ''); }}
            className="min-w-0 flex-1 text-left truncate text-blue-600 hover:underline font-semibold"
            title={row[field] || ''}
          >
            {display || '-'}
          </button>
          {isPassword && (
            <button
              type="button"
              onClick={() => setVendorPasswordVisible(prev => ({ ...prev, [row.row_key]: !prev[row.row_key] }))}
              className={`flex-shrink-0 p-1.5 rounded-lg ${darkMode ? 'hover:bg-gray-700 text-gray-300' : 'hover:bg-slate-100 text-slate-500'}`}
              title={visible ? 'Hide password' : 'Show password'}
            >
              {visible ? <EyeOff className="w-4 h-4"/> : <Eye className="w-4 h-4"/>}
            </button>
          )}
        </div>
      );
    };

    return (
      <div className={`${card} overflow-hidden`}>
        <div className={`px-5 py-4 border-b ${darkMode?'border-gray-700':'border-gray-100'} flex flex-wrap justify-between items-center gap-3`}>
          <div>
            <h2 className={`text-lg font-bold ${txt}`}>Vendor Control</h2>
            <p className={`text-xs ${txt2}`}>Complete vendor login list from Google Sheet. Last update: {fmtDateTime(vendorControlLastUpdated)}</p>
          </div>
          <button
            onClick={() => fetchVendorControl(vendorControlPage, vendorControlPerPage, vendorControlAppliedSearch, true, vendorControlAppliedVendors)}
            className={`flex items-center gap-2 px-3 py-2 rounded-xl text-sm font-semibold ${darkMode?'bg-gray-700 text-gray-100 hover:bg-gray-600':'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}
          >
            <RotateCcw className="w-4 h-4"/>Refresh Sheet
          </button>
        </div>

        <FilterPanel darkMode={darkMode}>
          <div className="grid grid-cols-1 md:grid-cols-[minmax(260px,1fr)_110px_90px] gap-2 items-end">
            <div className="min-w-0 relative">
              <label className={`block text-xs font-semibold mb-1 ${txt2}`}>Search Vendor</label>
              <input
                ref={vendorControlSuggestFloat.triggerRef}
                value={vendorControlSearch}
                autoComplete="off"
                onChange={e => {
                  const next = e.target.value;
                  setVendorControlSearch(next);
                  setVendorControlSuggestOpen(next.trim().length >= 2);
                }}
                onFocus={() => setVendorControlSuggestOpen(vendorControlSearch.trim().length >= 2 && vendorControlSuggestions.length > 0)}
                onBlur={() => setTimeout(() => setVendorControlSuggestOpen(false), 120)}
                onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); setVendorControlSuggestOpen(false); handleSearch(); } }}
                placeholder="Type vendor name or ID, then select/add..."
                className={`w-full h-10 px-3 py-2 rounded-xl text-sm border ${darkMode?'bg-gray-700 border-gray-600 text-white placeholder:text-gray-400':'bg-white border-gray-200 text-gray-800 placeholder:text-gray-400'}`}
              />
              {vendorControlSuggestOpen && vendorControlSuggestions.length > 0 && (
                <div
                  style={vendorControlSuggestFloat.menuPos.style}
                  className={`max-h-64 overflow-auto rounded-xl border shadow-xl ${darkMode ? 'bg-gray-800 border-gray-700 text-gray-100' : 'bg-white border-gray-200 text-gray-800'}`}
                >
                  {vendorControlSuggestions.map(name => (
                    <button
                      key={name}
                      type="button"
                      onMouseDown={e => e.preventDefault()}
                      onClick={() => {
                        addVendorFilter(name);
                      }}
                      className={`block w-full px-3 py-2 text-left text-sm font-semibold ${darkMode ? 'hover:bg-gray-700' : 'hover:bg-blue-50'}`}
                    >
                      {name}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <button onClick={handleSearch} className="w-full h-10 px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold shadow-sm flex items-center justify-center gap-2">
              <Search className="w-4 h-4"/>{vendorControlSearch.trim() ? 'Add' : 'Search'}
            </button>
            <button onClick={handleClear} className={`w-full h-10 px-3 py-2 rounded-xl text-sm font-medium shadow-sm ${darkMode?'bg-gray-500 text-gray-100 hover:bg-gray-400':'bg-gray-400 text-white hover:bg-gray-500'}`}>
              Clear
            </button>
          </div>
          {vendorControlSelectedVendors.length > 0 && (
            <div className="mt-3 flex flex-wrap gap-2">
              {vendorControlSelectedVendors.map(vendor => (
                <span key={vendor} className={`inline-flex max-w-full items-center gap-2 rounded-full px-3 py-1.5 text-xs font-semibold ${darkMode ? 'bg-blue-900/40 text-blue-100 border border-blue-800' : 'bg-blue-50 text-blue-700 border border-blue-200'}`}>
                  <span className="max-w-[260px] truncate" title={vendor}>{vendor}</span>
                  <button type="button" onClick={() => removeVendorFilter(vendor)} className={`rounded-full p-0.5 ${darkMode ? 'hover:bg-blue-800' : 'hover:bg-blue-100'}`} title="Remove vendor filter">
                    <X className="w-3.5 h-3.5"/>
                  </button>
                </span>
              ))}
            </div>
          )}
        </FilterPanel>

        <DataTableScroll darkMode={darkMode}>
          <table className="w-full text-sm min-w-[860px]">
            <thead className={tblHd}>
              <tr>
                {['Vendor Name', 'Vendor ID', 'Password', 'Action'].map(label => (
                  <th key={label} className={`px-3 py-3 text-center font-bold ${txt2}`}>{label}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {vendorControlData.length === 0 ? (
                <tr><td colSpan={4} className={`px-4 py-12 text-center ${txt2}`}><Building2 className="w-10 h-10 mx-auto mb-2 opacity-40"/>No complete vendor login data</td></tr>
              ) : vendorControlData.map(row => (
                <tr key={row.row_key} className={trHov}>
                  <td className={`px-3 py-3 font-semibold ${txt}`}>{row.vendor_name}</td>
                  <td className="px-3 py-3 min-w-[180px]">{renderEditableVendorCell(row, 'vendor_id')}</td>
                  <td className="px-3 py-3 min-w-[220px]">{renderEditableVendorCell(row, 'password', true)}</td>
                  <td className="px-3 py-3 text-center">
                    <button
                      type="button"
                      onClick={() => openVendorLogin(row)}
                      className="inline-flex items-center justify-center gap-2 px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-bold shadow-sm"
                    >
                      <LinkIcon className="w-4 h-4"/>Login
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </DataTableScroll>

        <PagePagination
          darkMode={darkMode}
          txt2={txt2}
          page={vendorControlPage}
          totalPages={totalPages}
          total={vendorControlTotal}
          perPage={vendorControlPerPage}
          onPageChange={(p) => { setVendorControlPage(p); fetchVendorControl(p, vendorControlPerPage, vendorControlAppliedSearch, false, vendorControlAppliedVendors); }}
          onPerPageChange={(next) => { setVendorControlPerPage(next); setVendorControlPage(1); fetchVendorControl(1, next, vendorControlAppliedSearch, false, vendorControlAppliedVendors); }}
        />
      </div>
    );
  };

  const renderAllRegisteredItems = () => {
    const totalPages = Math.max(1, Math.ceil(registeredItemsTotal / registeredItemsPerPage));
    // NOTE: column order MUST match <colgroup> + <td> order below.
    // "Category ID" is placed to the LEFT of "Category" per spec.
    // "Vendor Name" column removed — source Excel has no Vendor column for
    // product master data.
    const columns = [
      ['Product ID', 'prod_id'],
      ['Category ID', 'category_id'],
      ['Category', 'category'],
      ['PIC', 'pic'],
      ['Product Name', 'prod_name'],
      ['Specification', 'spec'],
      ['Manufacturer Name', 'mfr_name'],
      ['Order Unit', 'odr_unit'],
      ['HUB Handling Check', 'hub_handling_check'],
      ['Tax Type', 'tax_type'],
      ['Registration Date', 'registration_date'],
      ['Product Registry PIC', 'product_registry_pic'],
    ];

    // Apply PIC dropdown filter (locally selected) — sends to backend which
    // resolves PIC by category_id OR category_name (same logic as Item
    // Registration / Pending Delivery).
    const applyPicFilter = (nextPic) => {
      setRegisteredItemsPicFilter(nextPic);
      setRegisteredItemsAppliedPicFilter(nextPic);
      setRegisteredItemsPage(1);
      fetchRegisteredItems(1, registeredItemsPerPage, registeredItemsAppliedSearch, registeredItemsAppliedProdIds, registeredItemsFilters, nextPic);
    };

    const handleClear = () => {
      setRegisteredItemsSearch('');
      setRegisteredItemsAppliedSearch('');
      setRegisteredItemsProdIds([]);
      setRegisteredItemsAppliedProdIds([]);
      setRegisteredItemsPicFilter('');
      setRegisteredItemsAppliedPicFilter('');
      const emptyFilters = { mfr_names: [] };
      setRegisteredItemsFilters(emptyFilters);
      setRegisteredItemsPage(1);
      fetchRegisteredItems(1, registeredItemsPerPage, '', [], emptyFilters, '');
    };

    const downloadRegisteredExcel = () => {
      const p = new URLSearchParams();
      if (registeredItemsAppliedSearch) p.append('search', registeredItemsAppliedSearch);
      (registeredItemsAppliedProdIds || []).forEach(v => p.append('prod_id', v));
      resolveFilter(registeredItemsFilters.mfr_names).forEach(v => p.append('mfr_name', v));
      if (registeredItemsAppliedPicFilter) p.append('pic_name', registeredItemsAppliedPicFilter);
      downloadBlob(`/api/export/all-registered-items?${p}`, `All_Registered_Items_${new Date().toISOString().slice(0,10)}.xlsx`, 'All Registered Items');
    };

    const fmtDateShort = (d) => {
      if (!d) return '-';
      try { return d.slice(0, 10); } catch { return d; }
    };
    const fmtProductId = (v) => {
      if (v == null || v === '') return '-';
      // Strip trailing ".0" that pandas may add when reading numeric columns
      // as floats. Do NOT strip leading zeros (Product IDs in source are
      // pure integers without leading zeros, but be defensive).
      return String(v).replace(/\.0+$/, '');
    };
    const fmtCategoryId = (v) => {
      if (v == null || v === '') return '-';
      // IMPORTANT: do NOT strip leading zeros. Category IDs in the source
      // Excel look like "000200030000100007" and the leading zeros are
      // significant (they encode hierarchy depth). Only strip the trailing
      // ".0" pandas appends when it accidentally reads the column as float.
      return String(v).replace(/\.0+$/, '');
    };

    return (
      <div className={`rounded-2xl overflow-hidden ${card}`}>
        <div className={`px-5 py-4 border-b ${darkMode ? 'border-gray-700' : 'border-gray-100'} flex flex-wrap justify-between items-center gap-3`}>
          <div className="flex items-center gap-2 min-w-0">
            <FileText className="w-5 h-5 text-blue-500 flex-shrink-0" />
            <h2 className={`text-lg font-bold ${txt}`}>All Registered Items</h2>
            <span className={`text-sm ${txt2}`}>({fmtNum(registeredItemsTotal)} records)</span>
          </div>
          <DownloadButton onClick={downloadRegisteredExcel} className="flex items-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-sm font-semibold shadow-sm">
            <Download className="w-4 h-4"/>Download Excel
          </DownloadButton>
        </div>

        <FilterPanel darkMode={darkMode}>
          {/* Grid: Search | Prod ID | PIC | Mfr Name | Search btn | Clear btn
              Vendor Name filter removed — source has no Vendor column. */}
          <div className="grid grid-cols-1 gap-2 md:grid-cols-2 xl:grid-cols-[minmax(240px,1fr)_170px_180px_minmax(180px,1fr)_90px_110px] items-end">
            <div className="min-w-0">
              <RFQMultiSearch value={registeredItemsSearch} onChange={setRegisteredItemsSearch} onSearch={(next) => { setRegisteredItemsAppliedSearch(next); setRegisteredItemsPage(1); fetchRegisteredItems(1, registeredItemsPerPage, next, registeredItemsAppliedProdIds, registeredItemsFilters, registeredItemsAppliedPicFilter); }} darkMode={darkMode} txt2={txt2} label="Search" description="Enter Product ID, Product Name, Specification, or Manufacturer per line. Results match any entered value." placeholder={'8381684\nBearing SKF\nJTC'} />
            </div>
            <div className="min-w-0">
              <label className={`block text-xs font-semibold mb-1 ${txt2}`}>Search Prod ID</label>
              <SearchInput key={`registered-prod-id-${registeredItemsProdIds.join('|')}`} placeholder={'8381684\n8382076'} label="Prod ID" darkMode={darkMode} txt2={txt2} onSearch={(nums) => { setRegisteredItemsProdIds(nums); setRegisteredItemsAppliedProdIds(nums); setRegisteredItemsPage(1); fetchRegisteredItems(1, registeredItemsPerPage, registeredItemsAppliedSearch, nums, registeredItemsFilters, registeredItemsAppliedPicFilter); }} />
            </div>
            {/* PIC dropdown — resolved server-side using MasterPIC by category_id
                OR category_name, identical to Item Registration / Pending Delivery. */}
            <div className="min-w-0">
              <label className={`block text-xs font-semibold mb-1 ${txt2}`}>PIC</label>
              <select
                value={registeredItemsPicFilter}
                onChange={(e) => applyPicFilter(e.target.value)}
                className={`w-full h-10 px-3 rounded-xl border text-sm ${darkMode ? 'bg-gray-700 border-gray-600 text-gray-100' : 'bg-white border-gray-300 text-gray-700'} focus:outline-none focus:ring-2 focus:ring-blue-500`}
              >
                <option value="">All PICs</option>
                {(registeredItemsOptions.pic_options || []).map(pic => (
                  <option key={pic} value={pic}>{pic}</option>
                ))}
              </select>
            </div>
            <MultiSelect label="Manufacturer Name" options={registeredItemsOptions.mfr_names || []} selected={registeredItemsFilters.mfr_names} onChange={v => { const next={...registeredItemsFilters, mfr_names:v}; setRegisteredItemsFilters(next); setRegisteredItemsPage(1); fetchRegisteredItems(1, registeredItemsPerPage, registeredItemsAppliedSearch, registeredItemsAppliedProdIds, next, registeredItemsAppliedPicFilter); }} darkMode={darkMode} txt2={txt2} />
            <button onClick={() => { setRegisteredItemsAppliedSearch(registeredItemsSearch); setRegisteredItemsPage(1); fetchRegisteredItems(1, registeredItemsPerPage, registeredItemsSearch, registeredItemsAppliedProdIds, registeredItemsFilters, registeredItemsAppliedPicFilter); }} className="w-full h-10 px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold shadow-sm">Search</button>
            <button onClick={handleClear} className={`w-full h-10 px-3 py-2 rounded-lg text-sm font-medium shadow-sm flex items-center justify-center whitespace-nowrap ${darkMode ? 'bg-gray-500 text-gray-100 hover:bg-gray-400' : 'bg-gray-400 text-white hover:bg-gray-500'}`}>Clear</button>
          </div>
        </FilterPanel>

        <DataTableScroll darkMode={darkMode}>
          <table className="freeze-table-all-registered-items w-full text-xs">
            <colgroup>
              <col style={{minWidth:'120px'}}/>
              <col style={{minWidth:'180px'}}/>
              <col style={{minWidth:'140px'}}/>
              <col style={{minWidth:'80px'}}/>
              <col style={{minWidth:'160px'}}/>
              <col style={{minWidth:'280px'}}/>
              <col style={{minWidth:'180px'}}/>
              <col style={{minWidth:'80px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'140px'}}/>
              <col style={{minWidth:'160px'}}/>
            </colgroup>
            <thead className={tblHd}>
              <tr>
                {columns.map(([label], index) => (
                  <th key={label} data-col-index={index + 1} className={`px-2 py-2 text-center font-bold whitespace-nowrap ${txt2}`}>{renderFreezeHeader('all-registered-items', index + 1, label)}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {registeredItemsData.length === 0 ? (
                <tr>
                  <td colSpan={columns.length} className={`px-4 py-12 text-center ${txt2}`}>
                    <Package className="w-10 h-10 mx-auto mb-2 opacity-40" />No registered item data
                  </td>
                </tr>
              ) : registeredItemsData.map(row => (
                <tr key={row.id} className={`${trHov} transition-colors`}>
                  <td data-col-index="1" className="px-2 py-2 font-mono text-blue-600 whitespace-nowrap text-center">{fmtProductId(row.prod_id)}</td>
                  <td data-col-index="2" className={`px-2 py-2 font-mono text-center whitespace-nowrap ${txt2}`} title={String(row.category_id || '')}>{fmtCategoryId(row.category_id)}</td>
                  <td data-col-index="3" className={`px-2 py-2 ${txt2}`} title={row.category}>{row.category || '-'}</td>
                  <td data-col-index="4" className="px-2 py-2 text-center whitespace-nowrap">
                    {row.pic ? (() => {
                      const c = getPicColor(row.pic);
                      return <span className={`px-1.5 py-0.5 rounded-full text-xs font-semibold ${c ? `${c.bg} ${c.text}` : 'bg-gray-100 text-gray-700'}`}>{row.pic}</span>;
                    })() : <span className={txt2}>-</span>}
                  </td>
                  <td data-col-index="5" className={`px-2 py-2 max-w-[160px] truncate ${txt}`} title={row.prod_name}>{row.prod_name || '-'}</td>
                  <td data-col-index="6" className={`px-2 py-2 max-w-[280px] truncate ${txt2}`} title={row.spec}>{row.spec || '-'}</td>
                  <td data-col-index="7" className={`px-2 py-2 max-w-[180px] truncate ${txt2}`} title={row.mfr_name}>{row.mfr_name || '-'}</td>
                  <td data-col-index="8" className={`px-2 py-2 text-center whitespace-nowrap ${txt2}`}>{row.odr_unit || '-'}</td>
                  <td data-col-index="9" className={`px-2 py-2 text-center whitespace-nowrap ${txt2}`}>{row.hub_handling_check || '-'}</td>
                  <td data-col-index="10" className={`px-2 py-2 text-center whitespace-nowrap ${txt2}`}>{row.tax_type || '-'}</td>
                  <td data-col-index="11" className={`px-2 py-2 whitespace-nowrap ${txt2}`}>{fmtDateShort(row.registration_date)}</td>
                  <td data-col-index="12" className={`px-2 py-2 max-w-[160px] truncate ${txt2}`} title={row.product_registry_pic}>{row.product_registry_pic || '-'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </DataTableScroll>

        <PagePagination
          darkMode={darkMode}
          txt2={txt2}
          page={registeredItemsPage}
          totalPages={totalPages}
          total={registeredItemsTotal}
          perPage={registeredItemsPerPage}
          onPageChange={(p) => { setRegisteredItemsPage(p); fetchRegisteredItems(p, registeredItemsPerPage, registeredItemsAppliedSearch, registeredItemsAppliedProdIds, registeredItemsFilters, registeredItemsAppliedPicFilter); }}
          onPerPageChange={(next) => { setRegisteredItemsPerPage(next); setRegisteredItemsPage(1); fetchRegisteredItems(1, next, registeredItemsAppliedSearch, registeredItemsAppliedProdIds, registeredItemsFilters, registeredItemsAppliedPicFilter); }}
        />
      </div>
    );
  };

  // Utility: convert any date string to YYYY-MM-DD for <input type="date">.
  // Defined at component scope so both renderRFQ and renderImport can use it.
  const toDateInputValue = (value) => {
    const raw = String(value || '').trim();
    if (!raw) return '';
    const ymd = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (ymd) return `${ymd[1]}-${String(ymd[2]).padStart(2, '0')}-${String(ymd[3]).padStart(2, '0')}`;
    const dmy = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (dmy && Number(dmy[2]) <= 12) {
      return `${dmy[3]}-${String(dmy[2]).padStart(2, '0')}-${String(dmy[1]).padStart(2, '0')}`;
    }
    if (dmy && Number(dmy[1]) <= 12) {
      return `${dmy[3]}-${String(dmy[1]).padStart(2, '0')}-${String(dmy[2]).padStart(2, '0')}`;
    }
    const monthMap = {
      jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
      jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
      january: 1, february: 2, march: 3, april: 4, june: 6,
      july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
    };
    const yearlessPatterns = [
      /^(\d{1,2})\s+([A-Za-z]+)$/,
      /^(\d{1,2})-([A-Za-z]+)$/,
      /^([A-Za-z]+)\s+(\d{1,2})$/,
      /^([A-Za-z]+)-(\d{1,2})$/,
    ];
    for (const pat of yearlessPatterns) {
      const m = raw.match(pat);
      if (!m) continue;
      let dayStr, monStr;
      if (Number.isFinite(Number(m[1]))) { dayStr = m[1]; monStr = m[2]; }
      else { dayStr = m[2]; monStr = m[1]; }
      const mon = monthMap[monStr.toLowerCase()];
      if (!mon) continue;
      const day = Number(dayStr);
      if (!day || day > 31) continue;
      const today = new Date();
      let year = today.getFullYear();
      const candidate = new Date(year, mon - 1, day);
      const todayMidnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      if (candidate < todayMidnight) year += 1;
      return `${year}-${String(mon).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
    const parsed = Date.parse(raw);
    if (!Number.isNaN(parsed)) {
      const d = new Date(parsed);
      if (!Number.isNaN(d.getTime())) {
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      }
    }
    return '';
  };

  const renderRFQ = () => {
    const totalPages = Math.max(1, Math.ceil(rfqTotal / rfqPerPage));
    const editableSet = new Set(rfqEditableFields || []);
    const baseColumns = (rfqColumns.length ? rfqColumns : [
      { field: 'check', label: 'Check' }, { field: 'sheet_status', label: 'Status' }, { field: 'days_left', label: 'Days Left' }, { field: 'no', label: 'No' }, { field: 'client_name', label: 'Nama Client' },
      { field: 'rfq_date', label: 'RFQ Date' }, { field: 'closing_date', label: 'Closing Date' }, { field: 'sales_pic', label: 'Sales PIC' },
      { field: 'category_name', label: 'Category Name' }, { field: 'purchase_pic', label: 'Purchase PIC' },
      { field: 'rfq_code', label: 'No. RFQ / KODE' }, { field: 'item_name', label: 'Item Name' }, { field: 'detail_spec', label: 'Detail Spec' }, { field: 'brand_manufacturer', label: 'Brand/Manufaktur' },
      { field: 'qty', label: 'Qty' }, { field: 'unit', label: 'Unit' }, { field: 'remark', label: 'Remark' },
      { field: 'product_id', label: 'Product ID' }, { field: 'request_number', label: 'Request Number' },
      { field: 'same_replacement', label: 'Same/Replacement' }, { field: 'vendor_name', label: 'Vendor Name' },
      { field: 'unit_price_idr', label: 'Unit Price (IDR)' }, { field: 'amt_idr', label: 'Amt (IDR)' }, { field: 'quoted_item_name', label: 'Item Name' },
      { field: 'quoted_spec', label: 'Spec' }, { field: 'quoted_brand', label: 'Brand' }, { field: 'quoted_unit', label: 'Unit' }, { field: 'moq', label: 'MOQ' },
      { field: 'lead_time_days', label: 'Lead Time (Days)' }, { field: 'valid_period', label: 'Valid period' }, { field: 'photo_url', label: 'Photo URL (optional)' },
      { field: 'remarks', label: 'Remarks' },
    ]).filter(col => col.field !== 'category_id');
    const similarityColumns = rfqSimilarityColumns.length ? rfqSimilarityColumns : [
      { field: 'similar_prod_ids', label: 'Similar Product ID' },
      { field: 'similar_prod_name', label: 'Similar Product Name' },
      { field: 'similar_spec', label: 'Similar Specification' },
      { field: 'similar_mfr_name', label: 'Similar Manufacturer' },
      { field: 'similar_odr_unit', label: 'Similar Unit' },
      { field: 'similar_score', label: '%Similarity' },
    ];
    const columns = rfqShowSimilarity ? [...baseColumns, ...similarityColumns] : baseColumns;
    // Shift+click multi-select helpers for the RFQ table (same pattern as
    // Import — see computeImportSelection for the canonical implementation).
    const rfqEditableFieldsList = columns.map(col => col.field);
    const rfqCellKey = (rowIndex, field) => `${rowIndex}|${field}`;
    const computeRfqSelection = (anchor, target) => {
      const startRow = Math.min(anchor.rowIndex, target.rowIndex);
      const endRow = Math.max(anchor.rowIndex, target.rowIndex);
      const startColIdx = Math.min(anchor.colIdx, target.colIdx);
      const endColIdx = Math.max(anchor.colIdx, target.colIdx);
      const result = new Set();
      for (let r = startRow; r <= endRow; r += 1) {
        for (let c = startColIdx; c <= endColIdx; c += 1) {
          result.add(rfqCellKey(r, rfqEditableFieldsList[c]));
        }
      }
      return result;
    };
    const rfqSourceStyleFields = new Set([
      'check', 'sheet_status', 'days_left', 'no', 'client_name', 'rfq_date', 'closing_date', 'sales_pic',
      'category_name', 'purchase_pic', 'rfq_code', 'item_name', 'detail_spec', 'brand_manufacturer', 'qty', 'unit', 'remark',
      'similar_prod_ids', 'similar_prod_name', 'similar_spec', 'similar_mfr_name', 'similar_odr_unit', 'similar_score'
    ]);
    const colWidth = (field) => ({
      check: 64, sheet_status: 90, days_left: 76, no: 70, client_name: 160, rfq_date: 110, closing_date: 110, sales_pic: 120,
      rfq_code: 150, item_name: 180, detail_spec: 620, brand_manufacturer: 160, qty: 80, unit: 80, remark: 380,
      category_id: 180, category_name: 150, product_id: 120, request_number: 150, purchase_pic: 120,
      same_replacement: 92, vendor_name: 200, unit_price_idr: 130, amt_idr: 130, quoted_item_name: 180,
      quoted_spec: 150, quoted_brand: 130, quoted_unit: 58, moq: 62, lead_time_days: 78,
      valid_period: 82, photo_url: 92, remarks: 360, private_remarks_1: 220, private_remarks_2: 220,
      similar_prod_ids: 150, similar_prod_name: 220, similar_spec: 280, similar_mfr_name: 140, similar_odr_unit: 78, similar_score: 96
    }[field] || 140);
    const colStyle = (field) => {
      const width = `${colWidth(field)}px`;
      return { width, minWidth: width, maxWidth: width };
    };
    const rfqTableWidth = columns.reduce((sum, col) => sum + colWidth(col.field), 0);
    const linkInfo = (value) => {
      const text = String(value || '').trim();
      const match = text.match(/https?:\/\/[^\s]+/i);
      if (!match) return null;
      try {
        const url = new URL(match[0]);
        const host = url.hostname.replace(/^www\./, '').toLowerCase();
        const known = [
          ['shopee', 'Shopee'], ['tokopedia', 'Tokopedia'], ['lazada', 'Lazada'], ['blibli', 'Blibli'],
          ['bukalapak', 'Bukalapak'], ['amazon', 'Amazon'], ['google', 'Google'], ['drive.google', 'Google Drive']
        ];
        const found = known.find(([key]) => host.includes(key));
        return { url: match[0], label: found ? found[1] : host.split('.')[0].replace(/^./, c => c.toUpperCase()) };
      } catch { return { url: match[0], label: 'Link' }; }
    };
    const renderValue = (value, extraClass = '') => {
      const link = linkInfo(value);
      if (link) {
        return <a href={link.url} target="_blank" rel="noreferrer" className="text-blue-600 hover:underline font-semibold" title={String(value || '')}>{link.label}</a>;
      }
      const display = value === 0 || value ? value : '-';
      return <span className={extraClass} title={String(value || '')}>{display}</span>;
    };
    const splitSimilarValues = (value, allowComma = false) => String(value || '')
      .split(allowComma ? /\r?\n|,\s*/ : /\r?\n/)
      .map(v => v.trim())
      .filter(Boolean);
    const renderSimilarLines = (value, className = '') => {
      const lines = splitSimilarValues(value);
      if (!lines.length) return <span>-</span>;
      return <div className="flex flex-col gap-1">{lines.map((line, index) => (
        <div key={`${line}-${index}`} className={`min-h-[22px] min-w-0 truncate ${className}`} title={line}>{line}</div>
      ))}</div>;
    };
    const isRFQClosingPast = (value) => {
      const raw = String(value || '').trim();
      if (!raw) return false;
      let d = null;
      const m = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
      else {
        const parsed = new Date(raw);
        if (!Number.isNaN(parsed.getTime())) d = parsed;
      }
      if (!d || Number.isNaN(d.getTime())) return false;
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      d.setHours(0, 0, 0, 0);
      return d < today;
    };
    const handleClear = () => {
      setRfqSearch('');
      setRfqAppliedSearch('');
      setRfqPicFilter('');
      const nextFilters = { checks: [], clients: [], rfq_numbers: [], brands: [], purchase_pics: [], vendors: [] };
      setRfqFilters(nextFilters);
      setRfqPage(1);
      fetchRFQData(1, rfqPerPage, '', false, nextFilters, '', rfqShowSimilarity);
    };
    const rfqParams = () => {
      const p = new URLSearchParams();
      if (rfqAppliedSearch) p.append('search', rfqAppliedSearch);
      if (rfqPicFilter) p.append('pic', rfqPicFilter);
      resolveFilter(rfqFilters.checks).forEach(v => p.append('check', v));
      resolveFilter(rfqFilters.clients).forEach(v => p.append('client_name', v));
      resolveFilter(rfqFilters.rfq_numbers).forEach(v => p.append('rfq_no', v));
      resolveFilter(rfqFilters.brands).forEach(v => p.append('brand_manufacturer', v));
      resolveFilter(rfqFilters.purchase_pics).forEach(v => p.append('purchase_pic', v));
      resolveFilter(rfqFilters.vendors).forEach(v => p.append('vendor_name', v));
      return p;
    };
    const downloadRFQTemplate = () => downloadBlob(`/api/rfq/template?${rfqParams()}`, `Template_RFQ_BatchUpload_${new Date().toISOString().slice(0,10)}.xlsx`, 'RFQ Batch Upload Template');
    const downloadRFQExcel = () => downloadBlob(`/api/export/rfq?${rfqParams()}`, `RFQ_${new Date().toISOString().slice(0,10)}.xlsx`, 'RFQ Excel');
    const handleRFQBatchUpload = async (e) => {
      const files = Array.from(e.target.files || []);
      if (!files.length) return;
      e.target.value = '';
      const fd = new FormData();
      files.forEach(file => fd.append('file', file));
      setUploadProgress({ label: files.length > 1 ? `RFQ Batch (${files.length} files)` : 'RFQ Batch', pct: 0 });
      try {
        const res = await api.post('/api/rfq/batch-upload', fd, {
          onUploadProgress: (ev) => setUploadProgress({ label: files.length > 1 ? `RFQ Batch (${files.length} files)` : 'RFQ Batch', pct: Math.round(ev.loaded * 100 / (ev.total || ev.loaded)) })
        });
        const syncInfo = res.data?.sheet_sync;
        if (syncInfo && syncInfo.synced === false) {
          addToast(`RFQ batch updated locally but Sheet sync failed: ${syncInfo.reason || 'unknown error'}`, 'warning');
        } else {
          addToast(`RFQ batch: ${res.data.updated || 0} cells updated${res.data.sheet_updates ? `, ${res.data.sheet_updates} synced to Sheet` : ''}${res.data.not_found ? `, ${res.data.not_found} No not found` : ''}`, 'success');
        }
        fetchRFQData(rfqPage, rfqPerPage, rfqAppliedSearch, false, rfqFilters, rfqPicFilter);
      } catch (err) {
        addToast(`Failed to upload RFQ batch: ${err.response?.data?.error || err.message}`, 'error');
      } finally {
        setUploadProgress(null);
      }
    };
    const visibleRfqPicKpis = rfqPicKpis.filter(row => row.pic && row.pic.toLowerCase() !== 'unassigned' && row.pic.trim() !== '');
    const totalPendingRFQ = visibleRfqPicKpis.reduce((sum, row) => sum + (Number(row.count) || 0), 0);
    const rfqKpis = [{ pic: 'Total Pending RFQ', count: totalPendingRFQ, isTotal: true }, ...visibleRfqPicKpis];
    const rfqKpiCols = Math.max(1, rfqKpis.length);
    const editableColumns = columns.filter(col => editableSet.has(col.field));
    const editableFieldOrder = editableColumns.map(col => col.field);
    const applyRFQPaste = async (startRowIndex, startField, text) => {
      const rows = String(text || '').replace(/\r/g, '').split('\n').filter((line, idx, arr) => line !== '' || idx < arr.length - 1);
      if (!rows.length) return;
      const startColIndex = editableFieldOrder.indexOf(startField);
      if (startColIndex < 0) return;
      const batchUpdates = [];
      for (let r = 0; r < rows.length; r += 1) {
        const targetRow = rfqData[startRowIndex + r];
        if (!targetRow) break;
        const values = rows[r].split('\t');
        for (let c = 0; c < values.length; c += 1) {
          const field = editableFieldOrder[startColIndex + c];
          if (!field) break;
          batchUpdates.push({ row_key: targetRow.row_key, field, value: values[c] });
        }
      }
      if (batchUpdates.length && await updateRFQCellsBatch(batchUpdates)) {
        addToast(`RFQ paste: ${batchUpdates.length} cells updated`, 'success');
      }
    };
    const fillRFQRange = async (startRowIndex, field, endRowIndex) => {
      if (endRowIndex === startRowIndex) return;
      const source = rfqData[startRowIndex];
      if (!source) return;
      const value = source[field] ?? '';
      const minRow = Math.max(0, Math.min(startRowIndex, endRowIndex));
      const maxRow = Math.min(rfqData.length - 1, Math.max(startRowIndex, endRowIndex));
      const batchUpdates = [];
      for (let i = minRow; i <= maxRow; i += 1) {
        if (i === startRowIndex || !rfqData[i]?.row_key) continue;
        batchUpdates.push({ row_key: rfqData[i].row_key, field, value });
      }
      if (batchUpdates.length && await updateRFQCellsBatch(batchUpdates)) {
        addToast(`RFQ drag-fill: ${batchUpdates.length} cells updated`, 'success');
      }
    };
    const startRFQFill = (event, rowIndex, field) => {
      event.preventDefault();
      event.stopPropagation();
      const getTargetCell = (evt) => document.elementFromPoint(evt.clientX, evt.clientY)?.closest('[data-rfq-cell="true"]');
      const updateRange = (evt) => {
        const target = getTargetCell(evt);
        const endRowIndex = Number(target?.getAttribute('data-row-index'));
        const targetField = target?.getAttribute('data-field');
        if (Number.isFinite(endRowIndex) && targetField === field) {
          setRfqFillRange({
            field,
            startRow: rowIndex,
            minRow: Math.min(rowIndex, endRowIndex),
            maxRow: Math.max(rowIndex, endRowIndex),
          });
        }
      };
      const cleanup = () => {
        document.body.classList.remove('rfq-fill-dragging');
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        setRfqFillRange(null);
      };
      const onMove = (moveEvent) => updateRange(moveEvent);
      const onUp = (upEvent) => {
        const target = getTargetCell(upEvent);
        const endRowIndex = Number(target?.getAttribute('data-row-index'));
        const targetField = target?.getAttribute('data-field');
        cleanup();
        if (Number.isFinite(endRowIndex) && targetField === field) fillRFQRange(rowIndex, field, endRowIndex);
      };
      document.body.classList.add('rfq-fill-dragging');
      setRfqFillRange({ field, startRow: rowIndex, minRow: rowIndex, maxRow: rowIndex });
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    };

    return (
      <div className={`rounded-2xl overflow-hidden ${card}`}>
        <div className={`px-5 py-4 border-b ${darkMode?'border-gray-700':'border-gray-100'} flex flex-wrap justify-between items-center gap-3`}>
          <div className="flex items-center gap-2 min-w-0">
            <Mail className="w-5 h-5 text-blue-500 flex-shrink-0" />
            <h2 className={`text-lg font-bold ${txt}`}>RFQ</h2>
            <span className={`text-sm ${txt2}`}>({fmtNum(rfqTotal)} records)</span>
            {rfqLastUpdated && <span className={`text-xs ${txt2}`}>Last sync: {fmtDate(rfqLastUpdated)}</span>}
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button onClick={downloadRFQTemplate} className={`flex items-center gap-2 px-3 py-2.5 rounded-xl text-sm font-semibold shadow-sm ${darkMode?'bg-gray-700 text-gray-100 hover:bg-gray-600':'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}>
              <Download className="w-4 h-4"/>Template
            </button>
            <label className="flex items-center gap-2 px-3 py-2.5 bg-slate-600 hover:bg-slate-700 text-white rounded-xl text-sm font-semibold shadow-sm cursor-pointer">
              <FileSpreadsheet className="w-4 h-4"/>Batch Upload
              <input type="file" accept=".xlsx,.xls" multiple onChange={handleRFQBatchUpload} className="hidden"/>
            </label>
            <DownloadButton onClick={downloadRFQExcel} className="flex items-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-sm font-semibold shadow-sm">
              <Download className="w-4 h-4"/>Download Excel
            </DownloadButton>
            <button onClick={() => { const next = !rfqShowSimilarity; setRfqShowSimilarity(next); setRfqPage(1); fetchRFQData(1, rfqPerPage, rfqAppliedSearch, false, rfqFilters, rfqPicFilter, next); }} className={`flex items-center gap-2 px-3 py-2.5 rounded-xl text-sm font-semibold shadow-sm ${rfqShowSimilarity ? 'bg-amber-100 text-amber-700 border border-amber-200 hover:bg-amber-200' : darkMode?'bg-gray-700 text-gray-100 hover:bg-gray-600':'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}>
              {rfqShowSimilarity ? <EyeOff className="w-4 h-4"/> : <Eye className="w-4 h-4"/>}{rfqShowSimilarity ? 'Hide Similarity' : 'Show Similarity'}
            </button>
            <button onClick={() => fetchRFQData(rfqPage, rfqPerPage, rfqAppliedSearch, true, rfqFilters, rfqPicFilter)} className={`flex items-center gap-2 px-3 py-2.5 rounded-xl text-sm font-semibold shadow-sm ${darkMode?'bg-gray-700 text-gray-100 hover:bg-gray-600':'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}>
              <RotateCcw className="w-4 h-4"/>Refresh Sheet
            </button>
          </div>
        </div>

        <div className={`px-5 py-3 border-b ${darkMode?'border-gray-700':'border-gray-100'}`}>
          <div className="grid grid-flow-col gap-2" style={{ gridTemplateColumns: `repeat(${rfqKpiCols}, minmax(0, 1fr))` }}>
            {rfqKpis.map((row) => {
              const activePic = !row.isTotal && rfqPicFilter === row.pic;
              const activePicColor = activePic ? getPicColor(row.pic) : null;
              const applyRFQPicFilter = () => {
                const nextPic = row.isTotal || activePic ? '' : row.pic;
                const nextFilters = { ...rfqFilters, purchase_pics: nextPic ? [nextPic] : [] };
                setRfqPicFilter(nextPic);
                setRfqFilters(nextFilters);
                setRfqPage(1);
                fetchRFQData(1, rfqPerPage, rfqAppliedSearch, false, nextFilters, nextPic, rfqShowSimilarity);
              };
              return (
                <button key={row.pic} type="button" onClick={applyRFQPicFilter} className={`min-w-0 p-3 rounded-xl text-left transition-all ${activePic ? (darkMode ? 'bg-amber-900/30 border border-amber-500 ring-2 ring-amber-400' : 'bg-amber-50 border border-amber-300 ring-2 ring-amber-200') : row.isTotal ? (darkMode ? 'bg-gray-800 border border-gray-700' : 'bg-gray-50 border border-gray-200') : card} ${row.isTotal ? 'hover:border-slate-300' : 'hover:border-amber-300'}`}>
                  <div className="flex items-start justify-between gap-2">
                    <div className="min-w-0">
                      <p className={`text-xs font-semibold truncate ${activePic ? 'text-amber-700' : row.isTotal ? (darkMode ? 'text-gray-200' : 'text-gray-700') : txt2}`} title={row.pic}>{row.pic}</p>
                      <h3 className={`text-xl font-bold leading-tight ${activePic ? 'text-amber-700' : row.isTotal ? (darkMode ? 'text-gray-100' : 'text-gray-800') : kpiValue}`}>{fmtNum(row.count)}</h3>
                      <p className={`text-[11px] leading-tight whitespace-nowrap ${txt2}`}>No Prod/Price</p>
                    </div>
                    <div className={`p-1.5 rounded-lg flex-shrink-0 ${activePic ? 'bg-amber-100 text-amber-700' : row.isTotal ? (darkMode ? 'bg-gray-700 text-gray-200' : 'bg-gray-100 text-gray-600') : neutralIcon}`}>
                      <Mail className="w-3.5 h-3.5" />
                    </div>
                  </div>
                </button>
              );
            })}
          </div>
        </div>

        <FilterPanel darkMode={darkMode}>
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-4 2xl:grid-cols-[minmax(160px,1fr)_105px_minmax(130px,0.95fr)_minmax(130px,0.95fr)_repeat(3,minmax(115px,0.9fr))_84px] items-end">
            <div className="min-w-0">
              <RFQMultiSearch
                value={rfqSearch}
                onChange={setRfqSearch}
                onSearch={(searchValue) => {
                  setRfqAppliedSearch(searchValue);
                  setRfqPage(1);
                  fetchRFQData(1, rfqPerPage, searchValue, false, rfqFilters, rfqPicFilter, rfqShowSimilarity);
                }}
                darkMode={darkMode}
                txt2={txt2}
              />
            </div>
            <div className="min-w-0">
              <MultiSelect label="Check" options={rfqOptions.checks || []} selected={rfqFilters.checks}
                onChange={v=>{ const next={...rfqFilters, checks:v}; setRfqFilters(next); setRfqPage(1); fetchRFQData(1, rfqPerPage, rfqAppliedSearch, false, next, rfqPicFilter, rfqShowSimilarity); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="Nama Client" options={rfqOptions.clients || []} selected={rfqFilters.clients}
                onChange={v=>{ const next={...rfqFilters, clients:v}; setRfqFilters(next); setRfqPage(1); fetchRFQData(1, rfqPerPage, rfqAppliedSearch, false, next, rfqPicFilter, rfqShowSimilarity); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="No. RFQ" options={rfqOptions.rfq_numbers || []} selected={rfqFilters.rfq_numbers}
                onChange={v=>{ const next={...rfqFilters, rfq_numbers:v}; setRfqFilters(next); setRfqPage(1); fetchRFQData(1, rfqPerPage, rfqAppliedSearch, false, next, rfqPicFilter, rfqShowSimilarity); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="Brand/Manufaktur" options={rfqOptions.brands || []} selected={rfqFilters.brands}
                onChange={v=>{ const next={...rfqFilters, brands:v}; setRfqFilters(next); setRfqPage(1); fetchRFQData(1, rfqPerPage, rfqAppliedSearch, false, next, rfqPicFilter, rfqShowSimilarity); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="Purchase PIC" options={rfqOptions.purchase_pics || []} selected={rfqFilters.purchase_pics}
                onChange={v=>{ const next={...rfqFilters, purchase_pics:v}; setRfqPicFilter(''); setRfqFilters(next); setRfqPage(1); fetchRFQData(1, rfqPerPage, rfqAppliedSearch, false, next, '', rfqShowSimilarity); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="Vendor Name" options={rfqOptions.vendors || []} selected={rfqFilters.vendors}
                onChange={v=>{ const next={...rfqFilters, vendors:v}; setRfqFilters(next); setRfqPage(1); fetchRFQData(1, rfqPerPage, rfqAppliedSearch, false, next, rfqPicFilter, rfqShowSimilarity); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <button onClick={handleClear} className={`w-full h-10 px-3 py-2 rounded-lg text-sm font-medium shadow-sm flex items-center justify-center whitespace-nowrap ${darkMode?'bg-gray-500 text-gray-100 hover:bg-gray-400':'bg-gray-400 text-white hover:bg-gray-500'}`}>
              Clear
            </button>
          </div>
        </FilterPanel>

        <DataTableScroll darkMode={darkMode}>
          <table className="freeze-table-rfq table-fixed text-xs border-collapse" style={{ width: `${rfqTableWidth}px`, minWidth: `${rfqTableWidth}px` }}>
            <colgroup>{columns.map(col => <col key={col.field} style={colStyle(col.field)}/>)}</colgroup>
            <thead className={tblHd}>
              <tr>{columns.map((col, index) => {
                const darkHeaderCols = ['check', 'sheet_status', 'days_left', 'no', 'client_name', 'rfq_date', 'closing_date', 'sales_pic', 'category_name', 'purchase_pic', 'rfq_code', 'item_name', 'detail_spec', 'brand_manufacturer', 'qty', 'unit', 'remark', 'similar_prod_ids', 'similar_prod_name', 'similar_spec', 'similar_mfr_name', 'similar_odr_unit', 'similar_score'];
                const isDarkHeader = darkHeaderCols.includes(col.field);
                return <th key={col.field} data-col-index={index + 1} className={`px-2 py-2 text-center font-bold whitespace-nowrap border-r ${isDarkHeader ? 'bg-slate-200 text-slate-700' : darkMode ? 'bg-gray-800/60 border-gray-700 text-gray-200' : 'bg-slate-50 border-gray-200 text-gray-700'} ${darkMode ? 'border-gray-700' : 'border-gray-200'}`}>{renderFreezeHeader('rfq', index + 1, col.label)}</th>;
              })}</tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {rfqData.length === 0 ? (
                <tr><td colSpan={columns.length} className={`px-4 py-12 text-center ${txt2}`}><Mail className="w-10 h-10 mx-auto mb-2 opacity-40"/>No RFQ data</td></tr>
              ) : rfqData.map((row, rowIndex) => {
                return (
                <tr key={row.row_key} className={`${trHov} transition-colors${rfqPicFilter && rfqEditedRowKeys.has(row.row_key) ? ' ring-1 ring-inset ring-amber-400/60' : ''}`}>
                  {columns.map((col, colIdx) => {
                    const field = col.field;
                    const value = row[field] ?? '';
                    const isEditable = editableSet.has(field);
                    const isEditing = editingCell?.id === row.row_key && editingCell.field === `rfq_${field}`;
                    if (field === 'check') {
                      const checkValue = String(row.check || '').toLowerCase();
                      if (checkValue === 'complete') {
                        return <td key={field} data-col-index={colIdx + 1} className={`px-2 py-2 text-center border-r ${darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200'}`} title="Complete"><span className="inline-flex h-6 w-6 items-center justify-center rounded-full bg-[#20B71F]"><Check className="w-4 h-4 text-white stroke-[4]"/></span></td>;
                      }
                      if (checkValue === 'reject') {
                        return <td key={field} data-col-index={colIdx + 1} className={`px-2 py-2 text-center border-r ${darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200'}`} title="Reject"><span className="inline-flex h-6 w-6 items-center justify-center rounded-full bg-[#EA0D0D]"><X className="w-4 h-4 text-white stroke-[4]"/></span></td>;
                      }
                      const closed = checkValue === 'closed' || (!row.product_id && isRFQClosingPast(row.closing_date));
                      return <td key={field} data-col-index={colIdx + 1} className={`px-2 py-2 text-center border-r ${darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200'}`} title={closed ? 'Closed' : 'Open'}><span className={`inline-flex h-6 w-6 rounded-full border ${closed ? (darkMode ? 'bg-gray-500 border-gray-400' : 'bg-gray-300 border-gray-400') : darkMode ? 'bg-gray-700 border-gray-500' : 'bg-white border-gray-300'}`}/></td>;
                    }
                    if (field === 'days_left') {
                      return <td key={field} data-col-index={colIdx + 1} className={`px-2 py-2 text-center border-r ${darkMode ? 'bg-gray-800/60 border-gray-700 text-gray-100' : 'bg-slate-50 border-gray-200 text-black'}`}>{row.days_left === 0 || row.days_left ? fmtNum(row.days_left) : '-'}</td>;
                    }
                    if (isEditable && isEditing) {
                      const tall = ['quoted_spec', 'remarks', 'photo_url'].includes(field);
                      const Control = tall ? 'textarea' : 'input';
                      // Same pattern as Import: td shows the blue outer outline,
                      // input/textarea/select inside has NO outline (data-no-focus-ring
                      // + inline style + global CSS rule). This avoids the
                      // double-blue-border bug.
                      const rfqInputCls = `block w-full min-h-8 px-2 py-1 text-xs border-0 rounded-none ring-0 outline-none focus:outline-none focus:ring-0 focus-visible:outline-none shadow-none ${darkMode?'bg-gray-700 text-white':'bg-white text-gray-900'}`;
                      const rfqInputStyle = { outline: 'none', outlineStyle: 'none', outlineWidth: 0, boxShadow: 'none', borderColor: 'transparent', borderWidth: 0 };
                      const rfqTdCls = `relative p-0 align-top border-r outline outline-2 outline-blue-500 outline-offset-[-2px] ${darkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200'}`;
                      if (['rfq_date', 'closing_date'].includes(field)) {
                        return <td key={field} data-col-index={colIdx + 1} data-rfq-cell="true" data-row-index={rowIndex} data-field={field} className={rfqTdCls}>
                          <input
                            type="date"
                            data-no-focus-ring=""
                            style={{ outline: 'none', outlineStyle: 'none', outlineWidth: 0, borderColor: 'transparent', borderWidth: 0 }}
                            value={toDateInputValue(editValue)}
                            className={`block w-full min-h-8 px-2 py-1 text-xs rounded-none outline-none focus:outline-none focus:ring-0 focus-visible:outline-none ${darkMode?'bg-gray-700 text-white':'bg-white text-gray-900'}`}
                            onChange={e => setEditValue(e.target.value)}
                            onBlur={() => updateRFQCell(row.row_key, field, editValue)}
                            onKeyDown={e => {
                              if (e.key === 'Enter') {
                                e.preventDefault();
                                updateRFQCell(row.row_key, field, editValue);
                              }
                              if (e.key === 'Escape') setEditingCell(null);
                            }}
                            autoFocus
                          />
                        </td>;
                      }
                      if (field === 'same_replacement') {
                        return <td key={field} data-col-index={colIdx + 1} data-rfq-cell="true" data-row-index={rowIndex} data-field={field} className={rfqTdCls}>
                          <select
                            data-no-focus-ring=""
                            style={rfqInputStyle}
                            value={editValue}
                            className={`${rfqInputCls} px-1.5`}
                            onChange={e => { setEditValue(e.target.value); updateRFQCell(row.row_key, field, e.target.value); }}
                            onBlur={() => setEditingCell(null)}
                            onKeyDown={e => { if (e.key === 'Escape') setEditingCell(null); }}
                            autoFocus
                          >
                            <option value=""></option>
                            <option value="Same">Same</option>
                            <option value="Replacement">Replacement</option>
                          </select>
                        </td>;
                      }
                      return <td key={field} data-col-index={colIdx + 1} data-rfq-cell="true" data-row-index={rowIndex} data-field={field} className={rfqTdCls}>
                        <Control
                          data-no-focus-ring=""
                          style={rfqInputStyle}
                          value={editValue}
                          rows={tall ? 3 : undefined}
                          className={`${rfqInputCls} ${tall ? 'resize-y' : ''}`}
                          onChange={e => setEditValue(e.target.value)}
                          onBlur={() => updateRFQCell(row.row_key, field, editValue)}
                          onPaste={e => {
                            const text = e.clipboardData.getData('text/plain');
                            if (text.includes('\t') || text.includes('\n')) {
                              e.preventDefault();
                              setEditingCell(null);
                              applyRFQPaste(rowIndex, field, text);
                            }
                          }}
                          onKeyDown={e => {
                            if (e.key === 'Enter' && !tall) {
                              e.preventDefault();
                              updateRFQCell(row.row_key, field, editValue);
                            }
                            if (e.key === 'Escape') setEditingCell(null);
                          }}
                          autoFocus
                        />
                      </td>;
                    }
                    if (isEditable) {
                      const hasValue = value === 0 || value;
                      const selected = rfqSelectedCell?.rowKey === row.row_key && rfqSelectedCell?.field === field;
                      const rfqCellKeyStr = rfqCellKey(rowIndex, field);
                      const inMultiSelection = rfqSelectedCells?.has(rfqCellKeyStr);
                      const sourceStyle = rfqSourceStyleFields.has(field);
                      const fillHighlighted = rfqFillRange?.field === field && rowIndex >= rfqFillRange.minRow && rowIndex <= rfqFillRange.maxRow && rowIndex !== rfqFillRange.startRow;
                      return <td key={field} data-col-index={colIdx + 1} data-rfq-cell="true" data-row-index={rowIndex} data-field={field}
                        tabIndex={0}
                        onFocus={() => setRfqSelectedCell({ rowKey: row.row_key, field })}
                        onClick={(e) => {
                          setRfqSelectedCell({ rowKey: row.row_key, field });
                          // Shift+click multi-select (Excel-like).
                          const clickedColIdx = rfqEditableFieldsList.indexOf(field);
                          if (e.shiftKey && rfqSelectionAnchor && clickedColIdx >= 0) {
                            const newSelection = computeRfqSelection(rfqSelectionAnchor, { rowIndex, colIdx: clickedColIdx });
                            setRfqSelectedCells(newSelection);
                          } else {
                            setRfqSelectionAnchor({ rowIndex, colIdx: clickedColIdx });
                            setRfqSelectedCells(new Set([rfqCellKeyStr]));
                            setEditingCell({ id: row.row_key, field: `rfq_${field}` });
                            if (field === 'unit_price_idr') {
                              setEditValue(String(value ?? '').replace(/[^0-9.-]/g, ''));
                            } else {
                              setEditValue(value ?? '');
                            }
                          }
                        }}
                        onPaste={e => { e.preventDefault(); applyRFQPaste(rowIndex, field, e.clipboardData.getData('text/plain')); }}
                        className={`group relative px-2 py-1 align-top border-r cursor-pointer ${sourceStyle ? (darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200') : (darkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200')} ${fillHighlighted ? 'outline outline-2 outline-blue-300 outline-offset-[-2px]' : inMultiSelection ? 'outline outline-2 outline-blue-500 outline-offset-[-2px] bg-blue-50/50' : selected ? 'outline outline-2 outline-blue-500 outline-offset-[-2px]' : 'hover:outline hover:outline-2 hover:outline-blue-400 hover:outline-offset-[-2px]'} ${['qty','unit_price_idr','moq','lead_time_days'].includes(field) ? 'text-right font-semibold' : ''}`}>
                        <div className={`min-h-7 min-w-0 truncate ${sourceStyle ? txt2 : 'text-blue-600'} ${field === 'photo_url' ? 'flex items-center gap-1 justify-center' : ''} ${field === 'purchase_pic' ? 'text-center' : ''}`}>
                          {field === 'photo_url' && <LinkIcon className={`w-3.5 h-3.5 flex-shrink-0 ${hasValue ? 'text-blue-600' : 'text-blue-400'}`} />}
                          {hasValue && field === 'purchase_pic' ? (() => {
                            const c = getPicColor(value);
                            return <span className={`inline-flex max-w-full truncate px-2 py-0.5 rounded-full text-[11px] font-semibold ${c ? `${c.bg} ${c.text}` : 'bg-gray-100 text-gray-700'}`}>{value}</span>;
                          })() : hasValue ? renderValue(value, sourceStyle ? txt2 : 'text-blue-600') : <span>{field === 'photo_url' ? '' : '\u00a0'}</span>}
                        </div>
                        <button type="button" aria-label="Fill down" title="Drag to copy this value" onClick={e => e.stopPropagation()} onMouseDown={e => startRFQFill(e, rowIndex, field)} className="rfq-fill-handle absolute bottom-0 right-0 h-3 w-3 translate-x-1/2 translate-y-1/2 border border-blue-600 bg-blue-600 opacity-0 group-hover:opacity-100 focus:opacity-100" />
                      </td>;
                    }
                    if (field === 'amt_idr') {
                      return <td key={field} className={`px-2 py-1 align-top text-right font-semibold border-r ${darkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200'} ${txt2}`}>
                        <div className="min-h-7 min-w-0 truncate">{renderValue(value)}</div>
                      </td>;
                    }
                    if (field === 'similar_prod_ids') {
                      const ids = splitSimilarValues(value, true);
                      const noSimilar = ids.length === 1 && ids[0].toLowerCase() === 'no similar item';
                      return <td key={field} className={`px-2 py-1 align-top border-r ${darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200'} ${txt2}`} title={String(value || '')}>
                        {ids.length ? (
                          <div className="flex flex-col gap-1">
                            {ids.map(id => (
                              <div key={id} className="flex items-center gap-1 min-w-0">
                                <button
                                  type="button"
                                  disabled={noSimilar}
                                  onClick={() => !noSimilar && setRfqSimilarAction(prev => (prev?.rowKey === row.row_key && prev?.productId === id ? null : { rowKey: row.row_key, productId: id }))}
                                  className={`min-h-[22px] min-w-0 flex-1 truncate text-left ${noSimilar ? 'font-semibold text-slate-500 cursor-default' : 'font-mono text-blue-600 hover:underline'}`}
                                  title={id}
                                >{id}</button>
                                {!row.product_id && !noSimilar && rfqSimilarAction?.rowKey === row.row_key && rfqSimilarAction?.productId === id && (
                                  <button
                                    type="button"
                                    onClick={() => { setRfqSimilarAction(null); updateRFQCell(row.row_key, 'product_id', id); }}
                                    className={`flex-shrink-0 px-1.5 py-0.5 rounded text-[10px] font-semibold ${darkMode ? 'bg-blue-900/50 text-blue-200 hover:bg-blue-800' : 'bg-blue-50 text-blue-700 hover:bg-blue-100 border border-blue-200'}`}
                                  >
                                    Use this ID
                                  </button>
                                )}
                              </div>
                            ))}
                          </div>
                        ) : <span>-</span>}
                      </td>;
                    }
                    if (['similar_prod_name', 'similar_spec', 'similar_mfr_name', 'similar_odr_unit'].includes(field)) {
                      return <td key={field} className={`px-2 py-1 align-top border-r ${darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200'} ${txt2}`} title={String(value || '')}>
                        {renderSimilarLines(value)}
                      </td>;
                    }
                    if (field === 'similar_score') {
                      const scores = splitSimilarValues(value, true).map(score => score ? `${String(score).replace(/%$/, '')}%` : '-');
                      return <td key={field} className={`px-2 py-1 align-top text-right font-semibold border-r ${darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200'} ${txt2}`} title={String(value || '')}>
                        {scores.length ? <div className="flex flex-col gap-1">{scores.map((score, index) => <div key={`${score}-${index}`} className="min-h-[22px]">{score}</div>)}</div> : '-'}
                      </td>;
                    }
                    if (field === 'purchase_pic') {
                      const c = getPicColor(value);
                      return <td key={field} className={`px-2 py-2 text-center truncate border-r ${darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200'}`}>{value ? <span className={`inline-flex max-w-full truncate px-2 py-0.5 rounded-full text-[11px] font-semibold ${c ? `${c.bg} ${c.text}` : 'bg-gray-100 text-gray-700'}`}>{value}</span> : <span className={txt2}>-</span>}</td>;
                    }
                    const darkDataCols = ['sheet_status', 'no', 'client_name', 'rfq_date', 'closing_date', 'sales_pic', 'category_name', 'purchase_pic', 'rfq_code', 'item_name', 'detail_spec', 'brand_manufacturer', 'qty', 'unit', 'remark', 'similar_prod_ids', 'similar_prod_name', 'similar_spec', 'similar_mfr_name', 'similar_odr_unit', 'similar_score'];
                    const isDarkDataCol = darkDataCols.includes(field);
                    return <td key={field} className={`px-2 py-2 align-top border-r ${isDarkDataCol ? (darkMode ? 'bg-gray-800/60 border-gray-700 text-gray-100' : 'bg-slate-50 border-gray-200 text-black') : (darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-white border-gray-200')} ${['detail_spec','remark','category_name','similar_spec'].includes(field) ? '' : 'truncate'} ${['qty','amt_idr','similar_score'].includes(field) ? 'text-right font-semibold' : ''} ${isDarkDataCol ? '' : txt2}`}>
                      {renderValue(value)}
                    </td>;
                  })}
                </tr>
              );})}
            </tbody>
          </table>
        </DataTableScroll>

        <PagePagination
          darkMode={darkMode}
          txt2={txt2}
          page={rfqPage}
          totalPages={totalPages}
          total={rfqTotal}
          perPage={rfqPerPage}
          onPageChange={(p) => { setRfqPage(p); fetchRFQData(p, rfqPerPage, rfqAppliedSearch, false, rfqFilters, rfqPicFilter); }}
          onPerPageChange={(next) => { setRfqPerPage(next); setRfqPage(1); fetchRFQData(1, next, rfqAppliedSearch, false, rfqFilters, rfqPicFilter); }}
        />
      </div>
    );
  };

  const renderImport = () => {
    const totalPages = Math.max(1, Math.ceil(importTotal / importPerPage));
    const columns = importColumns || [];
    const checklistFields = new Set(columns.filter(col => isImportHideableChecklistColumn(col)).map(col => col.field));
    const checklistCount = checklistFields.size;
    // "Detail" block = per-item columns from SO through PURCHASE AMOUNT.
    // Toggled by the Show/Hide Detail button. When hidden, the table
    // becomes much narrower (1 line per row) — useful for quick overview
    // without the long Spec / Item Name / Remark columns eating space.
    const IMPORT_DETAIL_FIELDS = new Set([
      'so', 'group', 'po_date_by_email', 'po_sementara', 'item_yupi',
      'item_name', 'spec', 'remark_yupi', 'reschedule', 'ord_qty', 'unit',
      'unit_price', 'amount', 'purchase_price', 'currency', 'purchase_amount',
    ]);
    const detailCount = columns.filter(col => IMPORT_DETAIL_FIELDS.has(col.field)).length;
    // Apply both filters: hide checklist (if toggled off) + hide detail (if toggled off).
    const visibleColumns = columns.filter(col => {
      if (!showImportChecklist && checklistFields.has(col.field)) return false;
      if (!showImportDetail && IMPORT_DETAIL_FIELDS.has(col.field)) return false;
      return true;
    });
    // A "group" column repeats the same value for every item that belongs to
    // the same PO (Status, Vendor, dates, checklist docs, ...). Per-item
    // columns (SO, PO Sementara, Item Name, Spec, Qty, prices, ...) carry
    // `group_per_item: true` from the backend and are never merged.
    const isImportGroupColumn = (col) => !col.group_per_item;
    // Merge key: rows merge into one visual group when YUPI PO + Req Dlv Date
    // BOTH match (user spec). Other group columns (status, days_left, vendor,
    // etc.) are NOT part of the merge key — they may differ between items in
    // the same delivery group, and that's fine. The merged cell will show
    // the first row's value for those columns.
    const IMPORT_MERGE_KEY_FIELDS = ['yupi_po', 'req_dlv_date'];
    const importRowSpans = (() => {
      const spans = new Array(importData.length).fill(null);
      let groupStart = 0;
      const sameGroup = (a, b) => IMPORT_MERGE_KEY_FIELDS.every(f => String(a?.[f] ?? '').trim() === String(b?.[f] ?? '').trim());
      // Require YUPI PO to be non-blank for merging — blank YUPI PO rows
      // should never merge (they're individual items without a PO).
      const hasMergeKey = (row) => String(row?.yupi_po ?? '').trim() !== '';
      for (let i = 1; i <= importData.length; i++) {
        const continues = i < importData.length && hasMergeKey(importData[i]) && hasMergeKey(importData[groupStart]) && sameGroup(importData[i], importData[groupStart]);
        if (!continues) {
          spans[groupStart] = i - groupStart;
          groupStart = i;
        }
      }
      return spans;
    })();
    // For every row, which row index "owns" the merged group cell (i.e. the
    // row whose <td> actually renders, with rowSpan covering the rest).
    const importGroupStartIndexFor = (() => {
      const owners = new Array(importData.length).fill(0);
      let currentStart = 0;
      for (let i = 0; i < importData.length; i++) {
        if (importRowSpans[i] != null) currentStart = i;
        owners[i] = currentStart;
      }
      return owners;
    })();
    // Zebra striping by GROUP (not by row). Every group gets an alternating
    // color so merged rows share one shade, the next group gets the other.
    // Group index = count of group starts up to and including this row's
    // group. Even groups → shade A, odd groups → shade B.
    const importGroupIndexFor = (() => {
      const indices = new Array(importData.length).fill(0);
      let groupIdx = 0;
      for (let i = 0; i < importData.length; i++) {
        if (importRowSpans[i] != null) groupIdx += 1;
        indices[i] = groupIdx;
      }
      return indices;
    })();
    const colWidth = (col) => {
      if (col.field === 'days_left') return 64;
      // Use the backend-provided width (already tightened to match content),
      // clamped to a sensible min/max. Spec column is allowed to exceed the
      // max so long item specs stay readable.
      if (Number(col.width)) {
        const isSpec = String(col.field || '').toLowerCase() === 'spec' || String(col.label || '').toLowerCase().includes('spec');
        if (isSpec) return Math.max(340, Math.min(Number(col.width), 420));
        return Math.max(56, Math.min(Number(col.width), 320));
      }
      const label = String(col.label || '').toLowerCase();
      if (isImportChecklistColumn(col)) return 70;
      if (label.includes('spec') || label.includes('remark')) return 280;
      if (label.includes('item name')) return 200;
      if (label.includes('vendor')) return 140;
      if (label.includes('status')) return 110;
      if (label.includes('date') || label.includes('actual') || ['etd', 'eta'].includes(col.field)) return 100;
      if (isImportHyperlinkColumn(col)) return 150;
      return 110;
    };
    const tableWidth = Math.max(1100, visibleColumns.reduce((sum, col) => sum + colWidth(col), 0));
    const handleVendorUpload = async (e) => {
      const files = Array.from(e.target.files || []);
      e.target.value = '';
      if (!files.length) return;
      const fd = new FormData();
      files.forEach(file => fd.append('file', file));
      try {
        const res = await api.post('/api/import/vendors/upload', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
        addToast(res.data?.message || 'Import vendors updated', 'success');
        setImportPage(1);
        fetchImportData(1, importPerPage, importAppliedSearch, true, importFilters, importReqDlvSort, importYupiPoSort);
      } catch (err) {
        addToast(`Failed to upload import vendors: ${err.response?.data?.error || err.message}`, 'error');
      }
    };
    const startImportEdit = (row, col) => {
      const key = `${row._row_key}:${col.field}`;
      setImportEditingCell(key);
      // For date fields, normalize the value to YYYY-MM-DD up front so the
      // <input type="date"> picker can parse it and show the calendar with
      // the correct day pre-selected. Without this, a value like "2 Jun"
      // would make the picker open blank (the input can't parse it).
      const isDateField = ['po_send_date', 'po_date_by_email', 'req_dlv_date', 'source_req_dlv_date', 'reschedule', 'etd', 'eta', 'payment_date'].includes(col.field) || String(col.label || '').toLowerCase().includes('date');
      const rawValue = String(row[col.field] ?? '');
      // For payment_date, the backend may set a sentinel "Overdue" value
      // when no real date has been entered. When opening the editor, treat
      // "Overdue" as empty so the date picker opens blank — the user can
      // then enter an actual date.
      const effectiveValue = col.field === 'payment_date' && rawValue === 'Overdue' ? '' : rawValue;
      setImportEditValue(isDateField ? toDateInputValue(effectiveValue) : effectiveValue);
    };
    const renderImportCell = (row, col) => {
      const key = `${row._row_key}:${col.field}`;
      const editing = importEditingCell === key;
      const value = row[col.field] ?? '';

      if (col.field === 'status' || col.type === 'status') {
        const poSendDate = String(row.po_send_date || '').trim();
        const isNewImport = !poSendDate;
        const current = String(isNewImport ? 'NEW' : (value && String(value).toUpperCase() !== 'NEW' ? value : 'ON PROCESS')).toUpperCase();
        // Always render a <select> dropdown — even for NEW status — so the
        // user can change it. Previously NEW was a static <span> badge with
        // no way to advance the status. Now NEW appears as the selected
        // option, and the user can pick ON PROCESS / ON DELIVERY / etc.
        return (
          <select
            value={current}
            onChange={(e) => updateImportCell(row._row_key, col.field, e.target.value)}
            className={`w-full h-7 rounded-lg border px-2 py-0 text-[11px] font-bold outline-none cursor-pointer ${importStatusClass(current, darkMode)}`}
          >
            {/* Include NEW as a selectable option so the user can see the
                current value AND optionally revert to NEW (rare, but
                possible if a PO Send Date was entered by mistake). */}
            <option value="NEW" style={importStatusOptionStyle('NEW')}>NEW</option>
            {IMPORT_STATUS_OPTIONS.filter(opt => String(opt).toUpperCase() !== 'NEW').map(opt => <option key={opt} value={opt} style={importStatusOptionStyle(opt)}>{opt}</option>)}
          </select>
        );
      }

      if (isImportChecklistColumn(col)) {
        const checked = importCheckboxChecked(value);
        return (
          <button
            type="button"
            onClick={() => updateImportCell(row._row_key, col.field, checked ? 'FALSE' : 'TRUE')}
            className="flex w-full items-center justify-center py-1"
            title={checked ? 'Checked' : 'Unchecked'}
          >
            {checked ? (
              <span className={`inline-flex h-6 w-6 items-center justify-center rounded-full ${darkMode ? 'bg-emerald-600' : 'bg-emerald-500'}`}>
                <Check className="w-4 h-4 text-white stroke-[4]" />
              </span>
            ) : (
              <span className={`inline-flex h-6 w-6 rounded-full border ${darkMode ? 'bg-gray-700 border-gray-500' : 'bg-white border-gray-300'}`} />
            )}
          </button>
        );
      }

      // ── Payment dropdown (empty / "DONE") ──────────────────────────────
      if (col.field === 'payment' || col.payment_dropdown) {
        const current = String(value || '').trim().toUpperCase();
        return (
          <select
            value={current === 'DONE' ? 'DONE' : ''}
            onChange={(e) => updateImportCell(row._row_key, col.field, e.target.value)}
            className={`w-full h-7 rounded-lg border px-1 py-0 text-[11px] font-bold outline-none cursor-pointer ${current === 'DONE' ? (darkMode ? 'bg-emerald-700 text-emerald-100 border-emerald-600' : 'bg-emerald-100 text-emerald-700 border-emerald-300') : (darkMode ? 'bg-gray-700 text-gray-200 border-gray-600' : 'bg-white text-gray-700 border-gray-300')}`}
          >
            <option value=""></option>
            <option value="DONE">DONE</option>
          </select>
        );
      }

      // ── Payment Date — date picker with auto-"Overdue" red badge ───────
      // Backend sets payment_date = "Overdue" when today > ETA + TOP days
      // AND no real date has been entered yet. Render "Overdue" as a red
      // pill; otherwise render the actual date (click to open date picker).
      if (col.field === 'payment_date' || col.payment_date) {
        const str = String(value ?? '').trim();
        if (editing) {
          // Date picker — uses the same date-input styling as other date fields.
          const dateInputCls = `block w-full min-h-8 px-2 py-1 text-xs rounded-none outline-none focus:outline-none ${darkMode ? 'bg-gray-700 text-white' : 'bg-white text-gray-900'}`;
          const dateInputStyle = { outline: 'none', outlineStyle: 'none', outlineWidth: 0, borderColor: 'transparent', borderWidth: 0 };
          return (
            <input
              type="date"
              autoFocus
              data-no-focus-ring=""
              className={dateInputCls}
              style={dateInputStyle}
              value={toDateInputValue(importEditValue)}
              onFocus={e => { try { e.currentTarget.showPicker?.(); } catch {} }}
              onClick={e => { try { e.currentTarget.showPicker?.(); } catch {} }}
              onChange={e => setImportEditValue(e.target.value)}
              onBlur={e => updateImportCell(row._row_key, col.field, e.target.value)}
              onKeyDown={e => {
                if (e.key === 'Enter') updateImportCell(row._row_key, col.field, e.currentTarget.value);
                if (e.key === 'Escape') setImportEditingCell(null);
              }}
            />
          );
        }
        if (str === 'Overdue') {
          return (
            <button
              type="button"
              onClick={() => startImportEdit(row, col)}
              className="block w-full text-center"
              title="Overdue — click to enter a payment date"
            >
              <span className={`inline-flex h-6 items-center rounded-full px-2 text-[11px] font-bold ${darkMode ? 'bg-red-900/60 text-red-200 border border-red-700' : 'bg-red-100 text-red-700 border border-red-300'}`}>
                Overdue
              </span>
            </button>
          );
        }
        const display = importDisplayValue(value);
        return (
          <button
            type="button"
            className={`block w-full truncate whitespace-nowrap leading-6 text-left ${darkMode ? 'text-gray-200' : 'text-slate-700'} hover:underline decoration-dotted`}
            title={display || 'Click to enter payment date'}
            onClick={() => startImportEdit(row, col)}
          >
            {display || <span className="text-slate-400">-</span>}
          </button>
        );
      }

      if (editing) {
        const isLong = ['spec', 'remark_yupi', 'import_remarks', 'soft_copy_doc'].includes(col.field);
        const isDateField = ['po_send_date', 'po_date_by_email', 'req_dlv_date', 'source_req_dlv_date', 'reschedule', 'etd', 'eta', 'payment_date'].includes(col.field) || String(col.label || '').toLowerCase().includes('date');
        // Input fills the td edge-to-edge with NO inner ring/outline — the td
        // itself already shows the blue outer outline when editingCellNow is
        // true (see the td className above). A second inner ring here would
        // produce the double blue border the user reported.
        //
        // CRITICAL: we also pass inline `style={{ outline: 'none', boxShadow: 'none' }}`
        // because Tailwind's `focus:outline-none` / `focus:ring-0` classes are
        // NOT enough to suppress the browser's default focus ring on autoFocus
        // inputs in Chrome/Safari/Edge. The inline style wins over the
        // user-agent stylesheet, which is what actually kills the second blue
        // border. The `data-no-focus-ring` attribute is also added so a CSS
        // rule in <style> can target it as a backstop (see the global CSS
        // block at the top of this file for `input[data-no-focus-ring]:focus`).
        const inputCls = `block w-full min-h-8 px-2 py-1 text-xs border-0 rounded-none ring-0 outline-none focus:outline-none focus:ring-0 focus-visible:outline-none shadow-none ${darkMode ? 'bg-gray-700 text-white' : 'bg-white text-gray-900'}`;
        const inputStyle = { outline: 'none', outlineStyle: 'none', outlineWidth: 0, boxShadow: 'none', borderColor: 'transparent', borderWidth: 0 };
        // Date input: clean style without box-shadow/appearance suppression — these break the native calendar picker.
        const dateInputCls = `block w-full min-h-8 px-2 py-1 text-xs rounded-none outline-none focus:outline-none ${darkMode ? 'bg-gray-700 text-white' : 'bg-white text-gray-900'}`;
        const dateInputStyle = { outline: 'none', outlineStyle: 'none', outlineWidth: 0, borderColor: 'transparent', borderWidth: 0 };
        if (isDateField) {
          return (
            <input
              type="date"
              autoFocus
              data-no-focus-ring=""
              className={dateInputCls}
              style={dateInputStyle}
              value={toDateInputValue(importEditValue)}
              onFocus={e => { try { e.currentTarget.showPicker?.(); } catch {} }}
              onClick={e => { try { e.currentTarget.showPicker?.(); } catch {} }}
              onChange={e => setImportEditValue(e.target.value)}
              onBlur={e => updateImportCell(row._row_key, col.field, e.target.value)}
              onPaste={e => {
                const text = e.clipboardData.getData('text/plain');
                if (text.includes('\t') || text.includes('\n')) {
                  e.preventDefault();
                  setImportEditingCell(null);
                  const rowIndex = importData.findIndex(r => r._row_key === row._row_key);
                  applyImportPaste(rowIndex, col.field, text);
                }
              }}
              onKeyDown={e => {
                if (e.key === 'Enter') updateImportCell(row._row_key, col.field, e.currentTarget.value);
                if (e.key === 'Escape') setImportEditingCell(null);
              }}
            />
          );
        }
        if (isLong) {
          return (
            <textarea
              autoFocus
              data-no-focus-ring=""
              rows={3}
              className={`${inputCls} resize-y min-h-[74px]`}
              style={inputStyle}
              value={importEditValue}
              onChange={e => setImportEditValue(e.target.value)}
              onBlur={() => updateImportCell(row._row_key, col.field, importEditValue)}
              onPaste={e => {
                const text = e.clipboardData.getData('text/plain');
                if (text.includes('\t') || text.includes('\n')) {
                  e.preventDefault();
                  setImportEditingCell(null);
                  const rowIndex = importData.findIndex(r => r._row_key === row._row_key);
                  applyImportPaste(rowIndex, col.field, text);
                }
              }}
              onKeyDown={e => {
                if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) updateImportCell(row._row_key, col.field, importEditValue);
                if (e.key === 'Escape') setImportEditingCell(null);
              }}
            />
          );
        }
        return (
          <input
            autoFocus
            data-no-focus-ring=""
            className={inputCls}
            style={inputStyle}
            value={importEditValue}
            onChange={e => setImportEditValue(e.target.value)}
            onBlur={() => updateImportCell(row._row_key, col.field, importEditValue)}
            onPaste={e => {
              const text = e.clipboardData.getData('text/plain');
              if (text.includes('\t') || text.includes('\n')) {
                e.preventDefault();
                setImportEditingCell(null);
                const rowIndex = importData.findIndex(r => r._row_key === row._row_key);
                applyImportPaste(rowIndex, col.field, text);
              }
            }}
            onKeyDown={e => {
              if (e.key === 'Enter') updateImportCell(row._row_key, col.field, importEditValue);
              if (e.key === 'Escape') setImportEditingCell(null);
            }}
          />
        );
      }

      if (isImportHyperlinkColumn(col)) {
        const directUrl = row[`${col.field}__url`] || row[`${col.field}_url`] || '';
        const driveUrl = extractGDriveUrl(directUrl || value);
        const label = String(value || '').replace(driveUrl, '').replace(/[|\n]+$/g, '').trim() || (driveUrl ? gDriveChipLabel(driveUrl) : '-');
        // Single link element only — no separate gray edit button. The cell
        // itself is clickable (td onClick → startImportEdit) so the user can
        // still edit by clicking the cell area outside the link chip.
        return (
          <div className="flex items-center gap-1 w-full">
            {driveUrl ? (
              <a href={driveUrl} target="_blank" rel="noopener noreferrer" title={driveUrl} onClick={e => e.stopPropagation()}
                 className={`flex min-w-0 max-w-full h-6 items-center gap-1 px-2 py-0 rounded-full text-[11px] font-medium truncate border ${darkMode ? 'bg-blue-900/40 text-blue-300 border-blue-800 hover:bg-blue-900/70' : 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100'}`}>
                <LinkIcon className="w-3.5 h-3.5 flex-shrink-0" />
                <span className="truncate">{label}</span>
              </a>
            ) : (
              <button type="button" onClick={(e) => { e.stopPropagation(); startImportEdit(row, col); }} className="block min-w-0 truncate text-left hover:underline text-blue-600" title={String(value || '-')}>{label}</button>
            )}
          </div>
        );
      }

      if (col.field === 'days_left') {
        const dayValue = String(value ?? '').trim();
        if (dayValue === '✅') {
          return <span className="inline-flex h-6 w-6 items-center justify-center rounded-full bg-emerald-500 mx-auto" title="Delivered"><Check className="w-4 h-4 text-white stroke-[4]" /></span>;
        }
        if (dayValue === '❌') {
          return <span className="inline-flex h-6 w-6 items-center justify-center rounded-full bg-red-500 mx-auto" title="Canceled"><X className="w-4 h-4 text-white stroke-[4]" /></span>;
        }
        // Color coding (final thresholds per user):
        //   red    → days <= 7  (urgent) OR negative / minus (overdue)
        //   yellow → 8 <= days <= 29  (moderate)
        //   green  → days >= 30  (plenty of time)
        //   black  → exactly 0   (today is the delivery day)
        // Empty / '-' / non-numeric → no styling.
        // Colors use soft Tailwind palette tones that match the existing
        // dashboard theme (emerald / amber / red soft tones).
        const dayNum = Number(dayValue);
        if (dayValue !== '' && dayValue !== '-' && Number.isFinite(dayNum)) {
          let bg, text;
          if (dayNum === 0) {
            // TODAY — dark pill with white text (matches dashboard accent)
            bg = darkMode ? 'bg-gray-900' : 'bg-slate-800';
            text = 'text-white';
          } else if (dayNum < 0 || dayNum <= 7) {
            // Red — urgent (<=7) or overdue (negative / minus)
            bg = darkMode ? 'bg-red-900/55 text-red-100' : 'bg-red-100 text-red-700';
            text = '';
          } else if (dayNum >= 30) {
            // Green — plenty of time
            bg = darkMode ? 'bg-emerald-900/55 text-emerald-100' : 'bg-emerald-100 text-emerald-700';
            text = '';
          } else {
            // Yellow — moderate (8-29 days)
            bg = darkMode ? 'bg-amber-900/55 text-amber-100' : 'bg-amber-100 text-amber-700';
            text = '';
          }
          return (
            <span className={`inline-flex h-6 min-w-[28px] items-center justify-center rounded-md px-1.5 text-[11px] font-bold ${bg} ${text}`} title={`Days Left: ${dayNum}`}>
              {dayNum === 0 ? 'TODAY' : fmtNum(dayNum)}
            </span>
          );
        }
        // Empty fallback
        return <span className="text-slate-400">-</span>;
      }

      if (col.field === 'arrival_check') {
        const text = importDisplayValue(value);
        return <span className={`inline-flex max-w-full h-6 items-center rounded-full border px-2 py-0 text-[11px] font-semibold truncate ${importArrivalClass(value, darkMode)}`} title={text}>{text}</span>;
      }

      const isFormula = isImportFormulaColumn(col);
      const isCenter = ['days_left'].includes(col.field);
      const isNumeric = !isCenter && (Boolean(col.number) || ['ord_qty', 'unit_price', 'amount', 'purchase_price', 'purchase_amount', 'lt_days'].includes(col.field));
      const display = importDisplayValue(value);
      const cellInnerClass = `${isCenter ? 'text-center tabular-nums font-bold' : isNumeric ? 'text-right tabular-nums font-semibold' : 'text-left'} ${isFormula ? (darkMode ? 'text-gray-200' : 'text-slate-700') : ''} block w-full truncate whitespace-nowrap leading-6`;
      if (col.field === 'req_dlv_date' && String(row.reschedule || '').trim()) {
        return (
          <div className="group/req flex min-w-0 items-center gap-1">
            <button type="button" className={`block min-w-0 flex-1 ${cellInnerClass} hover:underline decoration-dotted`} title={display} onClick={() => startImportEdit(row, col)}>{display}</button>
            <button
              type="button"
              title={`Update Req Dlv Date to ${row.reschedule}`}
              onClick={(e) => { e.stopPropagation(); updateImportCellsBatch([{ row_key: row._row_key, field: 'req_dlv_date', value: row.reschedule }, { row_key: row._row_key, field: 'reschedule', value: '' }]); }}
              className="hidden group-hover/req:inline-flex flex-shrink-0 rounded-md bg-amber-500 px-1.5 py-0.5 text-[10px] font-bold text-white hover:bg-amber-600"
            >Update</button>
          </div>
        );
      }
      return <button type="button" className={`block w-full ${cellInnerClass} hover:underline decoration-dotted`} title={display} onClick={() => startImportEdit(row, col)}>{display}</button>;
    };

    const importEditableFields = visibleColumns.map(col => col.field);

    // ── Shift+click multi-select (Excel-like) ─────────────────────────────
    // Click a cell → set anchor + select just that cell.
    // Shift+click another cell → select the rectangle from anchor to clicked
    //   cell. Works across rows AND columns: e.g. anchor at (row 2, col C),
    //   Shift+click (row 5, col E) selects rows 2–5 × columns C–E.
    // Shift+click in the same column → selects a vertical range (rows only).
    // Click without Shift → reset to single-cell selection.
    const importCellKey = (rowIndex, field) => `${rowIndex}|${field}`;
    const computeImportSelection = (anchor, target) => {
      const startRow = Math.min(anchor.rowIndex, target.rowIndex);
      const endRow = Math.max(anchor.rowIndex, target.rowIndex);
      const startColIdx = Math.min(anchor.colIdx, target.colIdx);
      const endColIdx = Math.max(anchor.colIdx, target.colIdx);
      const result = new Set();
      for (let r = startRow; r <= endRow; r += 1) {
        for (let c = startColIdx; c <= endColIdx; c += 1) {
          result.add(importCellKey(r, importEditableFields[c]));
        }
      }
      return result;
    };

    const applyImportPaste = async (startRowIndex, startField, text) => {
      const rows = String(text || '').replace(/\r/g, '').split('\n').filter((line, idx, arr) => line !== '' || idx < arr.length - 1);
      if (!rows.length) return;
      const startColIndex = importEditableFields.indexOf(startField);
      if (startColIndex < 0) return;
      const batchUpdates = [];
      for (let r = 0; r < rows.length; r += 1) {
        const targetRow = importData[startRowIndex + r];
        if (!targetRow) break;
        const values = rows[r].split('\t');
        for (let c = 0; c < values.length; c += 1) {
          const field = importEditableFields[startColIndex + c];
          if (!field) break;
          batchUpdates.push({ row_key: targetRow._row_key, field, value: values[c] });
        }
      }
      if (batchUpdates.length && await updateImportCellsBatch(batchUpdates)) {
        addToast(`Import paste: ${batchUpdates.length} cells updated`, 'success');
      }
    };
    const fillImportRange = async (startRowIndex, startField, endRowIndex, endField) => {
      // Determine drag direction: vertical (same column, different row) or
      // horizontal (same row, different column). If the user dragged both,
      // pick the dominant axis so the fill is predictable.
      const startColIndex = importEditableFields.indexOf(startField);
      const endColIndex = importEditableFields.indexOf(endField);
      if (startColIndex < 0 || endColIndex < 0) return;
      const rowDelta = Math.abs(endRowIndex - startRowIndex);
      const colDelta = Math.abs(endColIndex - startColIndex);
      const source = importData[startRowIndex];
      if (!source) return;
      const batchUpdates = [];
      if (rowDelta >= colDelta) {
        // Vertical fill (same as before): copy startField's value down/up
        // across all rows from minRow to maxRow.
        const value = source[startField] ?? '';
        const minRow = Math.max(0, Math.min(startRowIndex, endRowIndex));
        const maxRow = Math.min(importData.length - 1, Math.max(startRowIndex, endRowIndex));
        for (let i = minRow; i <= maxRow; i += 1) {
          if (i === startRowIndex || !importData[i]?._row_key) continue;
          batchUpdates.push({ row_key: importData[i]._row_key, field: startField, value });
        }
      } else {
        // Horizontal fill: copy startField's value left/right across all
        // columns from minCol to maxCol on the SAME row. This includes
        // checklist columns — the user can drag a check across SAP INPUT,
        // BL/AWB, INVOICE, PL, HC, MSDS, COA, COO in one motion.
        const value = source[startField] ?? '';
        const minCol = Math.min(startColIndex, endColIndex);
        const maxCol = Math.max(startColIndex, endColIndex);
        for (let c = minCol; c <= maxCol; c += 1) {
          if (c === startColIndex) continue;
          const field = importEditableFields[c];
          if (!field) continue;
          batchUpdates.push({ row_key: source._row_key, field, value });
        }
      }
      if (batchUpdates.length && await updateImportCellsBatch(batchUpdates)) {
        addToast(`Import drag-fill: ${batchUpdates.length} cells updated`, 'success');
      }
    };
    const startImportFill = (event, rowIndex, field) => {
      event.preventDefault();
      event.stopPropagation();
      const getTargetCell = (evt) => document.elementFromPoint(evt.clientX, evt.clientY)?.closest('[data-import-cell="true"]');
      const updateRange = (evt) => {
        const target = getTargetCell(evt);
        if (!target) return;
        const endRowIndex = Number(target.getAttribute('data-row-index'));
        const endField = target.getAttribute('data-field');
        if (!Number.isFinite(endRowIndex) || !endField) return;
        // Determine direction based on which axis moved more — this gives
        // the user intuitive "drag down to fill down, drag right to fill
        // right" behavior without needing a separate handle.
        const startColIndex = importEditableFields.indexOf(field);
        const endColIndex = importEditableFields.indexOf(endField);
        const rowDelta = Math.abs(endRowIndex - rowIndex);
        const colDelta = Math.abs(endColIndex - startColIndex);
        if (rowDelta >= colDelta) {
          // Vertical: highlight same column, rows between start and end
          if (endField === field) {
            setImportFillRange({ startField: field, startRow: rowIndex, minRow: Math.min(rowIndex, endRowIndex), maxRow: Math.max(rowIndex, endRowIndex), direction: 'vertical' });
          }
        } else {
          // Horizontal: highlight same row, columns between start and end
          if (endRowIndex === rowIndex) {
            setImportFillRange({ startField: field, startRow: rowIndex, minCol: Math.min(startColIndex, endColIndex), maxCol: Math.max(startColIndex, endColIndex), direction: 'horizontal' });
          }
        }
      };
      const cleanup = () => {
        document.body.classList.remove('rfq-fill-dragging');
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        setImportFillRange(null);
      };
      const onMove = (moveEvent) => updateRange(moveEvent);
      const onUp = (upEvent) => {
        const target = getTargetCell(upEvent);
        if (!target) { cleanup(); return; }
        const endRowIndex = Number(target.getAttribute('data-row-index'));
        const endField = target.getAttribute('data-field');
        cleanup();
        if (!Number.isFinite(endRowIndex) || !endField) return;
        if (endField === field && endRowIndex === rowIndex) return; // no drag
        fillImportRange(rowIndex, field, endRowIndex, endField);
      };
      document.body.classList.add('rfq-fill-dragging');
      setImportFillRange({ startField: field, startRow: rowIndex, minRow: rowIndex, maxRow: rowIndex, direction: 'vertical' });
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    };

    return (
      <div className={`rounded-2xl overflow-hidden ${card}`}>
        <div className={`px-5 py-4 border-b ${darkMode ? 'border-gray-700' : 'border-gray-100'} flex flex-wrap items-center gap-3`}>
          <div className="flex items-center gap-2 min-w-0 flex-shrink-0">
            <Ship className="w-5 h-5 text-blue-500 flex-shrink-0" />
            <h2 className={`text-lg font-bold ${txt}`}>Import</h2>
            <span className={`text-sm ${txt2}`}>({fmtNum(importTotal)} records)</span>
            <span className={`text-xs ${txt2}`} title={`Vendor Import: ${fmtNum(importVendorCount)}`}>Last Copy: {fmtWibDateTime(importLastCopyAt)}</span>
          </div>
          {/* Action buttons — kept on a SINGLE line (no wrap) and right-aligned.
              `ml-auto` pushes this whole group to the RIGHT edge of the header,
              even when the title wraps onto its own line on narrow viewports.
              `flex-nowrap` keeps the buttons themselves on one line (with
              horizontal scroll if needed) so they never "menumpuk" (stack). */}
          <div className="flex flex-nowrap items-center justify-end gap-2 ml-auto overflow-x-auto max-w-full -mx-1 px-1">
            {/* Vendor Import — single dropdown combining Template + Upload.
                Keeps the header tidy and matches the "1 menu per logical
                action" pattern used in other tables (RFQ, Item Registration). */}
            <div className="relative flex-shrink-0">
              <button
                ref={importVendorDropdown.triggerRef}
                type="button"
                onClick={() => setImportVendorMenuOpen(v => !v)}
                onBlur={() => setTimeout(() => setImportVendorMenuOpen(false), 150)}
                className={`flex items-center gap-2 px-3 py-2.5 rounded-xl text-sm font-semibold shadow-sm whitespace-nowrap ${darkMode ? 'bg-gray-700 text-gray-100 hover:bg-gray-600' : 'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}
              >
                <FileSpreadsheet className="w-4 h-4"/>Vendor Import
                <svg className={`w-3 h-3 transition-transform ${importVendorMenuOpen ? 'rotate-180' : ''}`} viewBox="0 0 12 12" fill="none"><path d="M3 4.5L6 7.5L9 4.5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg>
              </button>
              {importVendorMenuOpen && (
                <div
                  style={importVendorDropdown.menuPos.style}
                  className={`rounded-xl border shadow-2xl overflow-hidden ${darkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200'}`}
                >
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); downloadBlob('/api/import/vendor-template', `Import_Vendor_Template_${new Date().toISOString().slice(0,10)}.xlsx`, 'Import Vendor Template'); setImportVendorMenuOpen(false); }}
                    className={`flex items-center gap-2 w-full px-3 py-2.5 text-left text-sm font-medium ${darkMode ? 'text-gray-100 hover:bg-gray-700' : 'text-gray-700 hover:bg-blue-50'}`}
                  >
                    <Download className="w-4 h-4"/>Template Vendor
                  </button>
                  <label className={`flex items-center gap-2 w-full px-3 py-2.5 text-sm font-medium cursor-pointer ${darkMode ? 'text-gray-100 hover:bg-gray-700' : 'text-gray-700 hover:bg-blue-50'}`}>
                    <FileSpreadsheet className="w-4 h-4"/>Upload Vendor Import
                    <input type="file" accept=".xlsx,.xls,.csv" multiple onChange={(e) => { handleVendorUpload(e); setImportVendorMenuOpen(false); }} className="hidden"/>
                  </label>
                </div>
              )}
            </div>
            {/* Copy Sheet — triggers a fresh sync from the live Import tracker sheet. */}
            <button onClick={() => fetchImportData(importPage, importPerPage, importAppliedSearch, true, importFilters, importReqDlvSort, importYupiPoSort)} className={`flex-shrink-0 flex items-center gap-2 px-3 py-2.5 rounded-xl text-sm font-semibold shadow-sm whitespace-nowrap ${darkMode ? 'bg-gray-700 text-gray-100 hover:bg-gray-600' : 'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}><RotateCcw className="w-4 h-4"/>Copy Sheet</button>
            {/* Hide/Show Checklist — toggle visibility of the 8 checklist columns (SAP INPUT, BL/AWB, …, COO). */}
            {checklistCount > 0 && (
              <button onClick={() => setShowImportChecklist(v => !v)} className={`flex-shrink-0 flex items-center gap-2 px-3 py-2.5 rounded-xl text-sm font-semibold shadow-sm whitespace-nowrap ${darkMode ? 'bg-gray-700 text-gray-100 hover:bg-gray-600' : 'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}>
                {showImportChecklist ? <EyeOff className="w-4 h-4"/> : <Eye className="w-4 h-4"/>}
                {showImportChecklist ? 'Hide Checklist' : 'Show Checklist'}
              </button>
            )}
            {/* Show/Hide Detail — toggles the per-item block (SO through
                PURCHASE AMOUNT). When hidden, the table collapses to one
                line per row (much narrower) for a quick overview without
                the long Spec / Item Name / Remark columns. */}
            {detailCount > 0 && (
              <button onClick={() => setShowImportDetail(v => !v)} className={`flex-shrink-0 flex items-center gap-2 px-3 py-2.5 rounded-xl text-sm font-semibold shadow-sm whitespace-nowrap ${darkMode ? 'bg-gray-700 text-gray-100 hover:bg-gray-600' : 'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}>
                {showImportDetail ? <EyeOff className="w-4 h-4"/> : <Eye className="w-4 h-4"/>}
                {showImportDetail ? 'Hide Detail' : 'Show Detail'}
              </button>
            )}
            {/* Print PO — opens the Serveone PO printing tool in a new tab. */}
            <a href="https://serveone.streamlit.app/" target="_blank" rel="noopener noreferrer" className={`flex-shrink-0 flex items-center gap-2 px-3 py-2.5 rounded-xl text-sm font-semibold shadow-sm whitespace-nowrap ${darkMode ? 'bg-gray-700 text-gray-100 hover:bg-gray-600' : 'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}>
              <Printer className="w-4 h-4"/>Print PO
            </a>
            {/* Download Excel — exports the current filtered+sorted view. */}
            <DownloadButton onClick={downloadImportExcel} className="flex-shrink-0 flex items-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-sm font-semibold shadow-sm whitespace-nowrap">
              <Download className="w-4 h-4"/>Download Excel
            </DownloadButton>
          </div>
        </div>

        {/* ── KPI section ───────────────────────────────────────────────────
            Five KPI cards across the top: Total PO, This Week Arrival (with
            sub-count of rows whose SAP INPUT is still unchecked), Sales Amount,
            PO Amount (IDR-converted by ETA-date FX rate), and Gross Margin
            (= Sales − PO IDR). All values are computed backend-side across
            the FULL filtered result set, so they don't change with pagination.
            Currency values use the same short-format (IDR + B/K/M/T suffix) as
            the Summary page KPIs for visual consistency. */}
        <div className={`px-5 py-3 border-b ${darkMode ? 'border-gray-700' : 'border-gray-100'}`}>
          <div className="grid gap-2" style={{ gridTemplateColumns: 'repeat(5, minmax(0, 1fr))' }}>
            {/* Total PO */}
            <div className={`p-3 rounded-xl ${darkMode ? 'bg-gray-800 border border-gray-700' : 'bg-gray-50 border border-gray-200'}`}>
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0">
                  <p className={`text-xs font-semibold truncate ${txt2}`} title="Total PO">Total PO</p>
                  <h3 className={`text-xl font-bold leading-tight ${darkMode ? 'text-gray-100' : 'text-gray-800'}`}>{fmtNum(importKpis.total_po)}</h3>
                  <p className={`text-[11px] leading-tight ${txt2}`}>&nbsp;</p>
                </div>
                <div className={`p-1.5 rounded-lg flex-shrink-0 ${darkMode ? 'bg-gray-700 text-gray-200' : 'bg-gray-100 text-gray-600'}`}>
                  <Ship className="w-3.5 h-3.5" />
                </div>
              </div>
            </div>
            {/* This Week Arrival — with sub-count of "no SAP input yet" */}
            <div className={`p-3 rounded-xl ${darkMode ? 'bg-gray-800 border border-gray-700' : 'bg-gray-50 border border-gray-200'}`}>
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0">
                  <p className={`text-xs font-semibold truncate ${txt2}`} title="This Week Arrival">This Week Arrival</p>
                  <h3 className={`text-xl font-bold leading-tight ${darkMode ? 'text-gray-100' : 'text-gray-800'}`}>{fmtNum(importKpis.this_week_arrival)}</h3>
                  <p className={`text-[11px] leading-tight ${importKpis.this_week_no_sap > 0 ? 'text-red-500 font-semibold' : txt2}`} title="Rows arriving this week with SAP INPUT not yet checked">
                    SAP not input: {fmtNum(importKpis.this_week_no_sap)}
                  </p>
                </div>
                <div className={`p-1.5 rounded-lg flex-shrink-0 ${darkMode ? 'bg-amber-900/40 text-amber-200' : 'bg-amber-100 text-amber-700'}`}>
                  <Calendar className="w-3.5 h-3.5" />
                </div>
              </div>
            </div>
            {/* Sales Amount */}
            <div className={`p-3 rounded-xl ${darkMode ? 'bg-gray-800 border border-gray-700' : 'bg-gray-50 border border-gray-200'}`}>
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0">
                  <p className={`text-xs font-semibold truncate ${txt2}`} title="Sales Amount">Sales Amount</p>
                  <h3 className={`text-xl font-bold leading-tight ${darkMode ? 'text-gray-100' : 'text-gray-800'}`}>{fmtCurShort(importKpis.sales_amount)}</h3>
                  <p className={`text-[11px] leading-tight ${txt2}`} title="Sum of AMOUNT column">{fmtCur(importKpis.sales_amount)}</p>
                </div>
                <div className={`p-1.5 rounded-lg flex-shrink-0 ${darkMode ? 'bg-emerald-900/40 text-emerald-200' : 'bg-emerald-100 text-emerald-700'}`}>
                  <DollarSign className="w-3.5 h-3.5" />
                </div>
              </div>
            </div>
            {/* PO Amount (IDR) */}
            <div className={`p-3 rounded-xl ${darkMode ? 'bg-gray-800 border border-gray-700' : 'bg-gray-50 border border-gray-200'}`}>
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0">
                  <p className={`text-xs font-semibold truncate ${txt2}`} title="PO Amount (IDR-converted by ETA date)">PO Amount (IDR)</p>
                  <h3 className={`text-xl font-bold leading-tight ${darkMode ? 'text-gray-100' : 'text-gray-800'}`}>{fmtCurShort(importKpis.po_amount_idr)}</h3>
                  <p className={`text-[11px] leading-tight ${txt2}`} title="Converted using ETA-date exchange rate">{fmtCur(importKpis.po_amount_idr)}</p>
                </div>
                <div className={`p-1.5 rounded-lg flex-shrink-0 ${darkMode ? 'bg-blue-900/40 text-blue-200' : 'bg-blue-100 text-blue-700'}`}>
                  <Wallet className="w-3.5 h-3.5" />
                </div>
              </div>
            </div>
            {/* Gross Margin = Sales − PO (IDR) */}
            <div className={`p-3 rounded-xl ${darkMode ? 'bg-gray-800 border border-gray-700' : 'bg-gray-50 border border-gray-200'}`}>
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0">
                  <p className={`text-xs font-semibold truncate ${txt2}`} title="Gross Margin = Sales Amount − PO Amount (IDR)">Gross Margin</p>
                  <h3 className={`text-xl font-bold leading-tight ${importKpis.gross_margin < 0 ? 'text-red-500' : (darkMode ? 'text-gray-100' : 'text-gray-800')}`}>{fmtCurShort(importKpis.gross_margin)}</h3>
                  <p className={`text-[11px] leading-tight ${txt2}`}>{fmtCur(importKpis.gross_margin)}</p>
                </div>
                <div className={`p-1.5 rounded-lg flex-shrink-0 ${importKpis.gross_margin < 0 ? (darkMode ? 'bg-red-900/50 text-red-200' : 'bg-red-100 text-red-700') : (darkMode ? 'bg-emerald-900/40 text-emerald-200' : 'bg-emerald-100 text-emerald-700')}`}>
                  <TrendingUp className="w-3.5 h-3.5" />
                </div>
              </div>
            </div>
          </div>
        </div>

        <FilterPanel darkMode={darkMode}>
          <div className="flex flex-wrap items-end gap-2">
            <div className="min-w-[110px] flex-shrink-0">
              <label className={`block text-xs font-medium mb-0.5 ${txt2}`}>↕ Req Dlv Date</label>
              <select
                className={`w-full h-10 px-2 py-2 rounded-lg text-sm border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                value={importReqDlvSort}
                onChange={e=>{ const nextSort=e.target.value; setImportReqDlvSort(nextSort); setImportPage(1); fetchImportData(1, importPerPage, importAppliedSearch, false, importFilters, nextSort, importYupiPoSort); }}
                title="Sort Req Dlv Date"
              >
                <option value="oldest">Oldest ↑</option>
                <option value="newest">Newest ↓</option>
              </select>
            </div>
            <div className="min-w-[110px] flex-shrink-0">
              <label className={`block text-xs font-medium mb-0.5 ${txt2}`}>↕ YUPI PO</label>
              <select
                className={`w-full h-10 px-2 py-2 rounded-lg text-sm border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                value={importYupiPoSort}
                onChange={e=>{ const nextSort=e.target.value; setImportYupiPoSort(nextSort); setImportPage(1); fetchImportData(1, importPerPage, importAppliedSearch, false, importFilters, importReqDlvSort, nextSort); }}
                title="Sort YUPI PO"
              >
                <option value="">Default</option>
                <option value="asc">A-Z ↑</option>
                <option value="desc">Z-A ↓</option>
              </select>
            </div>
            <div className="min-w-[180px] flex-1"><label className={`block text-xs font-semibold mb-1 ${txt2}`}>Search Import</label><input value={importSearch} onChange={e=>setImportSearch(e.target.value)} onKeyDown={e=>{ if(e.key==='Enter'){ setImportAppliedSearch(importSearch); setImportPage(1); fetchImportData(1, importPerPage, importSearch, false, importFilters, importReqDlvSort, importYupiPoSort); } }} placeholder="Search vendor, PO, item, BL, invoice..." className={`w-full h-10 px-3 py-2 rounded-xl text-sm border ${darkMode ? 'bg-gray-700 border-gray-600 text-white placeholder:text-gray-400' : 'bg-white border-gray-200 text-gray-800 placeholder:text-gray-400'}`}/></div>
            <div className="min-w-[130px] flex-shrink-0">
              <MultiSelect
                label="Status"
                options={importOptions.statuses || ['NEW', ...IMPORT_STATUS_OPTIONS]}
                selected={importFilters.statuses || []}
                onChange={v=>{ const next={...importFilters, statuses:v}; setImportFilters(next); setImportPage(1); fetchImportData(1, importPerPage, importAppliedSearch, false, next, importReqDlvSort, importYupiPoSort); }}
                darkMode={darkMode} txt2={txt2}
              />
            </div>
            <div className="min-w-[140px] flex-shrink-0">
              <MultiSelect
                label="Days Left"
                options={['Red (≤7 / overdue)', 'Yellow (8–29)', 'Green (≥30)', 'Today (0)']}
                selected={(Array.isArray(importFilters.daysLeft) ? importFilters.daysLeft : []).map(v => ({
                  red: 'Red (≤7 / overdue)',
                  yellow: 'Yellow (8–29)',
                  green: 'Green (≥30)',
                  today: 'Today (0)',
                }[v] || v)).filter(Boolean)}
                onChange={v => {
                  const mapping = {
                    'Red (≤7 / overdue)': 'red',
                    'Yellow (8–29)': 'yellow',
                    'Green (≥30)': 'green',
                    'Today (0)': 'today',
                  };
                  const next = { ...importFilters, daysLeft: v.map(label => mapping[label] || label) };
                  setImportFilters(next);
                  setImportPage(1);
                  fetchImportData(1, importPerPage, importAppliedSearch, false, next, importReqDlvSort, importYupiPoSort);
                }}
                darkMode={darkMode} txt2={txt2}
              />
            </div>
            <div className="min-w-[130px] flex-shrink-0">
              <MultiSelect label="YUPI PO" options={importOptions.yupi_po || []} selected={importFilters.yupi_po}
                onChange={v=>{ const next={...importFilters, yupi_po:v}; setImportFilters(next); setImportPage(1); fetchImportData(1, importPerPage, importAppliedSearch, false, next, importReqDlvSort, importYupiPoSort); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-[130px] flex-shrink-0">
              <MultiSelect label="Vendor" options={importOptions.vendors || []} selected={importFilters.vendors}
                onChange={v=>{ const next={...importFilters, vendors:v}; setImportFilters(next); setImportPage(1); fetchImportData(1, importPerPage, importAppliedSearch, false, next, importReqDlvSort, importYupiPoSort); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <button onClick={()=>{ setImportAppliedSearch(importSearch); setImportPage(1); fetchImportData(1, importPerPage, importSearch, false, importFilters, importReqDlvSort, importYupiPoSort); }} className="h-10 px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold shadow-sm flex-shrink-0">Search</button>
            <button onClick={()=>{ const next={ yupi_po: [], vendors: [], statuses: [], daysLeft: [] }; const nextSort='oldest'; const nextYupiSort=''; setImportSearch(''); setImportAppliedSearch(''); setImportFilters(next); setImportReqDlvSort(nextSort); setImportYupiPoSort(nextYupiSort); setImportPage(1); fetchImportData(1, importPerPage, '', false, next, nextSort, nextYupiSort); }} className={`h-10 px-3 py-2 rounded-lg text-sm font-medium shadow-sm flex items-center justify-center whitespace-nowrap flex-shrink-0 ${darkMode ? 'bg-gray-500 text-gray-100 hover:bg-gray-400' : 'bg-gray-400 text-white hover:bg-gray-500'}`}>Clear</button>
          </div>
        </FilterPanel>

        <DataTableScroll darkMode={darkMode}>
          <table className={`freeze-table-import table-fixed text-xs border-collapse border ${showImportDetail ? '' : 'import-detail-hidden'}`} style={{ width: `${tableWidth}px`, minWidth: `${tableWidth}px` }}>
            <colgroup>{visibleColumns.map(col => <col key={col.field} style={{ width: `${colWidth(col)}px` }} />)}</colgroup>
            <thead className={tblHd}>
              <tr>{visibleColumns.map((col, index) => (
                <th key={col.field} data-col-index={index + 1} className={`px-2 py-2 h-10 text-center align-middle font-bold border-r whitespace-pre-line leading-tight ${darkMode ? 'border-gray-700 text-gray-200' : 'border-gray-200 text-slate-700'}`} title={col.sheet_col ? `${col.sheet_col} - ${col.label}` : col.label}>{renderFreezeHeader('import', index + 1, col.label)}</th>
              ))}</tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {importData.length === 0 ? (
                <tr><td colSpan={Math.max(1, visibleColumns.length)} className={`px-4 py-12 text-center ${txt2}`}><Ship className="w-10 h-10 mx-auto mb-2 opacity-40"/>No import data</td></tr>
              ) : importData.map((row, rowIndex) => {
                const hasReschedule = String(row.reschedule || '').trim();
                // Zebra striping by GROUP: even groups → white, odd groups →
                // slate-100 (light) / gray-900/40 (dark). Rows in the same
                // merged group share the same shade. Reschedule rows override
                // with amber tint regardless of zebra.
                const groupIdx = importGroupIndexFor[rowIndex] || 1;
                const isEvenGroup = groupIdx % 2 === 0;
                const zebraClass = hasReschedule
                  ? (darkMode ? 'bg-amber-900/25 hover:bg-amber-900/35' : 'bg-amber-50 hover:bg-amber-100/70')
                  : isEvenGroup
                    ? (darkMode ? 'bg-gray-800/40 hover:bg-gray-700/50' : 'bg-white hover:bg-slate-50')
                    : (darkMode ? 'bg-gray-900/50 hover:bg-gray-800/60' : 'bg-slate-100 hover:bg-slate-200/70');
                return <tr key={row._row_key} className={`${trHov} ${zebraClass}`}>{visibleColumns.map((col, colIdx) => {
                  const isGroupCol = isImportGroupColumn(col);
                  const groupStartRow = importGroupStartIndexFor[rowIndex];
                  // Group columns only render their <td> on the group's first
                  // row, with rowSpan covering the rest — the merged cell.
                  // Every other row in the group skips rendering this column
                  // entirely (no empty <td>, exactly like the merged Excel cell).
                  if (isGroupCol && groupStartRow !== rowIndex) return null;
                  const ownerRow = isGroupCol ? importData[groupStartRow] : row;
                  const ownerRowIndex = isGroupCol ? groupStartRow : rowIndex;
                  const rowSpan = isGroupCol ? (importRowSpans[groupStartRow] || 1) : 1;
                  const formula = isImportFormulaColumn(col);
                  const selected = importSelectedCell?.rowKey === ownerRow._row_key && importSelectedCell?.field === col.field;
                  // Multi-select highlight: if Shift+click selected a range,
                  // highlight every cell in that range. Falls back to the
                  // single-cell `selected` outline when no range is active.
                  const cellKey = importCellKey(ownerRowIndex, col.field);
                  const inMultiSelection = importSelectedCells?.has(cellKey);
                  // Fill highlight: support both vertical (down/up, same
                  // column) and horizontal (left/right, same row) drag-fill.
                  // colIdx is now passed from the map callback directly (no need to findIndex)
                  // const colIdx = visibleColumns.findIndex(c => c.field === col.field);
                  let fillHighlighted = false;
                  if (importFillRange?.direction === 'vertical') {
                    fillHighlighted = importFillRange.startField === col.field && ownerRowIndex >= importFillRange.minRow && ownerRowIndex <= importFillRange.maxRow && ownerRowIndex !== importFillRange.startRow;
                  } else if (importFillRange?.direction === 'horizontal') {
                    fillHighlighted = importFillRange.startRow === ownerRowIndex && colIdx >= importFillRange.minCol && colIdx <= importFillRange.maxCol && col.field !== importFillRange.startField;
                  }
                  const editingCellNow = importEditingCell === `${ownerRow._row_key}:${col.field}`;
                  // directEdit: cell types that DON'T auto-open a text editor
                  // on click — they render their own interactive control
                  // (dropdown/checkbox/date picker) inside the cell.
                  const directEdit = !(col.field === 'status' || col.type === 'status'
                    || isImportChecklistColumn(col)
                    || col.field === 'payment' || col.payment_dropdown
                    || col.field === 'payment_date' || col.payment_date);
                  return <td
                    key={col.field}
                    rowSpan={rowSpan > 1 ? rowSpan : undefined}
                    data-col-index={colIdx + 1}
                    data-import-cell="true"
                    data-row-index={ownerRowIndex}
                    data-field={col.field}
                    tabIndex={0}
                    onFocus={() => setImportSelectedCell({ rowKey: ownerRow._row_key, field: col.field })}
                    onClick={(e) => {
                      setImportSelectedCell({ rowKey: ownerRow._row_key, field: col.field });
                      // Shift+click multi-select (Excel-like). Click without
                      // Shift resets to single cell + new anchor. Shift+click
                      // extends from the existing anchor to the clicked cell,
                      // forming a rectangle that can span rows AND columns.
                      const clickedColIdx = importEditableFields.indexOf(col.field);
                      if (e.shiftKey && importSelectionAnchor && clickedColIdx >= 0) {
                        const newSelection = computeImportSelection(importSelectionAnchor, { rowIndex: ownerRowIndex, colIdx: clickedColIdx });
                        setImportSelectedCells(newSelection);
                      } else {
                        setImportSelectionAnchor({ rowIndex: ownerRowIndex, colIdx: clickedColIdx });
                        setImportSelectedCells(new Set([cellKey]));
                        if (directEdit && !editingCellNow && !e.target.closest('a,input,textarea,select,button')) startImportEdit(ownerRow, col);
                      }
                    }}
                    onPaste={e => { e.preventDefault(); applyImportPaste(ownerRowIndex, col.field, e.clipboardData.getData('text/plain')); }}
                    className={`group relative h-8 max-h-8 ${editingCellNow ? 'is-editing p-0' : 'px-2 py-1'} align-middle border-r focus:outline-none cursor-pointer ${hasReschedule ? (darkMode ? 'bg-amber-900/20' : 'bg-amber-50') : ''} ${darkMode ? 'border-gray-700' : 'border-gray-200'} ${col.blue_text ? (darkMode ? 'text-blue-300' : 'text-blue-600') : ''} ${col.bold_text ? 'font-bold' : ''} ${editingCellNow ? 'outline outline-2 outline-blue-500 outline-offset-[-2px]' : fillHighlighted ? 'outline outline-2 outline-blue-300 outline-offset-[-2px]' : inMultiSelection ? 'outline outline-2 outline-blue-500 outline-offset-[-2px] bg-blue-50/50' : selected ? 'outline outline-2 outline-blue-500 outline-offset-[-2px]' : 'hover:outline hover:outline-2 hover:outline-blue-400 hover:outline-offset-[-2px]'} ${col.field === 'days_left' ? 'text-center' : ''} ${txt2}`}
                  >
                    {renderImportCell(ownerRow, col)}
                    {/* Fill handle — bottom-right corner. Drag in any direction:
                        down to fill the same column, right to fill the same row.
                        The drag logic auto-detects the dominant axis. */}
                    <button type="button" aria-label="Fill handle" title="Drag to copy value (down or right)" onClick={e => e.stopPropagation()} onMouseDown={e => startImportFill(e, ownerRowIndex, col.field)} className="rfq-fill-handle absolute bottom-0 right-0 h-3 w-3 translate-x-1/2 translate-y-1/2 border border-blue-600 bg-blue-600 opacity-0 group-hover:opacity-100 focus:opacity-100" />
                  </td>;
                })}</tr>;
              })}
            </tbody>
          </table>
        </DataTableScroll>

        <PagePagination darkMode={darkMode} txt2={txt2} page={importPage} totalPages={totalPages} total={importTotal} perPage={importPerPage} onPageChange={(p)=>{ setImportPage(p); fetchImportData(p, importPerPage, importAppliedSearch, false, importFilters, importReqDlvSort, importYupiPoSort); }} onPerPageChange={(next)=>{ setImportPerPage(next); setImportPage(1); fetchImportData(1, next, importAppliedSearch, false, importFilters, importReqDlvSort, importYupiPoSort); }} />
      </div>
    );
  };
  const renderItemRegistration = () => {
    const fmtDateShort = (d) => {
      if (!d) return '-';
      try { return String(d).slice(0, 10); } catch { return d; }
    };
    const baseColumns = [
      ['Proc. Status', 'proc_status'], ['Req. Date', 'req_date'], ['Client Nm.', 'client_name'], ['Category', 'category'], ['PIC', 'pic'],
      ['Req. No', 'req_no'], ['Prod. ID', 'prod_id'], ['Prod. Nm.', 'prod_name'],
      ['Spec.', 'spec'], ['Mfr. Nm.', 'mfr_name'], ['Unit', 'odr_unit'],
      ['Prod. Price', 'prod_price'], ['Curr.', 'curr']
    ];
    const columns = [...baseColumns, ['Remarks', 'remarks']];
    const statusClass = (status) => {
      const s = String(status || '').toLowerCase();
      if (s.includes('complete')) return 'bg-green-100 text-green-700 border-green-200';
      if (s.includes('waiting')) return 'bg-slate-100 text-slate-700 border-slate-200';
      if (s.includes('receive')) return 'bg-sky-100 text-sky-700 border-sky-200';
      if (s.includes('process') || s.includes('proc')) return 'bg-blue-100 text-blue-700 border-blue-200';
      if (s.includes('reject') || s.includes('cancel')) return 'bg-red-100 text-red-700 border-red-200';
      return 'bg-gray-100 text-gray-700 border-gray-200';
    };
    const visibleItemRegPicKpis = itemRegMissingPicKpis.filter(row => row.pic && row.pic.toLowerCase() !== 'unassigned' && row.pic.trim() !== '');
    const totalMissingProdId = visibleItemRegPicKpis.reduce((sum, row) => sum + (Number(row.count) || 0), 0);
    const itemRegPicKpis = [
      { pic: 'Total Pending Regist.', count: totalMissingProdId, sub: 'Items Pending', isTotal: true },
      ...visibleItemRegPicKpis.map(row => ({ pic: row.pic, count: row.count, sub: 'Items Pending' }))
    ];
    const itemRegKpiCols = Math.max(1, itemRegPicKpis.length);
    const colWidth = (key) => ({
      proc_status: 150, req_date: 110, client_name: 180, category: 170, pic: 90, req_no: 150, prod_id: 110,
      prod_name: 240, spec: 220, mfr_name: 150, odr_unit: 80,
      prod_price: 120, curr: 70, remarks: 560
    }[key] || 140);
    const colStyle = (key) => {
      const width = `${colWidth(key)}px`;
      return { width, minWidth: width, maxWidth: width };
    };
    const itemRegTableWidth = columns.reduce((sum, [, key]) => sum + colWidth(key), 0);

    return (
      <div className={`rounded-2xl overflow-hidden ${card}`}>
        <div className={`px-5 py-4 border-b ${darkMode?'border-gray-700':'border-gray-100'} flex flex-wrap justify-between items-center gap-3`}>
          <div className="flex items-center gap-2 min-w-0">
            <Wrench className="w-5 h-5 text-blue-500 flex-shrink-0"/>
            <h2 className={`text-lg font-bold ${txt}`}>Item Registration</h2>
            <span className={`text-sm ${txt2}`}>({fmtNum(itemRegTotal)} records)</span>
            {itemRegLastUpdated && <span className={`text-xs ${txt2}`}>Last update: {fmtDate(itemRegLastUpdated)}</span>}
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button onClick={downloadItemRegistrationTemplate} className={`flex items-center gap-2 px-3 py-2.5 rounded-xl text-sm font-semibold shadow-sm ${darkMode?'bg-gray-700 text-gray-100 hover:bg-gray-600':'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}>
              <Download className="w-4 h-4"/>Template
            </button>
            <label className="flex items-center gap-2 px-3 py-2.5 bg-slate-600 hover:bg-slate-700 text-white rounded-xl text-sm font-semibold shadow-sm cursor-pointer">
              <FileSpreadsheet className="w-4 h-4"/>Batch Upload
              <input type="file" accept=".xlsx,.xls" onChange={handleItemRegistrationBatchUpload} className="hidden"/>
            </label>
            <DownloadButton onClick={downloadItemRegistrationExcel} className="flex items-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-sm font-semibold shadow-sm">
              <Download className="w-4 h-4"/>Download Excel
            </DownloadButton>
          </div>
        </div>

        <div className={`px-5 py-3 border-b ${darkMode?'border-gray-700':'border-gray-100'}`}>
          <div className="grid grid-flow-col gap-2" style={{ gridTemplateColumns: `repeat(${itemRegKpiCols}, minmax(0, 1fr))` }}>
            {itemRegPicKpis.map((row) => {
              const activePic = !row.isTotal && itemRegPicHighlight === row.pic;
              const applyItemRegPicFilter = () => {
                const nextHighlight = row.isTotal || activePic ? '' : row.pic;
                const nextFilters = { ...itemRegFilters, pics: nextHighlight ? [nextHighlight] : [] };
                setItemRegPicHighlight(nextHighlight);
                setItemRegFilters(nextFilters);
                setItemRegPage(1);
                fetchItemRegistration(1, itemRegPerPage, itemRegAppliedSearch, nextFilters, nextHighlight);
              };
              return (
                <button key={row.pic} type="button" onClick={applyItemRegPicFilter} className={`min-w-0 p-3 rounded-xl text-left transition-all ${activePic ? (darkMode ? 'bg-amber-900/30 border border-amber-500 ring-2 ring-amber-400' : 'bg-amber-50 border border-amber-300 ring-2 ring-amber-200') : row.isTotal ? (darkMode ? 'bg-gray-800 border border-gray-700' : 'bg-gray-50 border border-gray-200') : card} ${row.isTotal ? 'hover:border-slate-300' : 'hover:border-amber-300'}`}>
                  <div className="flex items-start justify-between gap-2">
                    <div className="min-w-0">
                      <p className={`text-xs font-semibold truncate ${activePic ? 'text-amber-700' : row.isTotal ? (darkMode ? 'text-gray-200' : 'text-gray-700') : txt2}`} title={row.pic}>{row.pic}</p>
                      <h3 className={`text-xl font-bold leading-tight ${activePic ? 'text-amber-700' : row.isTotal ? (darkMode ? 'text-gray-100' : 'text-gray-800') : kpiValue}`}>{fmtNum(row.count)}</h3>
                      <p className={`text-[11px] leading-tight whitespace-nowrap ${txt2}`}>{row.sub}</p>
                    </div>
                    <div className={`p-1.5 rounded-lg flex-shrink-0 ${activePic ? 'bg-amber-100 text-amber-700' : row.isTotal ? (darkMode ? 'bg-gray-700 text-gray-200' : 'bg-gray-100 text-gray-600') : neutralIcon}`}>
                      <Package className="w-3.5 h-3.5" />
                    </div>
                  </div>
                </button>
              );
            })}
          </div>
        </div>

        <FilterPanel darkMode={darkMode}>
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-4 2xl:grid-cols-[170px_repeat(5,minmax(150px,1fr))_120px] items-end">
            <div className="min-w-0">
              <label className={`block text-xs font-semibold mb-1 ${txt2}`}>Search Req No.</label>
              <SearchInput
                key={`item-req-no-${itemRegSearch.join('|')}`}
                placeholder={'REQ001\nREQ002'}
                label="Req No."
                darkMode={darkMode}
                txt2={txt2}
                onSearch={(nums) => {
                  setItemRegSearch(nums);
                  setItemRegAppliedSearch(nums);
                  setItemRegPage(1);
                  fetchItemRegistration(1, itemRegPerPage, nums, itemRegFilters);
                }}
              />
            </div>
            <div className="min-w-0">
              <MultiSelect label="Proc. Status" options={itemRegOptions.proc_statuses} selected={itemRegFilters.proc_statuses}
                onChange={v=>{ const next={...itemRegFilters, proc_statuses:v}; setItemRegFilters(next); setItemRegPage(1); fetchItemRegistration(1,itemRegPerPage,itemRegAppliedSearch,next); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="Client Name" options={itemRegOptions.clients} selected={itemRegFilters.clients}
                onChange={v=>{ const next={...itemRegFilters, clients:v}; setItemRegFilters(next); setItemRegPage(1); fetchItemRegistration(1,itemRegPerPage,itemRegAppliedSearch,next); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="Category" options={itemRegOptions.categories} selected={itemRegFilters.categories}
                onChange={v=>{ const next={...itemRegFilters, categories:v}; setItemRegFilters(next); setItemRegPage(1); fetchItemRegistration(1,itemRegPerPage,itemRegAppliedSearch,next); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="PIC" options={itemRegOptions.pics} selected={itemRegFilters.pics}
                onChange={v=>{ const next={...itemRegFilters, pics:v}; setItemRegPicHighlight(''); setItemRegFilters(next); setItemRegPage(1); fetchItemRegistration(1,itemRegPerPage,itemRegAppliedSearch,next,''); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="Mfr. Nm." options={itemRegOptions.mfr_names} selected={itemRegFilters.mfr_names}
                onChange={v=>{ const next={...itemRegFilters, mfr_names:v}; setItemRegFilters(next); setItemRegPage(1); fetchItemRegistration(1,itemRegPerPage,itemRegAppliedSearch,next); }} darkMode={darkMode} txt2={txt2}/>
            </div>
            <button onClick={() => { const next={ clients: [], categories: [], pics: [], proc_statuses: [], mfr_names: [] }; setItemRegSearch([]); setItemRegAppliedSearch([]); setItemRegPicHighlight(''); setItemRegFilters(next); setItemRegPage(1); fetchItemRegistration(1, itemRegPerPage, [], next, ''); }}
              className={`w-full h-10 px-3 py-2 rounded-lg text-sm font-medium shadow-sm flex items-center justify-center whitespace-nowrap ${darkMode?'bg-gray-500 text-gray-100 hover:bg-gray-400':'bg-gray-400 text-white hover:bg-gray-500'}`}>Clear</button>
          </div>
        </FilterPanel>

        <DataTableScroll darkMode={darkMode}>
          <table className="freeze-table-item-registration table-fixed text-xs" style={{ width: `${itemRegTableWidth}px`, minWidth: `${itemRegTableWidth}px` }}>
            <colgroup>{columns.map(([, key]) => <col key={key} style={colStyle(key)}/>)}</colgroup>
            <thead className={tblHd}><tr>{columns.map(([label], index) => <th key={label} data-col-index={index + 1} className={`px-2 py-2 text-center font-bold whitespace-nowrap ${txt2}`}>{renderFreezeHeader('item-registration', index + 1, label)}</th>)}</tr></thead>
            <tbody className={`divide-y ${tblDv}`}>
              {itemRegData.length === 0 ? <tr><td colSpan={columns.length} className={`px-4 py-12 text-center ${txt2}`}><Wrench className="w-10 h-10 mx-auto mb-2 opacity-40"/>No Item Registration data</td></tr>
              : itemRegData.map(row => {
                return <tr key={row.id} className={`${trHov} transition-colors`}>
                {columns.map(([label, key], colIdx) => {
                  const value = key === 'prod_price' ? fmtNum(row[key]) : key === 'req_date' ? fmtDateShort(row[key]) : (row[key] || '-');
                  if (key === 'proc_status') return <td key={key} data-col-index={colIdx + 1} className="px-2 py-2"><span className={`inline-flex max-w-full items-center px-2 py-0.5 rounded-full border text-[11px] font-semibold leading-snug truncate ${statusClass(row[key])}`}>{value}</span></td>;
                  if (key === 'pic') {
                    const c = getPicColor(row.pic);
                    return <td key={key} data-col-index={colIdx + 1} className="px-2 py-2 text-center truncate">{row.pic ? <span className={`inline-flex max-w-full truncate px-2 py-0.5 rounded-full text-[11px] font-semibold ${c ? `${c.bg} ${c.text}` : 'bg-gray-100 text-gray-700'}`}>{row.pic}</span> : <span className={txt2}>-</span>}</td>;
                  }
                  if (key === 'remarks') return <td key={key} data-col-index={colIdx + 1} className="px-2 py-2 truncate" title={row.remarks}>{editingCell?.id===row.id && editingCell.field==='item_remarks' ? (
                    <input type="text" defaultValue={row.remarks}
                      className={`w-full px-2 py-1 rounded text-xs border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                      onChange={e=>setEditValue(e.target.value)}
                      onBlur={()=>updateItemRegistrationCell(row.id,'remarks',editValue)}
                      onKeyDown={e=>{ if(e.key==='Enter') updateItemRegistrationCell(row.id,'remarks',editValue); if(e.key==='Escape') setEditingCell(null); }}
                      autoFocus/>
                  ) : (
                    <span className="cursor-pointer text-blue-600 hover:underline" onClick={()=>{setEditingCell({id:row.id,field:'item_remarks'});setEditValue(row.remarks||'');}}>{row.remarks||'Add'}</span>
                  )}</td>;
                  return <td key={key} data-col-index={colIdx + 1} className={`px-2 py-2 ${['req_no','prod_name'].includes(key) ? '' : 'truncate'} ${key === 'prod_price' ? `text-right font-semibold ${kpiValue}` : txt2} ${['req_date','req_no','prod_id','prod_name','odr_unit','curr'].includes(key) ? 'whitespace-nowrap' : ''}`} title={row[key]}>{value}</td>;
                })}
              </tr>;})}
            </tbody>
          </table>
        </DataTableScroll>

        <PagePagination
          darkMode={darkMode}
          txt2={txt2}
          page={itemRegPage}
          totalPages={itemRegTotalPages}
          total={itemRegTotal}
          perPage={itemRegPerPage}
          onPageChange={(p) => { setItemRegPage(p); fetchItemRegistration(p, itemRegPerPage, itemRegAppliedSearch, itemRegFilters); }}
          onPerPageChange={(next) => { setItemRegPerPage(next); setItemRegPage(1); fetchItemRegistration(1, next, itemRegAppliedSearch, itemRegFilters); }}
        />
      </div>
    );
  };

  const renderPendingDeliverySummary = () => {
    const agingTotals = agingData.reduce((acc, v) => ({
      less_30: acc.less_30 + (v.less_30 || 0),
      days_30_90: acc.days_30_90 + (v.days_30_90 || 0),
      days_90_180: acc.days_90_180 + (v.days_90_180 || 0),
      more_180: acc.more_180 + (v.more_180 || 0),
    }), { less_30:0, days_30_90:0, days_90_180:0, more_180:0 });
    const pendingMonthly = (stats?.monthly_trend || []).map(m => ({...m, amount_idr: (m.amount || 0) * 1_000_000}));
    const statusMonths = stats?.status_months || [];
    const statusMonthlyRows = stats?.so_status_monthly || [];
    const itemRegProcStatus = stats?.item_registration_proc_status || [];
    const itemRegClients = stats?.item_registration_clients || [];
    const pendingVendors = stats?.top_vendors || [];
    const sumRows = (rows, key) => (rows || []).reduce((sum, row) => sum + (Number(row?.[key]) || 0), 0);
    const pendingMonthlyTotal = sumRows(pendingMonthly, 'so_count');
    const pendingMonthlyAmount = sumRows(pendingMonthly, 'amount') * 1_000_000;
    const pendingVendorTotal = sumRows(pendingVendors, 'so_count');
    const pendingVendorAmount = sumRows(pendingVendors, 'total_amount');
    const pendingStatusTotal = sumRows(statusMonthlyRows, 'total');
    const pendingStatusAmount = sumRows(statusMonthlyRows, 'amount');
    const itemRegCategoryChart = (rows, title) => {
      const sorted = [...(rows || [])]
        .map(r => ({ name: r.name || '-', value: Number(r.value || 0) }))
        .filter(r => r.value > 0)
        .sort((a, b) => b.value - a.value);
      const top = sorted.slice(0, 5);
      const rest = sorted.slice(5);
      const restValue = rest.reduce((sum, r) => sum + r.value, 0);
      const data = restValue > 0 ? [...top, { name: `Others (${rest.length})`, value: restValue, isOthers: true }] : top;
      const total = sumRows(data, 'value');
      const pieColors = ['#2563EB', '#14B8A6', '#DC2626', '#60A5FA', '#5EEAD4', '#94A3B8'];
      return (
        <div className={`p-4 rounded-2xl ${card}`}>
          <h3 className={`text-sm font-bold ${txt}`}>{title}</h3>
          <p className={`text-xs mb-2 ${txt2}`}>Total: {fmtNum(total)} Req. No</p>
          {data.length === 0 ? (
            <div className={`h-[180px] flex items-center justify-center text-sm ${txt2}`}>No Item Registration data</div>
          ) : (
            <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-center md:gap-8 xl:gap-10">
              <div className="h-[190px] w-full min-w-0 md:w-[220px] md:flex-none">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart margin={{ top: 4, right: 4, bottom: 4, left: 4 }}>
                    <Pie data={data} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={38} outerRadius={76} labelLine={false} label={renderPctLabel} isAnimationActive={false}>
                      {data.map((_, i) => <Cell key={i} fill={pieColors[i % pieColors.length]} />)}
                    </Pie>
                    <Tooltip formatter={(v, n) => [`${fmtNum(v)} Req. No`, n]} contentStyle={{background:darkMode?'#1F2937':'#fff',border:'none',borderRadius:8,fontSize:12}}/>
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div className="w-full md:w-[250px] space-y-2">
                {data.map((item, i) => (
                  <div key={item.name} className="flex items-start gap-2 text-xs">
                    <span className="mt-1 h-2.5 w-2.5 rounded-full flex-shrink-0" style={{ backgroundColor: pieColors[i % pieColors.length] }} />
                    <div className="min-w-0 flex-1">
                      <p className={`font-semibold leading-snug break-words ${txt}`} title={item.name}>{item.name}</p>
                      <p className={`leading-tight ${txt2}`}>{fmtNum(item.value)} | {total ? ((item.value / total) * 100).toFixed(1) : '0.0'}%</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      );
    };
    const agingCards = [
      { label:'0-30 Days', value: agingTotals.less_30, color:'#14B8A6' },
      { label:'30-90 Days', value: agingTotals.days_30_90, color:'#2563EB' },
      { label:'90-180 Days', value: agingTotals.days_90_180, color:'#FCA5A5' },
      { label:'180+ Days', value: agingTotals.more_180, color:'#DC2626' },
    ];
    return (
      <div className="space-y-4 mb-5">
        <h3 className={`text-base font-bold ${txt} flex items-center gap-2`}>
          <Calendar className="w-5 h-5 text-blue-600"/> Pending Delivery Aging
        </h3>
        <div className="grid grid-cols-2 xl:grid-cols-4 gap-3">
          {agingCards.map(cardItem => (
            <button key={cardItem.label} onClick={() => openPendingDeliveryWithAging(cardItem.label.replace(' Days','').replace('+','+'))}
              className={`text-left p-4 rounded-2xl border transition-all ${card} hover:border-slate-300`}>
              <p className={`text-xs font-semibold ${txt2}`}>{cardItem.label}</p>
              <div className="mt-1 flex items-end justify-between gap-3">
                <h3 className={`text-2xl font-bold ${kpiValue}`}>{fmtNum(cardItem.value)}</h3>
                <span className="h-3 w-3 rounded-full" style={{backgroundColor: cardItem.color}} />
              </div>
            </button>
          ))}
        </div>
        <div className="grid grid-cols-1 xl:grid-cols-3 gap-4">
          <div className={`xl:col-span-2 p-5 rounded-2xl ${card}`}>
            <h3 className={`text-base font-bold ${txt}`}>Total Pending per Month</h3>
            <p className={`text-xs mb-4 ${txt2}`}>Total: {fmtNum(pendingMonthlyTotal)} SO | {fmtCurShort(pendingMonthlyAmount)}</p>
            <ResponsiveContainer width="100%" height={280}>
              <ComposedChart data={pendingMonthly}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={darkMode?'#374151':'#E5E7EB'}/>
                <XAxis dataKey="month" stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={10}/>
                <YAxis yAxisId="left" stroke="#2563EB" fontSize={10}/>
                <YAxis yAxisId="right" orientation="right" stroke="#14B8A6" fontSize={10} tickFormatter={(v)=>fmtCurShort(v*1_000_000)}/>
                <Tooltip formatter={(v,name)=>name==='Pending SO'?[fmtNum(v),name]:[fmtCur(v*1_000_000),name]} contentStyle={{background:darkMode?'#1F2937':'#fff',border:'none',borderRadius:8,fontSize:12}}/>
                <Legend wrapperStyle={{fontSize:12}}/>
                <Bar yAxisId="left" dataKey="so_count" name="Pending SO" fill="#2563EB" radius={[5,5,0,0]} isAnimationActive={false}/>
                <Line yAxisId="right" type="monotone" dataKey="amount" name="Total Sales Amount" stroke="#14B8A6" strokeOpacity={0.45} strokeWidth={1.5} dot={{r:2.5,fill:'#14B8A6',opacity:0.45}} activeDot={{r:4,opacity:0.65}} isAnimationActive={false}/>
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <div className={`p-5 rounded-2xl ${card}`}>
            <h3 className={`text-base font-bold ${txt}`}>Top 5 Pending Vendor</h3>
            <p className={`text-xs mb-4 ${txt2}`}>Total: {fmtNum(pendingVendorTotal)} SO | {fmtCurShort(pendingVendorAmount)}</p>
            <div className="space-y-2">
              {pendingVendors.map((v,i)=>(
                <div key={v.vendor} className={`p-2.5 rounded-lg ${darkMode?'bg-gray-700':'bg-gray-100'}`}>
                  <div className="flex items-center justify-between gap-3">
                    <div className="flex items-center gap-2 min-w-0">
                      <span className={`text-xs font-bold ${darkMode?'text-gray-400':'text-gray-500'}`}>#{i+1}</span>
                      <p className={`text-sm font-semibold truncate ${darkMode?'text-white':'text-gray-900'}`} title={v.vendor}>{v.vendor}</p>
                    </div>
                    <span className={`text-sm font-bold whitespace-nowrap ${darkMode?'text-gray-100':'text-gray-900'}`}>{fmtCurShort(v.total_amount)}</span>
                  </div>
                  <p className={`text-xs mt-0.5 ${txt2}`}>{fmtNum(v.so_count)} SO</p>
                </div>
              ))}
            </div>
          </div>
        </div>
        <div className={`p-4 rounded-2xl ${card}`}>
          <h3 className={`text-base font-bold ${txt}`}>Pending Delivery - Status Distribution</h3>
          <p className={`text-xs mb-2 ${txt2}`}>Total: {fmtNum(pendingStatusTotal)} SO | {fmtCurShort(pendingStatusAmount)}</p>
          {(() => {
            const monthLabel = (m) => {
              try {
                const [y, mo] = m.split('-');
                return format(new Date(parseInt(y), parseInt(mo) - 1, 1), 'MMM yy');
              } catch {
                return m;
              }
            };
            const maxCell = Math.max(
              1,
              ...statusMonthlyRows.flatMap(s => statusMonths.map(m => s.monthly?.[m] || 0))
            );
            const totalByMonth = statusMonths.reduce((acc, m) => {
              acc[m] = statusMonthlyRows.reduce((sum, s) => sum + (s.monthly?.[m] || 0), 0);
              return acc;
            }, {});
            const grandTotal = statusMonthlyRows.reduce((sum, s) => sum + (s.total || 0), 0);
            const grandAmount = statusMonthlyRows.reduce((sum, s) => sum + (s.amount || 0), 0);
            const heatStyle = (value) => {
              if (!value) {
                return {
                  backgroundColor: darkMode ? 'rgba(59,130,246,0.08)' : 'rgba(219,234,254,0.35)',
                  color: darkMode ? '#64748B' : '#94A3B8'
                };
              }
              const logVal = Math.log1p(value);
              const logMax = Math.log1p(maxCell);
              const intensity = Math.max(0.08, Math.min(1, Math.pow(logVal / logMax, 0.9)));
              const lightness = darkMode ? 24 + intensity * 30 : 97 - intensity * 48;
              return {
                backgroundColor: `hsl(217, 88%, ${lightness}%)`,
                color: intensity > 0.55 ? '#FFFFFF' : darkMode ? '#DBEAFE' : '#1E3A8A'
              };
            };
            const statusDetailUrl = (status, month) => {
              const p = new URLSearchParams();
              if (status) p.append('status', status);
              if (month) p.append('month', month);
              return appendDateQuery(`/api/dashboard/status-detail?${p}`);
            };
            const openStatusDetail = (status, month, titlePrefix = 'Pending Status') => {
              const label = status ? `${titlePrefix}: ${status}` : titlePrefix;
              openModal(month ? `${label} - ${month}` : label, statusDetailUrl(status, month));
            };
            return (
              <div className="overflow-x-auto">
                <table className="w-full text-[11px] leading-tight" style={{minWidth: statusMonths.length > 4 ? `${200 + statusMonths.length * 62 + 190}px` : undefined}}>
                  <thead className={tblHd}>
                    <tr>
                      <th className={`px-2 py-1 text-left font-bold whitespace-nowrap sticky left-0 z-10 ${txt2} ${darkMode?'bg-gray-700':'bg-[#f6f6f4]'}`}>Status</th>
                      {statusMonths.map(m => (
                        <th key={m} className={`px-1.5 py-1 text-center font-bold whitespace-nowrap ${txt2}`}>{monthLabel(m)}</th>
                      ))}
                      <th className={`px-1.5 py-1 text-center font-bold whitespace-nowrap ${txt2}`}>Total</th>
                      <th className={`px-1.5 py-1 text-center font-bold whitespace-nowrap ${txt2}`}>%</th>
                      <th className={`px-1.5 py-1 text-center font-bold whitespace-nowrap ${txt2}`}>Sales Amount</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${tblDv}`}>
                    {statusMonthlyRows.map(s => {
                      const percentage = grandTotal > 0 ? ((s.total / grandTotal) * 100).toFixed(1) : '0.0';
                      return (
                        <tr key={s.name} className={trHov}>
                          <td className={`px-2 py-0.5 font-semibold whitespace-nowrap sticky left-0 z-10 ${txt} ${darkMode?'bg-gray-800':'bg-white'}`}>
                            {s.name}
                          </td>
                          {statusMonths.map(m => {
                            const value = s.monthly?.[m] || 0;
                            return (
                              <td key={m} className="px-1 py-0.5 text-center font-bold rounded-sm" style={heatStyle(value)}>
                                {value ? (
                                  <button
                                    type="button"
                                    onClick={() => openStatusDetail(s.name, m)}
                                    className="w-full font-bold underline-offset-2 hover:underline cursor-pointer"
                                  >
                                    {fmtNum(value)}
                                  </button>
                                ) : ''}
                              </td>
                            );
                          })}
                          <td className="px-1.5 py-0.5 text-right font-bold text-blue-600">
                            <button type="button" onClick={() => openStatusDetail(s.name, null)} className="font-bold hover:underline cursor-pointer">
                              {fmtNum(s.total)}
                            </button>
                          </td>
                          <td className={`px-1.5 py-0.5 text-right ${txt2}`}>{percentage}%</td>
                          <td className={`px-1.5 py-0.5 text-right whitespace-nowrap ${kpiValue}`}>{fmtCurShort(s.amount)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot className={`${tblHd} font-bold`}>
                    <tr>
                      <td className={`px-2 py-0.5 sticky left-0 z-10 ${txt} ${darkMode?'bg-gray-700':'bg-[#f6f6f4]'}`}>TOTAL</td>
                      {statusMonths.map(m => (
                        <td key={m} className="px-1 py-0.5 text-center text-blue-600">
                          {totalByMonth[m] ? (
                            <button type="button" onClick={() => openStatusDetail(null, m, 'All Pending Status')} className="font-bold hover:underline cursor-pointer">
                              {fmtNum(totalByMonth[m])}
                            </button>
                          ) : ''}
                        </td>
                      ))}
                      <td className="px-1.5 py-0.5 text-right text-blue-600">
                        <button type="button" onClick={() => openStatusDetail(null, null, 'All Pending Status')} className="font-bold hover:underline cursor-pointer">
                          {fmtNum(grandTotal)}
                        </button>
                      </td>
                      <td className={`px-1.5 py-0.5 text-right ${txt2}`}>100%</td>
                      <td className={`px-1.5 py-0.5 text-right whitespace-nowrap ${kpiValue}`}>{fmtCurShort(grandAmount)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            );
          })()}
        </div>
        <div className="space-y-3">
          <h3 className={`text-base font-bold ${txt} flex items-center gap-2`}>
            <Wrench className="w-5 h-5 text-blue-600"/> Pending Item Registration
          </h3>
           <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
            {itemRegCategoryChart(itemRegProcStatus, 'Proc. Status')}
            {itemRegCategoryChart(itemRegClients, 'Client Nm.')}
          </div>
        </div>
      </div>
    );
  };

  const renderPendingDeliveryKpis = () => {
    const agingTotals = agingData.reduce((acc, v) => ({
      less_30: acc.less_30 + (v.less_30 || 0),
      days_30_90: acc.days_30_90 + (v.days_30_90 || 0),
      days_90_180: acc.days_90_180 + (v.days_90_180 || 0),
      more_180: acc.more_180 + (v.more_180 || 0),
    }), { less_30: 0, days_30_90: 0, days_90_180: 0, more_180: 0 });

    const agingCards = [
      { label: '0-30 Days', filter: '0-30', value: agingTotals.less_30, color: '#10B981' },
      { label: '30-90 Days', filter: '30-90', value: agingTotals.days_30_90, color: '#0EA5E9' },
      { label: '90-180 Days', filter: '90-180', value: agingTotals.days_90_180, color: '#F43F5E' },
      { label: '180+ Days', filter: '180+', value: agingTotals.more_180, color: '#EF4444' },
    ];

    // Use pic_aggregations from backend (aggregated from ALL filtered data, not just current page)
    // Filter out Unassigned entries
    const picKpis = (picAggregations || [])
      .filter(p => p.pic && p.pic !== 'Unassigned')
      .map(p => ({
        pic: p.pic,
        count: p.count || 0,
        amount: p.amount || 0
      }));
    const totalPendingCount = soTotal;
    const totalPendingAmount = soSubtotalAmount;
    const pendingKpis = [
      { pic: 'Total Pending', count: totalPendingCount, amount: totalPendingAmount, isTotal: true },
      ...picKpis
    ];
    const pendingKpiCols = Math.max(1, pendingKpis.length);

    return (
      <div className="mb-5">
        {pendingKpis.length > 0 ? (
          <div className="grid grid-flow-col gap-2" style={{ gridTemplateColumns: `repeat(${pendingKpiCols}, minmax(0, 1fr))` }}>
            {pendingKpis.map((p) => {
              const activePic = !p.isTotal && pendingPicHighlight === p.pic;
              const applyPicKpiFilter = () => {
                const nextHighlight = p.isTotal || activePic ? '' : p.pic;
                const nextFilters = { ...soFilters };
                setPendingPicHighlight(nextHighlight);
                setSoFilters(nextFilters);
                setSoPage(1);
                fetchSOData(nextFilters, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter, soSortOrder, nextHighlight);
              };
              return (
              <button key={p.pic} type="button" onClick={applyPicKpiFilter} className={`min-w-0 p-3 rounded-xl text-left transition-all ${activePic ? (darkMode ? 'bg-amber-900/30 border border-amber-500 ring-2 ring-amber-400' : 'bg-amber-50 border border-amber-300 ring-2 ring-amber-200') : p.isTotal ? (darkMode ? 'bg-gray-800 border border-gray-700' : 'bg-gray-50 border border-gray-200') : card} ${p.isTotal ? 'hover:border-slate-300' : 'hover:border-amber-300'}`}>
                <div className="flex items-start justify-between gap-2">
                  <div className="min-w-0">
                    <p className={`text-xs font-semibold truncate ${activePic ? 'text-amber-700' : p.isTotal ? (darkMode ? 'text-gray-200' : 'text-gray-700') : txt2}`} title={p.pic}>{p.pic}</p>
                    <h3 className={`text-xl font-bold leading-tight ${activePic ? 'text-amber-700' : p.isTotal ? (darkMode ? 'text-gray-100' : 'text-gray-800') : kpiValue}`}>{fmtNum(p.count)}</h3>
                    <p className={`text-[11px] leading-tight whitespace-nowrap ${txt2}`}>{fmtCurShort(p.amount)}</p>
                  </div>
                  <div className={`p-1.5 rounded-lg flex-shrink-0 ${activePic ? 'bg-amber-100 text-amber-700' : p.isTotal ? (darkMode ? 'bg-gray-700 text-gray-200' : 'bg-gray-100 text-gray-600') : neutralIcon}`}>
                    <Package className="w-3.5 h-3.5" />
                  </div>
                </div>
              </button>
            );})}
          </div>
        ) : (
          <div className={`py-6 text-center text-sm rounded-2xl ${card} ${txt2}`}>No pending PIC data</div>
        )}
      </div>
    );
  };

  const renderAllSO = () => (
    <div>
      {renderPendingDeliveryKpis()}
      <div className={`p-4 rounded-2xl shadow mb-5 ${card}`}>
        <div className="flex flex-wrap justify-between items-center gap-3 mb-3">
          <div>
            <h2 className={`text-xl font-bold ${txt} flex items-center gap-2`}><Clock className="w-5 h-5 text-blue-500 flex-shrink-0"/>Detail Pending Delivery</h2>
            <p className={`text-sm ${txt2}`}>{fmtNum(soTotal)} total records — page {soPage} of {soTotalPages}</p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <DownloadButton onClick={downloadSOTemplate} className={`flex items-center gap-2 px-3 py-2.5 rounded-xl text-sm font-semibold shadow-sm ${darkMode?'bg-gray-700 text-gray-100 hover:bg-gray-600':'bg-white text-gray-700 hover:bg-gray-50 border border-gray-200'}`}>
              <Download className="w-4 h-4"/>Template
            </DownloadButton>
            <label className="flex items-center gap-2 px-3 py-2.5 bg-slate-600 hover:bg-slate-700 text-white rounded-xl text-sm font-semibold shadow-sm cursor-pointer">
              <FileSpreadsheet className="w-4 h-4"/>Batch Upload
              <input type="file" accept=".xlsx,.xls" onChange={handleBatchUpload} className="hidden"/>
            </label>
            <DownloadButton onClick={downloadSOExcel} className="flex items-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-sm font-semibold shadow-sm">
              <Download className="w-4 h-4"/>Download Excel
            </DownloadButton>
          </div>
        </div>
        {/* PIC upload feedback */}
        {picUploadMsg && (
          <div className={`mb-3 px-4 py-2 rounded-lg text-sm font-medium flex items-center justify-between gap-2 ${picUploadMsg.startsWith('✅') ? (darkMode?'bg-green-900/40 text-green-300':'bg-green-50 text-green-700') : picUploadMsg.startsWith('❌') ? (darkMode?'bg-red-900/40 text-red-300':'bg-red-50 text-red-700') : (darkMode?'bg-blue-900/40 text-blue-300':'bg-blue-50 text-blue-700')}`}>
            <span>{picUploadMsg}</span>
            <button onClick={()=>setPicUploadMsg('')} className="opacity-60 hover:opacity-100 font-bold text-lg leading-none">×</button>
          </div>
        )}
        <div className="mb-2 flex flex-wrap gap-2 items-center">
          <span className={`text-xs font-medium ${txt2}`}>Filter by Aging:</span>
          {AGING_LABELS.map(label => {
            const active = soFilters.aging.includes(label);
            return (
              <button key={label} onClick={()=>toggleAgingFilter(label)}
                className={`px-3 py-1 rounded-full text-xs font-semibold border transition-all ${active?'text-white border-transparent':'border-gray-200 text-gray-400 bg-gray-100'}`}
                style={active ? {backgroundColor: AGING_COLORS[label], borderColor: AGING_COLORS[label]} : {}}>
                {label} working days
              </button>
            );
          })}
          {soFilters.aging.length > 0 && (
            <button onClick={()=>setSoFilters(f=>({...f,aging:[]}))}
              className={`px-2 py-1 rounded text-xs ${txt2} hover:text-red-500`}>Reset Aging</button>
          )}
        </div>

        {/* Multi-select filters */}
        <FilterPanel darkMode={darkMode} className="mx-0 my-3">
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-4 2xl:grid-cols-[105px_150px_repeat(6,minmax(105px,1fr))_105px] items-end">
            <div className="min-w-0">
              <label className={`block text-xs font-medium mb-0.5 ${txt2}`}>↕ SO Date</label>
              <select className={`w-full h-10 px-2 py-2 rounded-lg text-sm border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                value={soSortOrder} onChange={e=>{ setSoSortOrder(e.target.value); setSoPage(1); }} title="Sort SO Date">
                <option value="oldest">Oldest ↑</option>
                <option value="newest">Newest ↓</option>
              </select>
            </div>
            <div className="min-w-0">
              <label className={`block text-xs font-medium mb-0.5 ${txt2}`}>Search SO Item</label>
              <SearchInput
                label="SO Item"
                placeholder={"e.g.\n1234-10\n1234-20"}
                onSearch={(nums) => {
                  setSoSearchNums(nums);
                  setSoPage(1);
                  fetchSOData(soFilters, 1, soPerPage, nums, soMarginFilter, soDateFilter);
                }}
                darkMode={darkMode} txt2={txt2}
              />
            </div>
            <div className="min-w-0">
              <MultiSelect label="PIC" options={soFilterOptions.pics || []}
                selected={soFilters.pics} onChange={v=>{
                  const next = {...soFilters, pics: v};
                  setPendingPicHighlight('');
                  setSoFilters(next); setSoPage(1);
                  fetchSOData(next, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter, soSortOrder, '');
                }}
                darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="SO Status" options={soFilterOptions.statuses}
                selected={soFilters.statuses} onChange={v=>{
                  const next = {...soFilters, statuses: v};
                  setSoFilters(next); setSoPage(1);
                  fetchSOData(next, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
                }}
                darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="Operation Unit" options={soFilterOptions.op_units}
                selected={soFilters.op_units} onChange={v=>{
                  const next = {...soFilters, op_units: v};
                  setSoFilters(next); setSoPage(1);
                  fetchSOData(next, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
                }}
                darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="Vendor Name" options={soFilterOptions.vendors}
                selected={soFilters.vendors} onChange={v=>{
                  const next = {...soFilters, vendors: v};
                  setSoFilters(next); setSoPage(1);
                  fetchSOData(next, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
                }}
                darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <MultiSelect label="Manufacturer Name" options={soFilterOptions.manufacturers || []}
                selected={soFilters.manufacturers} onChange={v=>{
                  const next = {...soFilters, manufacturers: v};
                  setSoFilters(next); setSoPage(1);
                  fetchSOData(next, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
                }}
                darkMode={darkMode} txt2={txt2}/>
            </div>
            <div className="min-w-0">
              <label className={`block text-xs font-medium mb-0.5 ${txt2}`}>Margin</label>
              <select className={`w-full h-10 px-2 py-2 rounded-lg text-sm border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                value={soMarginFilter} onChange={e=>{
                  const next = e.target.value;
                  setSoMarginFilter(next); setSoPage(1);
                  fetchSOData(soFilters, 1, soPerPage, soSearchNums, next, soDateFilter);
                }}>
                <option value="all">All</option>
                <option value="positive">≥ 0</option>
                <option value="negative">Below 0</option>
              </select>
            </div>
            <div className="min-w-0">
              <label className={`block text-xs font-medium mb-0.5 ${txt2} opacity-0`}>.</label>
              <button onClick={()=>{
                const f={op_units:[],vendors:[],manufacturers:[],statuses:[],aging:[],pics:[]};
                setSoFilters(f); setPendingPicHighlight(''); setSoSearchNums([]); setSoMarginFilter('all'); setSoPage(1);
                fetchSOData(f,1,soPerPage,[],'all',soDateFilter,soSortOrder,'');
              }}
                className={`w-full h-10 px-3 py-2 rounded-lg text-sm font-medium shadow-sm flex items-center justify-center whitespace-nowrap ${darkMode?'bg-gray-500 text-gray-100 hover:bg-gray-400':'bg-gray-400 text-white hover:bg-gray-500'}`}>Clear</button>
            </div>
          </div>
          {/* Active filter tags */}
          {(soSearchNums.length + filterValues(soFilters.op_units).length + filterValues(soFilters.vendors).length + filterValues(soFilters.manufacturers).length + filterValues(soFilters.statuses).length + filterValues(soFilters.pics).length) > 0 && (
            <div className="mt-3 flex flex-wrap gap-1.5">
              {soSearchNums.map(v=>(
                <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-indigo-100 text-indigo-700 rounded-full text-xs">
                  SO: {v}<button onClick={()=>{ const next=soSearchNums.filter(x=>x!==v); setSoSearchNums(next); setSoPage(1); fetchSOData(soFilters,1,soPerPage,next,soMarginFilter,soDateFilter); }} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {filterValues(soFilters.op_units).map(v=>(
                <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-blue-100 text-blue-700 rounded-full text-xs">
                  {v}<button onClick={()=>{ const next={...soFilters,op_units:filterValues(soFilters.op_units).filter(x=>x!==v)}; setSoFilters(next); setSoPage(1); fetchSOData(next,1,soPerPage,soSearchNums,soMarginFilter,soDateFilter); }} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {filterValues(soFilters.vendors).map(v=>(
                <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-blue-100 text-blue-700 rounded-full text-xs">
                  {v}<button onClick={()=>{ const next={...soFilters,vendors:filterValues(soFilters.vendors).filter(x=>x!==v)}; setSoFilters(next); setSoPage(1); fetchSOData(next,1,soPerPage,soSearchNums,soMarginFilter,soDateFilter); }} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {filterValues(soFilters.manufacturers).map(v=>(
                <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-cyan-100 text-cyan-700 rounded-full text-xs">
                  Mfr: {v}<button onClick={()=>{ const next={...soFilters,manufacturers:filterValues(soFilters.manufacturers).filter(x=>x!==v)}; setSoFilters(next); setSoPage(1); fetchSOData(next,1,soPerPage,soSearchNums,soMarginFilter,soDateFilter); }} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {filterValues(soFilters.statuses).map(v=>(
                <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-green-100 text-green-700 rounded-full text-xs">
                  {v}<button onClick={()=>{ const next={...soFilters,statuses:filterValues(soFilters.statuses).filter(x=>x!==v)}; setSoFilters(next); setSoPage(1); fetchSOData(next,1,soPerPage,soSearchNums,soMarginFilter,soDateFilter); }} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {filterValues(soFilters.pics).map(v=>{
                const c = getPicColor(v);
                return (
                  <span key={v} className={`flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold ${c ? c.bg+' '+c.text : 'bg-gray-100 text-gray-700'}`}>
                    PIC: {v}<button onClick={()=>{ const nextPics=filterValues(soFilters.pics).filter(x=>x!==v); const next={...soFilters,pics:nextPics}; setPendingPicHighlight(''); setSoFilters(next); setSoPage(1); fetchSOData(next,1,soPerPage,soSearchNums,soMarginFilter,soDateFilter,soSortOrder,''); }} className="hover:text-red-600 ml-0.5"><X className="w-3 h-3"/></button>
                  </span>
                );
              })}
            </div>
          )}
        </FilterPanel>

        {/* Detail Pending Delivery table follows the downloadable Excel layout. */}
        <DataTableScroll darkMode={darkMode} className="rounded-lg border border-gray-200">
          <table className="freeze-table-pending-delivery w-full text-sm">
            <colgroup>
              <col style={{minWidth:'76px', width:'76px', maxWidth:'76px'}}/>
              <col style={{minWidth:'60px'}}/>
              <col style={{minWidth:'110px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'170px', width:'170px', maxWidth:'170px'}}/>
              <col style={{minWidth:'120px'}}/>
              <col style={{minWidth:'80px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'180px'}}/>
              <col style={{minWidth:'260px'}}/>
              <col style={{minWidth:'140px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'90px'}}/>
              <col style={{minWidth:'140px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'160px'}}/>
              <col style={{minWidth:'80px'}}/>
              <col style={{minWidth:'140px'}}/>
              <col style={{minWidth:'140px'}}/>
              <col style={{minWidth:'130px'}}/>
              <col style={{minWidth:'130px'}}/>
              <col style={{minWidth:'130px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'90px'}}/>
              <col style={{minWidth:'200px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'560px'}}/>
            </colgroup>
            <thead className={tblHd}>
              <tr>
                {['Aging','Day','SO Create Date','SO Item','PO No.','SO Status','Category','PIC','Product ID','Product Name','Specification','Manufacturer Name','SO Quantity','Sales Unit','Operation Unit Name','Vendor ID','Vendor Name','Currency','Sales Price (Exclude Tax)','Sales Amount (Exclude Tax)','Purchasing Currency','Purchasing Price','Purchase Price (IDR)','Margin','%Margin','Delivery Memo','Plan Date','Remarks'].map((h, index)=>(
                  <th key={h} data-col-index={index + 1} className={`px-3 py-2.5 text-center font-bold ${txt2}`}>{renderFreezeHeader('pending-delivery', index + 1, h)}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {(() => {
                if (sortedSOData.length === 0) return (
                <tr><td colSpan={28} className={`px-4 py-10 text-center ${txt2}`}>
                    <FileText className="w-10 h-10 mx-auto mb-2 opacity-40"/>No data
                  </td></tr>
                );
                return sortedSOData.map((so) => {
                const isDeliveryCompleted = so.so_status === 'Delivery Completed';
                // Use IDR-converted purchase amount for margin (handles USD/EUR).
                // Backend provides purchase_amount_idr & purchase_price_idr.
                const poAmount = Number(so.purchase_amount_idr) || 0;
                const poPriceIdr = Number(so.purchase_price_idr) || 0;
                // Margin valid only when purchase (in IDR) is positive.
                // Invalid purchase → margin = null, displayed as '-'.
                const purchaseValid = poAmount > 0;
                const margin = purchaseValid ? (so.sales_amount || 0) - poAmount : null;
                const marginPct = purchaseValid ? (margin / poAmount) * 100 : null;
                const workingDays = Number.isFinite(Number(so.aging_days)) ? Number(so.aging_days) : workingDaysUntilToday(so.so_create_date);
                const marginColor = margin == null ? txt2 : margin < 0 ? 'text-red-600 font-semibold' : margin > 0 ? 'text-green-600 font-semibold' : txt2;
                return (
                <tr key={so.id} className={`${trHov} transition-colors`}>
                  <td className="px-2 py-2 text-center whitespace-nowrap">
                    {!isDeliveryCompleted && so.aging_label && so.aging_label !== 'No Date' ? (
                      <span className="px-2 py-0.5 rounded-full text-xs font-bold text-white"
                        style={{backgroundColor: AGING_COLORS[so.aging_label] || '#6B7280'}}>
                        {so.aging_label}
                      </span>
                    ) : null}
                  </td>
                  <td className={`px-3 py-2 text-center whitespace-nowrap ${workingDays !== null && workingDays > 180 ? 'text-red-600 font-bold' : workingDays !== null && workingDays > 30 ? 'text-slate-700 font-semibold' : 'text-green-600 font-semibold'}`}>
                    {workingDays !== null ? workingDays : '-'}
                  </td>
                  <td className={`px-3 py-2 text-center text-xs ${txt2} whitespace-nowrap`}>{so.so_create_date||'-'}</td>
                  <td className="px-3 py-2 text-blue-600 font-medium whitespace-nowrap">{so.so_item}</td>
                  <td className={`px-3 py-2 ${txt2} whitespace-nowrap`}>{so.svo_po || '-'}</td>
                  <td className="px-3 py-2 w-[170px] max-w-[170px] whitespace-nowrap">
                    <span title={so.so_status || '-'} className={`inline-block max-w-[150px] truncate align-middle px-2 py-0.5 rounded-full text-xs font-medium ${
                      so.so_status==='Delivery Completed'?'bg-green-100 text-green-700':
                      so.so_status==='SO Cancel'?'bg-red-100 text-red-700':'bg-blue-100 text-blue-700'}`}>
                      {so.so_status||'-'}
                    </span>
                  </td>
                  <td className={`px-3 py-2 ${txt2} whitespace-nowrap`}>{so.category_name || '-'}</td>
                  <td className={`px-3 py-2 whitespace-nowrap`}>
                    {so.pic_name ? (
                      (() => { const c = getPicColor(so.pic_name); return (
                        <span className={`px-2 py-0.5 rounded-full text-xs font-semibold ${c.bg} ${c.text}`}>{so.pic_name}</span>
                      ); })()
                    ) : <span className={`text-xs ${txt2}`}>-</span>}
                  </td>
                  <td className={`px-3 py-2 ${txt2} whitespace-nowrap`}>{so.product_id || '-'}</td>
                  <td className={`px-3 py-2 max-w-[180px] truncate ${txt2}`} title={so.product_name}>{so.product_name || '-'}</td>
                  <td className={`px-3 py-2 max-w-[260px] truncate ${txt2}`} title={so.specification}>{so.specification || '-'}</td>
                  <td className={`px-3 py-2 max-w-[180px] truncate ${txt2}`} title={so.manufacturer_name}>{so.manufacturer_name || '-'}</td>
                  <td className={`px-3 py-2 text-right ${txt2}`}>{fmtNum(so.so_qty)}</td>
                  <td className={`px-3 py-2 ${txt2} whitespace-nowrap`}>{so.sales_unit || '-'}</td>
                  <td className={`px-3 py-2 min-w-[180px] truncate ${txt2}`} title={so.operation_unit_name}>{so.operation_unit_name || '-'}</td>
                  <td className={`px-3 py-2 ${txt2} whitespace-nowrap`}>{so.vendor_id || '-'}</td>
                  <td className={`px-3 py-2 max-w-[160px] truncate ${txt2}`} title={so.vendor_name}>{so.vendor_name || '-'}</td>
                  <td className={`px-3 py-2 ${txt2} whitespace-nowrap`}>{so.currency || '-'}</td>
                  <td className={`px-3 py-2 text-right whitespace-nowrap min-w-[130px] ${txt}`}>{fmtCur(so.sales_price)}</td>
                  <td className={`px-3 py-2 text-center font-bold whitespace-nowrap min-w-[130px] ${kpiValue}`}>{fmtCur(so.sales_amount)}</td>
                  <td className={`px-3 py-2 ${txt2} whitespace-nowrap`}>{so.purchasing_currency || '-'}</td>
                  <td className={`px-3 py-2 text-right whitespace-nowrap min-w-[130px] ${txt}`}>{fmtCur(so.purchasing_price)}</td>
                  <td className={`px-3 py-2 text-right whitespace-nowrap min-w-[130px] ${txt}`}>{poPriceIdr > 0 ? fmtCur(poPriceIdr) : '-'}</td>
                  <td className={`px-3 py-2 text-right whitespace-nowrap min-w-[130px] ${marginColor}`}>{margin == null ? '-' : fmtCur(margin)}</td>
                  <td className={`px-3 py-2 text-right whitespace-nowrap ${marginColor}`}>
                    {marginPct !== null ? `${marginPct.toFixed(1)}%` : '-'}
                  </td>
                  <td className={`px-3 py-2 max-w-[200px] truncate ${txt2}`} title={so.delivery_memo}>{so.delivery_memo || '-'}</td>
                  <td className="px-3 py-2 text-center">
                    {editingCell?.id===so.id && editingCell.field==='delivery_plan_date' ? (
                      <div className="flex items-center gap-1">
                        <input type="date" defaultValue={so.delivery_plan_date}
                          className={`px-2 py-1 rounded text-xs border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                          onChange={e=>setEditValue(e.target.value)}
                          onKeyDown={e=>{
                            if(e.key==='Enter'){e.preventDefault();updateSOCell(so.id,'delivery_plan_date',editValue);}
                            if(e.key==='Escape'){e.preventDefault();setEditingCell(null);}
                          }}
                          autoFocus/>
                        <button onMouseDown={e=>{e.preventDefault();updateSOCell(so.id,'delivery_plan_date',editValue);}}
                          className="text-green-500 hover:text-green-700 p-0.5 text-xs font-bold">✓</button>
                        <button onMouseDown={e=>{e.preventDefault();setEditingCell(null);}}
                          className="text-red-400 hover:text-red-600 p-0.5"><X className="w-3.5 h-3.5"/></button>
                      </div>
                    ) : (
                      <div className="flex items-center justify-center gap-1 group">
                        <span className="cursor-pointer text-blue-600 hover:underline text-xs whitespace-nowrap"
                          onClick={()=>{setEditingCell({id:so.id,field:'delivery_plan_date'});setEditValue(so.delivery_plan_date||'');}}>
                          {so.delivery_plan_date||'Set'}
                        </span>
                        {so.delivery_plan_date && (
                          <button onClick={e=>{e.stopPropagation();updateSOCell(so.id,'delivery_plan_date','');}}
                            className="opacity-0 group-hover:opacity-100 text-red-400 hover:text-red-600 transition-all p-0.5"><X className="w-3 h-3"/></button>
                        )}
                      </div>
                    )}
                  </td>
                  <td className="px-3 py-2 min-w-[560px]">
                    {editingCell?.id===so.id && editingCell.field==='remarks' ? (
                      <input type="text" defaultValue={so.remarks}
                        className={`w-full px-2 py-1 rounded text-xs border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                        onChange={e=>setEditValue(e.target.value)}
                        onBlur={()=>updateSOCell(so.id,'remarks',editValue)}
                        onKeyDown={e=>e.key==='Enter'&&updateSOCell(so.id,'remarks',editValue)}
                        autoFocus/>
                    ) : (
                      <span className="cursor-pointer text-xs text-blue-600 hover:underline"
                        onClick={()=>{setEditingCell({id:so.id,field:'remarks'});setEditValue(so.remarks||'');}}>
                        {so.remarks||'Add'}
                      </span>
                    )}
                  </td>
                </tr>
                );
              })
              })()}
            </tbody>
          </table>
        </DataTableScroll>

        <PagePagination
          darkMode={darkMode}
          txt2={txt2}
          page={soPage}
          totalPages={soTotalPages}
          total={soTotal}
          perPage={soPerPage}
          onPageChange={(p) => { setSoPage(p); fetchSOData(soFilters, p, soPerPage, soSearchNums, soMarginFilter, soDateFilter); }}
          onPerPageChange={(next) => { setSoPerPage(next); setSoPage(1); fetchSOData(soFilters, 1, next, soSearchNums, soMarginFilter, soDateFilter); }}
        />
      </div>

    </div>
  );

  // ══════════════════════════════════════════════════════════════
  // MAIN RENDER
  // ══════════════════════════════════════════════════════════════
  return (
    <div
      className={`min-h-screen font-sans ${darkMode?'bg-gray-900':'bg-[#edf2f1]'} ${darkMode?'':'text-[#1f2937]'}`}
      style={{
        fontFamily: "'Inter', 'Plus Jakarta Sans', ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
        // Tabular figures only (digits occupy equal width so columns of
        // numbers align perfectly). We DO NOT enable the "zero" OpenType
        // feature — that adds a slash/dot inside 0, which the user explicitly
        // rejected. Plain 0 (Inter's default) matches what the Item
        // Registration page shows.
        fontFeatureSettings: '"tnum" 1',
        fontVariantNumeric: 'tabular-nums',
      }}
    >
    <style>{`
        /* ─────────────────────────────────────────────────────────────
           FONT BASELINE — Inter for everything, JetBrains Mono for
           monospace blocks. Loaded from Google Fonts. Matches the
           typography used on Item Registration so the entire dashboard
           looks unified. Plain 0 (no slash, no dot) — Inter's default
           digit zero is used everywhere.
           NOTE: @import statements MUST appear before any other rules
           in a stylesheet, so they live at the very top here.
           ───────────────────────────────────────────────────────────── */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&display=swap');

        /* Global typography baseline. tnum = tabular figures (digits
           occupy equal width so columns of numbers align perfectly).
           The "zero" feature is INTENTIONALLY NOT enabled — the user
           wants plain 0 (no slash, no dot) just like Item Registration.
           Applied to EVERY element so there is no drift between sidebar
           / dashboard / tables / modals. */
        *, *::before, *::after {
          font-family: 'Inter', 'Plus Jakarta Sans', ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
          font-feature-settings: "tnum" 1 !important;
          font-variant-numeric: tabular-nums !important;
        }
        /* Monospace blocks (formulas, code, ID columns) — JetBrains Mono
           with plain 0 as well (no slashed/dotted zero feature). */
        .font-mono, code, pre, kbd {
          font-family: 'JetBrains Mono', ui-monospace, 'SFMono-Regular', Menlo, Consolas, monospace !important;
          font-feature-settings: "tnum" 1 !important;
          font-variant-numeric: tabular-nums !important;
        }

        /* Suppress chart/card entry animations (they slow down the UI on re-render)
           but ALLOW spinner, pulse, and slide-in used by loading indicators.
           We define the keyframes here so they survive Tailwind's purge in prod. */
        *, *::before, *::after {
          transition: none !important;
          scroll-behavior: auto !important;
        }
        /* Kill recharts / radix / general UI entry animations */
        .recharts-wrapper *, [data-radix-popper-content-wrapper] * {
          animation: none !important;
        }
        /* ── Kill browser default focus ring on editable table cells ──
           When a user clicks a cell to edit, the <td> shows a blue outline
           (the "outer" border the user wants). The <input>/<textarea> inside
           the td would ALSO show the browser's default focus ring (the
           "inner" border), producing the double-blue-border bug. This rule
           kills the inner focus ring on any element tagged with
           data-no-focus-ring, in every state (:focus, :focus-visible,
           :focus-within) so Chrome/Safari/Edge can't sneak it back in. */
        [data-no-focus-ring], [data-no-focus-ring]:focus, [data-no-focus-ring]:focus-visible,
        [data-no-focus-ring]:focus-within, [data-no-focus-ring]:active {
          outline: none !important;
          outline-style: none !important;
          outline-width: 0 !important;
          box-shadow: none !important;
        }
        /* date inputs: keep the OUTER picker indicator visible but still
           kill the focus ring. The previous rule already set outline:none
           and box-shadow:none which is what we want. We only need to NOT
           restore a border-color here — restoring border-color would
           re-introduce the second blue border on date inputs in Chrome. */
        input[type="date"][data-no-focus-ring], input[type="date"][data-no-focus-ring]:focus,
        input[type="date"][data-no-focus-ring]:focus-visible {
          outline: none !important;
          outline-style: none !important;
          outline-width: 0 !important;
          box-shadow: none !important;
          /* borderColor stays transparent (set inline) so no second border
             appears. Don't restore it. */
        }
        /* ── Loading animation keyframes (must be explicit for Tailwind purge) ── */
        @keyframes spin {
          from { transform: rotate(0deg); }
          to   { transform: rotate(360deg); }
        }
        @keyframes pulse {
          0%, 100% { opacity: 1; }
          50%       { opacity: 0.4; }
        }
        @keyframes slide-in {
          from { opacity: 0; transform: translateX(24px); }
          to   { opacity: 1; transform: translateX(0); }
        }
        /* ── Apply animations to specific loading classes ── */
        .animate-spin {
          animation: spin 0.75s linear infinite !important;
        }
        .animate-pulse {
          animation: pulse 1.8s ease-in-out infinite !important;
        }
        .animate-slide-in {
          animation: slide-in 0.22s ease-out !important;
        }
        /* Global: all buttons, links, selects, labels with checkboxes → pointer cursor */
        button, [role="button"], select, label[for], a,
        input[type="checkbox"], input[type="radio"] {
          cursor: pointer !important;
        }
        button:disabled { cursor: not-allowed !important; opacity: 0.5; }
        .rfq-fill-handle { cursor: crosshair !important; }
        body.rfq-fill-dragging, body.rfq-fill-dragging * { cursor: crosshair !important; }
        .freeze-table-import td, .freeze-table-import th { box-shadow: inset 0 -1px 0 rgba(148, 163, 184, 0.22), inset -1px 0 0 rgba(148, 163, 184, 0.22); }
        .freeze-table-import tbody tr { height: 32px; }
        /* max-height only on non-date cells — date picker popup needs unrestricted height */
        .freeze-table-import td > *:not(input[type="date"]) { max-height: 28px; }
        /* When detail columns are hidden, FORCE every row to be exactly 1 line
           tall (32px) — no exceptions. This overrides any cell content that
           would normally push the row taller (long remarks, multi-line text).
           The user reported rows staying at original height even after Hide
           Detail; this rule enforces the 1-line height strictly. */
        .freeze-table-import.import-detail-hidden tbody tr { height: 32px !important; max-height: 32px !important; }
        .freeze-table-import.import-detail-hidden tbody td { height: 32px !important; max-height: 32px !important; padding-top: 2px !important; padding-bottom: 2px !important; overflow: hidden !important; }
        .freeze-table-import.import-detail-hidden tbody td > *:not(input[type="date"]) { max-height: 28px !important; overflow: hidden !important; }
        .freeze-table-import.import-detail-hidden tbody td button,
        .freeze-table-import.import-detail-hidden tbody td span,
        .freeze-table-import.import-detail-hidden tbody td div { max-height: 28px !important; overflow: hidden !important; line-height: 1.2 !important; }
        .freeze-table-import input, .freeze-table-import textarea, .freeze-table-import select {
          outline: none !important;
          box-shadow: none !important;
        }
        .freeze-table-import input:not([type="date"]):focus,
        .freeze-table-import textarea:focus,
        .freeze-table-import select:focus {
          outline: none !important;
          box-shadow: inset 0 0 0 2px #3b82f6 !important;
        }
        /* When a cell is being EDITED, the td already shows the outer blue
           outline. Suppress the inner inset box-shadow so we don't get a
           double blue border. This is the fix for the Import double-border
           bug. */
        .freeze-table-import td.is-editing input:focus,
        .freeze-table-import td.is-editing textarea:focus,
        .freeze-table-import td.is-editing select:focus {
          box-shadow: none !important;
        }
        /* date input: no box-shadow override so browser can render picker normally */
        .freeze-table-import input[type="date"] {
          box-shadow: none !important;
        }

        .backdrop-blur-sm, .backdrop-blur-xl {
          backdrop-filter: none !important;
          -webkit-backdrop-filter: none !important;
        }
        .reference-card {
          border-radius: 24px;
          box-shadow: 0 18px 45px rgba(15, 23, 42, 0.06);
        }
        .data-table-page table {
          border-collapse: collapse;
        }
        .data-table-page table th,
        .data-table-page table td {
          border-right: 1px solid rgba(148, 163, 184, 0.28);
        }
        .data-table-scroll thead th {
          position: relative !important;
          z-index: 56;
          box-shadow: 0 1px 0 rgba(148, 163, 184, 0.28);
          background-clip: padding-box;
        }
        .window-sticky-table-header {
          border-top: 1px solid rgba(148, 163, 184, 0.32);
          border-left: 1px solid rgba(148, 163, 184, 0.32);
          border-right: 1px solid rgba(148, 163, 184, 0.32);
          box-shadow: 0 8px 18px rgba(15, 23, 42, 0.16);
        }
        .window-sticky-table-header table {
          border-collapse: collapse !important;
          font-size: 0.75rem !important;
          line-height: 1rem !important;
        }
        .window-sticky-table-header thead th {
          box-shadow: inset 0 -1px 0 rgba(148, 163, 184, 0.28), inset -1px 0 0 rgba(148, 163, 184, 0.28) !important;
          white-space: normal !important;
          overflow-wrap: anywhere;
          word-break: normal;
          line-height: 1.15;
          vertical-align: middle;
        }
        .window-sticky-table-header .freeze-header {
          min-height: 2rem;
          height: auto;
          padding-right: 1.25rem;
          align-items: center;
        }
        .window-sticky-table-header .freeze-header-label {
          white-space: normal !important;
          overflow-wrap: anywhere;
          word-break: normal;
          line-height: 1.15;
        }
        .window-sticky-table-header button,
        .window-sticky-table-header select,
        .window-sticky-table-header input {
          pointer-events: none !important;
        }
        .data-table-scroll-frame {
          border: 1px solid rgba(148, 163, 184, 0.32);
          border-radius: 10px;
          background: ${darkMode ? '#111827' : '#ffffff'};
          min-height: 220px;
          overflow-x: auto !important;
          overflow-y: visible !important;
        }
        .floating-table-scrollbar {
          position: fixed;
          bottom: 12px;
          height: 24px;
          overflow-x: auto;
          overflow-y: hidden;
          z-index: 90;
          opacity: 0.72;
          border-radius: 999px;
          background: rgba(148, 163, 184, 0.18);
        }
        .floating-table-scrollbar-dark {
          background: rgba(255, 255, 255, 0.10);
        }
        .floating-table-scrollbar::-webkit-scrollbar {
          height: 22px;
        }
        .floating-table-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }
        .floating-table-scrollbar::-webkit-scrollbar-thumb {
          border-radius: 999px;
          background: rgba(71, 85, 105, 0.52);
          border: 4px solid transparent;
          background-clip: padding-box;
        }
        .floating-table-scrollbar-dark::-webkit-scrollbar-thumb {
          background: rgba(226, 232, 240, 0.52);
          border: 4px solid transparent;
          background-clip: padding-box;
        }
        .data-table-page table th {
          white-space: normal !important;
          overflow-wrap: anywhere;
          word-break: normal;
          line-height: 1.15;
          vertical-align: middle;
        }
        .data-table-page .freeze-header-label {
          display: -webkit-box;
          -webkit-line-clamp: 2;
          -webkit-box-orient: vertical;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: normal;
          word-break: normal;
          overflow-wrap: normal;
          hyphens: none;
        }
        /* Short labels like "Days Left" should NOT wrap to 3 lines — keep
           them on a single line so the header doesn't look cramped when
           hovered. Width 80px is enough for "Days Left" + the pin button. */
        .data-table-page .freeze-table-import th .freeze-header-label,
        .freeze-table-import th .freeze-header-label {
          white-space: nowrap !important;
          -webkit-line-clamp: 1 !important;
        }
        .data-table-page .freeze-header button {
          pointer-events: auto;
        }
        .data-table-page-dark table th,
        .data-table-page-dark table td {
          border-right-color: rgba(75, 85, 99, 0.85);
        }
        .data-table-page table th:last-child,
        .data-table-page table td:last-child {
          border-right: 0;
        }
        ${frozenColumnCss}
      `}</style>

      <div className="fixed top-5 right-5 z-[100] flex flex-col gap-2">
        {toasts.map(t=><Toast key={t.id} message={t.message} type={t.type} onClose={()=>removeToast(t.id)}/>)}
      </div>

      {/* Download progress toast */}
      {downloadToast && <DownloadToast message={downloadToast.message} />}

      {/* Sidebar */}
      <aside
        onMouseEnter={() => setSidebarExpanded(true)}
        onMouseLeave={() => setSidebarExpanded(false)}
        onFocusCapture={() => setSidebarExpanded(true)}
        onBlurCapture={(e) => { if (!e.currentTarget.contains(e.relatedTarget)) setSidebarExpanded(false); }}
        className={`fixed left-0 top-0 h-full ${sidebarExpanded?'lg:w-60':'lg:w-16'} w-16 flex flex-col items-stretch py-5 shadow-[0_8px_24px_rgba(15,23,42,0.08)] z-40 overflow-hidden transition-[width,background-color,border-color] duration-200 ease-out ${darkMode?'bg-gray-800 border-r border-gray-700':'bg-white/95 border-r border-gray-200/80 backdrop-blur-xl'}`}
      >
        <nav className={`flex-1 flex flex-col gap-2 w-full ${sidebarExpanded?'lg:px-3':'lg:px-2'} px-2 pt-0`}>
          <a href={PAGE_PATHS.dashboard} onClick={(e)=>openPage(e, 'dashboard')}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='dashboard'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="Summary">
            <BarChart3 className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-40 opacity-100':'max-w-0 opacity-0'}`}>Summary</span>
          </a>
          <a href={PAGE_PATHS['all-so']} data-tour="open-so-nav" onClick={(e)=>openPage(e, 'all-so', () => setSoPage(1))}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='all-so'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="Pending Delivery">
            <Clock className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-40 opacity-100':'max-w-0 opacity-0'}`}>Pending Delivery</span>
          </a>
          <a href={PAGE_PATHS['item-registration']} onClick={(e)=>openPage(e, 'item-registration', () => setItemRegPage(1))}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='item-registration'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="Item Registration">
            <Wrench className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-44 opacity-100':'max-w-0 opacity-0'}`}>Item Registration</span>
          </a>
          <a href={PAGE_PATHS.rfq} onClick={(e)=>openPage(e, 'rfq', () => setRfqPage(1))}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='rfq'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="RFQ">
            <Mail className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-44 opacity-100':'max-w-0 opacity-0'}`}>RFQ</span>
          </a>
          <a href={PAGE_PATHS.import} onClick={(e)=>openPage(e, 'import', () => setImportPage(1))}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='import'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="Import">
            <Ship className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-44 opacity-100':'max-w-0 opacity-0'}`}>Import</span>
          </a>
          <a href={PAGE_PATHS['vendor-control']} onClick={(e)=>openPage(e, 'vendor-control', () => setVendorControlPage(1))}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='vendor-control'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="Vendor Control">
            <Building2 className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-44 opacity-100':'max-w-0 opacity-0'}`}>Vendor Control</span>
          </a>
          <a href={PAGE_PATHS['all-registered-items']} onClick={(e)=>openPage(e, 'all-registered-items', () => setRegisteredItemsPage(1))}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='all-registered-items'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="All Registered Items">
            <FileText className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-44 opacity-100':'max-w-0 opacity-0'}`}>All Registered Items</span>
          </a>
        </nav>
        <div className={`flex flex-col gap-2 w-full ${sidebarExpanded?'lg:px-3':'lg:px-2'} px-2`}>
          <button onClick={()=>setDarkMode(d=>!d)} className={`p-3 rounded-xl transition-all flex items-center gap-3 justify-start whitespace-nowrap ${darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`}>
            {darkMode?<Sun className="w-5 h-5 flex-shrink-0"/>:<Moon className="w-5 h-5 flex-shrink-0"/>}
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-40 opacity-100':'max-w-0 opacity-0'}`}>{darkMode?'Light Mode':'Dark Mode'}</span>
          </button>
        </div>
      </aside>

      {/* Main */}
      <main className={`ml-16 ${sidebarExpanded?'lg:ml-60':'lg:ml-16'} p-4 lg:p-6 transition-[margin-left] duration-200 ease-out`}>
        <div className={`${darkMode?'bg-gray-900 border-gray-800':'bg-[#fbfbfa] border-gray-200/70'} ${activePage === 'dashboard' ? '' : darkMode ? 'data-table-page data-table-page-dark' : 'data-table-page'} min-h-[calc(100vh-32px)] lg:min-h-[calc(100vh-48px)] rounded-2xl border shadow-[0_12px_36px_rgba(15,23,42,0.07)] p-4 lg:p-6`}>
        <header className="mb-7 flex flex-wrap justify-between items-center gap-4">
          <div data-tour="page-title">
            <h1 className={`text-[28px] leading-tight font-bold tracking-[-0.02em] ${txt}`}>
              Serveone <span className="text-[#2563EB]">Dashboard</span>
            </h1>
            <p className={`mt-0.5 text-sm ${txt2}`}>
              {activePage==='dashboard'?'Purchase Orders & Sales Orders Summary'
               :activePage==='all-so'?'Pending Delivery monitoring and detail records'
               :activePage==='item-registration'?'Product Registration Status data'
               :activePage==='rfq'?'Sales Submit-RFQ live data and quotation updates'
               :activePage==='import'?'Import shipment and vendor import tracking'
               :activePage==='vendor-control'?'Vendor account access and credential control'
               :activePage==='all-registered-items'?'All registered product master data'
               :'Manage Pending Delivery records'}
            </p>
          </div>
          <div className="flex flex-col items-end gap-1">
            <div className="flex flex-wrap gap-2 justify-end">
              <div className="relative" ref={uploadDropdownRef}>
              <button
                data-tour="manual-update"
                onClick={() => setShowUploadDropdown(v => !v)}
                className="flex items-center gap-2 px-4 py-2.5 rounded-xl shadow-sm transition-all bg-blue-600 hover:bg-blue-700 text-white"
              >
                <Upload className="w-4 h-4"/>
                <span className="text-sm font-medium">Manual Update</span>
                <ChevronDown className={`w-4 h-4 transition-transform ${showUploadDropdown ? 'rotate-180' : ''}`}/>
              </button>
              {showUploadDropdown && (
                <div className={`absolute right-0 mt-2 z-50 rounded-xl shadow-2xl border min-w-[280px] overflow-hidden ${darkMode?'bg-gray-800 border-gray-700 text-white':'bg-white border-gray-200'}`}>
                  <label className={`flex items-center gap-2 px-4 py-3 cursor-pointer transition-all ${darkMode?'hover:bg-gray-700':'hover:bg-blue-50'}`}>
                    <Upload className="w-4 h-4 text-blue-500"/>
                    <span className={`text-sm font-medium ${txt}`}>Upload SO - Search Client Odr</span>
                    <input type="file" accept=".xlsx,.xls" multiple onChange={e=>{handleUpload(e,'scor'); setShowUploadDropdown(false);}} className="hidden"/>
                  </label>
                  <div className={`${darkMode?'border-t border-gray-700':'border-t border-gray-100'}`}></div>
                  <label className={`flex items-center gap-2 px-4 py-3 cursor-pointer transition-all ${darkMode?'hover:bg-gray-700':'hover:bg-blue-50'}`}>
                    <Upload className="w-4 h-4 text-blue-500"/>
                    <div>
                      <span className={`text-sm font-medium ${txt}`}>Upload Item Registration</span>
                      <p className={`text-xs ${txt2}`}>Upload SAP Process Pur. Info. Reg. only</p>
                    </div>
                    <input type="file" accept=".xlsx,.xls" multiple onChange={e=>{handleUploadItemRegistration(e); setShowUploadDropdown(false);}} className="hidden"/>
                  </label>
                  <div className={`${darkMode?'border-t border-gray-700':'border-t border-gray-100'}`}></div>
                  <label className={`flex items-center gap-2 px-4 py-3 cursor-pointer transition-all ${darkMode?'hover:bg-gray-700':'hover:bg-sky-50'}`}>
                    <Upload className="w-4 h-4 text-sky-500"/>
                    <div>
                      <span className={`text-sm font-medium ${txt}`}>Upload Prod ID (SAP)</span>
                      <p className={`text-xs ${txt2}`}>Update Product ID to category mapping</p>
                    </div>
                    <input type="file" accept=".xlsx,.xls" multiple onChange={e=>{handleUploadProductID(e); setShowUploadDropdown(false);}} className="hidden"/>
                  </label>
                  <div className={`${darkMode?'border-t border-gray-700':'border-t border-gray-100'}`}></div>
                  <button type="button" onClick={()=>{ downloadMasterPICTemplate(); setShowUploadDropdown(false); }} className={`w-full flex items-center gap-2 px-4 py-3 text-left transition-all ${darkMode?'hover:bg-gray-700':'hover:bg-indigo-50'}`}>
                    <Download className="w-4 h-4 text-indigo-500"/>
                    <div>
                      <span className={`text-sm font-medium ${txt}`}>Master PIC Update Template</span>
                      <p className={`text-xs ${txt2}`}>Category Name, PIC, Update New PIC</p>
                    </div>
                  </button>
                  <div className={`${darkMode?'border-t border-gray-700':'border-t border-gray-100'}`}></div>
                  <label className={`flex items-center gap-2 px-4 py-3 cursor-pointer transition-all ${darkMode?'hover:bg-gray-700':'hover:bg-indigo-50'}`}>
                    <Upload className="w-4 h-4 text-indigo-500"/>
                    <div>
                      <span className={`text-sm font-medium ${txt}`}>Update PIC</span>
                      <p className={`text-xs ${txt2}`}>Update PIC by Category Name</p>
                    </div>
                    <input type="file" accept=".xlsx,.xls" multiple onChange={e=>{handleUpdatePIC(e); setShowUploadDropdown(false);}} className="hidden"/>
                  </label>
                </div>
              )}
              </div>

            </div>
            <div className={`max-w-full text-right text-xs ${txt2}`}>
              <span className="font-semibold">Updates:</span>{' '}
              <span
                title={(() => {
                  const base = `Last Update SO: ${fmtDateTime(stats?.last_updated_smro)}`;
                  const updatedToday = stats?.so_updated_months_today;
                  const updateDate = stats?.so_updated_months_today_date;

                  if (!updatedToday || !Object.keys(updatedToday).length) {
                    return `${base}\n\nSO months updated today${updateDate ? ` (${updateDate})` : ''}: None`;
                  }

                  const MONTHS_FULL = ['January','February','March','April','May','June','July','August','September','October','November','December'];
                  const MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

                  const lines = Object.entries(updatedToday)
                    .sort(([a], [b]) => Number(a) - Number(b))
                    .map(([yr, months]) => {
                      const uniqueMonths = [...new Set(months || [])];
                      if (uniqueMonths.length === 12) return `${yr}: All month`;

                      const shorts = uniqueMonths
                        .map(month => {
                          const monthIndex = MONTHS_FULL.indexOf(month);
                          return monthIndex >= 0 ? MONTHS_SHORT[monthIndex] : month;
                        })
                        .filter(Boolean);

                      return `${yr}: ${shorts.join(', ')}`;
                    });

                  return `${base}\n\nSO months updated today${updateDate ? ` (${updateDate})` : ''}:\n${lines.join('\n')}`;
                })()}
                className="cursor-help"
              >SO {fmtUpdateShort(stats?.last_updated_smro)}</span>
              <span className="mx-1.5 opacity-50">·</span>
              <span title={`Last Update Regist.: ${fmtDateTime(stats?.last_updated_item_registration)}`}>Reg {fmtUpdateShort(stats?.last_updated_item_registration)}</span>
              <span className="mx-1.5 opacity-50">·</span>
              <span title={`Last Update Prod ID: ${fmtDateTime(picDbStatus?.last_product_id_upload)}`}>Prod ID {fmtUpdateShort(picDbStatus?.last_product_id_upload)}</span>
              {rfqLastUpdated && (<>
                <span className="mx-1.5 opacity-50">·</span>
                <span title={`Last Update RFQ: ${fmtDateTime(rfqLastUpdated)}`}>RFQ {fmtUpdateShort(rfqLastUpdated)}</span>
              </>)}
            </div>
          </div>
        </header>

        {renderGlobalSlicer()}

        {activePage==='dashboard' ? renderDashboardOverview()
          : activePage==='item-registration' ? renderItemRegistration()
          : activePage==='rfq' ? renderRFQ()
          : activePage==='import' ? renderImport()
          : activePage==='vendor-control' ? renderVendorControl()
          : activePage==='all-registered-items' ? renderAllRegisteredItems()
          : renderAllSO()}
        </div>
      </main>

      {modal && <SOModal title={modal.title} data={modal.data} darkMode={darkMode} onClose={()=>setModal(null)} onUpdateCell={updateSOCell}/>} 

      {marginDetailModal && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm" onClick={()=>setMarginDetailModal(null)}>
          <div className={`rounded-2xl overflow-hidden shadow-2xl w-full max-w-5xl max-h-[85vh] flex flex-col ${darkMode?'bg-gray-800 text-white':'bg-white'}`} onClick={e=>e.stopPropagation()}>
            <div className={`flex justify-between items-center px-6 py-4 border-b ${darkMode?'border-gray-700':'border-gray-100'}`}>
              <h3 className="font-bold text-lg">Margin Detail — {marginDetailModal.category}
                <span className={`text-sm font-normal ml-2 ${txt2}`}>({fmtNum(marginDetailModal.data?.length)} records)</span>
              </h3>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => downloadStyledExcel({ columns: MARGIN_DETAIL_COLUMNS, rows: marginDetailModal.data || [], filename: `Margin Detail - ${marginDetailModal.category}`, sheetName: 'Margin Detail' })}
                  className="flex items-center gap-1 px-3 py-1.5 bg-green-600 hover:bg-green-700 text-white rounded-lg text-sm"
                >
                  <FileSpreadsheet className="w-4 h-4"/>Excel
                </button>
                <button onClick={()=>setMarginDetailModal(null)} className={`p-1.5 rounded-lg ${darkMode?'hover:bg-gray-700':'hover:bg-gray-100'}`}><X className="w-5 h-5"/></button>
              </div>
            </div>
            <div className="overflow-auto flex-1 rounded-b-2xl">
              <table className="w-full text-xs">
                <thead className={`sticky top-0 ${darkMode?'bg-gray-700':'bg-blue-50'}`}>
                  <tr>{['SO Item','Product','Vendor','Sales','Purchase','Margin','%','Date'].map(h=>(
                    <th key={h} className={`px-3 py-2 text-center font-bold ${darkMode?'text-gray-200':'text-gray-700'}`}>{h}</th>
                  ))}</tr>
                </thead>
                <tbody className={`divide-y ${darkMode?'divide-gray-700':'divide-gray-100'}`}>
                  {(marginDetailModal.data||[]).map((t,i)=>(
                    <tr key={i} className={darkMode?'hover:bg-gray-700':'hover:bg-blue-50'}>
                      <td className="px-3 py-2 text-blue-600 font-medium whitespace-nowrap">{t.so_item||'-'}</td>
                      <td className={`px-3 py-2 max-w-[160px] truncate ${txt}`}>{t.product||'-'}</td>
                      <td className={`px-3 py-2 max-w-[120px] truncate ${txt2}`}>{t.vendor||'-'}</td>
                      <td className="px-3 py-2 text-right text-blue-600 whitespace-nowrap">{fmtCur(t.sales_amount)}</td>
                      <td className="px-3 py-2 text-right text-slate-700 whitespace-nowrap">{fmtCur(t.purchase_amount)}</td>
                      <td className={`px-3 py-2 text-right font-bold whitespace-nowrap ${t.margin<0?'text-red-600':t.margin>0?'text-green-600':'text-gray-400'}`}>{fmtCur(t.margin)}</td>
                      <td className={`px-3 py-2 text-right whitespace-nowrap ${t.margin<0?'text-red-500':t.margin>0?'text-green-500':'text-gray-400'}`}>{t.margin_pct!=null?`${t.margin_pct}%`:'—'}</td>
                      <td className={`px-3 py-2 ${txt2} whitespace-nowrap`}>{t.date?fmtDate(t.date):'-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ── Full-page loading overlay ────────────────────────────────────────
          Dashboard: only waits for the lightweight KPI stats. The heavier
          completed summary has its own inline loader so first paint is faster.
          Other pages: only on first visit when the page data array is still
          empty AND a fetch is in-flight — prevents overlay flashing on every
          filter / pagination change after initial load.                     */}
      {(() => {
        const PAGE_LABELS = {
          dashboard: 'Dashboard',
          'all-so': 'Pending Delivery',
          'item-registration': 'Item Registration',
          rfq: 'RFQ',
          import: 'Import',
          'vendor-control': 'Vendor Control',
          'all-registered-items': 'Registered Items',
        };

        // Dashboard: unblock the page as soon as the lightweight KPI stats are ready.
        // Completed summary keeps its own section-level loader, so the whole page
        // no longer stays covered while the heavier margin analytics are loading.
        const isDashboardLoading = activePage === 'dashboard' && (initialPageLoading || pageLoading || stats === null);

        // Other pages: only show when the page has never loaded data yet
        // (data array is still empty) AND a fetch is actively in-flight.
        // This prevents the overlay from re-appearing on pagination / filter.
        const isOtherPageFirstLoad = activePage !== 'dashboard' && pageLoading && (() => {
          if (activePage === 'all-so')               return allSOData.length === 0;
          if (activePage === 'item-registration')    return itemRegData.length === 0;
          if (activePage === 'rfq')                  return rfqData.length === 0;
          if (activePage === 'import')               return importData.length === 0;
          if (activePage === 'vendor-control')       return vendorControlData.length === 0;
          if (activePage === 'all-registered-items') return registeredItemsData.length === 0;
          return false;
        })();

        // Show the same loading popup on every page while its data request is running.
        // Dashboard still waits only for lightweight KPI stats, while table pages use
        // pageLoading from their own fetch function, so users get clear feedback when
        // changing page, pagination, filter, or search.
        const isOtherPageLoading = activePage !== 'dashboard' && pageLoading;
        const shouldShow = isDashboardLoading || isOtherPageLoading;
        if (!shouldShow) return null;
        const pageLabel = PAGE_LABELS[activePage] || 'Data';
        return (
          <div className="fixed inset-0 bg-black/50 z-[9999] flex items-center justify-center backdrop-blur-sm">
            <div className={`${darkMode?'bg-gray-800':'bg-white'} p-8 rounded-2xl shadow-2xl flex flex-col items-center gap-5 w-80 text-center`}>
              <div className="relative w-16 h-16">
                <div className="w-16 h-16 border-4 border-blue-200 rounded-full"/>
                <div className="absolute inset-0 w-16 h-16 border-4 border-blue-600 border-t-transparent rounded-full animate-spin"/>
              </div>
              <div>
                <p className={`font-bold text-lg mb-1 ${darkMode?'text-white':'text-gray-900'}`}>Memuat {pageLabel}</p>
                <p className={`text-sm ${darkMode?'text-gray-400':'text-gray-500'}`}>Sedang mengambil data dari server…</p>
                <p className={`text-xs mt-2 ${darkMode?'text-gray-500':'text-gray-400'}`}>Mohon tunggu sebentar</p>
              </div>
            </div>
          </div>
        );
      })()}

      {uploadProgress && (
        <div className="fixed inset-0 bg-black/60 z-[10000] flex items-center justify-center backdrop-blur-sm">
          <div className={`${darkMode?'bg-gray-800':'bg-white'} p-8 rounded-2xl shadow-2xl flex flex-col items-center gap-4 w-80`}>
            <div className="w-14 h-14 border-4 border-blue-600 border-t-transparent rounded-full animate-spin"/>
            <div className="w-full text-center">
              <p className={`font-bold text-lg mb-1 ${txt}`}>Uploading {uploadProgress.label}...</p>
              <p className={`text-xs mb-3 ${txt2}`}>Please wait, do not close the browser</p>
              <div className={`w-full rounded-full h-3 ${darkMode?'bg-gray-700':'bg-gray-200'}`}>
                <div className="bg-gradient-to-r from-blue-600 to-blue-400 h-3 rounded-full transition-all duration-300" style={{width:`${uploadProgress.pct}%`}}/>
              </div>
              <p className="text-blue-600 font-semibold mt-2">{uploadProgress.pct}%</p>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};

export default App;
