import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
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
  Clock, Wrench, Check, Link as LinkIcon, Pin, PinOff
} from 'lucide-react';
import axios from 'axios';
import { format, parseISO } from 'date-fns';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const BACKEND = import.meta.env.VITE_API_URL || 'http://127.0.0.1:5001';
const api = axios.create({ baseURL: BACKEND, timeout: 600000 });

const PIE_COLORS = ['#2563EB','#14B8A6','#22C55E','#EF4444','#06B6D4',
                    '#84CC16','#EC4899','#0EA5E9','#F43F5E','#94A3B8'];

const AGING_LABELS = ['0-30','30-90','90-180','180+'];
const AGING_COLORS = { '0-30':'#10B981','30-90':'#0EA5E9','90-180':'#F43F5E','180+':'#EF4444' };

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

// ─── Download Toast ────────────────────────────────────────────────────────
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
  const [pressed, setPressed] = useState(false);
  const handleClick = () => {
    setPressed(true);
    setTimeout(() => setPressed(false), 200);
    onClick && onClick();
  };
  return (
    <button
      onClick={handleClick}
      disabled={disabled}
      className={`${className} transition-all duration-100 ${pressed ? 'scale-95 brightness-90 shadow-inner' : 'scale-100'}`}
      style={{ transform: pressed ? 'scale(0.95)' : 'scale(1)' }}
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

// ─── MultiSelect dropdown — Excel-style (all checked by default) ─────────
const MultiSelect = ({ label, options, selected, onChange, darkMode, txt2, hideLabel = false }) => {
  const [open, setOpen] = useState(false);
  const [draftSelected, setDraftSelected] = useState([]);
  const [draftNone, setDraftNone] = useState(false);
  const [searchText, setSearchText] = useState('');
  const ref = useRef(null);

  const noSelection = selected === '__NONE__';
  const safeSelected = Array.isArray(selected) ? selected : [];
  const noneSelected = !noSelection && safeSelected.length === 0;
  const currentSelected = open ? draftSelected : safeSelected;
  const currentNone = open ? draftNone : noSelection;
  const currentAll = !currentNone && currentSelected.length === 0;
  const someSelected = !currentNone && currentSelected.length > 0 && currentSelected.length < options.length;

  const closeDropdown = () => {
    setOpen(false);
    setDraftSelected([]);
    setDraftNone(false);
    setSearchText('');
  };

  useEffect(() => {
    const handler = (e) => {
      if (ref.current && !ref.current.contains(e.target)) {
        closeDropdown();
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  useEffect(() => {
    if (!open) return;
    setDraftSelected(safeSelected);
    setDraftNone(noSelection);
  }, [open, selected]);

  const applySelection = () => {
    if (searchText.trim()) {
      const next = filteredOptions.length === 0
        ? '__NONE__'
        : filteredOptions.length === options.length
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
      // Uncheck all only changes the temporary dropdown state.
      // It is applied only when the user clicks Apply.
      setDraftSelected([]);
      setDraftNone(true);
    } else {
      setDraftSelected([]);
      setDraftNone(false);
    }
  };

  const toggle = (val) => {
    if (currentAll) {
      const next = options.filter(x => x !== val);
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
      const next = [val];
      setDraftSelected(next);
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
      const normalized = next.length === options.length ? [] : next;
      setDraftSelected(normalized);
      setDraftNone(false);
    }
  };

  const isChecked = (val) => {
    if (currentNone) return false;
    if (currentAll) return true;
    return currentSelected.includes(val);
  };

  const isAllChecked = currentAll;

  const filteredOptions = searchText.trim()
    ? options.filter(opt => String(opt).toLowerCase().includes(searchText.trim().toLowerCase()))
    : options;

  const displayLabel = currentNone
    ? `0 selected`
    : noneSelected
    ? `All ${label}`
    : safeSelected.length === 1
    ? String(safeSelected[0])
    : `${safeSelected.length} selected`;
  const hasActiveFilter = noSelection || !noneSelected;

  return (
    <div className="relative w-full min-w-0" ref={ref}>
      {!hideLabel && <label className={`block text-xs font-medium mb-1 ${txt2}`}>{label}</label>}
      <button onClick={()=>setOpen(o=>!o)} style={{cursor:'pointer'}}
        className={`w-full h-10 px-3 py-2 rounded-lg text-sm border text-left flex justify-between items-center transition-colors
          ${darkMode
            ? hasActiveFilter
              ? 'bg-amber-900/30 border-amber-500 text-amber-100 hover:bg-amber-900/40'
              : 'bg-gray-600 border-gray-500 text-white hover:bg-gray-500'
            : hasActiveFilter
              ? 'bg-amber-50 border-amber-300 text-amber-800 hover:bg-amber-100'
              : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'}`}>
        <span className={`truncate ${hasActiveFilter ? 'font-semibold' : ''}`}>{displayLabel}</span>
        <ChevronDown className="w-4 h-4 flex-shrink-0 ml-1"/>
      </button>
      {open && (
        <div
          className={`absolute z-50 mt-1 rounded-lg shadow-xl border overflow-hidden ${darkMode?'bg-gray-700 border-gray-600':'bg-white border-gray-200'}`}
          style={{ width: 'max(100%, 320px)', maxWidth: 'min(520px, calc(100vw - 32px))' }}
        >
          {/* Search input */}
          <div className={`px-2 pt-2 pb-1 border-b ${darkMode?'border-gray-600':'border-gray-100'}`}>
            <input
              type="text"
              value={searchText}
              onChange={e => setSearchText(e.target.value)}
              placeholder={`Search ${label}...`}
              autoFocus
              className={`w-full px-2 py-1.5 rounded text-xs border ${darkMode?'bg-gray-600 border-gray-500 text-white placeholder-gray-400':'bg-gray-50 border-gray-200 text-gray-800 placeholder-gray-400'}`}
              onClick={e => e.stopPropagation()}
            />
          </div>
          <div className="max-h-48 overflow-auto">
            <label style={{cursor:'pointer'}} className={`flex items-center gap-2 px-3 py-2 text-xs font-semibold border-b
              ${darkMode?'border-gray-600 hover:bg-gray-600 text-white':'border-gray-100 hover:bg-blue-50 text-gray-700'}`}>
              <input type="checkbox"
                checked={isAllChecked}
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
      )}
    </div>
  );
};

// ─── Search Input for SO / PO numbers ─────────────────────────────────────
const SearchInput = ({ placeholder, onSearch, darkMode, txt2, label }) => {
  const [open, setOpen] = useState(false);
  const [value, setValue] = useState('');
  const ref = useRef(null);

  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
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
        <div className={`absolute left-0 top-full mt-1 z-50 rounded-xl shadow-2xl border p-3 w-64 ${darkMode?'bg-gray-800 border-gray-700':'bg-white border-gray-200'}`}>
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

const FilterPanel = ({ children, darkMode, className = '' }) => (
  <div className={`mx-5 my-3 rounded-xl border p-3 ${darkMode ? 'border-gray-700 bg-gray-800/70' : 'border-gray-100 bg-[#f6f6f4]'} ${className}`}>
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


// ─── Delete Request Modal ──────────────────────────────────────────────────
const DeleteRequestModal = ({ darkMode, onClose, deleteForm, setDeleteForm, deleteFormError, onSubmit }) => {
  const bg = darkMode ? 'bg-gray-800 text-white' : 'bg-white text-gray-900';
  const inp = darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-gray-50 border-gray-300 text-gray-800';
  return (
    <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm" onClick={onClose}>
      <div className={`rounded-2xl shadow-2xl w-full max-w-md ${bg}`} onClick={e=>e.stopPropagation()}>
        <div className={`flex justify-between items-center px-6 py-4 border-b ${darkMode?'border-gray-700':'border-gray-200'}`}>
          <div className="flex items-center gap-2">
            <EyeOff className="w-5 h-5 text-slate-600"/>
            <h3 className="font-bold text-base">Hide from Summary</h3>
          </div>
          <button onClick={onClose} className={`p-1.5 rounded-lg ${darkMode?'hover:bg-gray-700':'hover:bg-gray-100'}`}><X className="w-5 h-5"/></button>
        </div>
        <div className="px-6 py-5 space-y-4">
          <div>
            <label className={`block text-xs font-semibold mb-1.5 ${darkMode?'text-gray-300':'text-gray-600'}`}>Data Type</label>
            <div className="flex gap-3">
              {['PO','SO'].map(t=>(
                <label key={t} className="flex items-center gap-2 cursor-pointer">
                  <input type="radio" name="ref_type" value={t} checked={deleteForm.ref_type===t}
                    onChange={()=>setDeleteForm(f=>({...f,ref_type:t}))} className="accent-blue-600"/>
                  <span className="text-sm font-medium">{t === 'PO' ? 'PO' : 'SO (Sales Order)'}</span>
                </label>
              ))}
            </div>
          </div>
          <div>
            <label className={`block text-xs font-semibold mb-1.5 ${darkMode?'text-gray-300':'text-gray-600'}`}>
              {deleteForm.ref_type === 'PO' ? 'PO Number' : 'SO Number / SO Item'}
            </label>
            <input
              type="text"
              value={deleteForm.ref_number}
              onChange={e=>setDeleteForm(f=>({...f,ref_number:e.target.value}))}
              placeholder={deleteForm.ref_type==='PO' ? 'e.g. 4570226161' : 'e.g. 9008988017-10'}
              className={`w-full px-3 py-2 rounded-lg text-sm border ${inp}`}
            />
          </div>
          <div>
            <label className={`block text-xs font-semibold mb-1.5 ${darkMode?'text-gray-300':'text-gray-600'}`}>Reason</label>
            <textarea
              value={deleteForm.reason}
              onChange={e=>setDeleteForm(f=>({...f,reason:e.target.value}))}
              placeholder="Enter reason why this data should be hidden from dashboard..."
              rows={3}
              className={`w-full px-3 py-2 rounded-lg text-sm border resize-none ${inp}`}
            />
          </div>
          {deleteFormError && (
            <div className="flex items-center gap-2 text-red-500 text-sm bg-red-50 rounded-lg px-3 py-2">
              <AlertCircle className="w-4 h-4 flex-shrink-0"/>{deleteFormError}
            </div>
          )}
        </div>
        <div className={`px-6 py-4 border-t flex justify-end gap-3 ${darkMode?'border-gray-700':'border-gray-200'}`}>
          <button onClick={onClose} className={`px-4 py-2 rounded-lg text-sm font-medium ${darkMode?'bg-gray-600 text-gray-200 hover:bg-gray-500':'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>Cancel</button>
          <button onClick={onSubmit} className="px-5 py-2 bg-slate-600 hover:bg-slate-700 text-white rounded-lg text-sm font-semibold flex items-center gap-2">
            <EyeOff className="w-4 h-4"/>Hide
          </button>
        </div>
      </div>
    </div>
  );
};

// ─── Hidden Items Panel ────────────────────────────────────────────────────
const HiddenItemsPanel = ({ darkMode, requests, onRestore, onClose }) => {
  const bg = darkMode ? 'bg-gray-800 text-white' : 'bg-white text-gray-900';
  const txt2 = darkMode ? 'text-gray-400' : 'text-gray-500';
  const hidden = requests.filter(r=>r.is_hidden);
  const fmtDt = (iso) => { try { return new Date(iso).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'}); } catch { return iso; } };
  return (
    <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm" onClick={onClose}>
      <div className={`rounded-2xl shadow-2xl w-full max-w-2xl max-h-[80vh] flex flex-col ${bg}`} onClick={e=>e.stopPropagation()}>
        <div className={`flex justify-between items-center px-6 py-4 border-b ${darkMode?'border-gray-700':'border-gray-200'}`}>
          <div className="flex items-center gap-2">
            <Eye className="w-5 h-5 text-blue-500"/>
            <h3 className="font-bold text-base">Items Hidden from Summary</h3>
            <span className={`text-xs font-medium px-2 py-0.5 rounded-full ${darkMode?'bg-gray-700 text-gray-300':'bg-gray-100 text-gray-600'}`}>{hidden.length} item</span>
          </div>
          <button onClick={onClose} className={`p-1.5 rounded-lg ${darkMode?'hover:bg-gray-700':'hover:bg-gray-100'}`}><X className="w-5 h-5"/></button>
        </div>
        <div className="overflow-auto flex-1 p-4">
          {hidden.length === 0 ? (
            <div className={`text-center py-12 ${txt2}`}>
              <Eye className="w-10 h-10 mx-auto mb-2 opacity-40"/>
              <p className="text-sm">No hidden data</p>
            </div>
          ) : (
            <div className="space-y-3">
              {hidden.map(r=>(
                <div key={r.id} className={`flex items-start justify-between gap-4 p-4 rounded-xl border ${darkMode?'bg-gray-700 border-gray-600':'bg-gray-50 border-gray-200'}`}>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2 mb-1">
                      <span className={`px-2 py-0.5 rounded text-xs font-bold ${r.ref_type==='PO'?'bg-red-100 text-red-700':'bg-slate-100 text-slate-700'}`}>{r.ref_type}</span>
                      <span className="font-semibold text-sm">{r.ref_number}</span>
                    </div>
                    <p className={`text-xs ${txt2} mb-1`}><span className="font-medium">Reason:</span> {r.reason}</p>
                    <p className={`text-xs ${txt2}`}>📅 {fmtDt(r.requested_at)}</p>
                  </div>
                  <button onClick={()=>onRestore(r)}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-green-600 hover:bg-green-700 text-white rounded-lg text-xs font-semibold flex-shrink-0">
                    <RotateCcw className="w-3.5 h-3.5"/>Restore
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

// ─── Date Range Filter ────────────────────────────────────────────────────
const DateRangeFilter = ({ darkMode, txt, txt2, card, onFilter, value, label = 'Filter SO Create Date', compact = false }) => {
  const [mode, setMode] = useState(value?.mode || 'all'); // all | today | week | month | year | range
  const [startDate, setStartDate] = useState(value?.start || '');
  const [endDate, setEndDate] = useState(value?.end || '');

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
    onFilter({ mode: 'all' });
  };

  return (
    <div data-tour="date-filter" className={`relative flex min-h-[64px] flex-wrap items-center gap-3 px-5 py-3 rounded-xl ${card} shadow ${compact ? 'mb-0' : 'mb-4'}`}>
      <Calendar className="w-4 h-4 text-blue-500 flex-shrink-0"/>
      <span className={`text-sm font-semibold ${txt} flex-shrink-0`}>{label}:</span>
      {/* Mode selector */}
      <div className="relative flex flex-wrap gap-1">
        {[
          ['all','All'], ['today','Today'], ['week','This Week'],
          ['month','This Month'], ['year','This Year'], ['range','Custom Date Range']
        ].map(([m, lbl]) => (
          <button key={m} onClick={() => setMode(m)}
            className={`px-3 py-1 rounded-full text-xs font-semibold transition-all
              ${mode === m ? 'bg-blue-600 text-white shadow' : darkMode ? 'bg-gray-700 text-gray-300 hover:bg-gray-600' : 'bg-gray-100 text-gray-600 hover:bg-blue-100'}`}>
            {lbl}
          </button>
        ))}
        {mode === 'range' && (
          <div className={`absolute left-0 top-full z-50 mt-2 flex items-center gap-2 rounded-xl border p-3 shadow-xl ${darkMode ? 'bg-gray-800 border-gray-600' : 'bg-white border-gray-200'}`}>
            <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)}
              className={`px-3 py-1.5 rounded-lg text-sm border ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-white border-gray-300'}`}/>
            <span className={`text-xs ${txt2}`}>to</span>
            <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)}
              className={`px-3 py-1.5 rounded-lg text-sm border ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-white border-gray-300'}`}/>
          </div>
        )}
      </div>
      {mode !== 'all' && (
        <button onClick={reset} className={`px-3 py-1.5 rounded-lg text-xs font-medium ${darkMode ? 'bg-gray-600 text-gray-200 hover:bg-gray-500' : 'bg-gray-200 text-gray-600 hover:bg-gray-300'}`}>
          Reset
        </button>
      )}
    </div>
  );
};

// ═══════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════
const App = () => {
  const [darkMode, setDarkMode] = useState(false);
  const [activePage, setActivePage] = useState('dashboard');
  const [showUploadDropdown, setShowUploadDropdown] = useState(false);
  const [sidebarExpanded, setSidebarExpanded] = useState(false);
  const uploadDropdownRef = useRef(null);
  const [frozenColumns, setFrozenColumns] = useState({});

  const [stats, setStats] = useState(null);
  const [summaryPendingTotal, setSummaryPendingTotal] = useState(null);
  const [agingData, setAgingData] = useState([]);
  const [allSOData, setAllSOData] = useState([]);
  const [approvalSOData, setApprovalSOData] = useState([]);
  const [picAggregations, setPicAggregations] = useState([]); // PIC aggregations from backend (all filtered data)
  const [soTotal, setSoTotal] = useState(0);
  const [soSubtotalAmount, setSoSubtotalAmount] = useState(0);
  const [soFilterOptions, setSoFilterOptions] = useState({ op_units: [], vendors: [], manufacturers: [], statuses: [], pics: [] });

  // SO filters
  const [soFilters, setSoFilters] = useState({ op_units: [], vendors: [], manufacturers: [], statuses: [], aging: [], pics: [] });
  const [soSearchNums, setSoSearchNums] = useState([]); // search SO Item
  const [soMarginFilter, setSoMarginFilter] = useState('all'); // 'all' | 'positive' | 'negative'
  const [soSortOrder, setSoSortOrder] = useState('oldest'); // 'oldest' | 'newest'
  const [soPage, setSoPage] = useState(1);
  const [soPerPage, setSoPerPage] = useState(10);
  const [pendingPicHighlight, setPendingPicHighlight] = useState('');

  // SO Approval Status filters (same as Open SO except Vendor Name)
  const [approvalFilters, setApprovalFilters] = useState({ op_units: [], statuses: [], aging: [] });
  const [approvalSearchNums, setApprovalSearchNums] = useState([]);
  const [approvalPage, setApprovalPage] = useState(1);
  const [approvalPerPage, setApprovalPerPage] = useState(10);

  // Item Registration
  const [itemRegData, setItemRegData] = useState([]);
  const [itemRegTotal, setItemRegTotal] = useState(0);
  const [itemRegPage, setItemRegPage] = useState(1);
  const [itemRegPerPage, setItemRegPerPage] = useState(10);
  const [itemRegSearch, setItemRegSearch] = useState([]);
  const [itemRegAppliedSearch, setItemRegAppliedSearch] = useState([]);
  const [itemRegLastUpdated, setItemRegLastUpdated] = useState(null);
  const [itemRegFilters, setItemRegFilters] = useState({ clients: [], categories: [], pics: [], proc_statuses: [], mfr_names: [], existing_owners: [] });
  const [itemRegOptions, setItemRegOptions] = useState({ clients: [], categories: [], pics: [], proc_statuses: [], mfr_names: [], existing_owners: [] });
  const [itemRegMissingPicKpis, setItemRegMissingPicKpis] = useState([]);
  const [itemRegPicHighlight, setItemRegPicHighlight] = useState('');

  // RFQ
  const [rfqData, setRfqData] = useState([]);
  const [rfqTotal, setRfqTotal] = useState(0);
  const [rfqPage, setRfqPage] = useState(1);
  const [rfqPerPage, setRfqPerPage] = useState(10);
  const [rfqSortOrder, setRfqSortOrder] = useState('newest');
  const [rfqSearch, setRfqSearch] = useState('');
  const [rfqAppliedSearch, setRfqAppliedSearch] = useState('');
  const [rfqColumns, setRfqColumns] = useState([]);
  const [rfqSimilarityColumns, setRfqSimilarityColumns] = useState([]);
  const [rfqShowSimilarity, setRfqShowSimilarity] = useState(false);
  const [rfqEditableFields, setRfqEditableFields] = useState([]);
  const [rfqPicKpis, setRfqPicKpis] = useState([]);
  const [rfqPicFilter, setRfqPicFilter] = useState('');
  const [rfqFilters, setRfqFilters] = useState({ checks: [], clients: [], brands: [], purchase_pics: [], vendors: [] });
  const [rfqOptions, setRfqOptions] = useState({ checks: [], clients: [], brands: [], purchase_pics: [], vendors: [] });
  const [rfqSelectedCell, setRfqSelectedCell] = useState(null);
  const [rfqSimilarAction, setRfqSimilarAction] = useState(null);
  const [rfqLastUpdated, setRfqLastUpdated] = useState(null);
  const [rfqEditedRowKeys, setRfqEditedRowKeys] = useState(new Set());

  // All Registered Items
  const [registeredItemsData, setRegisteredItemsData] = useState([]);
  const [registeredItemsTotal, setRegisteredItemsTotal] = useState(0);
  const [registeredItemsPage, setRegisteredItemsPage] = useState(1);
  const [registeredItemsPerPage, setRegisteredItemsPerPage] = useState(10);
  const [registeredItemsSearch, setRegisteredItemsSearch] = useState('');
  const [registeredItemsAppliedSearch, setRegisteredItemsAppliedSearch] = useState('');
  const [registeredItemsProdIds, setRegisteredItemsProdIds] = useState([]);
  const [registeredItemsAppliedProdIds, setRegisteredItemsAppliedProdIds] = useState([]);

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
  const [vendorControlLastUpdated, setVendorControlLastUpdated] = useState(null);
  const [vendorPasswordVisible, setVendorPasswordVisible] = useState({});

  const [loading, setLoading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(null);
  const [toasts, setToasts] = useState([]);
  const [modal, setModal] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState('');
  const [downloadToast, setDownloadToast] = useState(null);

  // Delete request / hide feature
  const [deleteRequests, setDeleteRequests] = useState([]);
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [showHiddenPanel, setShowHiddenPanel] = useState(false);
  const [deleteForm, setDeleteForm] = useState({ ref_type: 'PO', ref_number: '', reason: '' });
  const [deleteFormError, setDeleteFormError] = useState('');
  const [completedData, setCompletedData] = useState(null);
  const [completedYear, setCompletedYear] = useState('all');
  const [completedLoading, setCompletedLoading] = useState(false);
  const [completedLoaded, setCompletedLoaded] = useState(false);
  const [showHideMenu, setShowHideMenu] = useState(false);
  const [marginDetailModal, setMarginDetailModal] = useState(null); // {category, data}
  const [picDbStatus, setPicDbStatus] = useState(null); // {product_id_count, master_pic_count, last_product_id_upload, last_pic_update}
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
  const getFrozenIndex = (value) => (typeof value === 'object' && value ? value.index : value);
  const getFreezeLeft = (event) => {
    const headerCell = event?.currentTarget?.closest?.('th');
    const scrollBox = event?.currentTarget?.closest?.('.overflow-x-auto');
    if (!headerCell || !scrollBox) return 0;
    const headerRect = headerCell.getBoundingClientRect();
    const scrollRect = scrollBox.getBoundingClientRect();
    const maxLeft = Math.max(0, scrollBox.clientWidth - headerRect.width);
    return Math.max(0, Math.min(Math.round(headerRect.left - scrollRect.left), maxLeft));
  };
  const toggleFrozenColumn = useCallback((tableKey, colIndex, event) => {
    const left = getFreezeLeft(event);
    setFrozenColumns(prev => {
      const activeIndex = getFrozenIndex(prev[tableKey]);
      return { ...prev, [tableKey]: activeIndex === colIndex ? null : { index: colIndex, left } };
    });
  }, []);
  const renderFreezeHeader = (tableKey, colIndex, label) => {
    const active = getFrozenIndex(frozenColumns[tableKey]) === colIndex;
    return (
      <div className="freeze-header group relative flex min-h-8 w-full min-w-0 items-center justify-center">
        <span className="freeze-header-label max-w-full text-center leading-tight">{label}</span>
        <button
          type="button"
          aria-label={active ? `Unfreeze ${label}` : `Freeze ${label}`}
          title={active ? `Unfreeze ${label}` : `Freeze ${label}`}
          onClick={(e) => { e.stopPropagation(); toggleFrozenColumn(tableKey, colIndex, e); }}
          className={`absolute right-0 top-1/2 inline-flex h-5 w-5 -translate-y-1/2 items-center justify-center rounded-md border opacity-0 shadow-sm transition-all group-hover:opacity-100 group-focus-within:opacity-100 ${active ? 'border-amber-300 bg-amber-100 text-amber-700' : darkMode ? 'border-gray-600 bg-gray-700/90 text-gray-300 hover:bg-gray-600' : 'border-slate-200 bg-white/95 text-slate-500 hover:bg-slate-100'}`}
        >
          {active ? <PinOff className="h-3 w-3" /> : <Pin className="h-3 w-3" />}
        </button>
      </div>
    );
  };
  const frozenColumnCss = useMemo(() => Object.entries(frozenColumns)
    .map(([tableKey, spec]) => ({ tableKey, idx: getFrozenIndex(spec), left: (typeof spec === 'object' && spec ? spec.left : 0) }))
    .filter(({ idx }) => Number(idx) > 0)
    .map(({ tableKey, idx, left }) => `
      .freeze-table-${tableKey} th:nth-child(${idx}),
      .freeze-table-${tableKey} td:nth-child(${idx}) {
        position: sticky;
        left: ${Number(left) || 0}px;
        z-index: 25;
        box-shadow: 10px 0 14px -14px rgba(15, 23, 42, 0.85);
        background-clip: padding-box;
      }
      .freeze-table-${tableKey} thead th:nth-child(${idx}) {
        z-index: 45;
      }
      .data-table-page:not(.data-table-page-dark) .freeze-table-${tableKey} td:nth-child(${idx}):not([class*="bg-"]) {
        background: #ffffff;
      }
      .data-table-page:not(.data-table-page-dark) .freeze-table-${tableKey} thead th:nth-child(${idx}):not([class*="bg-"]) {
        background: #e2e8f0;
      }
      .data-table-page-dark .freeze-table-${tableKey} td:nth-child(${idx}):not([class*="bg-"]) {
        background: #1f2937;
      }
      .data-table-page-dark .freeze-table-${tableKey} thead th:nth-child(${idx}):not([class*="bg-"]) {
        background: #374151;
      }
    `)
    .join('\n'), [frozenColumns]);
  const hideMenuRef = useRef(null);

  // ── Global SO Create Date filter (shared across Dashboard / All SO / Delivery Completed)
  const [globalDateFilter, setGlobalDateFilter] = useState({ mode: 'all' });
  const [globalClientFilter, setGlobalClientFilter] = useState([]);
  const [globalPicFilter, setGlobalPicFilter] = useState([]);
  const [dashboardFilterOptions, setDashboardFilterOptions] = useState({ clients: [], pics: [] });
  // Aliases kept so existing references continue to compile.
  const dashDateFilter      = globalDateFilter;
  const setDashDateFilter   = setGlobalDateFilter;
  const soDateFilter        = globalDateFilter;
  const setSODateFilter     = setGlobalDateFilter;
  const completedDateFilter = globalDateFilter;
  const setCompletedDateFilter = setGlobalDateFilter;

  // Click-outside handlers
  useEffect(() => {
    const handler = (e) => { if (hideMenuRef.current && !hideMenuRef.current.contains(e.target)) setShowHideMenu(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  useEffect(() => {
    const handler = (e) => { if (uploadDropdownRef.current && !uploadDropdownRef.current.contains(e.target)) setShowUploadDropdown(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const addToast = useCallback((message, type='success') => {
    const id = Date.now(); setToasts(t => [...t, { id, message, type }]);
  }, []);
  const removeToast = useCallback((id) => setToasts(t => t.filter(x => x.id !== id)), []);

  function appendMultiParam(params, key, value) {
    if (value === '__NONE__') {
      params.append(key, '__NONE_PLACEHOLDER__');
      return;
    }
    (Array.isArray(value) ? value : []).forEach(v => params.append(key, v));
  }

  const fetchDashboard = useCallback(async (dateFilter) => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      const f = dateFilter || globalDateFilter;
      Object.entries(dateFilterParams(f)).forEach(([key, value]) => { if (value) params.append(key, value); });
      appendMultiParam(params, 'client', globalClientFilter);
      appendMultiParam(params, 'pic', globalPicFilter);
      const qs = params.toString() ? `?${params}` : '';
      const completedParams = new URLSearchParams();
      if (!f || f.mode === 'all') completedParams.set('date_year', String(new Date().getFullYear()));
      params.forEach((value, key) => completedParams.append(key, value));
      const completedQs = completedParams.toString();
      const pendingParams = new URLSearchParams();
      Object.entries(dateFilterParams(f)).forEach(([key, value]) => { if (value) pendingParams.append(key, value); });
      appendMultiParam(pendingParams, 'client', globalClientFilter);
      appendMultiParam(pendingParams, 'global_pic', globalPicFilter);
      pendingParams.set('page', '1');
      pendingParams.set('per_page', '1');
      const [sRes, aRes, cRes, pendingRes] = await Promise.all([
        api.get(`/api/dashboard/stats${qs}`),
        api.get(`/api/data/aging${qs}`),
        api.get(`/api/completed/summary?${completedQs}`),
        api.get(`/api/data/all-so?${pendingParams}`)
      ]);
      setStats(sRes.data);
      setSummaryPendingTotal(Number(pendingRes.data?.total) || 0);
      setDashboardFilterOptions(sRes.data?.filters || { clients: [], pics: [] });
      setAgingData(Array.isArray(aRes.data) ? aRes.data : []);
      setCompletedData(cRes.data);
      setCompletedLoaded(true);
    } catch (e) {
      addToast(`Error: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast, globalDateFilter, globalClientFilter, globalPicFilter]);

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

  const fetchItemRegistration = useCallback(async (page = itemRegPage, perPage = itemRegPerPage, search = itemRegAppliedSearch, filters = itemRegFilters, kpiPic = itemRegPicHighlight) => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      if (Array.isArray(search)) search.forEach(v => params.append('req_no', v));
      else if (search) params.append('search', search);
      resolveFilter(filters.clients).forEach(v => params.append('item_client', v));
      resolveFilter(filters.categories).forEach(v => params.append('category', v));
      resolveFilter(filters.pics).forEach(v => params.append('pic', v));
      resolveFilter(filters.proc_statuses).forEach(v => params.append('proc_status', v));
      resolveFilter(filters.mfr_names).forEach(v => params.append('mfr_name', v));
      resolveFilter(filters.existing_owners).forEach(v => params.append('existing_owner', v));
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
        mfr_names: res.data.mfr_name_options || [],
        existing_owners: res.data.existing_owner_options || []
      });
    } catch (e) {
      addToast(`Failed to load Item Registration: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast, itemRegPage, itemRegPerPage, itemRegAppliedSearch, itemRegFilters, itemRegPicHighlight]);

  const fetchRegisteredItems = useCallback(async (
    page = registeredItemsPage,
    perPage = registeredItemsPerPage,
    search = registeredItemsAppliedSearch,
    prodIds = registeredItemsAppliedProdIds
  ) => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      if (search) params.append('search', search);
      (prodIds || []).forEach(v => params.append('prod_id', v));
      const res = await api.get(`/api/all-registered-items?${params}`);
      setRegisteredItemsData(Array.isArray(res.data.data) ? res.data.data : []);
      setRegisteredItemsTotal(res.data.total || 0);
    } catch (e) {
      addToast(`Failed to load All Registered Items: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [
    addToast,
    registeredItemsPage,
    registeredItemsPerPage,
    registeredItemsAppliedSearch,
    registeredItemsAppliedProdIds
  ]);

  const fetchRFQData = useCallback(async (page = rfqPage, perPage = rfqPerPage, search = rfqAppliedSearch, refresh = false, filters = rfqFilters, pic = rfqPicFilter, showSimilarity = rfqShowSimilarity, sortOrder = rfqSortOrder) => {
    setRfqEditedRowKeys(new Set());
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      if (search) params.append('search', search);
      if (sortOrder) params.append('sort_order', sortOrder);
      if (refresh) params.append('refresh', '1');
      if (pic) params.append('pic', pic);
      if (showSimilarity) params.append('similarity', '1');
      resolveFilter(filters.checks).forEach(v => params.append('check', v));
      resolveFilter(filters.clients).forEach(v => params.append('client_name', v));
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
      const nextOptions = res.data.filters || { checks: [], clients: [], brands: [], purchase_pics: [], vendors: [] };
      setRfqOptions({
        ...nextOptions,
        purchase_pics: (nextOptions.purchase_pics || []).filter(v => String(v || '').trim().toLowerCase() !== 'unassigned')
      });
      setRfqLastUpdated(res.data.last_updated || null);
    } catch (e) {
      addToast(`Failed to load RFQ: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast, rfqPage, rfqPerPage, rfqAppliedSearch, rfqFilters, rfqPicFilter, rfqShowSimilarity, rfqSortOrder]);

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

  // ─── Delete Request API functions ────────────────────────────────────────
  const fetchDeleteRequests = useCallback(async () => {
    try {
      const res = await api.get('/api/delete-requests');
      setDeleteRequests(Array.isArray(res.data) ? res.data : []);
    } catch (e) { /* silent */ }
  }, []);

  const submitDeleteRequest = async () => {
    setDeleteFormError('');
    if (!deleteForm.ref_number.trim()) { setDeleteFormError('Reference number is required'); return; }
    if (!deleteForm.reason.trim()) { setDeleteFormError('Reason is required'); return; }
    try {
      await api.post('/api/delete-requests', deleteForm);
      addToast(`✅ ${deleteForm.ref_type} ${deleteForm.ref_number} successfully hidden from dashboard`, 'success');
      setDeleteForm({ ref_type: 'PO', ref_number: '', reason: '' });
      setShowDeleteModal(false);
      fetchDeleteRequests();
      fetchDashboard();
    } catch (e) {
      setDeleteFormError(e.response?.data?.error || e.message);
    }
  };

  const restoreDeleteRequest = async (req) => {
    try {
      await api.put(`/api/delete-requests/${req.id}/restore`);
      addToast(`✅ ${req.ref_type} ${req.ref_number} successfully restored`, 'success');
      fetchDeleteRequests();
      fetchDashboard();
    } catch (e) {
      addToast(`❌ Failed to restore: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  useEffect(() => { fetchDashboard(); fetchDeleteRequests(); fetchPicDbStatus(); }, []);
  useEffect(() => {
    if (activePage === 'all-so') {
      fetchSOData(soFilters, soPage, soPerPage, soSearchNums, soMarginFilter, soDateFilter, soSortOrder);
    }
  }, [activePage, soSortOrder, soPage, soPerPage, soFilters, soSearchNums, soMarginFilter, soDateFilter, globalClientFilter, globalPicFilter, fetchSOData]);

  useEffect(() => {
    if (activePage === 'item-registration') {
      fetchItemRegistration(itemRegPage, itemRegPerPage, itemRegAppliedSearch, itemRegFilters);
    }
  }, [activePage, itemRegPage, itemRegPerPage, itemRegAppliedSearch, itemRegFilters, itemRegPicHighlight, fetchItemRegistration]);

  useEffect(() => {
    if (activePage === 'rfq') {
      fetchRFQData(rfqPage, rfqPerPage, rfqAppliedSearch, false, rfqFilters, rfqPicFilter, rfqShowSimilarity);
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
        registeredItemsAppliedProdIds
      );
    }
  }, [
    activePage,
    registeredItemsPage,
    registeredItemsPerPage,
    registeredItemsAppliedSearch,
    registeredItemsAppliedProdIds,
    fetchRegisteredItems
  ]);

  // Refetch dashboard whenever the global SO Create Date filter changes
  // (skip the very first run since the mount effect above already fetched).
  const skipFirstFilterRefetch = useRef(true);
  useEffect(() => {
    if (skipFirstFilterRefetch.current) { skipFirstFilterRefetch.current = false; return; }
    fetchDashboard(globalDateFilter);
  }, [globalDateFilter, globalClientFilter, globalPicFilter, fetchDashboard]);

  const handleUpload = async (e, type) => {
    const files = Array.from(e.target.files || []); if (!files.length) return;
    e.target.value = '';
    const label = files.length > 1 ? `SO - Search Client Odr (${files.length} files)` : 'SO - Search Client Odr';
    const endpoint = '/api/upload/smro';

    // ── Client-side header validation ──────────────────────────────────
    const REQUIRED_HEADERS = {
      scor: {
        'SO Number':      ['so number','so no','so no.','so','sales order','sales order number','no so','nomor so'],
        'SO Item':        ['so item no','item no','line','so line','so item'],
        'SO Status':      ['so status','status','order status'],
        'Operation Unit': ['operation unit name','op unit','client name','client','operation unit'],
        'Vendor Name':    ['vendor name','vendor','supplier'],
        'Customer PO':    ['customer po number','customer po','po ref','po reference'],
        'Sales Amount':   ['sales amount(exclude tax)','sales amount','amount','total'],
        'SO Create Date': ['so create date','order date','so date','create date'],
      }
    };

    try {
      for (const file of files) {
        const arrayBuffer = await file.arrayBuffer();
        const wb = XLSX.read(arrayBuffer, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const headerRow = (jsonData[0] || []).map(h => String(h || '').trim().toLowerCase());

        const reqHeaders = REQUIRED_HEADERS[type];
        const missing = [];
        for (const [friendlyName, aliases] of Object.entries(reqHeaders)) {
          const found = aliases.some(alias => headerRow.includes(alias.toLowerCase()));
          if (!found) missing.push(friendlyName);
        }
        if (missing.length > 3) {
          addToast(
            `❌ Invalid file ${file.name} — ${missing.length} required columns not found: ${missing.join(', ')}. Please check the ${label} file is correct and try again.`,
            'error'
          );
          return;
        }
      }
    } catch (readErr) {
      addToast(`❌ Failed to read file: ${readErr.message}`, 'error');
      return;
    }
    // ── End client-side header validation ──────────────────────────────

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
      fetchDashboard();
      if (activePage === 'all-so') fetchSOData(soFilters, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
      setSoPage(1);
    } catch (e) {
      setUploadProgress(null);
      addToast(`❌ Failed to upload ${label}: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  const fetchPicDbStatus = async () => {
    try {
      const res = await api.get('/api/master-pic/status');
      setPicDbStatus(res.data);
    } catch {}
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
      fetchPicDbStatus();
      fetchSOData(soFilters, soPage, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
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
      setPicUploadMsg(`✅ Master PIC (${d.files || files.length} file): +${d.added} added, ${d.updated} updated (total categories: ${d.total_categories}). SO rows updated: ${d.so_pic_refreshed}.`);
      fetchPicDbStatus();
      fetchSOData(soFilters, soPage, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
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
      setItemRegPage(1);
      fetchDashboard();
      fetchItemRegistration(1, itemRegPerPage, itemRegAppliedSearch, itemRegFilters);
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
      fetchItemRegistration(itemRegPage, itemRegPerPage, itemRegAppliedSearch, itemRegFilters);
    } catch (e) {
      setUploadProgress(null);
      addToast(`Failed to upload Item Registration batch: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  const downloadBlob = async (url, filename, label) => {
    const toastId = Date.now();
    setDownloadToast({ id: toastId, message: `Downloading ${label || filename}...` });
    try {
      const res = await api.get(url, { responseType: 'blob' });
      const link = document.createElement('a');
      link.href = window.URL.createObjectURL(new Blob([res.data]));
      link.setAttribute('download', filename);
      document.body.appendChild(link); link.click(); link.remove();
      setDownloadToast(null);
      addToast(`✅ File "${filename}" downloaded successfully`, 'success');
    } catch (e) {
      setDownloadToast(null);
      addToast('❌ Failed to download file', 'error');
    }
  };

  const downloadItemRegistrationTemplate = () => {
    const p = new URLSearchParams();
    (itemRegAppliedSearch || []).forEach(v => p.append('req_no', v));
    resolveFilter(itemRegFilters.clients).forEach(v => p.append('item_client', v));
    resolveFilter(itemRegFilters.categories).forEach(v => p.append('category', v));
    resolveFilter(itemRegFilters.pics).forEach(v => p.append('pic', v));
    resolveFilter(itemRegFilters.proc_statuses).forEach(v => p.append('proc_status', v));
    resolveFilter(itemRegFilters.mfr_names).forEach(v => p.append('mfr_name', v));
    resolveFilter(itemRegFilters.existing_owners).forEach(v => p.append('existing_owner', v));
    if (itemRegPicHighlight) p.append('kpi_pic', itemRegPicHighlight);
    downloadBlob(`/api/item-registration/template?${p}`, `Template_ItemRegistration_BatchUpload_${new Date().toISOString().slice(0,10)}.xlsx`, 'Item Registration Batch Upload Template');
  };

  const downloadItemRegistrationExcel = () => {
    const p = new URLSearchParams();
    (itemRegAppliedSearch || []).forEach(v => p.append('req_no', v));
    resolveFilter(itemRegFilters.clients).forEach(v => p.append('item_client', v));
    resolveFilter(itemRegFilters.categories).forEach(v => p.append('category', v));
    resolveFilter(itemRegFilters.pics).forEach(v => p.append('pic', v));
    resolveFilter(itemRegFilters.proc_statuses).forEach(v => p.append('proc_status', v));
    resolveFilter(itemRegFilters.mfr_names).forEach(v => p.append('mfr_name', v));
    resolveFilter(itemRegFilters.existing_owners).forEach(v => p.append('existing_owner', v));
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

  const downloadHideTemplate = (type) => {
    setShowHideMenu(false);
    downloadBlob(`/api/template/hide?type=${type}`, `Template_Hide_${type === 'SO' ? 'SO' : 'PO'}.xlsx`, `Template Hide ${type}`);
  };

  const handleHideBatchUpload = async (e, type) => {
    const file = e.target.files[0]; if (!file) return;
    e.target.value = '';
    setShowHideMenu(false);
    const fd = new FormData();
    fd.append('file', file);
    fd.append('type', type);
    setUploadProgress({ label: `Hide ${type} Batch`, pct: 0 });
    try {
      const res = await api.post('/api/upload/hide-batch', fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        onUploadProgress: (ev) => setUploadProgress({ label: `Hide ${type} Batch`, pct: Math.round(ev.loaded*100/(ev.total||ev.loaded)) })
      });
      setUploadProgress(null);
      addToast(`✅ ${res.data.message}`, 'success');
      fetchDeleteRequests();
      fetchDashboard();
    } catch (e) {
      setUploadProgress(null);
      addToast(`❌ Failed to upload hide batch: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

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
      const margin = (Number(so.sales_amount) || 0) - poAmount;
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
      if (!quiet && res.data?.sheet_sync && res.data.sheet_sync.synced === false) {
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
      setRfqData(previousRows);
      if (!quiet) addToast(`Failed to update RFQ: ${e.response?.data?.error || e.message}`, 'error');
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
      if (res.data?.sheet_sync && res.data.sheet_sync.synced === false) {
        addToast(`RFQ batch updated locally. Sheet sync not active: ${res.data.sheet_sync.reason}`, 'warning');
      }
      if (res.data?.skipped?.length) {
        addToast(`RFQ batch skipped ${res.data.skipped.length} cells`, 'warning');
      }
      return true;
    } catch (e) {
      setRfqData(previousRows);
      addToast(`Failed to update RFQ batch: ${e.response?.data?.error || e.message}`, 'error');
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
    fetchSOData(next, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
    window.scrollTo({ top: 0, behavior: 'smooth' });
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
  const tblHd = darkMode ? 'bg-gray-800/60' : 'bg-slate-50';
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
  const globalSlicerPages = new Set(['dashboard', 'all-so']);
  const renderGlobalSlicer = () => {
    if (!globalSlicerPages.has(activePage)) return null;
    return (
      <div className="mb-5 grid grid-cols-1 gap-3 2xl:grid-cols-[minmax(560px,1fr)_minmax(560px,1fr)]">
        <DateRangeFilter
          darkMode={darkMode}
          txt={txt}
          txt2={txt2}
          card={card}
          value={globalDateFilter}
          label="Filter SO Create Date"
          compact
          onFilter={(f) => setGlobalDateFilter(f)}
        />
        <div className={`grid min-h-[64px] grid-cols-1 gap-3 px-5 py-3 rounded-xl ${card} shadow sm:grid-cols-[minmax(220px,1fr)_minmax(200px,0.85fr)_120px] sm:items-end`}>
          <div className="min-w-0">
            <label className={`mb-1 block text-xs font-semibold ${txt}`}>Client Nm.</label>
            <MultiSelect label="Client Nm." options={dashboardFilterOptions.clients || []} selected={globalClientFilter} onChange={setGlobalClientFilter} darkMode={darkMode} txt2={txt2} hideLabel />
          </div>
          <div className="min-w-0">
            <label className={`mb-1 block text-xs font-semibold ${txt}`}>PIC Name</label>
            <MultiSelect label="PIC Name" options={dashboardFilterOptions.pics || []} selected={globalPicFilter} onChange={setGlobalPicFilter} darkMode={darkMode} txt2={txt2} hideLabel />
          </div>
          <button
            type="button"
            onClick={() => { setGlobalDateFilter({ mode: 'all' }); setGlobalClientFilter([]); setGlobalPicFilter([]); }}
            className={`h-10 w-full px-4 rounded-lg text-sm font-medium shadow-sm flex items-center justify-center whitespace-nowrap ${darkMode ? 'bg-gray-500 text-gray-100 hover:bg-gray-400' : 'bg-gray-400 text-white hover:bg-gray-500'}`}
          >
            Clear
          </button>
        </div>
      </div>
    );
  };

  const renderDashboardOverview = () => {
    const d = completedData || {};
    
    // Create full 12-month array for current year
    const currentYear = new Date().getFullYear();
    const monthlyDataMap = {};
    (d.monthly_trend || []).forEach(m => {
      if (m.month) monthlyDataMap[m.month] = m;
    });
    
    const monthlyCompleted = Array.from({ length: 12 }, (_, i) => {
      const monthKey = `${currentYear}-${String(i + 1).padStart(2, '0')}`;
      const monthName = format(new Date(currentYear, i, 1), 'MMMM');
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
    const completedPoCountThisYear = monthlyCompleted.reduce((sum, m) => sum + (Number(m.count) || 0), 0);
    const completedPoAmountThisYear = monthlyCompleted.reduce((sum, m) => sum + (Number(m.purchase_amount) || 0), 0);
    
    const marginPct = d.total_sales ? ((d.total_margin || 0) / d.total_sales * 100) : null;
    const fmtM = (v) => v >= 1e9 ? `${(v/1e9).toFixed(1)}B` : v >= 1e6 ? `${(v/1e6).toFixed(1)}M` : v >= 1e3 ? `${(v/1e3).toFixed(0)}K` : String(Math.round(v || 0));
    const purchaseYoyYears = (d.purchase_yoy_years && d.purchase_yoy_years.length)
      ? d.purchase_yoy_years
      : [currentYear - 1, currentYear - 2];
    const purchaseYoyData = (d.purchase_yoy_trend && d.purchase_yoy_trend.length)
      ? d.purchase_yoy_trend
      : Array.from({ length: 12 }, (_, i) => ({
          month: i + 1,
          month_label: format(new Date(currentYear, i, 1), 'MMMM'),
          ...purchaseYoyYears.reduce((acc, year) => ({ ...acc, [`purchase_${year}`]: 0 }), {})
        }));
    const activePurchaseYoyYears = purchaseYoyYears.filter(year =>
      purchaseYoyData.some(row => Number(row[`purchase_${year}`] || 0) > 0)
    );
    const completedTrendData = monthlyCompleted.map((m, i) => {
      const yoyRow = purchaseYoyData[i] || {};
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
    const barList = (rows, labelKey, valueKey, label, color) => (
      <ResponsiveContainer width="100%" height={260}>
        <BarChart data={(rows || []).map(r => ({...r, shortLabel: String(r[labelKey] || '-').slice(0, 32)}))} layout="vertical" margin={{top: 8, right: 18, left: 8, bottom: 8}}>
          <CartesianGrid strokeDasharray="3 3" horizontal={true} stroke={darkMode?'#374151':'#E5E7EB'}/>
          <XAxis type="number" stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={12} tickFormatter={fmtM}/>
          <YAxis type="category" dataKey="shortLabel" width={190} stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={12} tick={{fontSize: 12, textAnchor: 'end'}} tickMargin={8}/>
          <Tooltip formatter={(v)=>[fmtCur(v), label]} labelFormatter={(_, payload)=>payload?.[0]?.payload?.[labelKey] || '-'} contentStyle={{background:darkMode?'#1F2937':'#fff',border:'none',borderRadius:8,fontSize:12}}/>
          <Bar dataKey={valueKey} name={label} fill={color} radius={[0,6,6,0]} isAnimationActive={false}/>
        </BarChart>
      </ResponsiveContainer>
    );

    return (
      <div className="space-y-5">
        <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-5 gap-4">
          {[
            { label:'Total PO', value: fmtNum(completedPoCountThisYear), sub: `Delivery Complete ${currentYear}`, icon:<FileText className="w-5 h-5"/> },
            { label:'PO Amount', value: fmtCurShort(completedPoAmountThisYear), sub: fmtCur(completedPoAmountThisYear), icon:<Coins className="w-5 h-5"/> },
            { label:'Sales Amount', value: fmtCurShort(d.total_sales), sub: fmtCur(d.total_sales), icon:<Wallet className="w-5 h-5"/> },
            { label:'Margin', value: fmtCurShort(d.total_margin), sub: marginPct == null ? 'Avg margin -' : `Avg margin ${marginPct.toFixed(1)}%`, icon:<TrendingUp className="w-5 h-5"/> },
            { label:'Total Pending Delivery', value: fmtNum(summaryPendingTotal ?? stats?.total_so_count), sub: 'Pending delivery records', icon:<Clock className="w-5 h-5"/>, goPending:true },
          ].map((k,i)=>{
            const Wrapper = k.goPending ? 'button' : 'div';
            return <Wrapper key={i} type={k.goPending ? 'button' : undefined} onClick={k.goPending ? () => { setActivePage('all-so'); setSoPage(1); fetchSOData(soFilters, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter); window.scrollTo({top:0, behavior:'smooth'}); } : undefined} className={`p-5 rounded-2xl text-left ${card} ${k.goPending ? 'cursor-pointer transition-all hover:border-blue-300' : ''}`}><div className="flex items-start justify-between gap-3"><div className="min-w-0"><p className={`text-sm font-medium ${txt2}`}>{k.label}</p><h3 className={`text-2xl font-bold mt-1 ${kpiValue}`}>{k.value}</h3><p className={`text-xs mt-1 ${txt2}`}>{k.sub}</p></div><div className={`p-2.5 rounded-xl ${neutralIcon}`}>{k.icon}</div></div></Wrapper>;
          })}
        </div>
        <div className={`p-5 rounded-2xl ${card}`}>
          <h3 className={`text-base font-bold mb-1 ${txt}`}>Monthly Trend Delivery Complete</h3>
          <p className={`text-xs mb-4 ${txt2}`}>Current year sales and purchase amount, with YoY purchase amount lines.</p>
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
          <h3 className={`text-sm font-bold mb-1.5 ${txt}`}>Margin by Month</h3>
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
                <tr className={trHov}><td className={`px-2 py-0.5 font-semibold ${txt}`}>Margin</td>{monthlyCompleted.map((m,i)=><td key={i} className={`px-1.5 py-0.5 text-center font-semibold ${m.margin != null ? (m.margin < 0 ? 'text-red-600' : 'text-green-600') : txt2}`}>{m.margin != null ? fmtCurShort(m.margin) : '-'}</td>)}</tr>
                <tr className={trHov}><td className={`px-2 py-0.5 font-semibold ${txt}`}>Margin %</td>{monthlyCompleted.map((m,i)=><td key={i} className={`px-1.5 py-0.5 text-center font-semibold ${m.margin_pct != null ? (m.margin < 0 ? 'text-red-600' : 'text-green-600') : txt2}`}>{m.margin_pct != null ? `${m.margin_pct.toFixed(1)}%` : '-'}</td>)}</tr>
              </tbody>
            </table>
          </div>
        </div>
        <div className="grid grid-cols-1 xl:grid-cols-2 gap-5"><div className={`p-5 rounded-2xl ${card}`}><h3 className={`text-base font-bold ${txt}`}>Top 5 Vendor PO Amount</h3><p className={`text-xs mb-4 ${txt2}`}>Total: {fmtCurShort(sumRows(d.top_vendors, 'purchase_amount'))}</p>{barList(d.top_vendors, 'vendor', 'purchase_amount', 'PO Amount', '#2563EB')}</div><div className={`p-5 rounded-2xl ${card}`}><h3 className={`text-base font-bold ${txt}`}>Top 5 Client Sales Amount</h3><p className={`text-xs mb-4 ${txt2}`}>Total: {fmtCurShort(sumRows(d.top_clients, 'sales_amount'))}</p>{barList(d.top_clients, 'client', 'sales_amount', 'Sales Amount', '#14B8A6')}</div></div>
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
                <div className={`absolute left-0 right-0 top-full mt-1 z-50 max-h-64 overflow-auto rounded-xl border shadow-xl ${darkMode ? 'bg-gray-800 border-gray-700 text-gray-100' : 'bg-white border-gray-200 text-gray-800'}`}>
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

        <div className="overflow-x-auto">
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
        </div>

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
    const columns = [
      ['Product ID', 'prod_id'],
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

    const handleClear = () => {
      setRegisteredItemsSearch('');
      setRegisteredItemsAppliedSearch('');
      setRegisteredItemsProdIds([]);
      setRegisteredItemsAppliedProdIds([]);
      setRegisteredItemsPage(1);
      fetchRegisteredItems(1, registeredItemsPerPage, '', []);
    };

    const downloadRegisteredExcel = () => {
      const p = new URLSearchParams();
      if (registeredItemsAppliedSearch) p.append('search', registeredItemsAppliedSearch);
      (registeredItemsAppliedProdIds || []).forEach(v => p.append('prod_id', v));
      downloadBlob(`/api/export/all-registered-items?${p}`, `All_Registered_Items_${new Date().toISOString().slice(0,10)}.xlsx`, 'All Registered Items');
    };

    const fmtDateShort = (d) => {
      if (!d) return '-';
      try { return d.slice(0, 10); } catch { return d; }
    };
    const fmtProductId = (v) => {
      if (v == null || v === '') return '-';
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
          <div className="flex flex-wrap gap-2 items-end">
            <div className="w-full sm:w-[280px]">
              <label className={`block text-xs font-semibold mb-1 ${txt2}`}>Search</label>
              <input
                value={registeredItemsSearch}
                onChange={e => setRegisteredItemsSearch(e.target.value)}
                placeholder="Name, spec, manufacturer..."
                className={`w-full h-10 px-3 py-2 rounded-xl text-sm border ${darkMode ? 'bg-gray-700 border-gray-600 text-white placeholder:text-gray-400' : 'bg-white border-gray-200 text-gray-800 placeholder:text-gray-400'}`}
              />
            </div>
            <div className="w-full sm:w-[170px]">
              <label className={`block text-xs font-semibold mb-1 ${txt2}`}>Search Prod ID</label>
              <SearchInput
                key={`registered-prod-id-${registeredItemsProdIds.join('|')}`}
                placeholder={'8381684\n8382076'}
                label="Prod ID"
                darkMode={darkMode}
                txt2={txt2}
                onSearch={(nums) => {
                  setRegisteredItemsProdIds(nums);
                  setRegisteredItemsAppliedProdIds(nums);
                  setRegisteredItemsPage(1);
                  fetchRegisteredItems(1, registeredItemsPerPage, registeredItemsAppliedSearch, nums);
                }}
              />
            </div>
            <button onClick={() => { setRegisteredItemsAppliedSearch(registeredItemsSearch); setRegisteredItemsPage(1); fetchRegisteredItems(1, registeredItemsPerPage, registeredItemsSearch, registeredItemsAppliedProdIds); }} className="w-[90px] h-10 px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold shadow-sm">
              Search
            </button>
            <button onClick={handleClear} className={`w-[110px] h-10 px-3 py-2 rounded-lg text-sm font-medium shadow-sm flex items-center justify-center whitespace-nowrap ${darkMode ? 'bg-gray-500 text-gray-100 hover:bg-gray-400' : 'bg-gray-400 text-white hover:bg-gray-500'}`}>
              Clear
            </button>
          </div>
        </FilterPanel>

        <div className="overflow-x-auto">
          <table className="freeze-table-all-registered-items w-full text-xs">
            <colgroup>
              <col style={{minWidth:'120px'}}/>
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
                  <th key={label} className={`px-2 py-2 text-center font-bold whitespace-nowrap ${txt2}`}>{renderFreezeHeader('all-registered-items', index + 1, label)}</th>
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
                  <td className="px-2 py-2 font-mono text-blue-600 whitespace-nowrap text-center">{fmtProductId(row.prod_id)}</td>
                  <td className={`px-2 py-2 ${txt2}`} title={row.category}>{row.category || '-'}</td>
                  <td className="px-2 py-2 text-center whitespace-nowrap">
                    {row.pic ? (() => {
                      const c = getPicColor(row.pic);
                      return <span className={`px-1.5 py-0.5 rounded-full text-xs font-semibold ${c ? `${c.bg} ${c.text}` : 'bg-gray-100 text-gray-700'}`}>{row.pic}</span>;
                    })() : <span className={txt2}>-</span>}
                  </td>
                  <td className={`px-2 py-2 max-w-[160px] truncate ${txt}`} title={row.prod_name}>{row.prod_name || '-'}</td>
                  <td className={`px-2 py-2 max-w-[280px] truncate ${txt2}`} title={row.spec}>{row.spec || '-'}</td>
                  <td className={`px-2 py-2 max-w-[180px] truncate ${txt2}`} title={row.mfr_name}>{row.mfr_name || '-'}</td>
                  <td className={`px-2 py-2 text-center whitespace-nowrap ${txt2}`}>{row.odr_unit || '-'}</td>
                  <td className={`px-2 py-2 text-center whitespace-nowrap ${txt2}`}>{row.hub_handling_check || '-'}</td>
                  <td className={`px-2 py-2 text-center whitespace-nowrap ${txt2}`}>{row.tax_type || '-'}</td>
                  <td className={`px-2 py-2 whitespace-nowrap ${txt2}`}>{fmtDateShort(row.registration_date)}</td>
                  <td className={`px-2 py-2 max-w-[160px] truncate ${txt2}`} title={row.product_registry_pic}>{row.product_registry_pic || '-'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <PagePagination
          darkMode={darkMode}
          txt2={txt2}
          page={registeredItemsPage}
          totalPages={totalPages}
          total={registeredItemsTotal}
          perPage={registeredItemsPerPage}
          onPageChange={(p) => { setRegisteredItemsPage(p); fetchRegisteredItems(p, registeredItemsPerPage, registeredItemsAppliedSearch, registeredItemsAppliedProdIds); }}
          onPerPageChange={(next) => { setRegisteredItemsPerPage(next); setRegisteredItemsPage(1); fetchRegisteredItems(1, next, registeredItemsAppliedSearch, registeredItemsAppliedProdIds); }}
        />
      </div>
    );
  };

  const renderRFQ = () => {
    const totalPages = Math.max(1, Math.ceil(rfqTotal / rfqPerPage));
    const editableSet = new Set(rfqEditableFields || []);
    const baseColumns = (rfqColumns.length ? rfqColumns : [
      { field: 'check', label: 'Check' }, { field: 'sheet_status', label: 'Status' }, { field: 'days_left', label: 'Days Left' }, { field: 'no', label: 'No' }, { field: 'client_name', label: 'Nama Client' },
      { field: 'rfq_date', label: 'RFQ Date' }, { field: 'closing_date', label: 'Closing Date' }, { field: 'sales_pic', label: 'Sales PIC' },
      { field: 'category_name', label: 'Category Name' }, { field: 'purchase_pic', label: 'Purchase PIC' },
      { field: 'item_name', label: 'Item Name' }, { field: 'detail_spec', label: 'Detail Spec' }, { field: 'brand_manufacturer', label: 'Brand/Manufaktur' },
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
    const rfqSourceStyleFields = new Set([
      'check', 'sheet_status', 'days_left', 'no', 'client_name', 'rfq_date', 'closing_date', 'sales_pic',
      'category_name', 'purchase_pic', 'item_name', 'detail_spec', 'brand_manufacturer', 'qty', 'unit', 'remark',
      'similar_prod_ids', 'similar_prod_name', 'similar_spec', 'similar_mfr_name', 'similar_odr_unit', 'similar_score'
    ]);
    const colWidth = (field) => ({
      check: 64, sheet_status: 90, days_left: 76, no: 70, client_name: 160, rfq_date: 110, closing_date: 110, sales_pic: 120,
      item_name: 180, detail_spec: 620, brand_manufacturer: 160, qty: 80, unit: 80, remark: 380,
      category_id: 180, category_name: 150, product_id: 120, request_number: 150, purchase_pic: 120,
      same_replacement: 92, vendor_name: 200, unit_price_idr: 130, amt_idr: 130, quoted_item_name: 180,
      quoted_spec: 150, quoted_brand: 130, quoted_unit: 58, moq: 62, lead_time_days: 78,
      valid_period: 82, photo_url: 92, remarks: 360,
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
    const toDateInputValue = (value) => {
      const raw = String(value || '').trim();
      if (!raw) return '';
      const dmy = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (dmy) return `${dmy[3]}-${String(dmy[2]).padStart(2, '0')}-${String(dmy[1]).padStart(2, '0')}`;
      const ymd = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
      if (ymd) return `${ymd[1]}-${String(ymd[2]).padStart(2, '0')}-${String(ymd[3]).padStart(2, '0')}`;
      return raw;
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
      setRfqSortOrder('newest');
      const nextFilters = { checks: [], clients: [], brands: [], purchase_pics: [], vendors: [] };
      setRfqFilters(nextFilters);
      setRfqPage(1);
      fetchRFQData(1, rfqPerPage, '', false, nextFilters, '', rfqShowSimilarity, 'newest');
    };
    const rfqParams = () => {
      const p = new URLSearchParams();
      if (rfqAppliedSearch) p.append('search', rfqAppliedSearch);
      if (rfqPicFilter) p.append('pic', rfqPicFilter);
      if (rfqSortOrder) p.append('sort_order', rfqSortOrder);
      resolveFilter(rfqFilters.checks).forEach(v => p.append('check', v));
      resolveFilter(rfqFilters.clients).forEach(v => p.append('client_name', v));
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
        addToast(`RFQ batch: ${res.data.updated || 0} cells updated${res.data.not_found ? `, ${res.data.not_found} No not found` : ''}`, 'success');
        fetchRFQData(rfqPage, rfqPerPage, rfqAppliedSearch, true, rfqFilters, rfqPicFilter);
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
    const fillRFQDown = async (startRowIndex, field, endRowIndex) => {
      if (endRowIndex <= startRowIndex) return;
      const source = rfqData[startRowIndex];
      if (!source) return;
      const value = source[field] ?? '';
      const batchUpdates = [];
      for (let i = startRowIndex + 1; i <= endRowIndex && i < rfqData.length; i += 1) {
        batchUpdates.push({ row_key: rfqData[i].row_key, field, value });
      }
      if (batchUpdates.length && await updateRFQCellsBatch(batchUpdates)) {
        addToast(`RFQ fill down: ${batchUpdates.length} cells updated`, 'success');
      }
    };
    const startRFQFill = (event, rowIndex, field) => {
      event.preventDefault();
      event.stopPropagation();
      const onUp = (upEvent) => {
        document.removeEventListener('mouseup', onUp);
        const target = document.elementFromPoint(upEvent.clientX, upEvent.clientY)?.closest('[data-rfq-cell="true"]');
        const endRowIndex = Number(target?.getAttribute('data-row-index'));
        const targetField = target?.getAttribute('data-field');
        if (Number.isFinite(endRowIndex) && targetField === field) fillRFQDown(rowIndex, field, endRowIndex);
      };
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
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-4 2xl:grid-cols-[110px_minmax(170px,1fr)_115px_repeat(4,minmax(120px,1fr))_84px_84px] items-end">
            <div className="min-w-0">
              <label className={`block text-xs font-semibold mb-1 ${txt2}`}>↕ RFQ Date</label>
              <select
                value={rfqSortOrder}
                onChange={e => { const next = e.target.value; setRfqSortOrder(next); setRfqPage(1); fetchRFQData(1, rfqPerPage, rfqAppliedSearch, false, rfqFilters, rfqPicFilter, rfqShowSimilarity, next); }}
                title="Sort RFQ Date"
                className={`w-full h-10 px-2 py-2 rounded-xl text-sm border ${darkMode?'bg-gray-700 border-gray-600 text-white':'bg-white border-gray-200 text-gray-800'}`}
              >
                <option value="newest">Newest ↓</option>
                <option value="oldest">Oldest ↑</option>
              </select>
            </div>
            <div className="min-w-0">
              <label className={`block text-xs font-semibold mb-1 ${txt2}`}>Search RFQ</label>
              <input
                value={rfqSearch}
                onChange={e => setRfqSearch(e.target.value)}
                placeholder="Client, item, vendor, product ID..."
                className={`w-full h-10 px-3 py-2 rounded-xl text-sm border ${darkMode?'bg-gray-700 border-gray-600 text-white placeholder:text-gray-400':'bg-white border-gray-200 text-gray-800 placeholder:text-gray-400'}`}
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
            <button onClick={() => { setRfqAppliedSearch(rfqSearch); setRfqPage(1); fetchRFQData(1, rfqPerPage, rfqSearch, false, rfqFilters, rfqPicFilter, rfqShowSimilarity); }} className="w-full h-10 px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold shadow-sm">
              Search
            </button>
            <button onClick={handleClear} className={`w-full h-10 px-3 py-2 rounded-lg text-sm font-medium shadow-sm flex items-center justify-center whitespace-nowrap ${darkMode?'bg-gray-500 text-gray-100 hover:bg-gray-400':'bg-gray-400 text-white hover:bg-gray-500'}`}>
              Clear
            </button>
          </div>
        </FilterPanel>

        <div className="overflow-x-auto">
          <table className="freeze-table-rfq table-fixed text-xs border-collapse" style={{ width: `${rfqTableWidth}px`, minWidth: `${rfqTableWidth}px` }}>
            <colgroup>{columns.map(col => <col key={col.field} style={colStyle(col.field)}/>)}</colgroup>
            <thead className={tblHd}>
              <tr>{columns.map((col, index) => {
                const darkHeaderCols = ['check', 'sheet_status', 'days_left', 'no', 'client_name', 'rfq_date', 'closing_date', 'sales_pic', 'category_name', 'purchase_pic', 'item_name', 'detail_spec', 'brand_manufacturer', 'qty', 'unit', 'remark', 'similar_prod_ids', 'similar_prod_name', 'similar_spec', 'similar_mfr_name', 'similar_odr_unit', 'similar_score'];
                const isDarkHeader = darkHeaderCols.includes(col.field);
                return <th key={col.field} className={`px-2 py-2 text-center font-bold whitespace-nowrap border-r ${isDarkHeader ? 'bg-slate-200 text-slate-700' : darkMode ? 'bg-gray-800/60 border-gray-700 text-gray-200' : 'bg-slate-50 border-gray-200 text-gray-700'} ${darkMode ? 'border-gray-700' : 'border-gray-200'}`}>{renderFreezeHeader('rfq', index + 1, col.label)}</th>;
              })}</tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {rfqData.length === 0 ? (
                <tr><td colSpan={columns.length} className={`px-4 py-12 text-center ${txt2}`}><Mail className="w-10 h-10 mx-auto mb-2 opacity-40"/>No RFQ data</td></tr>
              ) : rfqData.map((row, rowIndex) => {
                return (
                <tr key={row.row_key} className={`${trHov} transition-colors${rfqPicFilter && rfqEditedRowKeys.has(row.row_key) ? ' ring-1 ring-inset ring-amber-400/60' : ''}`}>
                  {columns.map((col) => {
                    const field = col.field;
                    const value = row[field] ?? '';
                    const isEditable = editableSet.has(field);
                    const isEditing = editingCell?.id === row.row_key && editingCell.field === `rfq_${field}`;
                    if (field === 'check') {
                      const checkValue = String(row.check || '').toLowerCase();
                      if (checkValue === 'complete') {
                        return <td key={field} className={`px-2 py-2 text-center border-r ${darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200'}`} title="Complete"><span className="inline-flex h-6 w-6 items-center justify-center rounded-full bg-[#20B71F]"><Check className="w-4 h-4 text-white stroke-[4]"/></span></td>;
                      }
                      if (checkValue === 'reject') {
                        return <td key={field} className={`px-2 py-2 text-center border-r ${darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200'}`} title="Reject"><span className="inline-flex h-6 w-6 items-center justify-center rounded-full bg-[#EA0D0D]"><X className="w-4 h-4 text-white stroke-[4]"/></span></td>;
                      }
                      const closed = checkValue === 'closed' || (!row.product_id && isRFQClosingPast(row.closing_date));
                      return <td key={field} className={`px-2 py-2 text-center border-r ${darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200'}`} title={closed ? 'Closed' : 'Open'}><span className={`inline-flex h-6 w-6 rounded-full border ${closed ? (darkMode ? 'bg-gray-500 border-gray-400' : 'bg-gray-300 border-gray-400') : darkMode ? 'bg-gray-700 border-gray-500' : 'bg-white border-gray-300'}`}/></td>;
                    }
                    if (field === 'days_left') {
                      return <td key={field} className={`px-2 py-2 text-center border-r ${darkMode ? 'bg-gray-800/60 border-gray-700 text-gray-100' : 'bg-slate-50 border-gray-200 text-black'}`}>{row.days_left === 0 || row.days_left ? fmtNum(row.days_left) : '-'}</td>;
                    }
                    if (isEditable && isEditing) {
                      const tall = ['quoted_spec', 'remarks', 'photo_url'].includes(field);
                      const Control = tall ? 'textarea' : 'input';
                      if (['rfq_date', 'closing_date'].includes(field)) {
                        return <td key={field} data-rfq-cell="true" data-row-index={rowIndex} data-field={field} className={`relative p-0 align-top border-r ${darkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200'}`}>
                          <input
                            type="date"
                            value={toDateInputValue(editValue)}
                            className={`block w-full min-h-8 px-1.5 py-1 text-xs border-0 rounded-none outline outline-2 outline-blue-500 outline-offset-[-2px] ${darkMode?'bg-gray-700 text-white':'bg-white text-gray-900'}`}
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
                        return <td key={field} data-rfq-cell="true" data-row-index={rowIndex} data-field={field} className={`relative p-0 align-top border-r ${darkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200'}`}>
                          <select
                            value={editValue}
                            className={`block w-full min-h-8 px-1.5 py-1 text-xs border-0 rounded-none outline outline-2 outline-blue-500 outline-offset-[-2px] ${darkMode?'bg-gray-700 text-white':'bg-white text-gray-900'}`}
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
                      return <td key={field} data-rfq-cell="true" data-row-index={rowIndex} data-field={field} className={`relative p-0 align-top border-r ${darkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200'}`}>
                        <Control
                          value={editValue}
                          rows={tall ? 3 : undefined}
                          className={`block w-full min-h-8 px-2 py-1 text-xs border-0 rounded-none outline outline-2 outline-blue-500 outline-offset-[-2px] ${tall ? 'resize-y' : ''} ${darkMode?'bg-gray-700 text-white':'bg-white text-gray-900'}`}
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
                      const sourceStyle = rfqSourceStyleFields.has(field);
                      return <td key={field} data-rfq-cell="true" data-row-index={rowIndex} data-field={field}
                        tabIndex={0}
                        onFocus={() => setRfqSelectedCell({ rowKey: row.row_key, field })}
                        onClick={() => {
                          setRfqSelectedCell({ rowKey: row.row_key, field });
                          setEditingCell({ id: row.row_key, field: `rfq_${field}` });
                          if (field === 'unit_price_idr') {
                            setEditValue(String(value ?? '').replace(/[^0-9.-]/g, ''));
                          } else {
                            setEditValue(value ?? '');
                          }
                        }}
                        onPaste={e => { e.preventDefault(); applyRFQPaste(rowIndex, field, e.clipboardData.getData('text/plain')); }}
                        className={`group relative px-2 py-1 align-top border-r cursor-pointer ${sourceStyle ? (darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-slate-50 border-gray-200') : (darkMode ? 'bg-gray-800 border-gray-700' : 'bg-white border-gray-200')} ${selected ? 'outline outline-2 outline-blue-500 outline-offset-[-2px]' : 'hover:outline hover:outline-2 hover:outline-blue-400 hover:outline-offset-[-2px]'} ${['qty','unit_price_idr','moq','lead_time_days'].includes(field) ? 'text-right font-semibold' : ''}`}>
                        <div className={`min-h-7 min-w-0 truncate ${sourceStyle ? txt2 : 'text-blue-600'} ${field === 'photo_url' ? 'flex items-center gap-1 justify-center' : ''} ${field === 'purchase_pic' ? 'text-center' : ''}`}>
                          {field === 'photo_url' && <LinkIcon className={`w-3.5 h-3.5 flex-shrink-0 ${hasValue ? 'text-blue-600' : 'text-blue-400'}`} />}
                          {hasValue && field === 'purchase_pic' ? (() => {
                            const c = getPicColor(value);
                            return <span className={`inline-flex max-w-full truncate px-2 py-0.5 rounded-full text-[11px] font-semibold ${c ? `${c.bg} ${c.text}` : 'bg-gray-100 text-gray-700'}`}>{value}</span>;
                          })() : hasValue ? renderValue(value, sourceStyle ? txt2 : 'text-blue-600') : <span>{field === 'photo_url' ? '' : '\u00a0'}</span>}
                        </div>
                        <button type="button" aria-label="Fill down" title="Drag down to copy" onMouseDown={e => startRFQFill(e, rowIndex, field)} className="absolute bottom-0 right-0 h-2.5 w-2.5 translate-x-1/2 translate-y-1/2 border border-blue-600 bg-blue-600 opacity-0 group-hover:opacity-100 focus:opacity-100" />
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
                    const darkDataCols = ['sheet_status', 'no', 'client_name', 'rfq_date', 'closing_date', 'sales_pic', 'category_name', 'purchase_pic', 'item_name', 'detail_spec', 'brand_manufacturer', 'qty', 'unit', 'remark', 'similar_prod_ids', 'similar_prod_name', 'similar_spec', 'similar_mfr_name', 'similar_odr_unit', 'similar_score'];
                    const isDarkDataCol = darkDataCols.includes(field);
                    return <td key={field} className={`px-2 py-2 align-top border-r ${isDarkDataCol ? (darkMode ? 'bg-gray-800/60 border-gray-700 text-gray-100' : 'bg-slate-50 border-gray-200 text-black') : (darkMode ? 'bg-gray-800/60 border-gray-700' : 'bg-white border-gray-200')} ${['detail_spec','remark','category_name','similar_spec'].includes(field) ? '' : 'truncate'} ${['qty','amt_idr','similar_score'].includes(field) ? 'text-right font-semibold' : ''} ${isDarkDataCol ? '' : txt2}`}>
                      {renderValue(value)}
                    </td>;
                  })}
                </tr>
              );})}
            </tbody>
          </table>
        </div>

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

  const renderItemRegistration = () => {
    const baseColumns = [
      ['Proc. Status', 'proc_status'], ['Existing Owner', 'existing_owner'], ['Client Nm.', 'client_name'], ['Category', 'category'], ['PIC', 'pic'],
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
      proc_status: 150, existing_owner: 120, client_name: 180, category: 170, pic: 90, req_no: 150, prod_id: 110,
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
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-4 2xl:grid-cols-[170px_repeat(6,minmax(150px,1fr))_120px] items-end">
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
              <MultiSelect label="Existing Owner" options={itemRegOptions.existing_owners} selected={itemRegFilters.existing_owners}
                onChange={v=>{ const next={...itemRegFilters, existing_owners:v}; setItemRegFilters(next); setItemRegPage(1); fetchItemRegistration(1,itemRegPerPage,itemRegAppliedSearch,next); }} darkMode={darkMode} txt2={txt2}/>
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
            <button onClick={() => { const next={ clients: [], categories: [], pics: [], proc_statuses: [], mfr_names: [], existing_owners: [] }; setItemRegSearch([]); setItemRegAppliedSearch([]); setItemRegPicHighlight(''); setItemRegFilters(next); setItemRegPage(1); fetchItemRegistration(1, itemRegPerPage, [], next, ''); }}
              className={`w-full h-10 px-3 py-2 rounded-lg text-sm font-medium shadow-sm flex items-center justify-center whitespace-nowrap ${darkMode?'bg-gray-500 text-gray-100 hover:bg-gray-400':'bg-gray-400 text-white hover:bg-gray-500'}`}>Clear</button>
          </div>
        </FilterPanel>

        <div className="overflow-x-auto">
          <table className="freeze-table-item-registration table-fixed text-xs" style={{ width: `${itemRegTableWidth}px`, minWidth: `${itemRegTableWidth}px` }}>
            <colgroup>{columns.map(([, key]) => <col key={key} style={colStyle(key)}/>)}</colgroup>
            <thead className={tblHd}><tr>{columns.map(([label], index) => <th key={label} className={`px-2 py-2 text-center font-bold whitespace-nowrap ${txt2}`}>{renderFreezeHeader('item-registration', index + 1, label)}</th>)}</tr></thead>
            <tbody className={`divide-y ${tblDv}`}>
              {itemRegData.length === 0 ? <tr><td colSpan={columns.length} className={`px-4 py-12 text-center ${txt2}`}><Wrench className="w-10 h-10 mx-auto mb-2 opacity-40"/>No Item Registration data</td></tr>
              : itemRegData.map(row => {
                return <tr key={row.id} className={`${trHov} transition-colors`}>
                {columns.map(([, key]) => {
                  const value = key === 'prod_price' ? fmtNum(row[key]) : (row[key] || '-');
                  if (key === 'proc_status') return <td key={key} className="px-2 py-2"><span className={`inline-flex max-w-full items-center px-2 py-0.5 rounded-full border text-[11px] font-semibold leading-snug truncate ${statusClass(row[key])}`}>{value}</span></td>;
                  if (key === 'pic') {
                    const c = getPicColor(row.pic);
                    return <td key={key} className="px-2 py-2 text-center truncate">{row.pic ? <span className={`inline-flex max-w-full truncate px-2 py-0.5 rounded-full text-[11px] font-semibold ${c ? `${c.bg} ${c.text}` : 'bg-gray-100 text-gray-700'}`}>{row.pic}</span> : <span className={txt2}>-</span>}</td>;
                  }
                  if (key === 'remarks') return <td key={key} className="px-2 py-2 truncate" title={row.remarks}>{editingCell?.id===row.id && editingCell.field==='item_remarks' ? (
                    <input type="text" defaultValue={row.remarks}
                      className={`w-full px-2 py-1 rounded text-xs border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                      onChange={e=>setEditValue(e.target.value)}
                      onBlur={()=>updateItemRegistrationCell(row.id,'remarks',editValue)}
                      onKeyDown={e=>{ if(e.key==='Enter') updateItemRegistrationCell(row.id,'remarks',editValue); if(e.key==='Escape') setEditingCell(null); }}
                      autoFocus/>
                  ) : (
                    <span className="cursor-pointer text-blue-600 hover:underline" onClick={()=>{setEditingCell({id:row.id,field:'item_remarks'});setEditValue(row.remarks||'');}}>{row.remarks||'Add'}</span>
                  )}</td>;
                  return <td key={key} className={`px-2 py-2 ${['req_no','prod_name'].includes(key) ? '' : 'truncate'} ${key === 'prod_price' ? `text-right font-semibold ${kpiValue}` : txt2} ${['req_no','prod_id','prod_name','odr_unit','curr'].includes(key) ? 'whitespace-nowrap' : ''}`} title={row[key]}>{value}</td>;
                })}
              </tr>;})}
            </tbody>
          </table>
        </div>

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
    const itemRegExistingOwners = stats?.item_registration_existing_owners || [];
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
          <h3 className={`text-base font-bold ${txt}`}>{title}</h3>
          <p className={`text-xs mb-4 ${txt2}`}>Total: {fmtNum(total)} Req. No</p>
          {data.length === 0 ? (
            <div className={`h-[180px] flex items-center justify-center text-sm ${txt2}`}>No Item Registration data</div>
          ) : (
            <div className="flex flex-col gap-3 md:flex-row md:items-start">
              <div className="h-[200px] w-full min-w-0 pt-5 md:w-[210px] md:flex-none">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart margin={{ top: 4, right: 4, bottom: 4, left: 4 }}>
                    <Pie data={data} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={38} outerRadius={72} labelLine={false} label={renderPctLabel} isAnimationActive={false}>
                      {data.map((_, i) => <Cell key={i} fill={pieColors[i % pieColors.length]} />)}
                    </Pie>
                    <Tooltip formatter={(v, n) => [`${fmtNum(v)} Req. No`, n]} contentStyle={{background:darkMode?'#1F2937':'#fff',border:'none',borderRadius:8,fontSize:12}}/>
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div className="w-full md:w-[190px] space-y-2">
                {data.map((item, i) => (
                  <div key={item.name} className="flex items-start gap-2 text-xs">
                    <span className="mt-1 h-2.5 w-2.5 rounded-full flex-shrink-0" style={{ backgroundColor: pieColors[i % pieColors.length] }} />
                    <div className="min-w-0 flex-1">
                      <p className={`font-semibold leading-snug break-words ${txt}`} title={item.name}>{item.name}</p>
                      <p className={txt2}>{fmtNum(item.value)} | {total ? ((item.value / total) * 100).toFixed(1) : '0.0'}%</p>
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
           <div className="grid grid-cols-1 xl:grid-cols-3 gap-4">
            {itemRegCategoryChart(itemRegProcStatus, 'Proc. Status')}
            {itemRegCategoryChart(itemRegClients, 'Client Nm.')}
            {itemRegCategoryChart(itemRegExistingOwners, 'Existing Owner')}
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
                const nextFilters = { ...soFilters, pics: nextHighlight ? [nextHighlight] : [] };
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
        <div className="overflow-x-auto rounded-lg border border-gray-200">
          <table className="freeze-table-pending-delivery w-full text-sm">
            <colgroup>
              <col style={{minWidth:'76px', width:'76px', maxWidth:'76px'}}/>
              <col style={{minWidth:'60px'}}/>
              <col style={{minWidth:'110px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'130px'}}/>
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
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'90px'}}/>
              <col style={{minWidth:'200px'}}/>
              <col style={{minWidth:'100px'}}/>
              <col style={{minWidth:'560px'}}/>
            </colgroup>
            <thead className={tblHd}>
              <tr>
                {['Aging','Day','SO Create Date','SO Item','PO No.','SO Status','Category','PIC','Product ID','Product Name','Specification','Manufacturer Name','SO Quantity','Sales Unit','Operation Unit Name','Vendor ID','Vendor Name','Currency','Sales Price (Exclude Tax)','Sales Amount (Exclude Tax)','Purchasing Currency','Purchasing Price','Margin','%Margin','Delivery Memo','Plan Date','Remarks'].map((h, index)=>(
                  <th key={h} className={`px-3 py-2.5 text-center font-bold ${txt2}`}>{renderFreezeHeader('pending-delivery', index + 1, h)}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {(() => {
                if (sortedSOData.length === 0) return (
                <tr><td colSpan={27} className={`px-4 py-10 text-center ${txt2}`}>
                    <FileText className="w-10 h-10 mx-auto mb-2 opacity-40"/>No data
                  </td></tr>
                );
                return sortedSOData.map((so) => {
                const isDeliveryCompleted = so.so_status === 'Delivery Completed';
                const poAmount = Number(so.purchasing_amount) || ((Number(so.purchasing_price) || 0) * (Number(so.so_qty) || 0));
                const margin = (so.sales_amount || 0) - poAmount;
                const marginPct = poAmount !== 0 ? (margin / poAmount) * 100 : null;
                const workingDays = Number.isFinite(Number(so.aging_days)) ? Number(so.aging_days) : workingDaysUntilToday(so.so_create_date);
                const marginColor = margin < 0 ? 'text-red-600 font-semibold' : margin > 0 ? 'text-green-600 font-semibold' : txt2;
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
                  <td className="px-3 py-2 whitespace-nowrap">
                    <span className={`px-2 py-0.5 rounded-full text-xs font-medium ${
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
                  <td className={`px-3 py-2 text-right whitespace-nowrap min-w-[130px] ${marginColor}`}>{fmtCur(margin)}</td>
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
        </div>

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
    <div className={`min-h-screen font-sans ${darkMode?'bg-gray-900':'bg-[#edf2f1]'} ${darkMode?'':'text-[#1f2937]'}`} style={{fontFamily: "'Inter', 'Plus Jakarta Sans', ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"}}>
    <style>{`
        @keyframes slide-in {
          from { transform: translateX(100%); opacity: 0; }
          to   { transform: translateX(0);    opacity: 1; }
        }
        .animate-slide-in { animation: slide-in 0.25s ease-out forwards; }
        /* Global: all buttons, links, selects, labels with checkboxes → pointer cursor */
        button, [role="button"], select, label[for], a,
        input[type="checkbox"], input[type="radio"] {
          cursor: pointer !important;
        }
        button:disabled { cursor: not-allowed !important; opacity: 0.5; }
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
          <button onClick={()=>setActivePage('dashboard')}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='dashboard'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="Summary">
            <BarChart3 className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-40 opacity-100':'max-w-0 opacity-0'}`}>Summary</span>
          </button>
          <button data-tour="open-so-nav" onClick={()=>{ setActivePage('all-so'); setSoPage(1); fetchSOData(soFilters,1,soPerPage,soSearchNums,soMarginFilter,soDateFilter); window.scrollTo({top:0, behavior:'smooth'}); }}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='all-so'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="Pending Delivery">
            <Clock className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-40 opacity-100':'max-w-0 opacity-0'}`}>Pending Delivery</span>
          </button>
          <button onClick={()=>{ setActivePage('item-registration'); setItemRegPage(1); fetchItemRegistration(1,itemRegPerPage,itemRegAppliedSearch,itemRegFilters); window.scrollTo({top:0,behavior:'smooth'}); }}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='item-registration'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="Item Registration">
            <Wrench className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-44 opacity-100':'max-w-0 opacity-0'}`}>Item Registration</span>
          </button>
          <button onClick={()=>{ setActivePage('rfq'); setRfqPage(1); window.scrollTo({top:0,behavior:'smooth'}); }}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='rfq'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="RFQ">
            <Mail className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-44 opacity-100':'max-w-0 opacity-0'}`}>RFQ</span>
          </button>
          <button onClick={()=>{ setActivePage('vendor-control'); setVendorControlPage(1); window.scrollTo({top:0,behavior:'smooth'}); }}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='vendor-control'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="Vendor Control">
            <Building2 className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-44 opacity-100':'max-w-0 opacity-0'}`}>Vendor Control</span>
          </button>
          <button onClick={()=>{ setActivePage('all-registered-items'); setRegisteredItemsPage(1); fetchRegisteredItems(1,registeredItemsPerPage,registeredItemsAppliedSearch,registeredItemsAppliedProdIds); window.scrollTo({top:0,behavior:'smooth'}); }}
            className={`p-3 rounded-xl flex items-center gap-3 justify-start transition-all whitespace-nowrap ${activePage==='all-registered-items'?'bg-slate-600 text-white shadow-sm':darkMode?'text-gray-300 hover:bg-gray-700':'text-gray-600 hover:bg-[#f4f4f2]'}`} title="All Registered Items">
            <FileText className="w-5 h-5 flex-shrink-0"/>
            <span className={`hidden lg:inline overflow-hidden text-sm font-semibold transition-all duration-200 ${sidebarExpanded?'max-w-44 opacity-100':'max-w-0 opacity-0'}`}>All Registered Items</span>
          </button>
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
               :activePage==='item-registration'?'Process Purchase Info Registration data'
               :activePage==='rfq'?'Sales Submit-RFQ live data and quotation updates'
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
                      <p className={`text-xs ${txt2}`}>Update Process Purchase Info Registration</p>
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
                  <label className={`flex items-center gap-2 px-4 py-3 cursor-pointer transition-all ${darkMode?'hover:bg-gray-700':'hover:bg-indigo-50'}`}>
                    <Upload className="w-4 h-4 text-indigo-500"/>
                    <div>
                      <span className={`text-sm font-medium ${txt}`}>Update PIC</span>
                      <p className={`text-xs ${txt2}`}>Update PIC by category</p>
                    </div>
                    <input type="file" accept=".xlsx,.xls" multiple onChange={e=>{handleUpdatePIC(e); setShowUploadDropdown(false);}} className="hidden"/>
                  </label>
                </div>
              )}
              </div>

              <div className="relative" ref={hideMenuRef}>
              <button data-tour="hide-menu" onClick={()=>setShowHideMenu(o=>!o)}
                className="flex items-center gap-2 px-4 py-2.5 rounded-xl shadow-sm transition-all bg-slate-600 hover:bg-slate-700 text-white">
                <EyeOff className="w-4 h-4"/><span className="text-sm font-medium">Hide</span>
                <ChevronDown className="w-3.5 h-3.5"/>
                {deleteRequests.filter(r=>r.is_hidden).length > 0 && (
                  <span className="px-1.5 py-0.5 bg-white text-slate-700 rounded-full text-xs font-bold">
                    {deleteRequests.filter(r=>r.is_hidden).length}
                  </span>
                )}
              </button>
              {showHideMenu && (
                <div className={`absolute right-0 mt-2 z-50 rounded-xl shadow-2xl border w-80 p-3 ${darkMode?'bg-gray-800 border-gray-700 text-white':'bg-white border-gray-200'}`}>
                  {/* View Hidden History */}
                  <button onClick={()=>{ setShowHideMenu(false); fetchDeleteRequests(); setShowHiddenPanel(true); }}
                    className={`w-full flex items-center gap-2 px-3 py-2.5 rounded-lg text-sm font-semibold mb-3 ${darkMode?'bg-gray-700 hover:bg-gray-600 text-white':'bg-gray-100 hover:bg-gray-200 text-gray-700'}`}>
                    <Eye className="w-4 h-4 text-blue-500"/>
                    View Hide History
                    {deleteRequests.filter(r=>r.is_hidden).length > 0 && (
                      <span className="ml-auto px-2 py-0.5 bg-slate-600 text-white rounded-full text-xs font-bold">
                        {deleteRequests.filter(r=>r.is_hidden).length}
                      </span>
                    )}
                  </button>
                  <p className={`text-xs font-semibold mb-2 px-1 ${darkMode?'text-gray-300':'text-gray-600'}`}>
                    Hide data from dashboard via Excel template
                  </p>
                  {/* PO */}
                  <div className={`mb-2 p-3 rounded-lg ${darkMode?'bg-gray-700':'bg-slate-50'}`}>
                    <p className="text-xs font-bold mb-1 text-slate-700 flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-slate-700"/><span>PO HLI</span></p>
                    <p className={`text-xs mb-2 ${darkMode?'text-gray-400':'text-gray-500'}`}>Format: PO Number-Item No (e.g. 4502358819-10)</p>
                    <div className="flex gap-2">
                      <button onClick={()=>downloadHideTemplate('PO')}
                        className="flex-1 flex items-center justify-center gap-1 px-2 py-1.5 bg-slate-600 hover:bg-slate-700 text-white rounded-lg text-xs font-semibold">
                        <Download className="w-3 h-3"/>Download Template
                      </button>
                      <label className="flex-1 flex items-center justify-center gap-1 px-2 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold cursor-pointer">
                        <Upload className="w-3 h-3"/>Upload Filled
                        <input type="file" accept=".xlsx,.xls" onChange={e=>handleHideBatchUpload(e,'PO')} className="hidden"/>
                      </label>
                    </div>
                  </div>
                  {/* SO */}
                  <div className={`p-3 rounded-lg ${darkMode?'bg-gray-700':'bg-blue-50'}`}>
                    <p className="text-xs font-bold mb-1 text-blue-700 flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-blue-600"/><span>SO</span></p>
                    <p className={`text-xs mb-2 ${darkMode?'text-gray-400':'text-gray-500'}`}>Format: SO Number or SO Number-Item No</p>
                    <div className="flex gap-2">
                      <button onClick={()=>downloadHideTemplate('SO')}
                        className="flex-1 flex items-center justify-center gap-1 px-2 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold">
                        <Download className="w-3 h-3"/>Download Template
                      </button>
                      <label className="flex-1 flex items-center justify-center gap-1 px-2 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold cursor-pointer">
                        <Upload className="w-3 h-3"/>Upload Filled
                        <input type="file" accept=".xlsx,.xls" onChange={e=>handleHideBatchUpload(e,'SO')} className="hidden"/>
                      </label>
                    </div>
                  </div>
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

      {showHiddenPanel && (
        <HiddenItemsPanel
          darkMode={darkMode}
          requests={deleteRequests}
          onRestore={restoreDeleteRequest}
          onClose={()=>setShowHiddenPanel(false)}
        />
      )}

      {uploadProgress && (
        <div className="fixed inset-0 bg-black/60 z-[60] flex items-center justify-center backdrop-blur-sm">
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

      {loading && !uploadProgress && (
        <div className="fixed inset-0 bg-black/30 z-[55] flex items-center justify-center">
          <div className={`${darkMode?'bg-gray-800':'bg-white'} px-6 py-4 rounded-xl shadow-xl flex items-center gap-3`}>
            <div className="w-6 h-6 border-3 border-blue-600 border-t-transparent rounded-full animate-spin"/>
            <p className={`text-sm font-semibold ${txt}`}>Loading data...</p>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
