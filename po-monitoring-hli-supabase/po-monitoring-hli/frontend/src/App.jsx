import React, { useState, useEffect, useCallback, useRef } from 'react';
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, AreaChart, Area
} from 'recharts';
import {
  Upload, Download, AlertCircle, CheckCircle, XCircle,
  Package, DollarSign, TrendingUp, Calendar, ChevronLeft,
  ChevronRight, Moon, Sun, FileText, BarChart3, FileSpreadsheet,
  Filter, X, ChevronDown, ChevronUp, Building2, Search, Loader2,
  EyeOff, Eye, Trash2, RotateCcw, Plus
} from 'lucide-react';
import axios from 'axios';
import { format, parseISO } from 'date-fns';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

const BACKEND = import.meta.env.VITE_API_URL || 'http://127.0.0.1:5000';
const api = axios.create({ baseURL: BACKEND, timeout: 600000 });

const PIE_COLORS = ['#8B5CF6','#F97316','#10B981','#EF4444','#3B82F6',
                    '#EC4899','#14B8A6','#F59E0B','#6366F1','#84CC16'];

const AGING_LABELS = ['0-30','30-90','90-180','180+'];
const AGING_COLORS = { '0-30':'#10B981','30-90':'#F59E0B','90-180':'#F97316','180+':'#EF4444' };

// ─── Excluded from PO HLI without SO calculation ──────────────────────────
const EXCLUDED_OP_UNITS = new Set(['HLI GREEN POWER (CONSUMABLE)']);

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

// ─── Download Toast ────────────────────────────────────────────────────────
const DownloadToast = ({ message, onClose }) => {
  return (
    <div className="fixed top-5 right-5 z-[200] flex itemss-center gap-3 px-5 py-3 rounded-xl shadow-2xl text-white bg-purple-700 max-w-sm animate-slide-in">
      <Loader2 className="w-5 h-5 flex-shrink-0 animate-spin"/>
      <span className="text-sm font-medium">{message}</span>
    </div>
  );
};

const Toast = ({ message, type, onClose }) => {
  useEffect(() => { const t = setTimeout(onClose, 3000); return () => clearTimeout(t); }, [onClose]);
  const bg = type === 'success' ? 'bg-green-600' : type === 'error' ? 'bg-red-600' : 'bg-blue-600';
  return (
    <div className={`fixed top-5 right-5 z-[100] flex itemss-center gap-3 px-5 py-3 rounded-xl shadow-2xl text-white ${bg} max-w-sm`}>
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

const SOModal = ({ title, data, onClose, darkMode }) => {
  const [dlPage, setDlPage] = useState(1);
  const PER = 50;
  const pages = Math.ceil((data?.length || 0) / PER);
  const rows = (data || []).slice((dlPage-1)*PER, dlPage*PER);

  // Determine if SO Item column exists in data (show SO Number only when SO Item is absent)
  const hasSoItem = (data || []).some(s => s.so_items);

  const downloadExcel = () => {
    const ws = XLSX.utils.json_to_sheet(data.map(s => ({
      'SO Item': s.so_items,
      ...(!hasSoItem ? { 'SO Number': s.so_number } : {}),
      'Status': s.so_status,
      'Op Unit': s.operation_unit_name, 'Vendor': s.vendor_name, 'Product': s.product_name,
      'SO Qty': s.so_qty, 'Sales Price': s.sales_price, 'Sales Amount': s.sales_amount,
      'Customer PO': s.customer_po_number, 'Delivery Memo': s.delivery_memo,
      'SO Date': s.so_create_date, 'Delivery Plan Date': s.delivery_plan_date, 'Remarks': s.remarks
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Detail');
    saveAs(new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],
      {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),
      `${title.replace(/\s+/g,'_')}.xlsx`);
  };
  return (
    <div className="fixed inset-0 bg-black/60 z-50 flex itemss-center justify-center p-4 backdrop-blur-sm" onClick={onClose}>
      <div role="dialog" aria-modal="true" aria-label={title} className={`rounded-2xl shadow-2xl w-full max-w-6xl max-h-[85vh] flex flex-col ${darkMode?'bg-gray-800 text-white':'bg-white'}`} onClick={e=>e.stopPropagation()}>
        <div className={`flex justify-between itemss-center px-6 py-4 border-b ${darkMode?'border-gray-700':'border-gray-100'}`}>
          <h3 className="font-bold text-lg">{title} <span className={`text-sm font-normal ml-2 ${darkMode?'text-gray-400':'text-gray-500'}`}>({fmtNum(data?.length)} records)</span></h3>
          <div className="flex gap-2">
            <button onClick={downloadExcel} className="flex itemss-center gap-1 px-3 py-1.5 bg-green-600 hover:bg-green-700 text-white rounded-lg text-sm"><FileSpreadsheet className="w-4 h-4"/>Excel</button>
            <button onClick={onClose} className={`p-1.5 rounded-lg ${darkMode?'hover:bg-gray-700':'hover:bg-gray-100'}`}><X className="w-5 h-5"/></button>
          </div>
        </div>
        <div className="overflow-auto flex-1">
          <table className="w-full text-sm">
            <thead className={`sticky top-0 ${darkMode?'bg-gray-700':'bg-purple-50'}`}>
              <tr>{['SO Item', ...(!hasSoItem ? ['SO Number'] : []), 'Status','Op Unit','Vendor','Product','Qty','Sales Amount','Cust PO','Delivery Memo','SO Date','Plan Date','Remarks'].map(h=>(
                <th key={h} className={`px-3 py-2 text-left font-semibold whitespace-nowrap ${darkMode?'text-gray-200':'text-gray-700'}`}>{h}</th>
              ))}</tr>
            </thead>
            <tbody className={`divide-y ${darkMode?'divide-gray-700':'divide-gray-100'}`}>
              {rows.map((s,i)=>(
                <tr key={i} className={darkMode?'hover:bg-gray-700':'hover:bg-purple-50'}>
                  <td className="px-3 py-2 text-purple-600 font-medium whitespace-nowrap">{s.so_items||'-'}</td>
                  {!hasSoItem && <td className="px-3 py-2 whitespace-nowrap">{s.so_number}</td>}
                  <td className="px-3 py-2 whitespace-nowrap"><span className={`px-2 py-0.5 rounded-full text-xs font-medium ${s.so_status==='Delivery Completed'?'bg-green-100 text-green-700':s.so_status==='SO Cancel'?'bg-red-100 text-red-700':'bg-blue-100 text-blue-700'}`}>{s.so_status||'-'}</span></td>
                  <td className="px-3 py-2 whitespace-nowrap min-w-[180px]">{s.operation_unit_name}</td>
                  <td className="px-3 py-2 whitespace-nowrap max-w-[140px] truncate">{s.vendor_name}</td>
                  <td className="px-3 py-2 max-w-[160px] truncate">{s.product_name}</td>
                  <td className="px-3 py-2 text-right">{fmtNum(s.so_qty)}</td>
                  <td className="px-3 py-2 text-right font-semibold text-orange-600 whitespace-nowrap">{fmtCur(s.sales_amount)}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{s.customer_po_number||'-'}</td>
                  <td className="px-3 py-2 max-w-[160px] truncate">{s.delivery_memo||'-'}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{s.so_create_date||'-'}</td>
                  <td className="px-3 py-2 whitespace-nowrap text-purple-600">{s.delivery_plan_date||'-'}</td>
                  <td className="px-3 py-2 max-w-[140px] truncate">{s.remarks||'-'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {pages > 1 && (
          <div className={`flex justify-between itemss-center px-6 py-3 border-t ${darkMode?'border-gray-700':'border-gray-100'}`}>
            <span className={`text-sm ${darkMode?'text-gray-400':'text-gray-600'}`}>{(dlPage-1)*PER+1}–{Math.min(dlPage*PER,data.length)} / {fmtNum(data.length)}</span>
            <div className="flex gap-2">
              <button disabled={dlPage===1} onClick={()=>setDlPage(p=>p-1)} className={`p-1.5 rounded ${dlPage===1?'opacity-40':'hover:bg-gray-200'}`}><ChevronLeft className="w-4 h-4"/></button>
              <span className="px-3 py-1 bg-purple-100 rounded text-sm text-purple-700">{dlPage}/{pages}</span>
              <button disabled={dlPage===pages} onClick={()=>setDlPage(p=>p+1)} className={`p-1.5 rounded ${dlPage===pages?'opacity-40':'hover:bg-gray-200'}`}><ChevronRight className="w-4 h-4"/></button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

// ─── MultiSelect dropdown — Excel-style (all checked by default) ─────────
const MultiSelect = ({ label, options, selected, onChange, darkMode, txt2 }) => {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  // "all selected" = selected array is empty (no filter = show all)
  const noneSelected = selected.length === 0;
  const allSelected  = selected.length === options.length;
  const someSelected = !noneSelected && !allSelected;
  // visually "all checked" when no filter applied
  const allVisuallyChecked = noneSelected;

  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const toggleAll = () => {
    if (noneSelected) {
      // Currently all checked → uncheck all (set to explicit empty selection = nothing shown)
      // We use a sentinel: store all itemss as "selected" but display as "0 selected"
      // Better UX: uncheck all means filter passes nothing, so we store all as excluded
      // Actually Excel behavior: uncheck all → nothing shows. We store [] but invert logic.
      // Simplest: use null/special state — instead use a "noneMode" approach:
      // When noneSelected (was all-checked), clicking unchecks all → store special marker
      onChange('__NONE__');
    } else {
      // Currently some/all explicitly selected OR none-mode → check all → reset to []
      onChange([]);
    }
  };

  const toggle = (val) => {
    // If currently in all-checked visual state (noneSelected)
    if (noneSelected) {
      // Click one items: keep only that one checked (deselect all others)
      onChange([val]);
      return;
    }
    const currentSelected = selected === '__NONE__' ? [] : selected;
    if (currentSelected.includes(val)) {
      const next = currentSelected.filter(x => x !== val);
      onChange(next.length === 0 ? '__NONE__' : next);
    } else {
      const next = [...currentSelected, val];
      onChange(next.length === options.length ? [] : next);
    }
  };

  const isChecked = (val) => {
    if (selected === '__NONE__') return false;
    if (noneSelected) return true; // all visually checked
    return selected.includes(val);
  };

  const isAllChecked = selected !== '__NONE__' && noneSelected;
  const isNoneMode   = selected === '__NONE__';

  const displayLabel = isNoneMode
    ? `0 selected`
    : noneSelected
    ? `All ${label}`
    : `${selected.length} selected`;

  return (
    <div className="relative flex-1 min-w-[180px]" ref={ref}>
      <label className={`block text-xs font-medium mb-1 ${txt2}`}>{label}</label>
      <button onClick={()=>setOpen(o=>!o)} style={{cursor:'pointer'}}
        className={`w-full px-3 py-2 rounded-lg text-sm border text-left flex justify-between itemss-center transition-colors
          ${darkMode
            ? 'bg-gray-600 border-gray-500 text-white hover:bg-gray-500'
            : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'}`}>
        <span className="truncate">{displayLabel}</span>
        <ChevronDown className="w-4 h-4 flex-shrink-0 ml-1"/>
      </button>
      {open && (
        <div className={`absolute z-50 mt-1 w-full max-h-56 overflow-auto rounded-lg shadow-xl border ${darkMode?'bg-gray-700 border-gray-600':'bg-white border-gray-200'}`}>
          {/* Select All row — like Excel */}
          <label style={{cursor:'pointer'}} className={`flex itemss-center gap-2 px-3 py-2 text-xs font-semibold border-b
            ${darkMode?'border-gray-600 hover:bg-gray-600 text-white':'border-gray-100 hover:bg-purple-50 text-gray-700'}`}>
            <input type="checkbox"
              checked={isAllChecked}
              ref={el => { if (el) el.indeterminate = someSelected; }}
              onChange={toggleAll}
              className="accent-purple-600" style={{cursor:'pointer'}}/>
            <span>(Select All)</span>
          </label>
          {options.map(opt => (
            <label key={opt} style={{cursor:'pointer'}} className={`flex itemss-center gap-2 px-3 py-2 text-xs
              ${darkMode?'hover:bg-gray-600 text-white':'hover:bg-purple-50 text-gray-700'}`}>
              <input type="checkbox" checked={isChecked(opt)} onChange={()=>toggle(opt)}
                className="accent-purple-600" style={{cursor:'pointer'}}/>
              <span className="truncate" title={opt}>{opt}</span>
            </label>
          ))}
          {options.length === 0 && <div className={`px-3 py-2 text-xs ${txt2}`}>No options available</div>}
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
    <div className="relative" ref={ref}>
      <button
        onClick={() => setOpen(o => !o)}
        title={`Search ${label}`}
        className={`flex itemss-center gap-1.5 px-3 py-2 rounded-lg text-sm border font-medium transition-all
          ${darkMode ? 'bg-gray-600 border-gray-500 text-white hover:bg-gray-500' : 'bg-white border-gray-300 text-gray-700 hover:bg-purple-50 hover:border-purple-400'}`}
      >
        <Search className="w-4 h-4"/>
        <span>Search {label}</span>
        <ChevronDown className="w-3.5 h-3.5 ml-0.5 opacity-60"/>
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
              className="flex-1 px-3 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-xs font-semibold">
              Search
            </button>
            <button onClick={handleClear}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium ${darkMode?'bg-gray-600 text-gray-200 hover:bg-gray-500':'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>
              Reset
            </button>
          </div>
        </div>
      )}
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
          <Pie data={pieData} cx="50%" cy="42%" innerRadius={52} outerRadius={88}
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
          <div className="font-bold mb-1">Etc ({rest.length} status):</div>
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
    <div className="fixed inset-0 bg-black/60 z-50 flex itemss-center justify-center p-4 backdrop-blur-sm" onClick={onClose}>
      <div className={`rounded-2xl shadow-2xl w-full max-w-md ${bg}`} onClick={e=>e.stopPropagation()}>
        <div className={`flex justify-between itemss-center px-6 py-4 border-b ${darkMode?'border-gray-700':'border-gray-200'}`}>
          <div className="flex itemss-center gap-2">
            <EyeOff className="w-5 h-5 text-orange-500"/>
            <h3 className="font-bold text-base">Hide from Dashboard</h3>
          </div>
          <button onClick={onClose} className={`p-1.5 rounded-lg ${darkMode?'hover:bg-gray-700':'hover:bg-gray-100'}`}><X className="w-5 h-5"/></button>
        </div>
        <div className="px-6 py-5 space-y-4">
          <div>
            <label className={`block text-xs font-semibold mb-1.5 ${darkMode?'text-gray-300':'text-gray-600'}`}>Data Type</label>
            <div className="flex gap-3">
              {['PO','SO'].map(t=>(
                <label key={t} className="flex itemss-center gap-2 cursor-pointer">
                  <input type="radio" name="ref_type" value={t} checked={deleteForm.ref_type===t}
                    onChange={()=>setDeleteForm(f=>({...f,ref_type:t}))} className="accent-purple-600"/>
                  <span className="text-sm font-medium">{t === 'PO' ? 'PO HLI' : 'SO (Sales Order)'}</span>
                </label>
              ))}
            </div>
          </div>
          <div>
            <label className={`block text-xs font-semibold mb-1.5 ${darkMode?'text-gray-300':'text-gray-600'}`}>
              {deleteForm.ref_type === 'PO' ? 'PO HLI Number' : 'SO Number / SO Item'}
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
              placeholder="Enter reason for hiding this items from the dashboard..."
              rows={3}
              className={`w-full px-3 py-2 rounded-lg text-sm border resize-none ${inp}`}
            />
          </div>
          {deleteFormError && (
            <div className="flex itemss-center gap-2 text-red-500 text-sm bg-red-50 rounded-lg px-3 py-2">
              <AlertCircle className="w-4 h-4 flex-shrink-0"/>{deleteFormError}
            </div>
          )}
        </div>
        <div className={`px-6 py-4 border-t flex justify-end gap-3 ${darkMode?'border-gray-700':'border-gray-200'}`}>
          <button onClick={onClose} className={`px-4 py-2 rounded-lg text-sm font-medium ${darkMode?'bg-gray-600 text-gray-200 hover:bg-gray-500':'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>Cancel</button>
          <button onClick={onSubmit} className="px-5 py-2 bg-orange-600 hover:bg-orange-700 text-white rounded-lg text-sm font-semibold flex itemss-center gap-2">
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
  const fmtDt = (iso) => { try { return new Date(iso).toLocaleDateString('id-ID',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'}); } catch { return iso; } };
  return (
    <div className="fixed inset-0 bg-black/60 z-50 flex itemss-center justify-center p-4 backdrop-blur-sm" onClick={onClose}>
      <div className={`rounded-2xl shadow-2xl w-full max-w-2xl max-h-[80vh] flex flex-col ${bg}`} onClick={e=>e.stopPropagation()}>
        <div className={`flex justify-between itemss-center px-6 py-4 border-b ${darkMode?'border-gray-700':'border-gray-200'}`}>
          <div className="flex itemss-center gap-2">
            <Eye className="w-5 h-5 text-purple-500"/>
            <h3 className="font-bold text-base">Items Hidden from Dashboard</h3>
            <span className={`text-xs font-medium px-2 py-0.5 rounded-full ${darkMode?'bg-gray-700 text-gray-300':'bg-gray-100 text-gray-600'}`}>{hidden.length} items</span>
          </div>
          <button onClick={onClose} className={`p-1.5 rounded-lg ${darkMode?'hover:bg-gray-700':'hover:bg-gray-100'}`}><X className="w-5 h-5"/></button>
        </div>
        <div className="overflow-auto flex-1 p-4">
          {hidden.length === 0 ? (
            <div className={`text-center py-12 ${txt2}`}>
              <Eye className="w-10 h-10 mx-auto mb-2 opacity-40"/>
              <p className="text-sm">No hidden items</p>
            </div>
          ) : (
            <div className="space-y-3">
              {hidden.map(r=>(
                <div key={r.id} className={`flex itemss-start justify-between gap-4 p-4 rounded-xl border ${darkMode?'bg-gray-700 border-gray-600':'bg-gray-50 border-gray-200'}`}>
                  <div className="flex-1 min-w-0">
                    <div className="flex itemss-center gap-2 mb-1">
                      <span className={`px-2 py-0.5 rounded text-xs font-bold ${r.ref_type==='PO'?'bg-red-100 text-red-700':'bg-orange-100 text-orange-700'}`}>{r.ref_type}</span>
                      <span className="font-semibold text-sm">{r.ref_number}</span>
                    </div>
                    <p className={`text-xs ${txt2} mb-1`}><span className="font-medium">Reason:</span> {r.reason}</p>
                    <p className={`text-xs ${txt2}`}>📅 {fmtDt(r.requested_at)}</p>
                  </div>
                  <button onClick={()=>onRestore(r)}
                    className="flex itemss-center gap-1.5 px-3 py-1.5 bg-green-600 hover:bg-green-700 text-white rounded-lg text-xs font-semibold flex-shrink-0">
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

// ═══════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════
const App = () => {
  const [darkMode, setDarkMode] = useState(false);
  const [activePage, setActivePage] = useState('dashboard');

  const [stats, setStats] = useState(null);
  const [poWithoutSO, setPoWithoutSO] = useState([]);
  const [poFiltered, setPoFiltered] = useState([]); // after local filters
  const [agingData, setAgingData] = useState([]);
  const [allSOData, setAllSOData] = useState([]);
  const [soTotal, setSoTotal] = useState(0);
  const [soFilterOptions, setSoFilterOptions] = useState({ op_units: [], vendors: [], statuses: [] });

  // SO filters
  const [soFilters, setSoFilters] = useState({ op_units: [], vendors: [], statuses: [], aging: [] });
  const [soSearchNums, setSoSearchNums] = useState([]); // search SO Item
  const [soMarginFilter, setSoMarginFilter] = useState('all'); // 'all' | 'positive' | 'negative'
  const [soPage, setSoPage] = useState(1);
  const [soPerPage, setSoPerPage] = useState(20);

  // PO filters
  const [poPage, setPoPage] = useState(1);
  const [poPerPage, setPoPerPage] = useState(20);
  const [poSearchNums, setPoSearchNums] = useState([]); // search PO HLI Number
  const [poFilterItemType, setPoFilterItemType] = useState([]); // multi-select
  const [poFilterOpUnit, setPoFilterOpUnit] = useState([]);   // multi-select
  const [poItemTypeOptions, setPoItemTypeOptions] = useState([]);
  const [poOpUnitOptions, setPoOpUnitOptions] = useState([]);

  const [loading, setLoading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(null);
  const [toasts, setToasts] = useState([]);
  const [modal, setModal] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState('');
  const [downloadToast, setDownloadToast] = useState(null);
  const poTableRef = useRef(null);

  // Delete request / hide feature
  const [deleteRequests, setDeleteRequests] = useState([]);
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [showHiddenPanel, setShowHiddenPanel] = useState(false);
  const [deleteForm, setDeleteForm] = useState({ ref_type: 'PO', ref_number: '', reason: '' });
  const [deleteFormError, setDeleteFormError] = useState('');
  // Hide batch upload
  const [showHideMenu, setShowHideMenu] = useState(false);
  const hideMenuRef = useRef(null);
  useEffect(() => {
    const handler = (e) => { if (hideMenuRef.current && !hideMenuRef.current.contains(e.target)) setShowHideMenu(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const addToast = useCallback((message, type='success') => {
    const id = Date.now(); setToasts(t => [...t, { id, message, type }]);
  }, []);
  const removeToast = useCallback((id) => setToasts(t => t.filter(x => x.id !== id)), []);

  const fetchDashboard = useCallback(async () => {
    setLoading(true);
    try {
      const [sRes, pRes, aRes] = await Promise.all([
        api.get('/api/dashboard/stats'),
        api.get('/api/data/po-without-so'),
        api.get('/api/data/aging')
      ]);
      setStats(sRes.data);
      const poData = Array.isArray(pRes.data) ? pRes.data : [];
      // Filter out EXCLUDED op units on client side
      const poFiltered = poData.filter(p => !EXCLUDED_OP_UNITS.has(p.operation_unit));
      setPoWithoutSO(poFiltered);
      setPoFiltered(poFiltered);
      // Build filter options from PO data
      const itemsTypes = [...new Set(poFiltered.map(p=>p.po_items_type).filter(Boolean))].sort();
      const opUnits   = [...new Set(poFiltered.map(p=>p.operation_unit).filter(Boolean))].sort();
      setPoItemTypeOptions(itemsTypes);
      setPoOpUnitOptions(opUnits);
      setAgingData(Array.isArray(aRes.data) ? aRes.data : []);
    } catch (e) {
      addToast(`Error: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast]);

  // Helper: resolve filter array — __NONE__ means "nothing selected" (filter passes nothing)
  const resolveFilter = (val) => {
    if (val === '__NONE__') return ['__NONE_PLACEHOLDER__']; // backend will return 0 rows
    if (!Array.isArray(val) || val.length === 0) return []; // empty = no filter = all
    return val;
  };

  const fetchSOData = useCallback(async (filters, page, perPage, searchNums, marginFilter) => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      resolveFilter(filters.op_units).forEach(v => params.append('op_unit', v));
      resolveFilter(filters.vendors).forEach(v => params.append('vendor', v));
      resolveFilter(filters.statuses).forEach(v => params.append('status', v));
      (filters.aging || []).forEach(a => params.append('aging', a));
      (searchNums || []).forEach(n => params.append('so_items', n));
      if (marginFilter && marginFilter !== 'all') params.append('margin_filter', marginFilter);
      const res = await api.get(`/api/data/all-so?${params}`);
      setAllSOData(Array.isArray(res.data.data) ? res.data.data : []);
      setSoTotal(res.data.total || 0);
      setSoFilterOptions(res.data.filters || { op_units: [], vendors: [], statuses: [] });
    } catch (e) {
      addToast(`Failed to load SO data: ${e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast]);

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
    if (!deleteForm.reason.trim()) { setDeleteFormError('Reason wajib diisi'); return; }
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
      addToast(`✅ ${req.ref_type} ${req.ref_number} successfully restored to dashboard`, 'success');
      fetchDeleteRequests();
      fetchDashboard();
    } catch (e) {
      addToast(`❌ Restore failed: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  // Apply PO local filters whenever dependencies change

  // Apply PO local filters whenever dependencies change
  useEffect(() => {
    let filtered = [...poWithoutSO];
    if (poSearchNums.length > 0) {
      const nums = poSearchNums.map(n=>n.toLowerCase());
      filtered = filtered.filter(p => {
        const poHliKey = p.items_no ? `${p.po_no}-${p.items_no}`.toLowerCase() : (p.po_no||'').toLowerCase();
        return nums.some(n =>
          poHliKey.includes(n) ||
          (p.po_no||'').toLowerCase().includes(n)
        );
      });
    }
    // __NONE__ = nothing passes; [] = all pass; array = filter by those values
    if (poFilterItemType === '__NONE__') {
      filtered = [];
    } else if (Array.isArray(poFilterItemType) && poFilterItemType.length > 0) {
      filtered = filtered.filter(p => poFilterItemType.includes(p.po_items_type));
    }
    if (poFilterOpUnit === '__NONE__') {
      filtered = [];
    } else if (Array.isArray(poFilterOpUnit) && poFilterOpUnit.length > 0) {
      filtered = filtered.filter(p => poFilterOpUnit.includes(p.operation_unit));
    }
    setPoFiltered(filtered);
    setPoPage(1);
  }, [poWithoutSO, poSearchNums, poFilterItemType, poFilterOpUnit]);

  useEffect(() => { fetchDashboard(); fetchDeleteRequests(); }, []);
  useEffect(() => { if (activePage === 'all-so') fetchSOData(soFilters, soPage, soPerPage, soSearchNums, soMarginFilter); }, [activePage]);

  const handleUpload = async (e, type) => {
    const file = e.target.files[0]; if (!file) return;
    e.target.value = '';
    const label = type === 'po' ? 'HLI PO List (Item)' : 'SMRO - Search Client Odr';
    const endpoint = type === 'po' ? '/api/upload/po-list' : '/api/upload/smro';

    // ── Client-side header validation ──────────────────────────────────
    const REQUIRED_HEADERS = {
      po: {
        'PO Number':        ['po no.','po no','po number','po'],
        'Item No':          ['items no.','items no','items number','no. items'],
        'PO Item Type':     ['po items type','items type','type','po type'],
        'Supplier':         ['supplier','vendor','supplier name'],
        'Qty':              ['qty.','qty','quantity'],
        'Amount':           ['amount','total amount','total'],
        'PO Date':          ['po date','order date','tanggal po'],
        'Request Delivery': ['request delivery date','delivery date','req delivery'],
      },
      smro: {
        'SO Number':      ['so number','so no','so no.','so','sales order','sales order number','no so','nomor so'],
        'SO Item':        ['so items no','items no','line','so line','so items'],
        'SO Status':      ['so status','status','order status'],
        'Operation Unit': ['operation unit name','op unit','client name','client','operation unit'],
        'Vendor Name':    ['vendor name','vendor','supplier'],
        'Customer PO':    ['customer po number','customer po','po ref','po reference'],
        'Sales Amount':   ['sales amount(exclude tax)','sales amount','amount','total'],
        'SO Create Date': ['so create date','order date','so date','create date'],
      }
    };

    try {
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
      if (missing.length >= 3) {
        addToast(
          `❌ File tidak valid — ${missing.length} kolom penting tidak ditemsukan: ${missing.join(', ')}. Pastikan file ${label} yang benar, lalu cek kembali.`,
          'error'
        );
        return;
      }
    } catch (readErr) {
      addToast(`❌ Failed to read file: ${readErr.message}`, 'error');
      return;
    }
    // ── End client-side header validation ──────────────────────────────

    const fd = new FormData(); fd.append('file', file);
    setUploadProgress({ label, pct: 0 });
    try {
      const res = await api.post(endpoint, fd, {
        headers: { 'Content-Type': 'multipart/form-data' },
        onUploadProgress: (ev) => setUploadProgress({ label, pct: Math.round(ev.loaded*100/(ev.total||ev.loaded)) })
      });
      setUploadProgress(null);
      addToast(`✅ ${res.data.message}`, 'success');
      fetchDashboard();
      if (activePage === 'all-so') fetchSOData(soFilters, 1, soPerPage, soSearchNums, soMarginFilter);
      setSoPage(1);
    } catch (e) {
      setUploadProgress(null);
      addToast(`❌ Upload failed ${label}: ${e.response?.data?.error || e.message}`, 'error');
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
      fetchSOData(soFilters, soPage, soPerPage, soSearchNums, soMarginFilter);
    } catch (e) {
      setUploadProgress(null);
      addToast(`❌ Batch upload failed: ${e.response?.data?.error || e.message}`, 'error');
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
      addToast('❌ File download failed', 'error');
    }
  };

  const downloadHideTemplate = (type) => {
    setShowHideMenu(false);
    downloadBlob(`/api/template/hide?type=${type}`, `Template_Hide_${type === 'SO' ? 'SO' : 'PO_HLI'}.xlsx`, `Template Hide ${type}`);
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
      addToast(`❌ Upload failed hide batch: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  const downloadSOExcel = () => {
    const p = new URLSearchParams();
    (soFilters.op_units||[]).forEach(v => p.append('op_unit', v));
    (soFilters.vendors||[]).forEach(v => p.append('vendor', v));
    (soFilters.statuses||[]).forEach(v => p.append('status', v));
    downloadBlob(`/api/export/all-so?${p}`, `SO_List_${new Date().toISOString().slice(0,10)}.xlsx`, 'SO List');
  };
  const downloadPOExcel = () => downloadBlob('/api/export/po-without-so', `PO_Without_SO_${new Date().toISOString().slice(0,10)}.xlsx`, 'PO Without SO');
  const downloadSOTemplate = () => {
    setDownloadToast({ message: 'Preparing template...' });
    setTimeout(() => {
      const ws = XLSX.utils.json_to_sheet(allSOData.map(s=>({'SO Number':s.so_number,'Delivery Plan Date':s.delivery_plan_date||'','Remarks':s.remarks||''})));
      const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Template');
      saveAs(new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`SO_Template_${new Date().toISOString().slice(0,10)}.xlsx`);
      setDownloadToast(null);
      addToast('✅ Template downloaded successfully', 'success');
    }, 300);
  };

  const updateSOCell = async (soId, field, value) => {
    setEditingCell(null);
    try {
      await api.put(`/api/data/so/${soId}`, { [field]: value });
      setAllSOData(prev => prev.map(s => s.id === soId ? { ...s, [field]: value } : s));
    } catch (e) { addToast(`❌ Update failed: ${e.message}`, 'error'); }
  };

  const openModal = async (title, endpointOrData) => {
    if (Array.isArray(endpointOrData)) { setModal({ title, data: endpointOrData }); return; }
    try {
      const res = await api.get(endpointOrData);
      setModal({ title, data: Array.isArray(res.data) ? res.data : [] });
    } catch (e) { addToast(`❌ Failed to load details: ${e.message}`, 'error'); }
  };

  const toggleAgingFilter = (label) => {
    setSoFilters(f => {
      const aging = f.aging.includes(label) ? f.aging.filter(a=>a!==label) : [...f.aging, label];
      return {...f, aging};
    });
  };

  const poTotalPages = Math.max(1, Math.ceil(poFiltered.length / poPerPage));
  const poRows = poFiltered.slice((poPage-1)*poPerPage, poPage*poPerPage);
  const soTotalPages = Math.max(1, Math.ceil(soTotal / soPerPage));

  // PO KPI — unique PO numbers from filtered set (excluding consumable)
  const uniquePOCount = new Set(poFiltered.map(p=>p.po_no)).size;

  const card  = darkMode ? 'bg-gray-800 border border-gray-700' : 'bg-white border border-gray-100';
  const txt   = darkMode ? 'text-white' : txt;
  const txt2  = darkMode ? 'text-gray-400' : 'text-gray-600';
  const tblHd = darkMode ? 'bg-gray-700' : 'bg-purple-50';
  const tblDv = darkMode ? 'divide-gray-700' : 'divide-gray-100';
  const trHov = darkMode ? 'hover:bg-gray-700' : 'hover:bg-purple-50';

  const fmtDateRange = (range) => {
    if (!range?.min) return 'No data yet';
    return `${fmtDate(range.min)} — ${fmtDate(range.max)}`;
  };

  // ══════════════════════════════════════════════════════════════
  // RENDER DASHBOARD
  // ══════════════════════════════════════════════════════════════
  const renderDashboard = () => (
    <>
      {/* Date Range Info Bar */}
      <div className={`mb-4 px-5 py-3 rounded-xl flex flex-wrap gap-6 text-xs ${darkMode?'bg-gray-800 border border-gray-700':'bg-white border border-gray-100'} shadow`}>
        <div className="flex itemss-center gap-2">
          <Calendar className="w-4 h-4 text-purple-500"/>
          <span className={txt2}>PO Date Range:</span>
          <span className={`font-semibold ${txt}`}>{fmtDateRange(stats?.po_date_range)}</span>
        </div>
        <div className="flex itemss-center gap-2">
          <Calendar className="w-4 h-4 text-blue-500"/>
          <span className={txt2}>SO Create Date Range:</span>
          <span className={`font-semibold ${txt}`}>{fmtDateRange(stats?.so_date_range)}</span>
        </div>
      </div>

      {/* KPI Row */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <div className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all cursor-pointer ${card}`}
          onClick={() => {
            setActivePage('all-so');
            setSoPage(1);
            fetchSOData(soFilters, 1, soPerPage, soSearchNums, soMarginFilter);
            setTimeout(() => { poTableRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 300);
          }}>
          <div className="flex justify-between itemss-start">
            <div>
              <p className={`text-sm font-medium ${txt2}`}>PO HLI without SO</p>
              {/* Use client-side filtered count which excludes CONSUMABLE */}
              <h3 className="text-3xl font-bold mt-1 text-red-500">{fmtNum(uniquePOCount)}</h3>
              <p className={`text-xs mt-1 ${txt2}`}>unique PO numbers · click for details</p>
            </div>
            <div className="p-3 bg-red-100 rounded-xl"><AlertCircle className="w-6 h-6 text-red-500"/></div>
          </div>
        </div>

        <div className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all cursor-pointer ${card}`}
          onClick={() => openModal('SO Without PO HLI', '/api/data/so-without-po')}>
          <div className="flex justify-between itemss-start">
            <div>
              <p className={`text-sm font-medium ${txt2}`}>SO without PO HLI</p>
              <h3 className="text-3xl font-bold mt-1 text-orange-500">{fmtNum(stats?.so_without_po)}</h3>
              <p className={`text-xs mt-1 ${txt2}`}>Click for details</p>
            </div>
            <div className="p-3 bg-orange-100 rounded-xl"><XCircle className="w-6 h-6 text-orange-500"/></div>
          </div>
        </div>

        <div className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all ${card}`}>
          <div className="flex justify-between itemss-start">
            <div>
              <p className={`text-sm font-medium ${txt2}`}>Total PO HLI Amount</p>
              <h3 className={`text-xl font-bold mt-1 text-purple-600`}>{fmtCurShort(stats?.total_po_amount)}</h3>
              <p className={`text-xs mt-1 ${txt2}`}>{fmtCur(stats?.total_po_amount)}</p>
            </div>
            <div className="p-3 bg-purple-100 rounded-xl"><DollarSign className="w-6 h-6 text-purple-600"/></div>
          </div>
        </div>

        <div className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all cursor-pointer ${card}`}
          onClick={() => {
            setActivePage('all-so');
            setSoPage(1);
            fetchSOData(soFilters, 1, soPerPage, soSearchNums, soMarginFilter);
            window.scrollTo({ top: 0, behavior: 'smooth' });
          }}>
          <div className="flex justify-between itemss-start">
            <div>
              <p className={`text-sm font-medium ${txt2}`}>Total SO (Open)</p>
              <h3 className="text-3xl font-bold mt-1 text-green-600">{fmtNum(stats?.total_so_count)}</h3>
              <p className={`text-xs mt-1 ${txt2}`}>{stats?.so_date_range?.max ? fmtDate(stats.so_date_range.max) : 'No upload yet'} · click for details</p>
            </div>
            <div className="p-3 bg-green-100 rounded-xl"><CheckCircle className="w-6 h-6 text-green-600"/></div>
          </div>
        </div>
      </div>

      {/* Charts Row 1 */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6 itemss-start">
        <div className={`p-6 rounded-2xl shadow ${card}`}>
          <h3 className={`text-base font-bold mb-4 flex itemss-center gap-2 ${txt}`}>
            <TrendingUp className="w-5 h-5 text-purple-600"/> Monthly Open SO Trend
          </h3>
          <ResponsiveContainer width="100%" height={190}>
            <AreaChart data={stats?.monthly_trend||[]}>
              <defs>
                <linearGradient id="cSO" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#8B5CF6" stopOpacity={0.3}/><stop offset="95%" stopColor="#8B5CF6" stopOpacity={0}/>
                </linearGradient>
                <linearGradient id="cAmt" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#F97316" stopOpacity={0.3}/><stop offset="95%" stopColor="#F97316" stopOpacity={0}/>
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke={darkMode?'#374151':'#E5E7EB'}/>
              <XAxis dataKey="month" stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={10}/>
              <YAxis yAxisId="left" stroke="#8B5CF6" fontSize={10}/>
              <YAxis yAxisId="right" orientation="right" stroke="#F97316" fontSize={10}/>
              <Tooltip contentStyle={{backgroundColor:darkMode?'#1F2937':'#fff',borderRadius:'8px'}}/>
              <Legend iconSize={8} wrapperStyle={{fontSize:'11px'}}/>
              <Area yAxisId="left" type="monotone" dataKey="so_count" name="SO Count" stroke="#8B5CF6" strokeWidth={2} fill="url(#cSO)"/>
              <Area yAxisId="right" type="monotone" dataKey="amount" name="Value (IDR Million)" stroke="#F97316" strokeWidth={2} fill="url(#cAmt)"/>
            </AreaChart>
          </ResponsiveContainer>
        </div>

        <div className="flex flex-col gap-4">
          <div className={`p-5 rounded-2xl shadow ${card}`}>
            <h3 className={`text-sm font-bold mb-3 flex itemss-center gap-2 ${txt}`}>
              <BarChart3 className="w-4 h-4 text-blue-600"/> Top 5 Vendors (Open SO)
            </h3>
            <table className="w-full text-xs">
              <thead className={tblHd}>
                <tr>
                  <th className={`p-1.5 text-left font-semibold ${txt2}`}>#</th>
                  <th className={`p-1.5 text-left font-semibold ${txt2}`}>Vendor</th>
                  <th className={`p-1.5 text-right font-semibold ${txt2}`}>Open SO</th>
                  <th className={`p-1.5 text-right font-semibold ${txt2}`}>Amount</th>
                </tr>
              </thead>
              <tbody className={`divide-y ${tblDv}`}>
                {(stats?.top_vendors||[]).map((v,i)=>(
                  <tr key={i} className={`${trHov} cursor-pointer`}
                    onClick={()=>openModal(`Vendor: ${v.vendor}`, `/api/data/top-vendor-detail/${encodeURIComponent(v.vendor)}`)}>
                    <td className="p-1.5">
                      <span className={`inline-flex itemss-center justify-center w-6 h-6 rounded text-xs font-bold ${i===0?'bg-yellow-100 text-yellow-700':i===1?'bg-gray-200 text-gray-700':i===2?'bg-orange-100 text-orange-700':'bg-purple-100 text-purple-700'}`}>#{i+1}</span>
                    </td>
                    <td className={`p-1.5 font-medium ${txt} max-w-[120px] truncate`} title={v.vendor}>{v.vendor}</td>
                    <td className="p-1.5 text-right font-semibold text-purple-600">{fmtNum(v.so_count)}</td>
                    <td className="p-1.5 text-right font-semibold text-orange-600">{fmtCurShort(v.total_amount)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Charts Row 2 */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6 itemss-stretch">
        <div className={`p-6 rounded-2xl shadow flex flex-col ${card}`}>
          <h3 className={`text-base font-bold mb-4 flex itemss-center gap-2 ${txt}`}>
            <FileText className="w-5 h-5 text-green-600"/> SO Status Distribution
          </h3>
          {(() => {
            const months = stats?.status_months || [];
            const rows   = stats?.so_status_monthly || [];
            const totByMonth = months.reduce((acc, m) => { acc[m] = rows.reduce((s, r) => s + (r.monthly?.[m] || 0), 0); return acc; }, {});
            const grandTotal  = rows.reduce((s, r) => s + r.total, 0);
            const grandAmount = rows.reduce((s, r) => s + (r.amount || 0), 0);
            return (
              <div className="overflow-auto flex-1">
                <table className="w-full text-xs" style={{minWidth: months.length > 4 ? `${160 + months.length * 72 + 200}px` : undefined}}>
                  <thead className={`sticky top-0 ${tblHd}`}>
                    <tr>
                      <th className={`px-3 py-2 text-left font-semibold whitespace-nowrap ${txt2} sticky left-0 ${darkMode?'bg-gray-700':'bg-purple-50'}`}>Status</th>
                      {months.map(m => <th key={m} className={`px-2 py-2 text-center font-semibold whitespace-nowrap ${txt2}`}>{m}</th>)}
                      <th className={`px-3 py-2 text-right font-semibold whitespace-nowrap ${txt2}`}>Total</th>
                      <th className={`px-3 py-2 text-right font-semibold whitespace-nowrap ${txt2}`}>%</th>
                      <th className={`px-3 py-2 text-right font-semibold whitespace-nowrap ${txt2}`}>Sales Amount</th>
                    </tr>
                  </thead>
                  <tbody className={`divide-y ${tblDv}`}>
                    {rows.map((s, i) => (
                      <tr key={i} className={trHov}>
                        <td
                          className={`px-3 py-2 font-medium whitespace-nowrap sticky left-0 ${darkMode?'bg-gray-800':'bg-white'} ${txt} cursor-pointer hover:text-purple-600 hover:underline`}
                          onClick={() => openModal(`SO Status: ${s.name}`, `/api/data/so-status-detail/${encodeURIComponent(s.name)}`)}>
                          {s.name}
                        </td>
                        {months.map(m => {
                          const val = s.monthly?.[m];
                          return val ? (
                            <td key={m} className="px-2 py-2 text-center font-semibold text-white" style={{backgroundColor:'#7C3AED'}}>
                              <button
                                onClick={() => openModal(`SO Status: ${s.name} — ${m}`, `/api/data/so-status-detail/${encodeURIComponent(s.name)}?month=${encodeURIComponent(m)}`)}
                                className="font-semibold underline-offset-2 hover:underline cursor-pointer text-white w-full">
                                {fmtNum(val)}
                              </button>
                            </td>
                          ) : (
                            <td key={m} className="px-2 py-2 text-center" style={{backgroundColor: darkMode?'rgba(59,130,246,0.08)':'rgba(219,234,254,0.45)'}}></td>
                          );
                        })}
                        <td className="px-3 py-2 text-right font-bold text-purple-600">
                          <button
                            onClick={() => openModal(`SO Status: ${s.name}`, `/api/data/so-status-detail/${encodeURIComponent(s.name)}`)}
                            className="font-bold text-purple-600 hover:underline cursor-pointer">
                            {fmtNum(s.total)}
                          </button>
                        </td>
                        <td className="px-3 py-2 text-right text-green-600">{s.percentage}%</td>
                        <td className="px-3 py-2 text-right text-orange-600 whitespace-nowrap">{fmtCurShort(s.amount)}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot className={`${tblHd} font-bold`}>
                    <tr>
                      <td className={`px-3 py-2 sticky left-0 ${darkMode?'bg-gray-700':'bg-purple-50'} ${txt}`}>TOTAL</td>
                      {months.map(m => (
                        <td key={m} className="px-2 py-2 text-center">
                          {totByMonth[m] ? (
                            <button
                              onClick={() => openModal(`All Statuses — ${m}`, `/api/data/so-status-detail-all?month=${encodeURIComponent(m)}`)}
                              className="font-bold text-purple-600 hover:underline cursor-pointer">
                              {fmtNum(totByMonth[m])}
                            </button>
                          ) : ''}
                        </td>
                      ))}
                      <td className="px-3 py-2 text-right">
                        <button
                          onClick={() => openModal('All SO', '/api/data/so-status-detail-all')}
                          className="font-bold text-purple-600 hover:underline cursor-pointer">
                          {fmtNum(grandTotal)}
                        </button>
                      </td>
                      <td className="px-3 py-2 text-right text-green-600">100%</td>
                      <td className="px-3 py-2 text-right text-orange-600 whitespace-nowrap">{fmtCurShort(grandAmount)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            );
          })()}
        </div>

        <div className="flex flex-col gap-4">
          <div className={`p-5 rounded-2xl shadow ${card}`}>
            <h3 className={`text-sm font-bold mb-3 flex itemss-center gap-2 ${txt}`}>
              <Building2 className="w-4 h-4 text-green-600"/> Total Open SO per Operation Unit
            </h3>
            <div className="overflow-auto max-h-40">
              <table className="w-full text-xs">
                <thead className={`sticky top-0 ${tblHd}`}>
                  <tr>
                    <th className={`p-1.5 text-left font-semibold ${txt2}`}>Operation Unit</th>
                    <th className={`p-1.5 text-right font-semibold ${txt2}`}>Open SO</th>
                    <th className={`p-1.5 text-right font-semibold ${txt2}`}>Amount</th>
                  </tr>
                </thead>
                <tbody className={`divide-y ${tblDv}`}>
                  {(stats?.top_op_units||[]).map((u,i)=>(
                    <tr key={i} className={`${trHov} transition-colors`}>
                      <td className={`p-1.5 font-medium ${txt} max-w-[160px] truncate`} title={u.op_unit}>{u.op_unit}</td>
                      <td className="p-1.5 text-right font-semibold text-purple-600">{fmtNum(u.so_count)}</td>
                      <td className="p-1.5 text-right font-semibold text-orange-600">{fmtCurShort(u.total_amount)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="grid gap-4 flex-1" style={{gridTemplateColumns:'3fr 2fr'}}>
            <div className={`p-5 rounded-2xl shadow ${card}`}>
              <h3 className={`text-sm font-bold mb-2 flex itemss-center gap-2 ${txt}`}><BarChart3 className="w-4 h-4 text-orange-600"/> SO Status (Pie)</h3>
              <StatusPie data={stats?.so_status} darkMode={darkMode}/>
            </div>
            {(() => {
              const agingPieData = [
                { name:'< 30 Days', value:agingData.reduce((s,v)=>s+(v.less_30||0),0), fill:'#10B981' },
                { name:'30–90 Days', value:agingData.reduce((s,v)=>s+(v.days_30_90||0),0), fill:'#F59E0B' },
                { name:'90–180 Days', value:agingData.reduce((s,v)=>s+(v.days_90_180||0),0), fill:'#F97316' },
                { name:'> 180 Days', value:agingData.reduce((s,v)=>s+(v.more_180||0),0), fill:'#EF4444' },
              ].filter(d=>d.value>0);
              return (
                <div className={`p-5 rounded-2xl shadow ${card}`}>
                  <h3 className={`text-sm font-bold mb-2 flex itemss-center gap-2 ${txt}`}><Calendar className="w-4 h-4 text-red-500"/> SO Aging (Pie)</h3>
                  <ResponsiveContainer width="100%" height={300}>
                    <PieChart>
                      <Pie data={agingPieData} cx="50%" cy="40%" innerRadius={52} outerRadius={88} paddingAngle={0} dataKey="value" labelLine={false} label={renderPctLabel}>
                        {agingPieData.map((d,i)=><Cell key={i} fill={d.fill}/>)}
                      </Pie>
                      <Tooltip contentStyle={{backgroundColor:darkMode?'#1F2937':'#fff',borderRadius:'8px'}} formatter={(v,n)=>[fmtNum(v)+' SO',n]}/>
                      <Legend layout="horizontal" align="center" verticalAlign="bottom" iconSize={8} formatter={(v)=><span className="text-xs">{v}</span>}/>
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              );
            })()}
          </div>
        </div>
      </div>

      {/* SO Aging Table */}
      <div className={`p-6 rounded-2xl shadow mb-6 ${card}`}>
        <h3 className={`text-base font-bold mb-4 flex itemss-center gap-2 ${txt}`}>
          <Calendar className="w-5 h-5 text-red-600"/> SO Aging — Open SO by Vendor
        </h3>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className={tblHd}>
              <tr>{['Vendor (SMRO)','< 30 Days','30–90 Days','90–180 Days','> 180 Days','Total Open','Sales Amount'].map(h=>(
                <th key={h} className={`p-3 text-center font-semibold ${txt2} first:text-left`}>{h}</th>
              ))}</tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {agingData.slice(0,15).map((v,i)=>{
                const openDetail = (bucket) => {
                  const url = bucket
                    ? `/api/data/aging-detail/${encodeURIComponent(v.vendor)}?bucket=${encodeURIComponent(bucket)}`
                    : `/api/data/aging-detail/${encodeURIComponent(v.vendor)}`;
                  const label = bucket ? `${v.vendor} — ${bucket} days` : `Aging Detail: ${v.vendor}`;
                  openModal(label, url);
                };
                const cellBtn = (val, bucket, colorClass) => val > 0 ? (
                  <button onClick={e=>{e.stopPropagation();openDetail(bucket);}}
                    className={`font-semibold underline-offset-2 hover:underline cursor-pointer ${colorClass}`}>
                    {fmtNum(val)}
                  </button>
                ) : <span className={`${colorClass} opacity-40`}>0</span>;
                return (
                  <tr key={i} className={`${trHov}`}>
                    <td className={`p-3 font-medium text-xs cursor-pointer hover:text-purple-600 ${txt}`}
                      onClick={()=>openDetail(null)}>{v.vendor}</td>
                    <td className="p-3 text-center">{cellBtn(v.less_30,'0-30','text-green-600')}</td>
                    <td className="p-3 text-center">{cellBtn(v.days_30_90,'30-90','text-yellow-600')}</td>
                    <td className="p-3 text-center">{cellBtn(v.days_90_180,'90-180','text-orange-600')}</td>
                    <td className="p-3 text-center">{cellBtn(v.more_180,'180+','text-red-600')}</td>
                    <td className="p-3 text-center">
                      <button onClick={()=>openDetail(null)}
                        className="font-bold text-purple-600 hover:underline cursor-pointer">
                        {fmtNum(v.total_open)}
                      </button>
                    </td>
                    <td className="p-3 text-right font-semibold text-orange-600 text-xs">{fmtCurShort(v.sales_amount)}</td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot className={`${tblHd} font-bold text-sm`}>
              {(() => {
                const tot = agingData.reduce((acc,v)=>({
                  less_30:acc.less_30+(v.less_30||0), days_30_90:acc.days_30_90+(v.days_30_90||0),
                  days_90_180:acc.days_90_180+(v.days_90_180||0), more_180:acc.more_180+(v.more_180||0),
                  total_open:acc.total_open+(v.total_open||0), sales_amount:acc.sales_amount+(v.sales_amount||0),
                }), {less_30:0,days_30_90:0,days_90_180:0,more_180:0,total_open:0,sales_amount:0});
                const totCellBtn = (val, bucket, colorClass) => val > 0 ? (
                  <button
                    onClick={() => openModal(`All Vendors — ${bucket} days`, `/api/data/aging-detail-all?bucket=${encodeURIComponent(bucket)}`)}
                    className={`font-bold underline-offset-2 hover:underline cursor-pointer ${colorClass}`}>
                    {fmtNum(val)}
                  </button>
                ) : <span className={colorClass}>{fmtNum(val)}</span>;
                return (
                  <tr>
                    <td className={`p-3 font-bold ${txt}`}>TOTAL</td>
                    <td className="p-3 text-center">{totCellBtn(tot.less_30,'0-30','text-green-700')}</td>
                    <td className="p-3 text-center">{totCellBtn(tot.days_30_90,'30-90','text-yellow-700')}</td>
                    <td className="p-3 text-center">{totCellBtn(tot.days_90_180,'90-180','text-orange-700')}</td>
                    <td className="p-3 text-center">{totCellBtn(tot.more_180,'180+','text-red-700')}</td>
                    <td className="p-3 text-center">
                      <button
                        onClick={() => openModal('All Vendors — All Aging Buckets', '/api/data/aging-detail-all')}
                        className="font-bold text-purple-700 hover:underline cursor-pointer">
                        {fmtNum(tot.total_open)}
                      </button>
                    </td>
                    <td className="p-3 text-right font-bold text-orange-700 text-xs">{fmtCurShort(tot.sales_amount)}</td>
                  </tr>
                );
              })()}
            </tfoot>
          </table>
        </div>
      </div>
    </>
  );

  // ══════════════════════════════════════════════════════════════
  // RENDER ALL SO PAGE
  // ══════════════════════════════════════════════════════════════
  const renderAllSO = () => (
    <div>
      {/* All SO Table */}
      <div className={`p-6 rounded-2xl shadow mb-6 ${card}`}>
        <div className="flex flex-wrap justify-between itemss-center gap-3 mb-5">
          <div>
            <h2 className={`text-xl font-bold ${txt}`}>Open SO (Sales Order)</h2>
            <p className={`text-sm ${txt2}`}>{fmtNum(soTotal)} total records — page {soPage} of {soTotalPages}</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <label className="flex itemss-center gap-1 px-3 py-1.5 bg-green-700 hover:bg-green-800 text-white rounded-lg text-sm font-medium shadow-sm">
              <Upload className="w-4 h-4"/>Batch Upload
              <input type="file" accept=".xlsx,.xls" onChange={handleBatchUpload} className="hidden"/>
            </label>
            <DownloadButton onClick={downloadSOTemplate} className="flex itemss-center gap-1 px-3 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-sm font-medium shadow-sm">
              <FileSpreadsheet className="w-4 h-4"/>Template
            </DownloadButton>
            <DownloadButton onClick={downloadSOExcel} className="flex itemss-center gap-1 px-3 py-1.5 bg-purple-700 hover:bg-purple-800 text-white rounded-lg text-sm font-medium shadow-sm">
              <Download className="w-4 h-4"/>Download Excel
            </DownloadButton>
          </div>
        </div>

        {/* Aging Filter Chips */}
        <div className="mb-3 flex flex-wrap gap-2 itemss-center">
          <span className={`text-xs font-medium ${txt2}`}>Aging Filter:</span>
          {AGING_LABELS.map(label => {
            const active = soFilters.aging.includes(label);
            return (
              <button key={label} onClick={()=>toggleAgingFilter(label)}
                className={`px-3 py-1 rounded-full text-xs font-semibold border transition-all ${active?'text-white border-transparent':'border-gray-200 text-gray-400 bg-gray-100'}`}
                style={active ? {backgroundColor: AGING_COLORS[label], borderColor: AGING_COLORS[label]} : {}}>
                {label} days
              </button>
            );
          })}
          {soFilters.aging.length > 0 && (
            <button onClick={()=>setSoFilters(f=>({...f,aging:[]}))}
              className={`px-2 py-1 rounded text-xs ${txt2} hover:text-red-500`}>Reset Aging</button>
          )}
        </div>

        {/* Multi-select Filters row — Search SO leftmost */}
        <div className={`p-4 rounded-xl mb-4 ${darkMode?'bg-gray-700':'bg-gray-50'}`}>
          <div className="flex flex-wrap gap-3 itemss-end">
            {/* Search SO Item — paling kiri */}
            <div>
              <label className={`block text-xs font-medium mb-1 ${txt2}`}>Search SO Item</label>
              <SearchInput
                label="SO Item"
                placeholder={"e.g.\n1234-10\n1234-20"}
                onSearch={(nums) => {
                  setSoSearchNums(nums);
                  setSoPage(1);
                  fetchSOData(soFilters, 1, soPerPage, nums, soMarginFilter);
                }}
                darkMode={darkMode} txt2={txt2}
              />
            </div>
            <MultiSelect label="Operation Unit" options={soFilterOptions.op_units}
              selected={soFilters.op_units} onChange={v=>setSoFilters(f=>({...f,op_units:v}))}
              darkMode={darkMode} txt2={txt2}/>
            <MultiSelect label="Vendor Name" options={soFilterOptions.vendors}
              selected={soFilters.vendors} onChange={v=>setSoFilters(f=>({...f,vendors:v}))}
              darkMode={darkMode} txt2={txt2}/>
            <MultiSelect label="SO Status" options={soFilterOptions.statuses}
              selected={soFilters.statuses} onChange={v=>setSoFilters(f=>({...f,statuses:v}))}
              darkMode={darkMode} txt2={txt2}/>
            <div className="flex-1 min-w-[140px]">
              <label className={`block text-xs font-medium mb-1 ${txt2}`}>Margin Filter</label>
              <select className={`w-full px-3 py-2 rounded-lg text-sm border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                value={soMarginFilter} onChange={e=>setSoMarginFilter(e.target.value)}>
                <option value="all">All Data</option>
                <option value="positive">With Margin (≥ 0)</option>
                <option value="negative">Negative Margin (&lt; 0)</option>
              </select>
            </div>
            <div className="flex-1 min-w-[100px]">
              <label className={`block text-xs font-medium mb-1 ${txt2}`}>Rows per Page</label>
              <select className={`w-full px-3 py-2 rounded-lg text-sm border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                value={soPerPage} onChange={e=>setSoPerPage(Number(e.target.value))}>
                <option value={20}>20</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
                <option value={500}>500</option>
              </select>
            </div>
            <div className="flex gap-2">
              <button onClick={()=>{ setSoPage(1); fetchSOData(soFilters,1,soPerPage,soSearchNums,soMarginFilter); }}
                className="px-5 py-2 bg-purple-700 hover:bg-purple-800 text-white rounded-lg text-sm font-semibold shadow-sm">Apply</button>
              <button onClick={()=>{
                const f={op_units:[],vendors:[],statuses:[],aging:[]};
                setSoFilters(f); setSoSearchNums([]); setSoMarginFilter('all'); setSoPage(1);
                fetchSOData(f,1,soPerPage,[],'all');
              }}
                className={`px-4 py-2 rounded-lg text-sm font-medium shadow-sm ${darkMode?'bg-gray-500 text-gray-100 hover:bg-gray-400':'bg-gray-400 text-white hover:bg-gray-500'}`}>Reset</button>
            </div>
          </div>
          {/* Active filter tags */}
          {(soSearchNums.length + soFilters.op_units.length + soFilters.vendors.length + soFilters.statuses.length) > 0 && (
            <div className="mt-3 flex flex-wrap gap-1.5">
              {soSearchNums.map(v=>(
                <span key={v} className="flex itemss-center gap-1 px-2 py-0.5 bg-indigo-100 text-indigo-700 rounded-full text-xs">
                  SO: {v}<button onClick={()=>{ const next=soSearchNums.filter(x=>x!==v); setSoSearchNums(next); setSoPage(1); fetchSOData(soFilters,1,soPerPage,next,soMarginFilter); }} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {soFilters.op_units.map(v=>(
                <span key={v} className="flex itemss-center gap-1 px-2 py-0.5 bg-purple-100 text-purple-700 rounded-full text-xs">
                  {v}<button onClick={()=>setSoFilters(f=>({...f,op_units:f.op_units.filter(x=>x!==v)}))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {soFilters.vendors.map(v=>(
                <span key={v} className="flex itemss-center gap-1 px-2 py-0.5 bg-blue-100 text-blue-700 rounded-full text-xs">
                  {v}<button onClick={()=>setSoFilters(f=>({...f,vendors:f.vendors.filter(x=>x!==v)}))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {soFilters.statuses.map(v=>(
                <span key={v} className="flex itemss-center gap-1 px-2 py-0.5 bg-green-100 text-green-700 rounded-full text-xs">
                  {v}<button onClick={()=>setSoFilters(f=>({...f,statuses:f.statuses.filter(x=>x!==v)}))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
            </div>
          )}
        </div>

        {/* SO Table — removed SO Number column, SO Item is leftmost */}
        <div className="overflow-x-auto rounded-lg border border-gray-200">
          <table className="w-full text-sm">
            <thead className={tblHd}>
              <tr>
                {['Aging','SO Item','Item Name','Status','Op Unit','Vendor','Qty',
                  'Sales Price','Sales Amount','PO Price','PO Amount','Margin','%Margin',
                  'Possible Delivery','Plan Date','Remarks'].map(h=>(
                  <th key={h} className={`px-3 py-2.5 text-left font-semibold whitespace-nowrap ${txt2}`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {(() => {
                if (allSOData.length === 0) return (
                  <tr><td colSpan={16} className={`px-4 py-10 text-center ${txt2}`}>
                    <FileText className="w-10 h-10 mx-auto mb-2 opacity-40"/>No data available
                  </td></tr>
                );
                return allSOData.map((so) => {
                const isDeliveryCompleted = so.so_status === 'Delivery Completed';
                const poAmount = (so.purchasing_price || 0) * (so.so_qty || 0);
                const margin = (so.sales_amount || 0) - poAmount;
                const marginPct = poAmount !== 0 ? (margin / poAmount) * 100 : null;
                const marginColor = margin < 0 ? 'text-red-600 font-semibold' : margin > 0 ? 'text-green-600 font-semibold' : txt2;
                return (
                <tr key={so.id} className={`${trHov} transition-colors`}>
                  <td className="px-3 py-2 whitespace-nowrap">
                    {!isDeliveryCompleted && so.aging_label && so.aging_label !== 'No Date' ? (
                      <span className="px-2 py-0.5 rounded-full text-xs font-bold text-white"
                        style={{backgroundColor: AGING_COLORS[so.aging_label] || '#6B7280'}}>
                        {so.aging_label}
                      </span>
                    ) : null}
                  </td>
                  {/* SO Item first, no SO Number column */}
                  <td className="px-3 py-2 text-purple-600 font-medium whitespace-nowrap">{so.so_items}</td>
                  <td className={`px-3 py-2 max-w-[160px] truncate ${txt2}`} title={so.product_name}>{so.product_name}</td>
                  <td className="px-3 py-2 whitespace-nowrap">
                    <span className={`px-2 py-0.5 rounded-full text-xs font-medium ${
                      so.so_status==='Delivery Completed'?'bg-green-100 text-green-700':
                      so.so_status==='SO Cancel'?'bg-red-100 text-red-700':'bg-blue-100 text-blue-700'}`}>
                      {so.so_status||'-'}
                    </span>
                  </td>
                  <td className={`px-3 py-2 min-w-[180px] truncate ${txt2}`} title={so.operation_unit_name}>{so.operation_unit_name}</td>
                  <td className={`px-3 py-2 max-w-[120px] truncate ${txt2}`} title={so.vendor_name}>{so.vendor_name}</td>
                  <td className={`px-3 py-2 text-right ${txt2}`}>{fmtNum(so.so_qty)}</td>
                  <td className="px-3 py-2 text-right whitespace-nowrap min-w-[130px]">{fmtCur(so.sales_price)}</td>
                  <td className="px-3 py-2 text-right font-semibold text-orange-600 whitespace-nowrap min-w-[130px]">{fmtCur(so.sales_amount)}</td>
                  <td className="px-3 py-2 text-right whitespace-nowrap min-w-[130px]">{fmtCur(so.purchasing_price)}</td>
                  <td className="px-3 py-2 text-right font-semibold text-green-600 whitespace-nowrap min-w-[130px]">{fmtCur(poAmount)}</td>
                  <td className={`px-3 py-2 text-right whitespace-nowrap min-w-[130px] ${marginColor}`}>{fmtCur(margin)}</td>
                  <td className={`px-3 py-2 text-right whitespace-nowrap ${marginColor}`}>
                    {marginPct !== null ? `${marginPct.toFixed(1)}%` : '-'}
                  </td>
                  <td className={`px-3 py-2 text-center text-xs ${txt2}`}>{so.delivery_possible_date||'-'}</td>
                  <td className="px-3 py-2 text-center">
                    {editingCell?.id===so.id && editingCell.field==='delivery_plan_date' ? (
                      <div className="flex itemss-center gap-1">
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
                      <div className="flex itemss-center justify-center gap-1 group">
                        <span className="cursor-pointer text-purple-600 hover:underline text-xs whitespace-nowrap"
                          onClick={()=>{setEditingCell({id:so.id,field:'delivery_plan_date'});setEditValue(so.delivery_plan_date||'');}}>
                          {so.delivery_plan_date||'✏️ Set'}
                        </span>
                        {so.delivery_plan_date && (
                          <button onClick={e=>{e.stopPropagation();updateSOCell(so.id,'delivery_plan_date','');}}
                            className="opacity-0 group-hover:opacity-100 text-red-400 hover:text-red-600 transition-all p-0.5"><X className="w-3 h-3"/></button>
                        )}
                      </div>
                    )}
                  </td>
                  <td className="px-3 py-2">
                    {editingCell?.id===so.id && editingCell.field==='remarks' ? (
                      <input type="text" defaultValue={so.remarks}
                        className={`w-full px-2 py-1 rounded text-xs border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                        onChange={e=>setEditValue(e.target.value)}
                        onBlur={()=>updateSOCell(so.id,'remarks',editValue)}
                        onKeyDown={e=>e.key==='Enter'&&updateSOCell(so.id,'remarks',editValue)}
                        autoFocus/>
                    ) : (
                      <span className={`cursor-pointer text-xs ${so.remarks?txt2:'text-orange-500 hover:underline'}`}
                        onClick={()=>{setEditingCell({id:so.id,field:'remarks'});setEditValue(so.remarks||'');}}>
                        {so.remarks||'✏️ Add'}
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

        {/* Pagination */}
        <div className={`mt-4 pt-3 border-t ${darkMode?'border-gray-700':'border-gray-200'} flex justify-between itemss-center`}>
          <span className={`text-sm ${txt2}`}>
            Showing {((soPage-1)*soPerPage)+1}–{Math.min(soPage*soPerPage,soTotal)} of {fmtNum(soTotal)}
          </span>
          <div className="flex gap-1 itemss-center">
            <button disabled={soPage===1} onClick={()=>{ const p=soPage-1; setSoPage(p); fetchSOData(soFilters,p,soPerPage,soSearchNums,soMarginFilter); }}
              className={`p-1.5 rounded ${soPage===1?'opacity-40':'hover:bg-purple-100'}`}><ChevronLeft className="w-4 h-4"/></button>
            <span className={`px-3 py-1 rounded text-sm font-semibold ${darkMode?'bg-gray-700 text-white':'bg-purple-100 text-purple-700'}`}>{soPage}/{soTotalPages}</span>
            <button disabled={soPage===soTotalPages} onClick={()=>{ const p=soPage+1; setSoPage(p); fetchSOData(soFilters,p,soPerPage,soSearchNums,soMarginFilter); }}
              className={`p-1.5 rounded ${soPage===soTotalPages?'opacity-40':'hover:bg-purple-100'}`}><ChevronRight className="w-4 h-4"/></button>
          </div>
        </div>
      </div>

      {/* PO HLI Without SO Table */}
      <div ref={poTableRef} className={`rounded-2xl shadow overflow-hidden ${card}`}>
        <div className={`p-5 border-b ${darkMode?'border-gray-700':'border-gray-100'} flex flex-wrap justify-between itemss-center gap-3`}>
          <div className="flex itemss-center gap-2">
            <AlertCircle className="w-5 h-5 text-yellow-600"/>
            <h3 className={`text-base font-bold ${txt}`}>PO HLI Without Matching SO</h3>
            <span className={`text-sm ${txt2}`}>
              ({fmtNum(new Set(poFiltered.map(p=>p.po_no)).size)} POs · {fmtNum(poFiltered.length)} line itemss)
            </span>
          </div>
          <div className="flex gap-2 itemss-center">
            <DownloadButton onClick={downloadPOExcel} className="flex itemss-center gap-1 px-4 py-1.5 bg-purple-700 hover:bg-purple-800 text-white rounded-lg text-sm font-medium shadow-sm">
              <Download className="w-4 h-4"/>Download Excel
            </DownloadButton>
          </div>
        </div>

        {/* PO Filters row */}
        <div className={`px-5 py-3 border-b ${darkMode?'border-gray-700 bg-gray-750':'border-gray-100 bg-gray-50'} flex flex-wrap gap-3 itemss-end`}>
          {/* Search PO HLI Number — paling kiri */}
          <div>
            <label className={`block text-xs font-medium mb-1 ${txt2}`}>Search PO Number</label>
            <SearchInput
              label="PO HLI Number"
              placeholder={"e.g.\n4502358819\n4502358819-10"}
              onSearch={(nums) => { setPoSearchNums(nums); setPoPage(1); }}
              darkMode={darkMode} txt2={txt2}
            />
          </div>
          <MultiSelect
            label="PO Item Type"
            options={poItemTypeOptions}
            selected={poFilterItemType}
            onChange={v => { setPoFilterItemType(v); setPoPage(1); }}
            darkMode={darkMode} txt2={txt2}
          />
          <MultiSelect
            label="Operation Unit"
            options={poOpUnitOptions}
            selected={poFilterOpUnit}
            onChange={v => { setPoFilterOpUnit(v); setPoPage(1); }}
            darkMode={darkMode} txt2={txt2}
          />
          <div className="flex-1 min-w-[120px]">
            <label className={`block text-xs font-medium mb-1 ${txt2}`}>Rows per Page</label>
            <select className={`w-full px-3 py-2 rounded-lg text-sm border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
              value={poPerPage} onChange={e=>setPoPerPage(Number(e.target.value))}>
              <option value={20}>20</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
              <option value={500}>500</option>
            </select>
          </div>
          <div className="flex gap-2 itemss-end">
            <button onClick={()=>setPoPage(1)}
              className="px-5 py-2 bg-purple-700 hover:bg-purple-800 text-white rounded-lg text-sm font-semibold shadow-sm">Apply</button>
            <button onClick={()=>{ setPoSearchNums([]); setPoFilterItemType([]); setPoFilterOpUnit([]); setPoPage(1); }}
              className={`px-4 py-2 rounded-lg text-sm font-medium shadow-sm ${darkMode?'bg-gray-500 text-gray-100 hover:bg-gray-400':'bg-gray-400 text-white hover:bg-gray-500'}`}>Reset</button>
          </div>
        </div>

        {/* Active PO filter tags */}
        {(poSearchNums.length > 0 || poFilterItemType.length > 0 || poFilterOpUnit.length > 0) && (
          <div className={`px-5 py-2 flex flex-wrap gap-1.5 border-b ${darkMode?'border-gray-700':'border-gray-100'}`}>
            {poSearchNums.map(v=>(
              <span key={v} className="flex itemss-center gap-1 px-2 py-0.5 bg-indigo-100 text-indigo-700 rounded-full text-xs">
                PO: {v}<button onClick={()=>setPoSearchNums(prev=>prev.filter(x=>x!==v))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
              </span>
            ))}
            {poFilterItemType.map(v=>(
              <span key={v} className="flex itemss-center gap-1 px-2 py-0.5 bg-blue-100 text-blue-700 rounded-full text-xs">
                {v}<button onClick={()=>setPoFilterItemType(prev=>prev.filter(x=>x!==v))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
              </span>
            ))}
            {poFilterOpUnit.map(v=>(
              <span key={v} className="flex itemss-center gap-1 px-2 py-0.5 bg-purple-100 text-purple-700 rounded-full text-xs">
                {v}<button onClick={()=>setPoFilterOpUnit(prev=>prev.filter(x=>x!==v))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
              </span>
            ))}
          </div>
        )}

        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className={tblHd}>
              <tr>
                {['PO HLI NUMBER','PO ITEM TYPE','ITEM CODE','OPERATION UNIT','DESCRIPTION','QTY','UNIT','PRICE','AMOUNT','CURRENCY','PO DATE','PURCHASE MEMBER','REQ. DELIVERY','BDAYS REMAINING'].map(h=>(
                  <th key={h} className={`px-4 py-3 text-left font-semibold whitespace-nowrap ${txt2} ${h==='PRICE'||h==='AMOUNT'?'min-w-[140px]':''}`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {poRows.length === 0 ? (
                <tr><td colSpan={16} className={`px-4 py-10 text-center ${txt2}`}>
                  <Package className="w-10 h-10 mx-auto mb-2 opacity-40"/>No data available
                </td></tr>
              ) : poRows.map((row,i)=>{
                const daysLeft = row.days_remaining;
                const daysColor = daysLeft === null ? txt2 : daysLeft < 0 ? 'text-red-600 font-bold' : daysLeft <= 7 ? 'text-orange-600 font-semibold' : daysLeft <= 30 ? 'text-yellow-600' : 'text-green-600';
                return (
                  <tr key={i} className={`${trHov} transition-colors`}>
                    <td className="px-4 py-3 text-purple-600 font-medium whitespace-nowrap">
                      {row.items_no ? `${row.po_no}-${row.items_no}` : row.po_no}
                    </td>
                    <td className={`px-4 py-3 whitespace-nowrap`}>
                      {row.po_items_type ? (
                        <span className={`px-2 py-0.5 rounded-full text-xs font-medium ${
                          row.po_items_type.toUpperCase()==='MRO' ? 'bg-blue-100 text-blue-700' :
                          row.po_items_type.toUpperCase()==='EQUIPMENT' ? 'bg-green-100 text-green-700' :
                          'bg-gray-100 text-gray-700'}`}>{row.po_items_type}</span>
                      ) : <span className={`${txt2} text-xs`}>-</span>}
                    </td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.items_code||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap text-xs`} title={row.operation_unit}>{row.operation_unit||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} max-w-xs truncate`} title={row.description}>{row.description}</td>
                    <td className={`px-4 py-3 text-right ${txt2}`}>{fmtNum(row.qty)}</td>
                    <td className={`px-4 py-3 ${txt2}`}>{row.unit||'-'}</td>
                    <td className="px-4 py-3 text-right whitespace-nowrap min-w-[140px]">{fmtCur(row.price)}</td>
                    <td className="px-4 py-3 text-right font-semibold text-orange-600 whitespace-nowrap min-w-[140px]">{fmtCur(row.amount)}</td>
                    <td className={`px-4 py-3 ${txt2}`}>{row.currency||'IDR'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.po_date||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.purchase_member||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.req_delivery||'-'}</td>
                    <td className={`px-4 py-3 text-center whitespace-nowrap ${daysColor}`}>
                      {daysLeft === null ? '-' : daysLeft < 0 ? `${Math.abs(daysLeft)} biz days overdue` : `${daysLeft} biz days`}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className={`p-4 border-t ${darkMode?'border-gray-700':'border-gray-100'} flex justify-between itemss-center`}>
          <span className={`text-sm ${txt2}`}>Showing {(poPage-1)*poPerPage+1}–{Math.min(poPage*poPerPage,poFiltered.length)} of {fmtNum(poFiltered.length)}</span>
          <div className="flex gap-1 itemss-center">
            <button disabled={poPage===1} onClick={()=>setPoPage(p=>p-1)} className={`p-1.5 rounded ${poPage===1?'opacity-40':'hover:bg-purple-100'}`}><ChevronLeft className="w-4 h-4"/></button>
            <span className={`px-3 py-1 rounded text-sm font-semibold ${darkMode?'bg-gray-700 text-white':'bg-purple-100 text-purple-700'}`}>{poPage}/{poTotalPages}</span>
            <button disabled={poPage===poTotalPages} onClick={()=>setPoPage(p=>p+1)} className={`p-1.5 rounded ${poPage===poTotalPages?'opacity-40':'hover:bg-purple-100'}`}><ChevronRight className="w-4 h-4"/></button>
          </div>
        </div>
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════════════════
  // MAIN RENDER
  // ══════════════════════════════════════════════════════════════
  return (
    <div className={`min-h-screen font-sans ${darkMode?'bg-gray-900 dark-root':'bg-gray-50'}`}>
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
        /* Dark mode guard: ensure no black text */
        .dark-root input, .dark-root select, .dark-root textarea,
        .dark-root option { color: #e5e7eb !important; background-color: #374151 !important; }
      `}</style>

      <div className="fixed top-5 right-5 z-[100] flex flex-col gap-2">
        {toasts.map(t=><Toast key={t.id} message={t.message} type={t.type} onClose={()=>removeToast(t.id)}/>)}
      </div>

      {/* Download progress toast */}
      {downloadToast && <DownloadToast message={downloadToast.message} />}

      {/* Sidebar */}
      <aside className={`fixed left-0 top-0 h-full w-20 flex flex-col itemss-center py-8 shadow-2xl z-40 ${darkMode?'bg-gray-800 border-r border-gray-700':'bg-gradient-to-b from-purple-600 to-purple-700'}`}>
        <nav className="flex-1 flex flex-col gap-4 w-full px-2 pt-0">
          <button onClick={()=>setActivePage('dashboard')}
            className={`p-3 rounded-xl flex justify-center transition-all ${activePage==='dashboard'?'bg-white/30 text-white shadow-lg':'text-purple-100 hover:bg-white/20'}`} title="Dashboard">
            <BarChart3 className="w-6 h-6"/>
          </button>
          <button onClick={()=>{ setActivePage('all-so'); setSoPage(1); fetchSOData(soFilters,1,soPerPage,soSearchNums,soMarginFilter); window.scrollTo({top:0, behavior:'smooth'}); }}
            className={`p-3 rounded-xl flex justify-center transition-all ${activePage==='all-so'?'bg-white/30 text-white shadow-lg':'text-purple-100 hover:bg-white/20'}`} title="Open SO (Sales Order)">
            <FileText className="w-6 h-6"/>
          </button>
        </nav>
        <button onClick={()=>setDarkMode(d=>!d)} className="p-3 rounded-xl text-white hover:bg-white/20 transition-all">
          {darkMode?<Sun className="w-6 h-6"/>:<Moon className="w-6 h-6"/>}
        </button>
      </aside>

      {/* Main */}
      <main className="ml-20 p-6">
        <header className="mb-6 flex flex-wrap justify-between itemss-center gap-4">
          <div>
            <h1 className={`text-2xl font-bold tracking-tight ${txt}`}>
              HLI PO Monitoring <span className="text-purple-600">Dashboard</span>
            </h1>
            <p className={`mt-0.5 text-sm ${txt2}`}>
              {activePage==='dashboard'?'Purchase Orders & Sales Orders Overview':'Manage Open SO (Sales Order) & PO Without SO'}
            </p>
          </div>
          <div className="flex flex-wrap gap-2">
            <label className={`flex itemss-center gap-2 px-4 py-2 rounded-xl shadow hover:shadow-md transition-all ${darkMode?'bg-purple-700 hover:bg-purple-800 text-white':'bg-purple-700 hover:bg-purple-800 text-white'}`}>
              <Upload className="w-4 h-4"/><span className="text-sm font-medium">Upload HLI PO List (Item)</span>
              <input type="file" accept=".xlsx,.xls" onChange={e=>handleUpload(e,'po')} className="hidden"/>
            </label>
            <label className={`flex itemss-center gap-2 px-4 py-2 rounded-xl shadow hover:shadow-md transition-all ${darkMode?'bg-blue-700 hover:bg-blue-800 text-white':'bg-blue-700 hover:bg-blue-800 text-white'}`}>
              <Upload className="w-4 h-4"/><span className="text-sm font-medium">Upload SMRO - Search Client Odr</span>
              <input type="file" accept=".xlsx,.xls" onChange={e=>handleUpload(e,'smro')} className="hidden"/>
            </label>
            <div className="relative" ref={hideMenuRef}>
              <button onClick={()=>setShowHideMenu(o=>!o)}
                className="flex itemss-center gap-2 px-4 py-2 rounded-xl shadow hover:shadow-md transition-all bg-orange-600 hover:bg-orange-700 text-white">
                <EyeOff className="w-4 h-4"/><span className="text-sm font-medium">Hide</span>
                <ChevronDown className="w-3.5 h-3.5"/>
              </button>
              {showHideMenu && (
                <div className={`absolute right-0 mt-2 z-50 rounded-xl shadow-2xl border w-72 p-3 ${darkMode?'bg-gray-800 border-gray-700 text-white':'bg-white border-gray-200'}`}>
                  <p className={`text-xs font-semibold mb-3 px-1 ${darkMode?'text-gray-300':'text-gray-600'}`}>
                    Hide items from dashboard using Excel template
                  </p>
                  {/* PO HLI */}
                  <div className={`mb-2 p-3 rounded-lg ${darkMode?'bg-gray-700':'bg-orange-50'}`}>
                    <p className="text-xs font-bold mb-1 text-orange-700">🔵 PO HLI (from PO List)</p>
                    <p className={`text-xs mb-2 ${darkMode?'text-gray-400':'text-gray-500'}`}>Format: PO Number-Item No (e.g. 4502358819-10)</p>
                    <div className="flex gap-2">
                      <button onClick={()=>downloadHideTemplate('PO')}
                        className="flex-1 flex itemss-center justify-center gap-1 px-2 py-1.5 bg-orange-600 hover:bg-orange-700 text-white rounded-lg text-xs font-semibold">
                        <Download className="w-3 h-3"/>Download Template
                      </button>
                      <label className="flex-1 flex itemss-center justify-center gap-1 px-2 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-xs font-semibold cursor-pointer">
                        <Upload className="w-3 h-3"/>Upload Filled
                        <input type="file" accept=".xlsx,.xls" onChange={e=>handleHideBatchUpload(e,'PO')} className="hidden"/>
                      </label>
                    </div>
                  </div>
                  {/* SO */}
                  <div className={`p-3 rounded-lg ${darkMode?'bg-gray-700':'bg-blue-50'}`}>
                    <p className="text-xs font-bold mb-1 text-blue-700">🟠 SO (from SMRO)</p>
                    <p className={`text-xs mb-2 ${darkMode?'text-gray-400':'text-gray-500'}`}>Format: SO Number or SO Number-Item No</p>
                    <div className="flex gap-2">
                      <button onClick={()=>downloadHideTemplate('SO')}
                        className="flex-1 flex itemss-center justify-center gap-1 px-2 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold">
                        <Download className="w-3 h-3"/>Download Template
                      </button>
                      <label className="flex-1 flex itemss-center justify-center gap-1 px-2 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-xs font-semibold cursor-pointer">
                        <Upload className="w-3 h-3"/>Upload Filled
                        <input type="file" accept=".xlsx,.xls" onChange={e=>handleHideBatchUpload(e,'SO')} className="hidden"/>
                      </label>
                    </div>
                  </div>
                </div>
              )}
            </div>

            <button onClick={()=>{ fetchDeleteRequests(); setShowHiddenPanel(true); }}
              className={`flex itemss-center gap-2 px-4 py-2 rounded-xl shadow hover:shadow-md transition-all ${darkMode?'bg-gray-600 hover:bg-gray-500 text-white':'bg-gray-200 hover:bg-gray-300 text-gray-700'}`}>
              <Eye className="w-4 h-4"/>
              <span className="text-sm font-medium">Hide History</span>
              {deleteRequests.filter(r=>r.is_hidden).length > 0 && (
                <span className="px-1.5 py-0.5 bg-orange-500 text-white rounded-full text-xs font-bold">
                  {deleteRequests.filter(r=>r.is_hidden).length}
                </span>
              )}
            </button>
          </div>
        </header>

        {activePage==='dashboard' ? renderDashboard() : renderAllSO()}
      </main>

      {modal && <SOModal title={modal.title} data={modal.data} darkMode={darkMode} onClose={()=>setModal(null)}/>}

      {showHiddenPanel && (
        <HiddenItemsPanel
          darkMode={darkMode}
          requests={deleteRequests}
          onRestore={restoreDeleteRequest}
          onClose={()=>setShowHiddenPanel(false)}
        />
      )}

      {uploadProgress && (
        <div className="fixed inset-0 bg-black/60 z-[60] flex itemss-center justify-center backdrop-blur-sm">
          <div className={`${darkMode?'bg-gray-800':'bg-white'} p-8 rounded-2xl shadow-2xl flex flex-col itemss-center gap-4 w-80`}>
            <div className="w-14 h-14 border-4 border-purple-600 border-t-transparent rounded-full animate-spin"/>
            <div className="w-full text-center">
              <p className={`font-bold text-lg mb-1 ${txt}`}>Uploading {uploadProgress.label}...</p>
              <p className={`text-xs mb-3 ${txt2}`}>Please wait, do not close the browser</p>
              <div className={`w-full rounded-full h-3 ${darkMode?'bg-gray-700':'bg-gray-200'}`}>
                <div className="bg-gradient-to-r from-purple-600 to-purple-400 h-3 rounded-full transition-all duration-300" style={{width:`${uploadProgress.pct}%`}}/>
              </div>
              <p className="text-purple-600 font-semibold mt-2">{uploadProgress.pct}%</p>
            </div>
          </div>
        </div>
      )}

      {loading && !uploadProgress && (
        <div className="fixed inset-0 bg-black/30 z-[55] flex itemss-center justify-center">
          <div className={`${darkMode?'bg-gray-800':'bg-white'} px-6 py-4 rounded-xl shadow-xl flex itemss-center gap-3`}>
            <div className="w-6 h-6 border-3 border-purple-600 border-t-transparent rounded-full animate-spin"/>
            <p className={`text-sm font-semibold ${txt}`}>Loading data...</p>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
