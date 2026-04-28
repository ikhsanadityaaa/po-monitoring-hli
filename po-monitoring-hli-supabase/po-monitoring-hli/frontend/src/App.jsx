import React, { useState, useEffect, useCallback, useRef } from 'react';
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, AreaChart, Area, ComposedChart
} from 'recharts';
import {
  Upload, Download, AlertCircle, CheckCircle, XCircle,
  Package, TrendingUp, TrendingDown, Award, Calendar, ChevronLeft,
  ChevronRight, Moon, Sun, FileText, BarChart3, FileSpreadsheet,
  Filter, X, ChevronDown, ChevronUp, Building2, Search, Loader2,
  EyeOff, Eye, Trash2, RotateCcw, Plus, Coins, Wallet
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
    <div className="fixed top-5 right-5 z-[200] flex items-center gap-3 px-5 py-3 rounded-xl shadow-2xl text-white bg-purple-700 max-w-sm animate-slide-in">
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

const SOModal = ({ title, data, onClose, darkMode }) => {
  const [dlPage, setDlPage] = useState(1);
  const PER = 50;
  const pages = Math.ceil((data?.length || 0) / PER);
  const rows = (data || []).slice((dlPage-1)*PER, dlPage*PER);

  // Determine if SO Item column exists in data (show SO Number only when SO Item is absent)
  const hasSoItem = (data || []).some(s => s.so_item);

  const downloadExcel = () => {
    const ws = XLSX.utils.json_to_sheet(data.map(s => ({
      'SO Item': s.so_item,
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
    <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm" onClick={onClose}>
      <div role="dialog" aria-modal="true" aria-label={title} className={`rounded-2xl shadow-2xl w-full max-w-6xl max-h-[85vh] flex flex-col ${darkMode?'bg-gray-800 text-white':'bg-white'}`} onClick={e=>e.stopPropagation()}>
        <div className={`flex justify-between items-center px-6 py-4 border-b ${darkMode?'border-gray-700':'border-gray-100'}`}>
          <h3 className="font-bold text-lg">{title} <span className={`text-sm font-normal ml-2 ${darkMode?'text-gray-400':'text-gray-500'}`}>({fmtNum(data?.length)} records)</span></h3>
          <div className="flex gap-2">
            <button onClick={downloadExcel} className="flex items-center gap-1 px-3 py-1.5 bg-green-600 hover:bg-green-700 text-white rounded-lg text-sm"><FileSpreadsheet className="w-4 h-4"/>Excel</button>
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
                  <td className="px-3 py-2 text-purple-600 font-medium whitespace-nowrap">{s.so_item||'-'}</td>
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
          <div className={`flex justify-between items-center px-6 py-3 border-t ${darkMode?'border-gray-700':'border-gray-100'}`}>
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
      // We use a sentinel: store all items as "selected" but display as "0 selected"
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
      // Click one item: keep only that one checked (deselect all others)
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
        className={`w-full px-3 py-2 rounded-lg text-sm border text-left flex justify-between items-center transition-colors
          ${darkMode
            ? 'bg-gray-600 border-gray-500 text-white hover:bg-gray-500'
            : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'}`}>
        <span className="truncate">{displayLabel}</span>
        <ChevronDown className="w-4 h-4 flex-shrink-0 ml-1"/>
      </button>
      {open && (
        <div className={`absolute z-50 mt-1 w-full max-h-56 overflow-auto rounded-lg shadow-xl border ${darkMode?'bg-gray-700 border-gray-600':'bg-white border-gray-200'}`}>
          {/* Select All row — like Excel */}
          <label style={{cursor:'pointer'}} className={`flex items-center gap-2 px-3 py-2 text-xs font-semibold border-b
            ${darkMode?'border-gray-600 hover:bg-gray-600 text-white':'border-gray-100 hover:bg-purple-50 text-gray-700'}`}>
            <input type="checkbox"
              checked={isAllChecked}
              ref={el => { if (el) el.indeterminate = someSelected; }}
              onChange={toggleAll}
              className="accent-purple-600" style={{cursor:'pointer'}}/>
            <span>(Select All)</span>
          </label>
          {options.map(opt => (
            <label key={opt} style={{cursor:'pointer'}} className={`flex items-center gap-2 px-3 py-2 text-xs
              ${darkMode?'hover:bg-gray-600 text-white':'hover:bg-purple-50 text-gray-700'}`}>
              <input type="checkbox" checked={isChecked(opt)} onChange={()=>toggle(opt)}
                className="accent-purple-600" style={{cursor:'pointer'}}/>
              <span className="truncate" title={opt}>{opt}</span>
            </label>
          ))}
          {options.length === 0 && <div className={`px-3 py-2 text-xs ${txt2}`}>No options</div>}
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
        className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-sm border font-medium transition-all
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
                <label key={t} className="flex items-center gap-2 cursor-pointer">
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
          <button onClick={onSubmit} className="px-5 py-2 bg-orange-600 hover:bg-orange-700 text-white rounded-lg text-sm font-semibold flex items-center gap-2">
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
            <Eye className="w-5 h-5 text-purple-500"/>
            <h3 className="font-bold text-base">Items Hidden from Dashboard</h3>
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
                      <span className={`px-2 py-0.5 rounded text-xs font-bold ${r.ref_type==='PO'?'bg-red-100 text-red-700':'bg-orange-100 text-orange-700'}`}>{r.ref_type}</span>
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
const DateRangeFilter = ({ darkMode, txt, txt2, card, onFilter, value, label = 'Filter SO Create Date' }) => {
  const currentYear = new Date().getFullYear();
  const years = Array.from({ length: 6 }, (_, i) => currentYear - i);
  const [mode, setMode] = useState(value?.mode || 'all'); // 'all' | 'year' | 'range'
  const [selectedYear, setSelectedYear] = useState(String(value?.year || currentYear));
  const [startDate, setStartDate] = useState(value?.start || '');
  const [endDate, setEndDate] = useState(value?.end || '');

  // Keep internal state in sync when the controlled `value` changes externally
  // (e.g. user changes filter on another page that shares the same global state).
  useEffect(() => {
    if (!value) return;
    setMode(value.mode || 'all');
    if (value.year)  setSelectedYear(String(value.year));
    if (value.start !== undefined) setStartDate(value.start || '');
    if (value.end   !== undefined) setEndDate(value.end || '');
  }, [value?.mode, value?.year, value?.start, value?.end]);

  const apply = () => {
    if (mode === 'all') onFilter({ mode: 'all' });
    else if (mode === 'year') onFilter({ mode: 'year', year: selectedYear });
    else onFilter({ mode: 'range', start: startDate, end: endDate });
  };

  const reset = () => {
    setMode('all');
    setSelectedYear(String(currentYear));
    setStartDate(''); setEndDate('');
    onFilter({ mode: 'all' });
  };

  return (
    <div className={`flex flex-wrap items-center gap-3 px-5 py-3 rounded-xl ${card} shadow mb-4`}>
      <Calendar className="w-4 h-4 text-purple-500 flex-shrink-0"/>
      <span className={`text-sm font-semibold ${txt} flex-shrink-0`}>{label}:</span>
      {/* Mode selector */}
      <div className="flex gap-1">
        {[['all','All'], ['year','Per Year'], ['range','Date Range']].map(([m, lbl]) => (
          <button key={m} onClick={() => setMode(m)}
            className={`px-3 py-1 rounded-full text-xs font-semibold transition-all
              ${mode === m ? 'bg-purple-600 text-white shadow' : darkMode ? 'bg-gray-700 text-gray-300 hover:bg-gray-600' : 'bg-gray-100 text-gray-600 hover:bg-purple-100'}`}>
            {lbl}
          </button>
        ))}
      </div>
      {mode === 'year' && (
        <select value={selectedYear} onChange={e => setSelectedYear(e.target.value)}
          className={`px-3 py-1.5 rounded-lg text-sm border ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-white border-gray-300'}`}>
          {years.map(y => <option key={y} value={y}>{y}</option>)}
        </select>
      )}
      {mode === 'range' && (
        <div className="flex items-center gap-2">
          <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)}
            className={`px-3 py-1.5 rounded-lg text-sm border ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-white border-gray-300'}`}/>
          <span className={`text-xs ${txt2}`}>to</span>
          <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)}
            className={`px-3 py-1.5 rounded-lg text-sm border ${darkMode ? 'bg-gray-700 border-gray-600 text-white' : 'bg-white border-gray-300'}`}/>
        </div>
      )}
      <button onClick={apply}
        className="px-4 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-xs font-semibold">
        Apply
      </button>
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
  const [completedData, setCompletedData] = useState(null);
  const [completedYear, setCompletedYear] = useState('all');
  const [completedLoading, setCompletedLoading] = useState(false);
  const [showHideMenu, setShowHideMenu] = useState(false);
  const [marginDetailModal, setMarginDetailModal] = useState(null); // {category, data}
  const hideMenuRef = useRef(null);

  // ── Global SO Create Date filter (shared across Dashboard / All SO / Delivery Completed)
  const [globalDateFilter, setGlobalDateFilter] = useState({ mode: 'all' });
  // Aliases kept so existing references continue to compile.
  const dashDateFilter      = globalDateFilter;
  const setDashDateFilter   = setGlobalDateFilter;
  const soDateFilter        = globalDateFilter;
  const setSODateFilter     = setGlobalDateFilter;
  const completedDateFilter = globalDateFilter;
  const setCompletedDateFilter = setGlobalDateFilter;
  useEffect(() => {
    const handler = (e) => { if (hideMenuRef.current && !hideMenuRef.current.contains(e.target)) setShowHideMenu(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const addToast = useCallback((message, type='success') => {
    const id = Date.now(); setToasts(t => [...t, { id, message, type }]);
  }, []);
  const removeToast = useCallback((id) => setToasts(t => t.filter(x => x.id !== id)), []);

  const fetchDashboard = useCallback(async (dateFilter) => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      const f = dateFilter || globalDateFilter;
      if (f && f.mode !== 'all') {
        if (f.mode === 'year') params.append('date_year', f.year);
        else if (f.mode === 'range') {
          if (f.start) params.append('date_from', f.start);
          if (f.end)   params.append('date_to',   f.end);
        }
      }
      const qs = params.toString() ? `?${params}` : '';
      const [sRes, pRes, aRes] = await Promise.all([
        api.get(`/api/dashboard/stats${qs}`),
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
      const itemTypes = [...new Set(poFiltered.map(p=>p.po_item_type).filter(Boolean))].sort();
      const opUnits   = [...new Set(poFiltered.map(p=>p.operation_unit).filter(Boolean))].sort();
      setPoItemTypeOptions(itemTypes);
      setPoOpUnitOptions(opUnits);
      setAgingData(Array.isArray(aRes.data) ? aRes.data : []);
    } catch (e) {
      addToast(`Error: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast, globalDateFilter]);

  // Helper: filter array of objects by date field using a DateRangeFilter config
  const applyDateFilter = useCallback((arr, dateField, filter) => {
    if (!filter || filter.mode === 'all') return arr;
    return arr.filter(item => {
      const d = item[dateField];
      if (!d) return false;
      const iso = d.slice(0, 10);
      if (filter.mode === 'year') return iso.startsWith(filter.year);
      if (filter.mode === 'range') {
        if (filter.start && iso < filter.start) return false;
        if (filter.end && iso > filter.end) return false;
        return true;
      }
      return true;
    });
  }, []);

  // Helper: build date query params for backend
  const dateFilterParams = (filter) => {
    if (!filter || filter.mode === 'all') return {};
    if (filter.mode === 'year') return { date_year: filter.year };
    if (filter.mode === 'range') return { date_from: filter.start || '', date_to: filter.end || '' };
    return {};
  };

  // Helper: resolve filter array
  const resolveFilter = (val) => {
    if (val === '__NONE__') return ['__NONE_PLACEHOLDER__']; // backend will return 0 rows
    if (!Array.isArray(val) || val.length === 0) return []; // empty = no filter = all
    return val;
  };

  const fetchSOData = useCallback(async (filters, page, perPage, searchNums, marginFilter, dateFilter) => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      resolveFilter(filters.op_units).forEach(v => params.append('op_unit', v));
      resolveFilter(filters.vendors).forEach(v => params.append('vendor', v));
      resolveFilter(filters.statuses).forEach(v => params.append('status', v));
      (filters.aging || []).forEach(a => params.append('aging', a));
      (searchNums || []).forEach(n => params.append('so_item', n));
      if (marginFilter && marginFilter !== 'all') params.append('margin_filter', marginFilter);
      // Date filter
      if (dateFilter && dateFilter.mode !== 'all') {
        if (dateFilter.mode === 'year') params.append('date_year', dateFilter.year);
        else if (dateFilter.mode === 'range') {
          if (dateFilter.start) params.append('date_from', dateFilter.start);
          if (dateFilter.end) params.append('date_to', dateFilter.end);
        }
      }
      const res = await api.get(`/api/data/all-so?${params}`);
      setAllSOData(Array.isArray(res.data.data) ? res.data.data : []);
      setSoTotal(res.data.total || 0);
      setSoFilterOptions(res.data.filters || { op_units: [], vendors: [], statuses: [] });
    } catch (e) {
      addToast(`Failed to load SO: ${e.message}`, 'error');
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

  // Apply PO local filters whenever dependencies change

  // Apply PO local filters whenever dependencies change
  useEffect(() => {
    let filtered = [...poWithoutSO];
    if (poSearchNums.length > 0) {
      const nums = poSearchNums.map(n=>n.toLowerCase());
      filtered = filtered.filter(p => {
        const poHliKey = p.item_no ? `${p.po_no}-${p.item_no}`.toLowerCase() : (p.po_no||'').toLowerCase();
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
      filtered = filtered.filter(p => poFilterItemType.includes(p.po_item_type));
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
  useEffect(() => { if (activePage === 'all-so') fetchSOData(soFilters, soPage, soPerPage, soSearchNums, soMarginFilter, soDateFilter); }, [activePage]);

  // Refetch dashboard whenever the global SO Create Date filter changes
  // (skip the very first run since the mount effect above already fetched).
  const skipFirstFilterRefetch = useRef(true);
  useEffect(() => {
    if (skipFirstFilterRefetch.current) { skipFirstFilterRefetch.current = false; return; }
    fetchDashboard(globalDateFilter);
  }, [globalDateFilter, fetchDashboard]);

  const handleUpload = async (e, type) => {
    const file = e.target.files[0]; if (!file) return;
    e.target.value = '';
    const label = type === 'po' ? 'HLI PO List (Item)' : 'SMRO - Search Client Odr';
    const endpoint = type === 'po' ? '/api/upload/po-list' : '/api/upload/smro';

    // ── Client-side header validation ──────────────────────────────────
    const REQUIRED_HEADERS = {
      po: {
        'PO Number':        ['po no.','po no','po number','po'],
        'Item No':          ['item no.','item no','item number','no. item'],
        'PO Item Type':     ['po item type','item type','type','po type'],
        'Supplier':         ['supplier','vendor','supplier name'],
        'Qty':              ['qty.','qty','quantity'],
        'Amount':           ['amount','total amount','total'],
        'PO Date':          ['po date','order date','tanggal po'],
        'Request Delivery': ['request delivery date','delivery date','req delivery'],
      },
      smro: {
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
          `❌ Invalid file — ${missing.length} required columns not found: ${missing.join(', ')}. Please check the ${label} file is correct and try again.`,
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
      if (activePage === 'all-so') fetchSOData(soFilters, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
      setSoPage(1);
    } catch (e) {
      setUploadProgress(null);
      addToast(`❌ Failed to upload ${label}: ${e.response?.data?.error || e.message}`, 'error');
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

  const fetchCompletedData = useCallback(async (year='all', dateFilter=null) => {
    setCompletedLoading(true);
    try {
      const params = new URLSearchParams({ year });
      if (dateFilter && dateFilter.mode !== 'all') {
        if (dateFilter.mode === 'year') params.append('date_year', dateFilter.year);
        else if (dateFilter.mode === 'range') {
          if (dateFilter.start) params.append('date_from', dateFilter.start);
          if (dateFilter.end) params.append('date_to', dateFilter.end);
        }
      }
      const res = await api.get(`/api/completed/summary?${params}`);
      setCompletedData(res.data);
    } catch(e) { addToast(`❌ Failed to load completed data: ${e.message}`, 'error'); }
    finally { setCompletedLoading(false); }
  }, []);

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
      addToast(`❌ Failed to upload hide batch: ${e.response?.data?.error || e.message}`, 'error');
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
    } catch (e) { addToast(`❌ Failed to update: ${e.message}`, 'error'); }
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
  const txt   = darkMode ? 'text-white' : 'text-gray-900';
  const txt2  = darkMode ? 'text-gray-400' : 'text-gray-600';
  const tblHd = darkMode ? 'bg-gray-700' : 'bg-purple-50';
  const tblDv = darkMode ? 'divide-gray-700' : 'divide-gray-100';
  const trHov = darkMode ? 'hover:bg-gray-700' : 'hover:bg-purple-50';

  // ══════════════════════════════════════════════════════════════
  // RENDER COMPLETED TRANSACTIONS PAGE
  // ══════════════════════════════════════════════════════════════
  const renderCompleted = () => {
    const d = completedData;
    const CPIE = ['#10B981','#EF4444','#9CA3AF'];
    const fmtM = (v) => v >= 1e9 ? `${(v/1e9).toFixed(1)}B` : v >= 1e6 ? `${(v/1e6).toFixed(1)}M` : v >= 1e3 ? `${(v/1e3).toFixed(0)}K` : String(Math.round(v));
    const mc = (m) => m > 0 ? 'text-green-600' : m < 0 ? 'text-red-600' : 'text-gray-400';
    const mcBg = (m) => m < 0 ? (darkMode?'bg-red-900/20':'bg-red-50') : (darkMode?'bg-gray-700':'bg-gray-50');

    if (completedLoading) return (
      <div className="flex items-center justify-center h-64">
        <div className="flex flex-col items-center gap-3">
          <div className="w-12 h-12 border-4 border-purple-600 border-t-transparent rounded-full animate-spin"/>
          <p className={`text-sm ${txt2}`}>Loading completed transactions...</p>
        </div>
      </div>
    );

    if (!d) return (
      <div className={`flex flex-col items-center justify-center h-64 rounded-2xl ${card}`}>
        <Coins className="w-16 h-16 text-gray-300 mb-4"/>
        <p className={`text-lg font-semibold ${txt}`}>No completed data yet</p>
        <p className={`text-sm ${txt2} mt-1`}>Upload SMRO data to see completed transactions</p>
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
            if (f.mode === 'year') { setCompletedYear(f.year); fetchCompletedData(f.year, f); }
            else { setCompletedYear('all'); fetchCompletedData('all', f); }
          }}
        />

        {/* ── KPI Cards ───────────────────────────────────────── */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
          {[
            { label:'Completed Transactions', value: fmtNum(d.total_count),
              icon:<CheckCircle className="w-6 h-6 text-green-500"/>, bg:'bg-green-100', color:'text-green-600' },
            { label:'Total Sales Amount', value: fmtCurShort(d.total_sales), sub: fmtCur(d.total_sales),
              icon:<Wallet className="w-6 h-6 text-blue-500"/>, bg:'bg-blue-100', color:'text-blue-600' },
            { label:'Total Purchase Amount', value: fmtCurShort(d.total_purchase), sub: fmtCur(d.total_purchase),
              icon:<Coins className="w-6 h-6 text-purple-500"/>, bg:'bg-purple-100', color:'text-purple-600' },
            { label:'Total Margin', value: fmtCurShort(d.total_margin), sub: fmtCur(d.total_margin),
              icon: d.total_margin>=0?<TrendingUp className="w-6 h-6 text-emerald-500"/>:<TrendingDown className="w-6 h-6 text-red-500"/>,
              bg: d.total_margin>=0?'bg-emerald-100':'bg-red-100', color: d.total_margin>=0?'text-emerald-600':'text-red-600' },
          ].map((k,i)=>(
            <div key={i} className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all ${card}`}>
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
                  <stop offset="5%" stopColor="#8B5CF6" stopOpacity={0.9}/><stop offset="95%" stopColor="#8B5CF6" stopOpacity={0.5}/>
                </linearGradient>
                <linearGradient id="cgPurchase" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#3B82F6" stopOpacity={0.9}/><stop offset="95%" stopColor="#3B82F6" stopOpacity={0.5}/>
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke={darkMode?'#374151':'#F3F4F6'}/>
              <XAxis dataKey="monthLabel" stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={10}/>
              <YAxis yAxisId="amt" stroke={darkMode?'#9CA3AF':'#6B7280'} fontSize={10} tickFormatter={fmtM}/>
              <YAxis yAxisId="cnt" orientation="right" stroke="#F97316" fontSize={10}/>
              <Tooltip
                formatter={(v, n) => n === 'Transactions' ? [fmtNum(v), n] : [fmtCur(v), n]}
                contentStyle={{background:darkMode?'#1F2937':'#fff',border:'none',borderRadius:8,fontSize:12}}/>
              <Legend wrapperStyle={{fontSize:12}}/>
              <Bar yAxisId="amt" dataKey="sales_amount" name="Sales Amount" fill="url(#cgSales)" radius={[4,4,0,0]}/>
              <Bar yAxisId="amt" dataKey="purchase_amount" name="Purchase Amount" fill="url(#cgPurchase)" radius={[4,4,0,0]}/>
              <Line yAxisId="cnt" type="monotone" dataKey="count" name="Transactions" stroke="#F97316" strokeWidth={3} dot={{r:3,fill:'#F97316'}} activeDot={{r:5}} z={10}/>
            </ComposedChart>
          </ResponsiveContainer>
        </div>

        {/* ── Top 5 Vendors  +  Margin Pie ────────────────────── */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">

          {/* Top 5 Vendors */}
          <div className={`p-5 rounded-2xl shadow ${card}`}>
            <h3 className={`text-base font-bold mb-4 ${txt} flex items-center gap-2`}>
              <Award className="w-5 h-5 text-purple-500"/> Top 5 Vendors — Completed Transactions
            </h3>
            <div className="space-y-4">
              {(d.top_vendors||[]).map((v,i)=>{
                const maxAmt = d.top_vendors[0]?.sales_amount || 1;
                const pct = Math.round(v.sales_amount / maxAmt * 100);
                const rankColors = ['bg-yellow-400','bg-gray-300','bg-orange-400','bg-purple-200','bg-purple-100'];
                return (
                  <div key={i}>
                    <div className="flex items-center justify-between mb-1">
                      <div className="flex items-center gap-2 min-w-0">
                        <span className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold text-gray-700 flex-shrink-0 ${rankColors[i]||'bg-gray-100'}`}>{i+1}</span>
                        <span className={`text-sm font-semibold truncate ${txt}`} title={v.vendor}>{v.vendor}</span>
                      </div>
                      <span className="text-xs font-bold text-purple-600 ml-2 flex-shrink-0">{fmtCurShort(v.sales_amount)}</span>
                    </div>
                    <div className={`w-full h-2 rounded-full ${darkMode?'bg-gray-700':'bg-gray-100'}`}>
                      <div className="h-2 rounded-full bg-gradient-to-r from-purple-600 to-purple-400" style={{width:`${pct}%`}}/>
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
                      if (completedDateFilter && completedDateFilter.mode !== 'all') {
                        if (completedDateFilter.mode === 'year') params.append('date_year', completedDateFilter.year);
                        else if (completedDateFilter.mode === 'range') {
                          if (completedDateFilter.start) params.append('date_from', completedDateFilter.start);
                          if (completedDateFilter.end) params.append('date_to', completedDateFilter.end);
                        }
                      }
                      const res = await api.get(`/api/completed/margin-detail?${params}`);
                      setMarginDetailModal({ category: label, data: Array.isArray(res.data) ? res.data : [] });
                    } catch(e) { addToast(`Failed to load detail: ${e.message}`, 'error'); }
                  }}
                    className={`text-left p-3 rounded-xl border cursor-pointer transition-all hover:shadow-md ${bg}`}>
                    <p className={`text-xs font-bold ${color} mb-1`}>{label}</p>
                    <p className={`text-lg font-bold ${color}`}>{fmtNum(count)} <span className="text-xs font-normal">PO</span></p>
                    <p className={`text-xs ${txt2} mt-0.5`}>{totalCompleted ? Math.round(count/totalCompleted*100) : 0}% of total</p>
                    <p className={`text-xs text-purple-500 font-semibold mt-1`}>Click for details →</p>
                  </button>
                ))}
              </div>
              {/* Pie chart */}
              <div className="w-full">
                <ResponsiveContainer width="100%" height={220}>
                  <PieChart>
                    <Pie data={marginPieData} cx="50%" cy="50%" innerRadius={55} outerRadius={88}
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

        {/* ── Top 20 Items ────────────────────────────────────── */}
        <div className={`p-5 rounded-2xl shadow ${card}`}>
          <h3 className={`text-base font-bold mb-4 ${txt} flex items-center gap-2`}>
            <Package className="w-5 h-5 text-orange-500"/> Top 20 Items by Sales Amount — Completed
          </h3>
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className={tblHd}>
                  {['#','Item / Product','Specification','Product ID','Order Freq','Sales Amount','Purchase Amount','Margin','Margin %'].map(h=>(
                    <th key={h} className={`px-3 py-2 text-left font-semibold ${darkMode?'text-purple-300':'text-purple-700'}`}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody className={`divide-y ${tblDv}`}>
                {(d.top_items||[]).map((item,i)=>{
                  const mPct = item.sales_amount ? (item.margin/item.sales_amount*100).toFixed(1) : '—';
                  return (
                    <tr key={i} className={trHov}>
                      <td className={`px-3 py-2 font-bold ${txt2}`}>{i+1}</td>
                      <td className={`px-3 py-2 font-medium ${txt} max-w-xs truncate`} title={item.item}>{item.item}</td>
                      <td className={`px-3 py-2 ${txt2} max-w-xs truncate`} title={item.specification||''}>{item.specification||'—'}</td>
                      <td className={`px-3 py-2 ${txt2} font-mono whitespace-nowrap`}>{item.product_id||'—'}</td>
                      <td className={`px-3 py-2 ${txt2}`}>{fmtNum(item.count)}</td>
                      <td className="px-3 py-2 text-purple-600 font-semibold">{fmtCurShort(item.sales_amount)}</td>
                      <td className="px-3 py-2 text-orange-600">{fmtCurShort(item.purchase_amount)}</td>
                      <td className={`px-3 py-2 font-bold ${mc(item.margin)}`}>{fmtCurShort(item.margin)}</td>
                      <td className={`px-3 py-2 font-semibold ${mc(item.margin)}`}>{mPct !== '—' ? `${mPct}%` : '—'}</td>
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
                      <div key={i} className={`p-3 rounded-xl ${mcBg(v.margin)}`}>
                        <div className="flex items-center justify-between mb-1">
                          <div className="flex items-center gap-2 min-w-0">
                            <span className="text-xs font-bold text-red-600 flex-shrink-0">#{i+1}</span>
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
                      </div>
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
                        {['#','Product ID','Product','Vendor','Sales','Purchase','Margin','%','Txns','Last Date'].map(h=>(
                          <th key={h} className={`px-2 py-2 text-left font-semibold ${darkMode?'text-purple-300':'text-purple-700'}`}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className={`divide-y ${tblDv}`}>
                      {d.worst_margin_transactions.map((t,i)=>(
                        <tr key={i} className={`${trHov} ${i===0?darkMode?'bg-red-900/20':'bg-red-50':''}`}>
                          <td className={`px-2 py-2 font-bold text-red-600`}>{i+1}</td>
                          <td className="px-2 py-2">
                            <p className="font-semibold text-purple-600 whitespace-nowrap">{t.item_code||'-'}</p>
                          </td>
                          <td className="px-2 py-2">
                            <p className={`truncate max-w-[120px] ${txt}`} title={t.product}>{t.product}</p>
                          </td>
                          <td className={`px-2 py-2 ${txt} truncate max-w-[90px]`} title={t.vendor}>{t.vendor}</td>
                          <td className="px-2 py-2 text-blue-600 whitespace-nowrap">{fmtCurShort(t.sales_amount)}</td>
                          <td className="px-2 py-2 text-orange-600 whitespace-nowrap">{fmtCurShort(t.purchase_amount)}</td>
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
  const renderDashboard = () => {
    // Apply client-side date filter to dashboard data
    const filteredMonthly = dashDateFilter.mode === 'all'
      ? (stats?.monthly_trend || [])
      : (stats?.monthly_trend || []).filter(m => {
          if (!m.month) return false;
          // month format: "Jan 2024"
          try {
            const d = new Date(m.month);
            const iso = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
            if (dashDateFilter.mode === 'year') return iso.startsWith(dashDateFilter.year);
            if (dashDateFilter.mode === 'range') {
              if (dashDateFilter.start && iso < dashDateFilter.start.slice(0,7)) return false;
              if (dashDateFilter.end && iso > dashDateFilter.end.slice(0,7)) return false;
              return true;
            }
          } catch { return true; }
          return true;
        });

    return (
    <>
      {/* Date Range Filter + Info Row */}
      <div className="mb-4">
        <DateRangeFilter
          darkMode={darkMode} txt={txt} txt2={txt2} card={card}
          value={globalDateFilter}
          label="Filter SO Create Date"
          onFilter={(f) => { setGlobalDateFilter(f); }}
        />
        {/* Date range info row */}
        <div className={`-mt-3 mb-4 px-5 py-2.5 rounded-b-xl flex flex-wrap gap-4 text-xs ${darkMode?'bg-gray-800/50':'bg-purple-50/80'} border-x border-b ${darkMode?'border-gray-700':'border-gray-100'}`}>
          <div className="flex items-center gap-1.5">
            <Calendar className="w-3.5 h-3.5 text-purple-400"/>
            <span className={txt2}>PO Range:</span>
            <span className={`font-semibold ${txt}`}>{fmtDateRange(stats?.po_date_range)}</span>
          </div>
          <div className="flex items-center gap-1.5">
            <Calendar className="w-3.5 h-3.5 text-blue-400"/>
            <span className={txt2}>SO Create Range:</span>
            <span className={`font-semibold ${txt}`}>{fmtDateRange(stats?.so_date_range)}</span>
          </div>
          {stats?.last_updated && (
            <div className="flex items-center gap-1.5 ml-auto">
              <span className={txt2}>Last Updated:</span>
              <span className={`font-semibold ${txt}`}>
                {(() => { try { return new Date(stats.last_updated).toLocaleString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'}); } catch { return stats.last_updated; } })()}
              </span>
            </div>
          )}
        </div>
      </div>

      {/* KPI Row */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <div className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all cursor-pointer ${card}`}
          onClick={() => {
            setActivePage('all-so');
            setSoPage(1);
            fetchSOData(soFilters, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
            setTimeout(() => { poTableRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' }); }, 300);
          }}>
          <div className="flex justify-between items-start">
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
          onClick={() => openModal('SO without PO HLI', `/api/data/so-without-po${(() => {
            const f = globalDateFilter; if (!f || f.mode === 'all') return '';
            const p = new URLSearchParams();
            if (f.mode === 'year') p.append('date_year', f.year);
            else if (f.mode === 'range') {
              if (f.start) p.append('date_from', f.start);
              if (f.end)   p.append('date_to',   f.end);
            }
            return p.toString() ? `?${p}` : '';
          })()}`)}>
          <div className="flex justify-between items-start">
            <div>
              <p className={`text-sm font-medium ${txt2}`}>SO without PO HLI</p>
              <h3 className="text-3xl font-bold mt-1 text-orange-500">{fmtNum(stats?.so_without_po)}</h3>
              <p className={`text-xs mt-1 ${txt2}`}>Click for details</p>
            </div>
            <div className="p-3 bg-orange-100 rounded-xl"><XCircle className="w-6 h-6 text-orange-500"/></div>
          </div>
        </div>

        <div className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all ${card}`}>
          <div className="flex justify-between items-start">
            <div>
              <p className={`text-sm font-medium ${txt2}`}>Total PO HLI Amount</p>
              <h3 className={`text-xl font-bold mt-1 text-purple-600`}>{fmtCurShort(stats?.total_po_amount)}</h3>
              <p className={`text-xs mt-1 ${txt2}`}>{fmtCur(stats?.total_po_amount)}</p>
            </div>
            <div className="p-3 bg-purple-100 rounded-xl"><Coins className="w-6 h-6 text-purple-600"/></div>
          </div>
        </div>

        <div className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all cursor-pointer ${card}`}
          onClick={() => {
            setActivePage('all-so');
            setSoPage(1);
            fetchSOData(soFilters, 1, soPerPage, soSearchNums, soMarginFilter, soDateFilter);
            window.scrollTo({ top: 0, behavior: 'smooth' });
          }}>
          <div className="flex justify-between items-start">
            <div>
              <p className={`text-sm font-medium ${txt2}`}>Total SO (Open)</p>
              <h3 className="text-3xl font-bold mt-1 text-green-600">{fmtNum(stats?.total_so_count)}</h3>
              <p className={`text-xs mt-1 ${txt2}`}>{stats?.so_date_range?.max ? fmtDate(stats.so_date_range.max) : 'No data uploaded'} · click for details</p>
            </div>
            <div className="p-3 bg-green-100 rounded-xl"><CheckCircle className="w-6 h-6 text-green-600"/></div>
          </div>
        </div>
      </div>

      {/* Charts Row 1 */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6 items-start">
        <div className={`p-6 rounded-2xl shadow ${card}`}>
          <h3 className={`text-base font-bold mb-4 flex items-center gap-2 ${txt}`}>
            <TrendingUp className="w-5 h-5 text-purple-600"/> Monthly Open SO Trend
          </h3>
          <ResponsiveContainer width="100%" height={190}>
            <AreaChart data={filteredMonthly}>              <defs>
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
              <Area yAxisId="right" type="monotone" dataKey="amount" name="Value (IDR Mil)" stroke="#F97316" strokeWidth={2} fill="url(#cAmt)"/>
            </AreaChart>
          </ResponsiveContainer>
        </div>

        <div className="flex flex-col gap-4">
          <div className={`p-5 rounded-2xl shadow ${card}`}>
            <h3 className={`text-sm font-bold mb-3 flex items-center gap-2 ${txt}`}>
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
                      <span className={`inline-flex items-center justify-center w-6 h-6 rounded text-xs font-bold ${i===0?'bg-yellow-100 text-yellow-700':i===1?'bg-gray-200 text-gray-700':i===2?'bg-orange-100 text-orange-700':'bg-purple-100 text-purple-700'}`}>#{i+1}</span>
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
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6 items-stretch">
        <div className={`p-6 rounded-2xl shadow flex flex-col ${card}`}>
          <h3 className={`text-base font-bold mb-4 flex items-center gap-2 ${txt}`}>
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
                              onClick={() => openModal(`All Status — ${m}`, `/api/data/so-status-detail-all?month=${encodeURIComponent(m)}`)}
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
            <h3 className={`text-sm font-bold mb-3 flex items-center gap-2 ${txt}`}>
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
              <h3 className={`text-sm font-bold mb-2 flex items-center gap-2 ${txt}`}><BarChart3 className="w-4 h-4 text-orange-600"/> SO Status (Pie)</h3>
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
                  <h3 className={`text-sm font-bold mb-2 flex items-center gap-2 ${txt}`}><Calendar className="w-4 h-4 text-red-500"/> SO Aging (Pie)</h3>
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
        <h3 className={`text-base font-bold mb-4 flex items-center gap-2 ${txt}`}>
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
                        onClick={() => openModal('All Vendors — All Aging', '/api/data/aging-detail-all')}
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
  };

  // ══════════════════════════════════════════════════════════════
  // RENDER ALL SO PAGE
  // ══════════════════════════════════════════════════════════════
  const renderAllSO = () => (
    <div>
      {/* Date Range Filter */}
      <DateRangeFilter
        darkMode={darkMode} txt={txt} txt2={txt2} card={card}
        value={globalDateFilter}
        label="Filter SO Create Date"
        onFilter={(f) => {
          setGlobalDateFilter(f);
          setSoPage(1);
          fetchSOData(soFilters, 1, soPerPage, soSearchNums, soMarginFilter, f);
        }}
      />
      <div className={`p-6 rounded-2xl shadow mb-6 ${card}`}>
        <div className="flex flex-wrap justify-between items-center gap-3 mb-5">
          <div>
            <h2 className={`text-xl font-bold ${txt}`}>Open SO (Sales Order)</h2>
            <p className={`text-sm ${txt2}`}>{fmtNum(soTotal)} total records — page {soPage} of {soTotalPages}</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <label className="flex items-center gap-1 px-3 py-1.5 bg-green-700 hover:bg-green-800 text-white rounded-lg text-sm font-medium shadow-sm">
              <Upload className="w-4 h-4"/>Batch Upload
              <input type="file" accept=".xlsx,.xls" onChange={handleBatchUpload} className="hidden"/>
            </label>
            <DownloadButton onClick={downloadSOTemplate} className="flex items-center gap-1 px-3 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-sm font-medium shadow-sm">
              <FileSpreadsheet className="w-4 h-4"/>Template
            </DownloadButton>
            <DownloadButton onClick={downloadSOExcel} className="flex items-center gap-1 px-3 py-1.5 bg-purple-700 hover:bg-purple-800 text-white rounded-lg text-sm font-medium shadow-sm">
              <Download className="w-4 h-4"/>Download Excel
            </DownloadButton>
          </div>
        </div>

        {/* Aging Filter Chips */}
        <div className="mb-3 flex flex-wrap gap-2 items-center">
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

        {/* Multi-select Filters row — Search SO leftmost */}
        <div className={`p-4 rounded-xl mb-4 ${darkMode?'bg-gray-700':'bg-gray-50'}`}>
          <div className="flex flex-wrap gap-3 items-end">
            {/* Search SO Item — leftmost */}
            <div>
              <label className={`block text-xs font-medium mb-1 ${txt2}`}>Search SO Item</label>
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
              <label className={`block text-xs font-medium mb-1 ${txt2}`}>Filter Margin</label>
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
              <button onClick={()=>{ setSoPage(1); fetchSOData(soFilters,1,soPerPage,soSearchNums,soMarginFilter,soDateFilter); }}
                className="px-5 py-2 bg-purple-700 hover:bg-purple-800 text-white rounded-lg text-sm font-semibold shadow-sm">Apply</button>
              <button onClick={()=>{
                const f={op_units:[],vendors:[],statuses:[],aging:[]};
                setSoFilters(f); setSoSearchNums([]); setSoMarginFilter('all'); setSoPage(1);
                fetchSOData(f,1,soPerPage,[],'all',soDateFilter);
              }}
                className={`px-4 py-2 rounded-lg text-sm font-medium shadow-sm ${darkMode?'bg-gray-500 text-gray-100 hover:bg-gray-400':'bg-gray-400 text-white hover:bg-gray-500'}`}>Reset</button>
            </div>
          </div>
          {/* Active filter tags */}
          {(soSearchNums.length + soFilters.op_units.length + soFilters.vendors.length + soFilters.statuses.length) > 0 && (
            <div className="mt-3 flex flex-wrap gap-1.5">
              {soSearchNums.map(v=>(
                <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-indigo-100 text-indigo-700 rounded-full text-xs">
                  SO: {v}<button onClick={()=>{ const next=soSearchNums.filter(x=>x!==v); setSoSearchNums(next); setSoPage(1); fetchSOData(soFilters,1,soPerPage,next,soMarginFilter,soDateFilter); }} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {soFilters.op_units.map(v=>(
                <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-purple-100 text-purple-700 rounded-full text-xs">
                  {v}<button onClick={()=>setSoFilters(f=>({...f,op_units:f.op_units.filter(x=>x!==v)}))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {soFilters.vendors.map(v=>(
                <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-blue-100 text-blue-700 rounded-full text-xs">
                  {v}<button onClick={()=>setSoFilters(f=>({...f,vendors:f.vendors.filter(x=>x!==v)}))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
                </span>
              ))}
              {soFilters.statuses.map(v=>(
                <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-green-100 text-green-700 rounded-full text-xs">
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
                    <FileText className="w-10 h-10 mx-auto mb-2 opacity-40"/>No data
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
                  <td className="px-3 py-2 text-purple-600 font-medium whitespace-nowrap">{so.so_item}</td>
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
                  <td className={`px-3 py-2 text-right whitespace-nowrap min-w-[130px] ${txt}`}>{fmtCur(so.sales_price)}</td>
                  <td className="px-3 py-2 text-right font-semibold text-orange-600 whitespace-nowrap min-w-[130px]">{fmtCur(so.sales_amount)}</td>
                  <td className={`px-3 py-2 text-right whitespace-nowrap min-w-[130px] ${txt}`}>{fmtCur(so.purchasing_price)}</td>
                  <td className="px-3 py-2 text-right font-semibold text-green-600 whitespace-nowrap min-w-[130px]">{fmtCur(poAmount)}</td>
                  <td className={`px-3 py-2 text-right whitespace-nowrap min-w-[130px] ${marginColor}`}>{fmtCur(margin)}</td>
                  <td className={`px-3 py-2 text-right whitespace-nowrap ${marginColor}`}>
                    {marginPct !== null ? `${marginPct.toFixed(1)}%` : '-'}
                  </td>
                  <td className={`px-3 py-2 text-center text-xs ${txt2}`}>{so.delivery_possible_date||'-'}</td>
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
        <div className={`mt-4 pt-3 border-t ${darkMode?'border-gray-700':'border-gray-200'} flex justify-between items-center`}>
          <span className={`text-sm ${txt2}`}>
            Showing {((soPage-1)*soPerPage)+1}–{Math.min(soPage*soPerPage,soTotal)} of {fmtNum(soTotal)}
          </span>
          <div className="flex gap-1 items-center">
            <button disabled={soPage===1} onClick={()=>{ const p=soPage-1; setSoPage(p); fetchSOData(soFilters,p,soPerPage,soSearchNums,soMarginFilter,soDateFilter); }}
              className={`p-1.5 rounded ${soPage===1?'opacity-40':'hover:bg-purple-100'}`}><ChevronLeft className="w-4 h-4"/></button>
            <span className={`px-3 py-1 rounded text-sm font-semibold ${darkMode?'bg-gray-700 text-white':'bg-purple-100 text-purple-700'}`}>{soPage}/{soTotalPages}</span>
            <button disabled={soPage===soTotalPages} onClick={()=>{ const p=soPage+1; setSoPage(p); fetchSOData(soFilters,p,soPerPage,soSearchNums,soMarginFilter,soDateFilter); }}
              className={`p-1.5 rounded ${soPage===soTotalPages?'opacity-40':'hover:bg-purple-100'}`}><ChevronRight className="w-4 h-4"/></button>
          </div>
        </div>
      </div>

      {/* PO HLI Without SO Table */}
      <div ref={poTableRef} className={`rounded-2xl shadow overflow-hidden ${card}`}>
        <div className={`p-5 border-b ${darkMode?'border-gray-700':'border-gray-100'} flex flex-wrap justify-between items-center gap-3`}>
          <div className="flex items-center gap-2">
            <AlertCircle className="w-5 h-5 text-yellow-600"/>
            <h3 className={`text-base font-bold ${txt}`}>PO HLI Without SO</h3>
            <span className={`text-sm ${txt2}`}>
              ({fmtNum(new Set(poFiltered.map(p=>p.po_no)).size)} PO · {fmtNum(poFiltered.length)} line items)
            </span>
          </div>
          <div className="flex gap-2 items-center">
            <DownloadButton onClick={downloadPOExcel} className="flex items-center gap-1 px-4 py-1.5 bg-purple-700 hover:bg-purple-800 text-white rounded-lg text-sm font-medium shadow-sm">
              <Download className="w-4 h-4"/>Download Excel
            </DownloadButton>
          </div>
        </div>

        {/* PO Filters row */}
        <div className={`px-5 py-3 border-b ${darkMode?'border-gray-700 bg-gray-750':'border-gray-100 bg-gray-50'} flex flex-wrap gap-3 items-end`}>
          {/* Search PO HLI Number — leftmost */}
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
          <div className="flex gap-2 items-end">
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
              <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-indigo-100 text-indigo-700 rounded-full text-xs">
                PO: {v}<button onClick={()=>setPoSearchNums(prev=>prev.filter(x=>x!==v))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
              </span>
            ))}
            {poFilterItemType.map(v=>(
              <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-blue-100 text-blue-700 rounded-full text-xs">
                {v}<button onClick={()=>setPoFilterItemType(prev=>prev.filter(x=>x!==v))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
              </span>
            ))}
            {poFilterOpUnit.map(v=>(
              <span key={v} className="flex items-center gap-1 px-2 py-0.5 bg-purple-100 text-purple-700 rounded-full text-xs">
                {v}<button onClick={()=>setPoFilterOpUnit(prev=>prev.filter(x=>x!==v))} className="hover:text-red-600"><X className="w-3 h-3"/></button>
              </span>
            ))}
          </div>
        )}

        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className={tblHd}>
              <tr>
                {['PO HLI NUMBER','PO ITEM TYPE','ITEM CODE','OPERATION UNIT','DESCRIPTION','QTY','UNIT','PRICE','AMOUNT','CURRENCY','PO DATE','PURCHASE MEMBER','REQ. DELIVERY','WORKING DAYS LEFT'].map(h=>(
                  <th key={h} className={`px-4 py-3 text-left font-semibold whitespace-nowrap ${txt2} ${h==='PRICE'||h==='AMOUNT'?'min-w-[140px]':''}`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {poRows.length === 0 ? (
                <tr><td colSpan={16} className={`px-4 py-10 text-center ${txt2}`}>
                  <Package className="w-10 h-10 mx-auto mb-2 opacity-40"/>No data
                </td></tr>
              ) : poRows.map((row,i)=>{
                const daysLeft = row.days_remaining;
                const daysColor = daysLeft === null ? txt2 : daysLeft < 0 ? 'text-red-600 font-bold' : daysLeft <= 7 ? 'text-orange-600 font-semibold' : daysLeft <= 30 ? 'text-yellow-600' : 'text-green-600';
                return (
                  <tr key={i} className={`${trHov} transition-colors`}>
                    <td className="px-4 py-3 text-purple-600 font-medium whitespace-nowrap">
                      {row.item_no ? `${row.po_no}-${row.item_no}` : row.po_no}
                    </td>
                    <td className={`px-4 py-3 whitespace-nowrap`}>
                      {row.po_item_type ? (
                        <span className={`px-2 py-0.5 rounded-full text-xs font-medium ${
                          row.po_item_type.toUpperCase()==='MRO' ? 'bg-blue-100 text-blue-700' :
                          row.po_item_type.toUpperCase()==='EQUIPMENT' ? 'bg-green-100 text-green-700' :
                          'bg-gray-100 text-gray-700'}`}>{row.po_item_type}</span>
                      ) : <span className={`${txt2} text-xs`}>-</span>}
                    </td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.item_code||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap text-xs`} title={row.operation_unit}>{row.operation_unit||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} max-w-xs truncate`} title={row.description}>{row.description}</td>
                    <td className={`px-4 py-3 text-right ${txt2}`}>{fmtNum(row.qty)}</td>
                    <td className={`px-4 py-3 ${txt2}`}>{row.unit||'-'}</td>
                    <td className={`px-4 py-3 text-right whitespace-nowrap min-w-[140px] ${txt}`}>{fmtCur(row.price)}</td>
                    <td className="px-4 py-3 text-right font-semibold text-orange-600 whitespace-nowrap min-w-[140px]">{fmtCur(row.amount)}</td>
                    <td className={`px-4 py-3 ${txt2}`}>{row.currency||'IDR'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.po_date||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.purchase_member||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.req_delivery||'-'}</td>
                    <td className={`px-4 py-3 text-center whitespace-nowrap ${daysColor}`}>
                      {daysLeft === null ? '-' : daysLeft < 0 ? `${Math.abs(daysLeft)} days overdue` : `${daysLeft} days`}                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className={`p-4 border-t ${darkMode?'border-gray-700':'border-gray-100'} flex justify-between items-center`}>
          <span className={`text-sm ${txt2}`}>Showing {(poPage-1)*poPerPage+1}–{Math.min(poPage*poPerPage,poFiltered.length)} of {fmtNum(poFiltered.length)}</span>
          <div className="flex gap-1 items-center">
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
    <div className={`min-h-screen font-sans ${darkMode?'bg-gray-900':'bg-gray-50'}`}>
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
      `}</style>

      <div className="fixed top-5 right-5 z-[100] flex flex-col gap-2">
        {toasts.map(t=><Toast key={t.id} message={t.message} type={t.type} onClose={()=>removeToast(t.id)}/>)}
      </div>

      {/* Download progress toast */}
      {downloadToast && <DownloadToast message={downloadToast.message} />}

      {/* Sidebar */}
      <aside className={`fixed left-0 top-0 h-full w-14 flex flex-col items-center py-6 shadow-2xl z-40 ${darkMode?'bg-gray-800 border-r border-gray-700':'bg-gradient-to-b from-purple-600 to-purple-700'}`}>
        <nav className="flex-1 flex flex-col gap-3 w-full px-1.5 pt-0">
          <button onClick={()=>setActivePage('dashboard')}
            className={`p-2 rounded-lg flex justify-center transition-all ${activePage==='dashboard'?'bg-white/30 text-white shadow-lg':'text-purple-100 hover:bg-white/20'}`} title="Dashboard">
            <BarChart3 className="w-5 h-5"/>
          </button>
          <button onClick={()=>{ setActivePage('all-so'); setSoPage(1); fetchSOData(soFilters,1,soPerPage,soSearchNums,soMarginFilter,soDateFilter); window.scrollTo({top:0, behavior:'smooth'}); }}
            className={`p-2 rounded-lg flex justify-center transition-all ${activePage==='all-so'?'bg-white/30 text-white shadow-lg':'text-purple-100 hover:bg-white/20'}`} title="Open SO (Sales Order)">
            <FileText className="w-5 h-5"/>
          </button>
          <button onClick={()=>{ setActivePage('completed'); fetchCompletedData(completedYear, completedDateFilter); window.scrollTo({top:0,behavior:'smooth'}); }}
            className={`p-2 rounded-lg flex justify-center transition-all ${activePage==='completed'?'bg-white/30 text-white shadow-lg':'text-purple-100 hover:bg-white/20'}`} title="Delivery Completed">
            <Coins className="w-5 h-5"/>
          </button>
        </nav>
        <button onClick={()=>setDarkMode(d=>!d)} className="p-2 rounded-lg text-white hover:bg-white/20 transition-all">
          {darkMode?<Sun className="w-5 h-5"/>:<Moon className="w-5 h-5"/>}
        </button>
      </aside>

      {/* Main */}
      <main className="ml-14 p-6">
        <header className="mb-6 flex flex-wrap justify-between items-center gap-4">
          <div>
            <h1 className={`text-2xl font-bold tracking-tight ${txt}`}>
              HLI PO Monitoring <span className="text-purple-600">Dashboard</span>
            </h1>
            <p className={`mt-0.5 text-sm ${txt2}`}>
              {activePage==='dashboard'?'Purchase Orders & Sales Orders Overview'
               :activePage==='completed'?'Delivery Completed Analytics'
               :'Manage Open SO (Sales Order) & PO Without SO'}
            </p>
          </div>
          <div className="flex flex-wrap gap-2">
            <label className={`flex items-center gap-2 px-4 py-2 rounded-xl shadow hover:shadow-md transition-all ${darkMode?'bg-purple-700 hover:bg-purple-800 text-white':'bg-purple-700 hover:bg-purple-800 text-white'}`}>
              <Upload className="w-4 h-4"/><span className="text-sm font-medium">Upload HLI PO List (Item)</span>
              <input type="file" accept=".xlsx,.xls" onChange={e=>handleUpload(e,'po')} className="hidden"/>
            </label>
            <label className={`flex items-center gap-2 px-4 py-2 rounded-xl shadow hover:shadow-md transition-all ${darkMode?'bg-blue-700 hover:bg-blue-800 text-white':'bg-blue-700 hover:bg-blue-800 text-white'}`}>
              <Upload className="w-4 h-4"/><span className="text-sm font-medium">Upload SMRO - Search Client Odr</span>
              <input type="file" accept=".xlsx,.xls" onChange={e=>handleUpload(e,'smro')} className="hidden"/>
            </label>
            <label title="Upload SMRO file to fill in missing Specification & Product ID data on existing records"
              className={`flex items-center gap-2 px-4 py-2 rounded-xl shadow hover:shadow-md transition-all ${darkMode?'bg-teal-700 hover:bg-teal-800 text-white':'bg-teal-600 hover:bg-teal-700 text-white'}`}>
              <Plus className="w-4 h-4"/><span className="text-sm font-medium">Backfill Spec & Product ID</span>
              <input type="file" accept=".xlsx,.xls" onChange={async e => {
                const file = e.target.files?.[0];
                if (!file) return;
                e.target.value = '';
                setDownloadToast({ message: 'Backfilling Specification & Product ID...' });
                try {
                  const fd = new FormData();
                  fd.append('file', file);
                  const res = await api.post('/api/upload/smro-backfill-spec', fd);
                  setDownloadToast(null);
                  addToast(`✅ ${res.data.message}`, 'success');
                  // Refresh dashboard data to show updated spec/pid
                  fetchDashboard(globalDateFilter);
                  if (activePage === 'completed') fetchCompletedData(completedYear, completedDateFilter);
                } catch (err) {
                  setDownloadToast(null);
                  addToast(`❌ Backfill gagal: ${err.response?.data?.error || err.message}`, 'error');
                }
              }} className="hidden"/>
            </label>
            <div className="relative" ref={hideMenuRef}>
              <button onClick={()=>setShowHideMenu(o=>!o)}
                className="flex items-center gap-2 px-4 py-2 rounded-xl shadow hover:shadow-md transition-all bg-orange-600 hover:bg-orange-700 text-white">
                <EyeOff className="w-4 h-4"/><span className="text-sm font-medium">Hide</span>
                <ChevronDown className="w-3.5 h-3.5"/>
                {deleteRequests.filter(r=>r.is_hidden).length > 0 && (
                  <span className="px-1.5 py-0.5 bg-white text-orange-600 rounded-full text-xs font-bold">
                    {deleteRequests.filter(r=>r.is_hidden).length}
                  </span>
                )}
              </button>
              {showHideMenu && (
                <div className={`absolute right-0 mt-2 z-50 rounded-xl shadow-2xl border w-80 p-3 ${darkMode?'bg-gray-800 border-gray-700 text-white':'bg-white border-gray-200'}`}>
                  {/* View Hidden History */}
                  <button onClick={()=>{ setShowHideMenu(false); fetchDeleteRequests(); setShowHiddenPanel(true); }}
                    className={`w-full flex items-center gap-2 px-3 py-2.5 rounded-lg text-sm font-semibold mb-3 ${darkMode?'bg-gray-700 hover:bg-gray-600 text-white':'bg-gray-100 hover:bg-gray-200 text-gray-700'}`}>
                    <Eye className="w-4 h-4 text-purple-500"/>
                    View Hide History
                    {deleteRequests.filter(r=>r.is_hidden).length > 0 && (
                      <span className="ml-auto px-2 py-0.5 bg-orange-500 text-white rounded-full text-xs font-bold">
                        {deleteRequests.filter(r=>r.is_hidden).length}
                      </span>
                    )}
                  </button>
                  <p className={`text-xs font-semibold mb-2 px-1 ${darkMode?'text-gray-300':'text-gray-600'}`}>
                    Hide data from dashboard via Excel template
                  </p>
                  {/* PO HLI */}
                  <div className={`mb-2 p-3 rounded-lg ${darkMode?'bg-gray-700':'bg-orange-50'}`}>
                    <p className="text-xs font-bold mb-1 text-orange-700">🔵 PO HLI (from PO List)</p>
                    <p className={`text-xs mb-2 ${darkMode?'text-gray-400':'text-gray-500'}`}>Format: PO Number-Item No (e.g. 4502358819-10)</p>
                    <div className="flex gap-2">
                      <button onClick={()=>downloadHideTemplate('PO')}
                        className="flex-1 flex items-center justify-center gap-1 px-2 py-1.5 bg-orange-600 hover:bg-orange-700 text-white rounded-lg text-xs font-semibold">
                        <Download className="w-3 h-3"/>Download Template
                      </button>
                      <label className="flex-1 flex items-center justify-center gap-1 px-2 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-xs font-semibold cursor-pointer">
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
                        className="flex-1 flex items-center justify-center gap-1 px-2 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold">
                        <Download className="w-3 h-3"/>Download Template
                      </button>
                      <label className="flex-1 flex items-center justify-center gap-1 px-2 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-xs font-semibold cursor-pointer">
                        <Upload className="w-3 h-3"/>Upload Filled
                        <input type="file" accept=".xlsx,.xls" onChange={e=>handleHideBatchUpload(e,'SO')} className="hidden"/>
                      </label>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        </header>

        {activePage==='dashboard' ? renderDashboard() : activePage==='completed' ? renderCompleted() : renderAllSO()}
      </main>

      {modal && <SOModal title={modal.title} data={modal.data} darkMode={darkMode} onClose={()=>setModal(null)}/>}

      {marginDetailModal && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm" onClick={()=>setMarginDetailModal(null)}>
          <div className={`rounded-2xl shadow-2xl w-full max-w-5xl max-h-[85vh] flex flex-col ${darkMode?'bg-gray-800 text-white':'bg-white'}`} onClick={e=>e.stopPropagation()}>
            <div className={`flex justify-between items-center px-6 py-4 border-b ${darkMode?'border-gray-700':'border-gray-100'}`}>
              <h3 className="font-bold text-lg">Margin Detail — {marginDetailModal.category}
                <span className={`text-sm font-normal ml-2 ${txt2}`}>({fmtNum(marginDetailModal.data?.length)} records)</span>
              </h3>
              <button onClick={()=>setMarginDetailModal(null)} className={`p-1.5 rounded-lg ${darkMode?'hover:bg-gray-700':'hover:bg-gray-100'}`}><X className="w-5 h-5"/></button>
            </div>
            <div className="overflow-auto flex-1">
              <table className="w-full text-xs">
                <thead className={`sticky top-0 ${darkMode?'bg-gray-700':'bg-purple-50'}`}>
                  <tr>{['SO Item','Product','Vendor','Sales','Purchase','Margin','%','Date'].map(h=>(
                    <th key={h} className={`px-3 py-2 text-left font-semibold ${darkMode?'text-gray-200':'text-gray-700'}`}>{h}</th>
                  ))}</tr>
                </thead>
                <tbody className={`divide-y ${darkMode?'divide-gray-700':'divide-gray-100'}`}>
                  {(marginDetailModal.data||[]).map((t,i)=>(
                    <tr key={i} className={darkMode?'hover:bg-gray-700':'hover:bg-purple-50'}>
                      <td className="px-3 py-2 text-purple-600 font-medium whitespace-nowrap">{t.so_item||'-'}</td>
                      <td className={`px-3 py-2 max-w-[160px] truncate ${txt}`}>{t.product||'-'}</td>
                      <td className={`px-3 py-2 max-w-[120px] truncate ${txt2}`}>{t.vendor||'-'}</td>
                      <td className="px-3 py-2 text-right text-blue-600 whitespace-nowrap">{fmtCur(t.sales_amount)}</td>
                      <td className="px-3 py-2 text-right text-orange-600 whitespace-nowrap">{fmtCur(t.purchase_amount)}</td>
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
        <div className="fixed inset-0 bg-black/30 z-[55] flex items-center justify-center">
          <div className={`${darkMode?'bg-gray-800':'bg-white'} px-6 py-4 rounded-xl shadow-xl flex items-center gap-3`}>
            <div className="w-6 h-6 border-3 border-purple-600 border-t-transparent rounded-full animate-spin"/>
            <p className={`text-sm font-semibold ${txt}`}>Loading data...</p>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
