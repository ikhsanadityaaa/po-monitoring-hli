import React, { useState, useEffect, useCallback } from 'react';
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, AreaChart, Area
} from 'recharts';
import {
  Upload, Download, AlertCircle, CheckCircle, XCircle,
  Package, DollarSign, TrendingUp, Calendar, ChevronLeft,
  ChevronRight, Moon, Sun, FileText, BarChart3, FileSpreadsheet,
  Filter, X, ChevronDown, ChevronUp, Building2
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

const SOModal = ({ title, data, onClose, darkMode }) => {
  const [dlPage, setDlPage] = useState(1);
  const PER = 50;
  const pages = Math.ceil((data?.length || 0) / PER);
  const rows = (data || []).slice((dlPage-1)*PER, dlPage*PER);
  const downloadExcel = () => {
    const ws = XLSX.utils.json_to_sheet(data.map(s => ({
      'SO Number': s.so_number, 'SO Item': s.so_item, 'Status': s.so_status,
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
              <tr>{['Aging','SO Number','SO Item','Status','Op Unit','Vendor','Product','Qty','Sales Amount','Cust PO','Delivery Memo','SO Date','Plan Date','Remarks'].map(h=>(
                <th key={h} className={`px-3 py-2 text-left font-semibold whitespace-nowrap ${darkMode?'text-gray-200':'text-gray-700'}`}>{h}</th>
              ))}</tr>
            </thead>
            <tbody className={`divide-y ${darkMode?'divide-gray-700':'divide-gray-100'}`}>
              {rows.map((s,i)=>(
                <tr key={i} className={darkMode?'hover:bg-gray-700':'hover:bg-purple-50'}>
                  <td className="px-3 py-2 whitespace-nowrap"><span className="px-2 py-0.5 rounded-full text-xs font-bold text-white" style={{backgroundColor: AGING_COLORS[s.aging_label] || '#6B7280'}}>{s.aging_label||'-'}</span></td>
                  <td className="px-3 py-2 text-purple-600 font-medium whitespace-nowrap">{s.so_number}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{s.so_item}</td>
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

// ═══════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════
const App = () => {
  const [darkMode, setDarkMode] = useState(false);
  const [activePage, setActivePage] = useState('dashboard');

  const [stats, setStats] = useState(null);
  const [poWithoutSO, setPoWithoutSO] = useState([]);
  const [agingData, setAgingData] = useState([]);
  const [allSOData, setAllSOData] = useState([]);
  const [soTotal, setSoTotal] = useState(0);
  const [soFilterOptions, setSoFilterOptions] = useState({ op_units: [], vendors: [], statuses: [] });

  // SO filters — semua multi-select
  const [soFilters, setSoFilters] = useState({ op_units: [], vendors: [], statuses: [], aging: [] });
  const [soPage, setSoPage] = useState(1);
  const [soPerPage, setSoPerPage] = useState(20);

  const [poPage, setPoPage] = useState(1);
  const [poPerPage, setPoPerPage] = useState(20);

  const [loading, setLoading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(null);
  const [toasts, setToasts] = useState([]);
  const [modal, setModal] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState('');

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
      setPoWithoutSO(Array.isArray(pRes.data) ? pRes.data : []);
      setAgingData(Array.isArray(aRes.data) ? aRes.data : []);
    } catch (e) {
      addToast(`Error: ${e.response?.data?.error || e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast]);

  const fetchSOData = useCallback(async (filters, page, perPage) => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, per_page: perPage });
      (filters.op_units || []).forEach(v => params.append('op_unit', v));
      (filters.vendors || []).forEach(v => params.append('vendor', v));
      (filters.statuses || []).forEach(v => params.append('status', v));
      (filters.aging || []).forEach(a => params.append('aging', a));
      const res = await api.get(`/api/data/all-so?${params}`);
      setAllSOData(Array.isArray(res.data.data) ? res.data.data : []);
      setSoTotal(res.data.total || 0);
      setSoFilterOptions(res.data.filters || { op_units: [], vendors: [], statuses: [] });
    } catch (e) {
      addToast(`Gagal memuat SO: ${e.message}`, 'error');
    } finally { setLoading(false); }
  }, [addToast]);

  useEffect(() => { fetchDashboard(); }, []);
  useEffect(() => { if (activePage === 'all-so') fetchSOData(soFilters, soPage, soPerPage); }, [activePage]);

  const handleUpload = async (e, type) => {
    const file = e.target.files[0]; if (!file) return;
    e.target.value = '';
    const label = type === 'po' ? 'PO List' : 'SMRO';
    const endpoint = type === 'po' ? '/api/upload/po-list' : '/api/upload/smro';
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
      if (activePage === 'all-so') fetchSOData(soFilters, 1, soPerPage);
      setSoPage(1);
    } catch (e) {
      setUploadProgress(null);
      addToast(`❌ Gagal upload ${label}: ${e.response?.data?.error || e.message}`, 'error');
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
      addToast(`✅ Batch update: ${res.data.updated} records diperbarui`, 'success');
      fetchSOData(soFilters, soPage, soPerPage);
    } catch (e) {
      setUploadProgress(null);
      addToast(`❌ Gagal batch upload: ${e.response?.data?.error || e.message}`, 'error');
    }
  };

  const downloadBlob = async (url, filename) => {
    try {
      const res = await api.get(url, { responseType: 'blob' });
      const link = document.createElement('a');
      link.href = window.URL.createObjectURL(new Blob([res.data]));
      link.setAttribute('download', filename);
      document.body.appendChild(link); link.click(); link.remove();
      addToast(`✅ File "${filename}" berhasil didownload`, 'success');
    } catch (e) { addToast('❌ Gagal download file', 'error'); }
  };

  const downloadSOExcel = () => {
    const p = new URLSearchParams();
    (soFilters.op_units||[]).forEach(v => p.append('op_unit', v));
    (soFilters.vendors||[]).forEach(v => p.append('vendor', v));
    (soFilters.statuses||[]).forEach(v => p.append('status', v));
    downloadBlob(`/api/export/all-so?${p}`, `SO_List_${new Date().toISOString().slice(0,10)}.xlsx`);
  };
  const downloadPOExcel = () => downloadBlob('/api/export/po-without-so', `PO_Without_SO_${new Date().toISOString().slice(0,10)}.xlsx`);
  const downloadSOTemplate = () => {
    const ws = XLSX.utils.json_to_sheet(allSOData.map(s=>({'SO Number':s.so_number,'Delivery Plan Date':s.delivery_plan_date||'','Remarks':s.remarks||''})));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Template');
    saveAs(new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`SO_Template_${new Date().toISOString().slice(0,10)}.xlsx`);
    addToast('✅ Template berhasil didownload', 'success');
  };

  const updateSOCell = async (soId, field, value) => {
    try {
      await api.put(`/api/data/so/${soId}`, { [field]: value });
      setEditingCell(null);
      setAllSOData(prev => prev.map(s => s.id === soId ? { ...s, [field]: value } : s));
    } catch (e) { addToast(`❌ Gagal update: ${e.message}`, 'error'); }
  };

  const openModal = async (title, endpointOrData) => {
    if (Array.isArray(endpointOrData)) { setModal({ title, data: endpointOrData }); return; }
    try {
      const res = await api.get(endpointOrData);
      setModal({ title, data: Array.isArray(res.data) ? res.data : [] });
    } catch (e) { addToast(`❌ Gagal memuat detail: ${e.message}`, 'error'); }
  };

  const toggleAgingFilter = (label) => {
    setSoFilters(f => {
      const aging = f.aging.includes(label) ? f.aging.filter(a=>a!==label) : [...f.aging, label];
      return {...f, aging};
    });
  };

  const poTotalPages = Math.max(1, Math.ceil(poWithoutSO.length / poPerPage));
  const poRows = poWithoutSO.slice((poPage-1)*poPerPage, poPage*poPerPage);
  const soTotalPages = Math.max(1, Math.ceil(soTotal / soPerPage));

  const card  = darkMode ? 'bg-gray-800 border border-gray-700' : 'bg-white border border-gray-100';
  const txt   = darkMode ? 'text-white' : 'text-gray-900';
  const txt2  = darkMode ? 'text-gray-400' : 'text-gray-600';
  const tblHd = darkMode ? 'bg-gray-700' : 'bg-purple-50';
  const tblDv = darkMode ? 'divide-gray-700' : 'divide-gray-100';
  const trHov = darkMode ? 'hover:bg-gray-700' : 'hover:bg-purple-50';

  // ── Date range helper ──
  const fmtDateRange = (range) => {
    if (!range?.min) return 'Belum ada data';
    return `${fmtDate(range.min)} — ${fmtDate(range.max)}`;
  };

  // ══════════════════════════════════════════════════════════════
  // RENDER DASHBOARD
  // ══════════════════════════════════════════════════════════════
  const renderDashboard = () => (
    <>
      {/* Date Range Info Bar */}
      <div className={`mb-4 px-5 py-3 rounded-xl flex flex-wrap gap-6 text-xs ${darkMode?'bg-gray-800 border border-gray-700':'bg-white border border-gray-100'} shadow`}>
        <div className="flex items-center gap-2">
          <Calendar className="w-4 h-4 text-purple-500"/>
          <span className={txt2}>PO Date Range:</span>
          <span className={`font-semibold ${txt}`}>{fmtDateRange(stats?.po_date_range)}</span>
        </div>
        <div className="flex items-center gap-2">
          <Calendar className="w-4 h-4 text-blue-500"/>
          <span className={txt2}>SO Create Date Range:</span>
          <span className={`font-semibold ${txt}`}>{fmtDateRange(stats?.so_date_range)}</span>
        </div>
      </div>

      {/* KPI Row */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <div className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all cursor-pointer ${card}`}
          onClick={() => setActivePage('all-so')}>
          <div className="flex justify-between items-start">
            <div>
              <p className={`text-sm font-medium ${txt2}`}>PO HLI tanpa SO</p>
              <h3 className="text-3xl font-bold mt-1 text-red-500">{fmtNum(stats?.po_without_so)}</h3>
              <p className={`text-xs mt-1 ${txt2}`}>Klik untuk detail</p>
            </div>
            <div className="p-3 bg-red-100 rounded-xl"><AlertCircle className="w-6 h-6 text-red-500"/></div>
          </div>
        </div>

        <div className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all cursor-pointer ${card}`}
          onClick={() => openModal('SO Tanpa PO HLI', '/api/data/so-without-po')}>
          <div className="flex justify-between items-start">
            <div>
              <p className={`text-sm font-medium ${txt2}`}>SO tanpa PO HLI</p>
              <h3 className="text-3xl font-bold mt-1 text-orange-500">{fmtNum(stats?.so_without_po)}</h3>
              <p className={`text-xs mt-1 ${txt2}`}>Klik untuk detail</p>
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
            <div className="p-3 bg-purple-100 rounded-xl"><DollarSign className="w-6 h-6 text-purple-600"/></div>
          </div>
        </div>

        <div className={`p-5 rounded-2xl shadow hover:shadow-lg transition-all ${card}`}>
          <div className="flex justify-between items-start">
            <div>
              <p className={`text-sm font-medium ${txt2}`}>Total SO (Open)</p>
              <h3 className="text-3xl font-bold mt-1 text-green-600">{fmtNum(stats?.total_so_count)}</h3>
              <p className={`text-xs mt-1 ${txt2}`}>{stats?.so_date_range?.max ? fmtDate(stats.so_date_range.max) : 'Belum ada upload'}</p>
            </div>
            <div className="p-3 bg-green-100 rounded-xl"><CheckCircle className="w-6 h-6 text-green-600"/></div>
          </div>
        </div>
      </div>

      {/* Charts Row 1 — Monthly Trend (kiri) | Top 5 Vendor + Top Op Unit (kanan) */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6 items-start">
        {/* Monthly Trend */}
        <div className={`p-6 rounded-2xl shadow ${card}`}>
          <h3 className={`text-base font-bold mb-4 flex items-center gap-2 ${txt}`}>
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
              <Area yAxisId="left" type="monotone" dataKey="so_count" name="Jumlah SO" stroke="#8B5CF6" strokeWidth={2} fill="url(#cSO)"/>
              <Area yAxisId="right" type="monotone" dataKey="amount" name="Nilai (IDR Juta)" stroke="#F97316" strokeWidth={2} fill="url(#cAmt)"/>
            </AreaChart>
          </ResponsiveContainer>
        </div>

        {/* Right column — Top 5 Vendor */}
        <div className="flex flex-col gap-4">
          {/* Top 5 Vendors */}
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

      {/* Charts Row 2 — SO Status Distribution (kiri, stretch penuh) | kolom kanan (Op Unit + 2 Pie) */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6 items-stretch">
        {/* SO Status Distribution — stretch mengikuti tinggi kolom kanan */}
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
                      <tr key={i} className={`${trHov} cursor-pointer`}
                        onClick={() => openModal(`SO Status: ${s.name}`, `/api/data/so-status-detail/${encodeURIComponent(s.name)}`)}>
                        <td className={`px-3 py-2 font-medium whitespace-nowrap sticky left-0 ${darkMode?'bg-gray-800':'bg-white'} ${txt}`}>{s.name}</td>
                        {months.map(m => {
                          const val = s.monthly?.[m];
                          return val ? (
                            <td key={m} className="px-2 py-2 text-center font-semibold text-white" style={{backgroundColor:'#7C3AED'}}>{fmtNum(val)}</td>
                          ) : (
                            <td key={m} className="px-2 py-2 text-center" style={{backgroundColor: darkMode?'rgba(59,130,246,0.08)':'rgba(219,234,254,0.45)'}}></td>
                          );
                        })}
                        <td className="px-3 py-2 text-right font-bold text-purple-600">{fmtNum(s.total)}</td>
                        <td className="px-3 py-2 text-right text-green-600">{s.percentage}%</td>
                        <td className="px-3 py-2 text-right text-orange-600 whitespace-nowrap">{fmtCurShort(s.amount)}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot className={`${tblHd} font-bold`}>
                    <tr>
                      <td className={`px-3 py-2 sticky left-0 ${darkMode?'bg-gray-700':'bg-purple-50'} ${txt}`}>TOTAL</td>
                      {months.map(m => <td key={m} className="px-2 py-2 text-center text-purple-600">{totByMonth[m]?fmtNum(totByMonth[m]):''}</td>)}
                      <td className="px-3 py-2 text-right text-purple-600">{fmtNum(grandTotal)}</td>
                      <td className="px-3 py-2 text-right text-green-600">100%</td>
                      <td className="px-3 py-2 text-right text-orange-600 whitespace-nowrap">{fmtCurShort(grandAmount)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            );
          })()}
        </div>

        {/* Kolom kanan: Op Unit (atas) + 2 Pie (bawah, side by side) */}
        <div className="flex flex-col gap-4">
          {/* Total Open SO per Operation Unit */}
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

          {/* ┌─────────────────────────────────────────────────────────────┐
               │  2 PIE CHARTS — PANDUAN KUSTOMISASI                         │
               │                                                              │
               │  LEBAR CHART  → ubah minHeight di kedua <div style>         │
               │  DIAMETER     → outerRadius (jari-jari luar donut)          │
               │  LUBANG TENGAH→ innerRadius (0 = pie penuh, > 0 = donut)   │
               │  POSISI VERTIKAL → cy="44%" (naik = kecilkan, turun = besar)│
               │  POSISI HORIZ → cx="50%" (biasanya biarkan 50%)            │
               │  JARAK ANTAR SLICE → paddingAngle                           │
               └─────────────────────────────────────────────────────────────┘ */}
          <div className="grid gap-4 flex-1" style={{gridTemplateColumns: '7fr 3fr'}}>

            {/* ── PIE 1: SO Status ── */}
            <div className={`p-5 rounded-2xl shadow ${card} flex flex-col`}>
              <h3 className={`text-sm font-bold mb-2 flex items-center gap-2 ${txt}`}><BarChart3 className="w-4 h-4 text-orange-600"/> SO Status (Pie)</h3>
              {/* minHeight = tinggi area chart dalam pixel — naikkan untuk chart lebih besar */}
              <div className="flex-1" style={{minHeight: 300}}>
                <ResponsiveContainer width="100%" height="100%">
                  {/* margin = jarak antara tepi chart dan tepi container, cegah label terpotong */}
                  <PieChart margin={{top: 10, right: 10, bottom: 10, left: 10}}>
                    <Pie
                      data={stats?.so_status||[]}
                      cx="50%"          /* posisi horizontal center donut */
                      cy="42%"          /* posisi vertikal center donut — kecilkan % = naik */
                      innerRadius={50}  /* jari-jari dalam (lubang tengah). 0 = pie penuh */
                      outerRadius={85}  /* jari-jari luar (ukuran donut). Naikkan = lebih besar */
                      paddingAngle={2}  /* jarak antar slice dalam derajat */
                      dataKey="value" labelLine={false} label={renderPctLabel}>
                      {(stats?.so_status||[]).map((_,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}
                    </Pie>
                    <Tooltip contentStyle={{backgroundColor:darkMode?'#1F2937':'#fff',borderRadius:'8px'}} formatter={(v,n)=>[fmtNum(v),n]}/>
                    <Legend layout="horizontal" align="center" verticalAlign="bottom" iconSize={8} formatter={(v)=><span className="text-xs">{v}</span>}/>
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* ── PIE 2: SO Aging ── */}
            {(() => {
              const agingPieData = [
                { name:'< 30 Hari', value:agingData.reduce((s,v)=>s+(v.less_30||0),0), fill:'#10B981' },
                { name:'30–90 Hari', value:agingData.reduce((s,v)=>s+(v.days_30_90||0),0), fill:'#F59E0B' },
                { name:'90–180 Hari', value:agingData.reduce((s,v)=>s+(v.days_90_180||0),0), fill:'#F97316' },
                { name:'> 180 Hari', value:agingData.reduce((s,v)=>s+(v.more_180||0),0), fill:'#EF4444' },
              ].filter(d=>d.value>0);
              return (
                <div className={`p-5 rounded-2xl shadow ${card} flex flex-col`}>
                  <h3 className={`text-sm font-bold mb-2 flex items-center gap-2 ${txt}`}><Calendar className="w-4 h-4 text-red-500"/> SO Aging (Pie)</h3>
                  {/* minHeight = tinggi area chart dalam pixel — samakan dengan Pie 1 */}
                  <div className="flex-1" style={{minHeight: 300}}>
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart margin={{top: 10, right: 10, bottom: 10, left: 10}}>
                        <Pie
                          data={agingPieData}
                          cx="50%"          /* posisi horizontal center donut */
                          cy="42%"          /* posisi vertikal center donut */
                          innerRadius={50}  /* jari-jari dalam (lubang tengah) */
                          outerRadius={85}  /* jari-jari luar — samakan dengan Pie 1 */
                          paddingAngle={2}
                          dataKey="value" labelLine={false} label={renderPctLabel}>
                          {agingPieData.map((d,i)=><Cell key={i} fill={d.fill}/>)}
                        </Pie>
                        <Tooltip contentStyle={{backgroundColor:darkMode?'#1F2937':'#fff',borderRadius:'8px'}} formatter={(v,n)=>[fmtNum(v)+' SO',n]}/>
                        <Legend layout="horizontal" align="center" verticalAlign="bottom" iconSize={8} formatter={(v)=><span className="text-xs">{v}</span>}/>
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
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
              <tr>{['Vendor (SMRO)','< 30 Hari','30–90 Hari','90–180 Hari','> 180 Hari','Total Open','Sales Amount'].map(h=>(
                <th key={h} className={`p-3 text-center font-semibold ${txt2} first:text-left`}>{h}</th>
              ))}</tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {agingData.slice(0,15).map((v,i)=>(
                <tr key={i} className={trHov}>
                  <td className={`p-3 font-medium text-xs ${txt}`}>{v.vendor}</td>
                  <td className="p-3 text-center font-semibold text-green-600 cursor-pointer hover:underline"
                    onClick={()=>openModal(`Aging Detail: ${v.vendor} — < 30 Hari`, `/api/data/aging-detail/${encodeURIComponent(v.vendor)}?bucket=0-30`)}>{fmtNum(v.less_30)}</td>
                  <td className="p-3 text-center font-semibold text-yellow-600 cursor-pointer hover:underline"
                    onClick={()=>openModal(`Aging Detail: ${v.vendor} — 30-90 Hari`, `/api/data/aging-detail/${encodeURIComponent(v.vendor)}?bucket=30-90`)}>{fmtNum(v.days_30_90)}</td>
                  <td className="p-3 text-center font-semibold text-orange-600 cursor-pointer hover:underline"
                    onClick={()=>openModal(`Aging Detail: ${v.vendor} — 90-180 Hari`, `/api/data/aging-detail/${encodeURIComponent(v.vendor)}?bucket=90-180`)}>{fmtNum(v.days_90_180)}</td>
                  <td className="p-3 text-center font-semibold text-red-600 cursor-pointer hover:underline"
                    onClick={()=>openModal(`Aging Detail: ${v.vendor} — > 180 Hari`, `/api/data/aging-detail/${encodeURIComponent(v.vendor)}?bucket=180%2B`)}>{fmtNum(v.more_180)}</td>
                  <td className="p-3 text-center font-bold text-purple-600 cursor-pointer hover:underline"
                    onClick={()=>openModal(`Aging Detail: ${v.vendor}`, `/api/data/aging-detail/${encodeURIComponent(v.vendor)}`)}>{fmtNum(v.total_open)}</td>
                  <td className="p-3 text-right font-semibold text-orange-600 text-xs">{fmtCurShort(v.sales_amount)}</td>
                </tr>
              ))}
            </tbody>
            <tfoot className={`${tblHd} font-bold text-sm`}>
              {(() => {
                const tot = agingData.reduce((acc,v)=>({
                  less_30:acc.less_30+(v.less_30||0), days_30_90:acc.days_30_90+(v.days_30_90||0),
                  days_90_180:acc.days_90_180+(v.days_90_180||0), more_180:acc.more_180+(v.more_180||0),
                  total_open:acc.total_open+(v.total_open||0), sales_amount:acc.sales_amount+(v.sales_amount||0),
                }), {less_30:0,days_30_90:0,days_90_180:0,more_180:0,total_open:0,sales_amount:0});
                return (
                  <tr>
                    <td className={`p-3 font-bold ${txt}`}>TOTAL</td>
                    <td className="p-3 text-center font-bold text-green-700 cursor-pointer hover:underline"
                      onClick={()=>openModal('TOTAL — < 30 Hari (Semua Vendor)', '/api/data/aging-detail-all?bucket=0-30')}>{fmtNum(tot.less_30)}</td>
                    <td className="p-3 text-center font-bold text-yellow-700 cursor-pointer hover:underline"
                      onClick={()=>openModal('TOTAL — 30-90 Hari (Semua Vendor)', '/api/data/aging-detail-all?bucket=30-90')}>{fmtNum(tot.days_30_90)}</td>
                    <td className="p-3 text-center font-bold text-orange-700 cursor-pointer hover:underline"
                      onClick={()=>openModal('TOTAL — 90-180 Hari (Semua Vendor)', '/api/data/aging-detail-all?bucket=90-180')}>{fmtNum(tot.days_90_180)}</td>
                    <td className="p-3 text-center font-bold text-red-700 cursor-pointer hover:underline"
                      onClick={()=>openModal('TOTAL — > 180 Hari (Semua Vendor)', '/api/data/aging-detail-all?bucket=180%2B')}>{fmtNum(tot.more_180)}</td>
                    <td className="p-3 text-center font-bold text-purple-700 cursor-pointer hover:underline"
                      onClick={()=>openModal('TOTAL — Semua Aging (Semua Vendor)', '/api/data/aging-detail-all')}>{fmtNum(tot.total_open)}</td>
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
        <div className="flex flex-wrap justify-between items-center gap-3 mb-5">
          <div>
            <h2 className={`text-xl font-bold ${txt}`}>All Sales Orders (SO)</h2>
            <p className={`text-sm ${txt2}`}>{fmtNum(soTotal)} total records — halaman {soPage} dari {soTotalPages}</p>
          </div>
          <div className="flex flex-wrap gap-2">
            <label className="flex items-center gap-1 px-3 py-1.5 bg-green-600 hover:bg-green-700 text-white rounded-lg text-sm cursor-pointer">
              <Upload className="w-4 h-4"/>Batch Upload
              <input type="file" accept=".xlsx,.xls" onChange={handleBatchUpload} className="hidden"/>
            </label>
            <button onClick={downloadSOTemplate} className="flex items-center gap-1 px-3 py-1.5 bg-yellow-500 hover:bg-yellow-600 text-white rounded-lg text-sm">
              <FileSpreadsheet className="w-4 h-4"/>Template
            </button>
            <button onClick={downloadSOExcel} className="flex items-center gap-1 px-3 py-1.5 bg-gradient-to-r from-purple-600 to-purple-500 text-white rounded-lg text-sm shadow">
              <Download className="w-4 h-4"/>Download Excel
            </button>
          </div>
        </div>

        {/* Aging Filter Chips */}
        <div className="mb-3 flex flex-wrap gap-2 items-center">
          <span className={`text-xs font-medium ${txt2}`}>Filter Aging:</span>
          {AGING_LABELS.map(label => {
            const active = soFilters.aging.includes(label);
            return (
              <button key={label} onClick={()=>toggleAgingFilter(label)}
                className={`px-3 py-1 rounded-full text-xs font-semibold border transition-all ${active?'text-white border-transparent':'border-gray-300 ' + txt2}`}
                style={active ? {backgroundColor: AGING_COLORS[label], borderColor: AGING_COLORS[label]} : {}}>
                {label} hari
              </button>
            );
          })}
          {soFilters.aging.length > 0 && (
            <button onClick={()=>setSoFilters(f=>({...f,aging:[]}))}
              className={`px-2 py-1 rounded text-xs ${txt2} hover:text-red-500`}>Reset Aging</button>
          )}
        </div>

        {/* Multi-select Filters */}
        {(() => {
          const MultiSelect = ({ label, options, selected, onChange }) => {
            const [open, setOpen] = React.useState(false);
            const ref = React.useRef(null);
            React.useEffect(() => {
              const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
              document.addEventListener('mousedown', handler);
              return () => document.removeEventListener('mousedown', handler);
            }, []);
            const toggle = (val) => onChange(selected.includes(val) ? selected.filter(x=>x!==val) : [...selected, val]);
            return (
              <div className="relative flex-1 min-w-[180px]" ref={ref}>
                <label className={`block text-xs font-medium mb-1 ${txt2}`}>{label}</label>
                <button onClick={()=>setOpen(o=>!o)}
                  className={`w-full px-3 py-2 rounded-lg text-sm border text-left flex justify-between items-center ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300 text-gray-700'}`}>
                  <span className="truncate">
                    {selected.length === 0 ? `All ${label}` : `${selected.length} dipilih`}
                  </span>
                  <ChevronDown className="w-4 h-4 flex-shrink-0 ml-1"/>
                </button>
                {open && (
                  <div className={`absolute z-50 mt-1 w-full max-h-56 overflow-auto rounded-lg shadow-xl border ${darkMode?'bg-gray-700 border-gray-600':'bg-white border-gray-200'}`}>
                    {selected.length > 0 && (
                      <button onClick={()=>onChange([])}
                        className={`w-full px-3 py-2 text-xs text-left text-red-500 hover:bg-red-50 border-b ${darkMode?'border-gray-600 hover:bg-gray-600':'border-gray-100'}`}>
                        ✕ Reset pilihan
                      </button>
                    )}
                    {options.map(opt => (
                      <label key={opt} className={`flex items-center gap-2 px-3 py-2 text-xs cursor-pointer ${darkMode?'hover:bg-gray-600 text-white':'hover:bg-purple-50 text-gray-700'}`}>
                        <input type="checkbox" checked={selected.includes(opt)} onChange={()=>toggle(opt)} className="accent-purple-600"/>
                        <span className="truncate" title={opt}>{opt}</span>
                      </label>
                    ))}
                    {options.length === 0 && <div className={`px-3 py-2 text-xs ${txt2}`}>Tidak ada opsi</div>}
                  </div>
                )}
              </div>
            );
          };
          return (
            <div className={`p-4 rounded-xl mb-4 ${darkMode?'bg-gray-700':'bg-gray-50'}`}>
              <div className="flex flex-wrap gap-3 items-end">
                <MultiSelect label="Operation Unit" options={soFilterOptions.op_units}
                  selected={soFilters.op_units} onChange={v=>setSoFilters(f=>({...f,op_units:v}))}/>
                <MultiSelect label="Vendor Name" options={soFilterOptions.vendors}
                  selected={soFilters.vendors} onChange={v=>setSoFilters(f=>({...f,vendors:v}))}/>
                <MultiSelect label="SO Status" options={soFilterOptions.statuses}
                  selected={soFilters.statuses} onChange={v=>setSoFilters(f=>({...f,statuses:v}))}/>
                <div className="flex-1 min-w-[100px]">
                  <label className={`block text-xs font-medium mb-1 ${txt2}`}>Baris per Halaman</label>
                  <select className={`w-full px-3 py-2 rounded-lg text-sm border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                    value={soPerPage} onChange={e=>setSoPerPage(Number(e.target.value))}>
                    <option value={20}>20</option>
                    <option value={50}>50</option>
                    <option value={100}>100</option>
                    <option value={200}>200</option>
                  </select>
                </div>
                <div className="flex gap-2">
                  <button onClick={()=>{ setSoPage(1); fetchSOData(soFilters,1,soPerPage); }}
                    className="px-5 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-sm font-medium">Apply</button>
                  <button onClick={()=>{ const f={op_units:[],vendors:[],statuses:[],aging:[]}; setSoFilters(f); setSoPage(1); fetchSOData(f,1,soPerPage); }}
                    className={`px-4 py-2 rounded-lg text-sm font-medium ${darkMode?'bg-gray-600 text-gray-200 hover:bg-gray-500':'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>Reset</button>
                </div>
              </div>
              {/* Active filter tags */}
              {(soFilters.op_units.length + soFilters.vendors.length + soFilters.statuses.length) > 0 && (
                <div className="mt-3 flex flex-wrap gap-1.5">
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
          );
        })()}

        {/* SO Table */}
        <div className="overflow-x-auto rounded-lg border border-gray-200">
          <table className="w-full text-sm">
            <thead className={tblHd}>
              <tr>
                {['Aging','SO Item','Item Name','Status','Op Unit','Vendor','Qty',
                  'Sales Price','Sales Amount','PO Price','PO Amount',
                  'Possible Delivery','Plan Date','Remarks'].map(h=>(
                  <th key={h} className={`px-3 py-2.5 text-left font-semibold whitespace-nowrap ${txt2}`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {allSOData.length === 0 ? (
                <tr><td colSpan={14} className={`px-4 py-10 text-center ${txt2}`}>
                  <FileText className="w-10 h-10 mx-auto mb-2 opacity-40"/>Tidak ada data
                </td></tr>
              ) : allSOData.map((so)=>(
                <tr key={so.id} className={`${trHov} transition-colors`}>
                  <td className="px-3 py-2 whitespace-nowrap">
                    <span className={`px-2 py-0.5 rounded-full text-xs font-bold text-white`}
                      style={{backgroundColor: AGING_COLORS[so.aging_label] || '#6B7280'}}>
                      {so.aging_label||'-'}
                    </span>
                  </td>
                  <td className={`px-3 py-2 whitespace-nowrap ${txt2}`}>{so.so_item}</td>
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
                  <td className="px-3 py-2 text-right whitespace-nowrap">{fmtCur(so.sales_price)}</td>
                  <td className="px-3 py-2 text-right font-semibold text-orange-600 whitespace-nowrap">{fmtCur(so.sales_amount)}</td>
                  <td className="px-3 py-2 text-right whitespace-nowrap">{fmtCur(so.purchasing_price)}</td>
                  <td className="px-3 py-2 text-right font-semibold text-green-600 whitespace-nowrap">{fmtCur(so.purchasing_amount)}</td>
                  <td className={`px-3 py-2 text-center text-xs ${txt2}`}>{so.delivery_possible_date||'-'}</td>
                  <td className="px-3 py-2 text-center">
                    {editingCell?.id===so.id && editingCell.field==='delivery_plan_date' ? (
                      <div className="flex items-center gap-1">
                        <input type="date" defaultValue={so.delivery_plan_date}
                          className={`px-2 py-1 rounded text-xs border ${darkMode?'bg-gray-600 border-gray-500 text-white':'bg-white border-gray-300'}`}
                          onChange={e=>setEditValue(e.target.value)}
                          onBlur={()=>updateSOCell(so.id,'delivery_plan_date',editValue)}
                          onKeyDown={e=>{if(e.key==='Enter')updateSOCell(so.id,'delivery_plan_date',editValue);if(e.key==='Escape')setEditingCell(null);}}
                          autoFocus/>
                        <button onClick={()=>updateSOCell(so.id,'delivery_plan_date','')} className="text-red-400 hover:text-red-600 p-0.5"><X className="w-3.5 h-3.5"/></button>
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
              ))}
            </tbody>
          </table>
        </div>

        {/* Pagination */}
        <div className={`mt-4 pt-3 border-t ${darkMode?'border-gray-700':'border-gray-200'} flex justify-between items-center`}>
          <span className={`text-sm ${txt2}`}>
            Menampilkan {((soPage-1)*soPerPage)+1}–{Math.min(soPage*soPerPage,soTotal)} dari {fmtNum(soTotal)}
          </span>
          <div className="flex gap-1 items-center">
            <button disabled={soPage===1} onClick={()=>{ const p=soPage-1; setSoPage(p); fetchSOData(soFilters,p,soPerPage); }}
              className={`p-1.5 rounded ${soPage===1?'opacity-40':'hover:bg-purple-100'}`}><ChevronLeft className="w-4 h-4"/></button>
            <span className={`px-3 py-1 rounded text-sm font-semibold ${darkMode?'bg-gray-700 text-white':'bg-purple-100 text-purple-700'}`}>{soPage}/{soTotalPages}</span>
            <button disabled={soPage===soTotalPages} onClick={()=>{ const p=soPage+1; setSoPage(p); fetchSOData(soFilters,p,soPerPage); }}
              className={`p-1.5 rounded ${soPage===soTotalPages?'opacity-40':'hover:bg-purple-100'}`}><ChevronRight className="w-4 h-4"/></button>
          </div>
        </div>
      </div>

      {/* PO HLI Without SO Table */}
      <div className={`rounded-2xl shadow overflow-hidden ${card}`}>
        <div className={`p-5 border-b ${darkMode?'border-gray-700':'border-gray-100'} flex flex-wrap justify-between items-center gap-3`}>
          <div className="flex items-center gap-2">
            <AlertCircle className="w-5 h-5 text-yellow-600"/>
            <h3 className={`text-base font-bold ${txt}`}>PO HLI yang Belum Ada SO-nya</h3>
            <span className={`text-sm ${txt2}`}>({fmtNum(poWithoutSO.length)} items)</span>
          </div>
          <div className="flex gap-2 items-center">
            <select className={`px-3 py-1.5 rounded-lg text-sm ${darkMode?'bg-gray-700 text-white':'bg-gray-100 text-gray-700'}`}
              value={poPerPage} onChange={e=>{ setPoPerPage(Number(e.target.value)); setPoPage(1); }}>
              <option value={20}>20 Baris</option>
              <option value={50}>50 Baris</option>
              <option value={100}>100 Baris</option>
              <option value={500}>500 Baris</option>
            </select>
            <button onClick={downloadPOExcel} className="flex items-center gap-1 px-4 py-1.5 bg-gradient-to-r from-purple-600 to-purple-500 text-white rounded-lg text-sm shadow">
              <Download className="w-4 h-4"/>Download Excel
            </button>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className={tblHd}>
              <tr>
                {['PO HLI NUMBER','ITEM NO','ITEM CODE','DESCRIPTION','QTY','UNIT','PRICE','AMOUNT','CURRENCY','PO DATE','PURCHASE MEMBER','REQ. DELIVERY','HARI TERSISA'].map(h=>(
                  <th key={h} className={`px-4 py-3 text-left font-semibold whitespace-nowrap ${txt2}`}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className={`divide-y ${tblDv}`}>
              {poRows.length === 0 ? (
                <tr><td colSpan={13} className={`px-4 py-10 text-center ${txt2}`}>
                  <Package className="w-10 h-10 mx-auto mb-2 opacity-40"/>Tidak ada data
                </td></tr>
              ) : poRows.map((row,i)=>{
                const daysLeft = row.days_remaining;
                const daysColor = daysLeft === null ? txt2 : daysLeft < 0 ? 'text-red-600 font-bold' : daysLeft <= 7 ? 'text-orange-600 font-semibold' : daysLeft <= 30 ? 'text-yellow-600' : 'text-green-600';
                return (
                  <tr key={i} className={`${trHov} transition-colors`}>
                    <td className="px-4 py-3 text-purple-600 font-medium whitespace-nowrap">{row.po_no}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.item_no||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.item_code||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} max-w-xs truncate`} title={row.description}>{row.description}</td>
                    <td className={`px-4 py-3 text-right ${txt2}`}>{fmtNum(row.qty)}</td>
                    <td className={`px-4 py-3 ${txt2}`}>{row.unit||'-'}</td>
                    <td className="px-4 py-3 text-right">{fmtCur(row.price)}</td>
                    <td className="px-4 py-3 text-right font-semibold text-orange-600">{fmtCur(row.amount)}</td>
                    <td className={`px-4 py-3 ${txt2}`}>{row.currency||'IDR'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.po_date||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.purchase_member||'-'}</td>
                    <td className={`px-4 py-3 ${txt2} whitespace-nowrap`}>{row.req_delivery||'-'}</td>
                    <td className={`px-4 py-3 text-center whitespace-nowrap ${daysColor}`}>
                      {daysLeft === null ? '-' : daysLeft < 0 ? `${Math.abs(daysLeft)} hari lewat` : `${daysLeft} hari`}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className={`p-4 border-t ${darkMode?'border-gray-700':'border-gray-100'} flex justify-between items-center`}>
          <span className={`text-sm ${txt2}`}>{(poPage-1)*poPerPage+1}–{Math.min(poPage*poPerPage,poWithoutSO.length)} dari {fmtNum(poWithoutSO.length)}</span>
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
      <div className="fixed top-4 right-4 z-[100] flex flex-col gap-2">
        {toasts.map(t=><Toast key={t.id} message={t.message} type={t.type} onClose={()=>removeToast(t.id)}/>)}
      </div>

      {/* Sidebar */}
      <aside className={`fixed left-0 top-0 h-full w-20 flex flex-col items-center py-8 shadow-2xl z-40 ${darkMode?'bg-gray-800 border-r border-gray-700':'bg-gradient-to-b from-purple-600 to-purple-700'}`}>
        <div className="mb-8 p-3 bg-white/20 rounded-2xl"><Package className="w-8 h-8 text-white"/></div>
        <nav className="flex-1 flex flex-col gap-4 w-full px-2">
          <button onClick={()=>setActivePage('dashboard')}
            className={`p-3 rounded-xl flex justify-center transition-all ${activePage==='dashboard'?'bg-white/30 text-white shadow-lg':'text-purple-100 hover:bg-white/20'}`} title="Dashboard">
            <BarChart3 className="w-6 h-6"/>
          </button>
          <button onClick={()=>{ setActivePage('all-so'); setSoPage(1); fetchSOData(soFilters,1,soPerPage); }}
            className={`p-3 rounded-xl flex justify-center transition-all ${activePage==='all-so'?'bg-white/30 text-white shadow-lg':'text-purple-100 hover:bg-white/20'}`} title="All Sales Orders">
            <FileText className="w-6 h-6"/>
          </button>
        </nav>
        <button onClick={()=>setDarkMode(d=>!d)} className="p-3 rounded-xl text-white hover:bg-white/20 transition-all">
          {darkMode?<Sun className="w-6 h-6"/>:<Moon className="w-6 h-6"/>}
        </button>
      </aside>

      {/* Main */}
      <main className="ml-20 p-6">
        <header className="mb-6 flex flex-wrap justify-between items-center gap-4">
          <div>
            <h1 className={`text-2xl font-bold tracking-tight ${txt}`}>
              PO HLI Monitoring <span className="text-purple-600">Dashboard</span>
            </h1>
            <p className={`mt-0.5 text-sm ${txt2}`}>
              {activePage==='dashboard'?'Purchase Orders & Sales Orders Overview':'Manage All Sales Orders & PO Without SO'}
            </p>
          </div>
          <div className="flex gap-3">
            <label className={`flex items-center gap-2 px-4 py-2 rounded-xl cursor-pointer shadow hover:shadow-md transition-all ${darkMode?'bg-purple-600 text-white':'bg-gradient-to-r from-purple-600 to-purple-500 text-white'}`}>
              <Upload className="w-4 h-4"/><span className="text-sm font-medium">Upload PO List</span>
              <input type="file" accept=".xlsx,.xls" onChange={e=>handleUpload(e,'po')} className="hidden"/>
            </label>
            <label className={`flex items-center gap-2 px-4 py-2 rounded-xl cursor-pointer shadow hover:shadow-md transition-all ${darkMode?'bg-blue-600 text-white':'bg-gradient-to-r from-blue-500 to-blue-600 text-white'}`}>
              <Upload className="w-4 h-4"/><span className="text-sm font-medium">Upload SMRO</span>
              <input type="file" accept=".xlsx,.xls" onChange={e=>handleUpload(e,'smro')} className="hidden"/>
            </label>
          </div>
        </header>

        {activePage==='dashboard' ? renderDashboard() : renderAllSO()}
      </main>

      {modal && <SOModal title={modal.title} data={modal.data} darkMode={darkMode} onClose={()=>setModal(null)}/>}

      {uploadProgress && (
        <div className="fixed inset-0 bg-black/60 z-[60] flex items-center justify-center backdrop-blur-sm">
          <div className={`${darkMode?'bg-gray-800':'bg-white'} p-8 rounded-2xl shadow-2xl flex flex-col items-center gap-4 w-80`}>
            <div className="w-14 h-14 border-4 border-purple-600 border-t-transparent rounded-full animate-spin"/>
            <div className="w-full text-center">
              <p className={`font-bold text-lg mb-1 ${txt}`}>Mengupload {uploadProgress.label}...</p>
              <p className={`text-xs mb-3 ${txt2}`}>Mohon tunggu, jangan tutup browser</p>
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
            <p className={`text-sm font-semibold ${txt}`}>Memuat data...</p>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;
