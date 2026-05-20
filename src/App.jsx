import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  FileSpreadsheet, Search, TrendingUp, TrendingDown,
  Upload, ArrowUpRight, ArrowDownRight, X, CalendarDays, Users,
} from 'lucide-react';

/* ─── responsive hook ─── */
function useBreakpoint() {
  const [w, setW] = useState(typeof window !== 'undefined' ? window.innerWidth : 1200);
  useEffect(() => {
    const fn = () => setW(window.innerWidth);
    window.addEventListener('resize', fn);
    return () => window.removeEventListener('resize', fn);
  }, []);
  return { isMobile: w < 640, isTablet: w < 1024, width: w };
}

/* ─── helpers ─── */
const cleanNum = (v) => {
  if (!v) return 0;
  if (typeof v === 'number') return v;
  return parseInt(String(v).replace(/[^0-9]/g, ''), 10) || 0;
};
const formatExcelDate = (v) => {
  if (!v) return '';
  if (typeof v === 'number') {
    const d = new Date((v - 25569) * 86400 * 1000);
    const months = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
    return `${d.getDate()} ${months[d.getMonth()]}`;
  }
  return v;
};
const toIDR = (n) =>
  new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(n || 0);
const toIDRShort = (n) => {
  if (!n) return 'Rp 0';
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 1_000_000_000) return `${sign}Rp ${(abs / 1_000_000_000).toFixed(2)}M`;
  if (abs >= 1_000_000) return `${sign}Rp ${(abs / 1_000_000).toFixed(1)}jt`;
  return toIDR(n);
};

const STATS = [
  { key: 'fktK', label: 'Faktur Keluaran', short: 'Fkt. Keluar', Icon: TrendingUp,     color: '#2563eb', dark:'#1e40af', bg: 'linear-gradient(135deg,#eff6ff,#dbeafe)', border: '#bfdbfe' },
  { key: 'fktM', label: 'Faktur Masukan',  short: 'Fkt. Masuk',  Icon: TrendingDown,   color: '#059669', dark:'#065f46', bg: 'linear-gradient(135deg,#ecfdf5,#a7f3d0)', border: '#6ee7b7' },
  { key: 'ppnK', label: 'PPN Keluaran',    short: 'PPN Keluar',  Icon: ArrowUpRight,   color: '#7c3aed', dark:'#4c1d95', bg: 'linear-gradient(135deg,#f5f3ff,#ddd6fe)', border: '#c4b5fd' },
  { key: 'ppnM', label: 'PPN Masukan',     short: 'PPN Masuk',   Icon: ArrowDownRight, color: '#d97706', dark:'#92400e', bg: 'linear-gradient(135deg,#fffbeb,#fde68a)', border: '#fcd34d' },
];

/* ═══════════════════ MAIN ═══════════════════ */
export default function TaxCoreDashboard() {
  const { isMobile, isTablet } = useBreakpoint();
  const [rawData, setRawData]         = useState([]);
  const [filterPIC, setFilterPIC]     = useState('ALL');
  const [filterBulan, setFilterBulan] = useState('ALL');
  const [searchTerm, setSearchTerm]   = useState('');
  const [isDragging, setIsDragging]   = useState(false);

  const processFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws);
      let curYear = '', curMonth = '';
      const cleaned = data
        .map((row) => {
          const nr = {};
          Object.keys(row).forEach((k) => { nr[k.trim().toLowerCase().replace(/\s+/g, '')] = row[k]; });
          if (nr.tahun) curYear = nr.tahun;
          if (nr.bulan) curMonth = nr.bulan;
          return {
            tahun: curYear, bulan: curMonth,
            tanggal: formatExcelDate(nr.tanggal),
            user: nr.user || '',
            fktK: cleanNum(nr.fakturkeluaran), fktM: cleanNum(nr.fakturmasukan),
            ppnK: cleanNum(nr.ppnkeluaran),    ppnM: cleanNum(nr.ppnmasukan),
            pic: nr.pic ? String(nr.pic).trim().toUpperCase() : null,
          };
        })
        .filter((item) => item.pic === 'JIS' || item.pic === 'DAIVA');
      setRawData(cleaned);
    };
    reader.readAsBinaryString(file);
  };

  const availableBulan = useMemo(() => {
    const seen = new Set();
    const list = [];
    rawData.forEach((item) => {
      const key = `${item.bulan} ${item.tahun}`;
      if (item.bulan && !seen.has(key)) { seen.add(key); list.push(key); }
    });
    return list;
  }, [rawData]);

  const filteredData = useMemo(() =>
    rawData.filter((item) =>
      (filterPIC === 'ALL' || item.pic === filterPIC) &&
      (filterBulan === 'ALL' || `${item.bulan} ${item.tahun}` === filterBulan) &&
      item.user.toLowerCase().includes(searchTerm.toLowerCase())
    ), [filterPIC, filterBulan, rawData, searchTerm]);

  const total = useMemo(() => {
    const t = filteredData.reduce(
      (acc, cur) => ({ fktK: acc.fktK+cur.fktK, fktM: acc.fktM+cur.fktM, ppnK: acc.ppnK+cur.ppnK, ppnM: acc.ppnM+cur.ppnM }),
      { fktK: 0, fktM: 0, ppnK: 0, ppnM: 0 }
    );
    return { ...t, selisih: t.ppnK - t.ppnM };
  }, [filteredData]);

  const px = isMobile ? '16px' : isTablet ? '28px' : '48px';

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:ital,wght@0,400;0,500;0,600;0,700;0,800&display=swap');
        html,body{margin:0!important;padding:0!important;width:100%!important;max-width:none!important}
        #root,#app,[data-reactroot]{width:100%!important;max-width:none!important;margin:0!important;padding:0!important}
        *{box-sizing:border-box}
        .tc-row:hover td{background:#f0f6ff!important}
        .tc-card{transition:transform .2s cubic-bezier(.34,1.56,.64,1),box-shadow .2s}
        .tc-card:hover{transform:translateY(-3px);box-shadow:0 16px 40px -10px rgba(0,0,0,.15)!important}
        .tc-upload:hover{background:#1d4ed8!important;box-shadow:0 4px 14px rgba(37,99,235,.4)!important}
        .tc-seg-btn{transition:all .15s}
        .tc-seg-btn:hover{color:#0f172a!important}
        .tc-search:focus{outline:none;border-color:#93c5fd!important;box-shadow:0 0 0 3px rgba(147,197,253,.35)!important}
        .tc-select:focus{outline:none;border-color:#93c5fd!important;box-shadow:0 0 0 3px rgba(147,197,253,.35)!important}
        .tc-reset:hover{background:#f1f5f9!important;border-color:#cbd5e1!important}
        ::-webkit-scrollbar{height:5px;width:5px}
        ::-webkit-scrollbar-track{background:transparent}
        ::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:99px}
        .tc-drop-zone{transition:background .15s,border-color .15s,transform .15s}
        .tc-drop-zone:hover{border-color:#93c5fd!important;background:#f8fbff!important}
      `}</style>

      <div style={{ position:'fixed', inset:0, overflowY:'auto', background:'#f1f5f9', fontFamily:"'Plus Jakarta Sans',system-ui,sans-serif", color:'#0f172a', zIndex:0 }}>

        {/* ══ NAV ══ */}
        <header style={{ background:'#fff', borderBottom:'1px solid #e2e8f0', position:'sticky', top:0, zIndex:30, boxShadow:'0 1px 0 #e2e8f0' }}>
          <div style={{ padding:`0 ${px}`, height:60, display:'flex', alignItems:'center', justifyContent:'space-between' }}>
            {/* Brand */}
            <div style={{ display:'flex', alignItems:'center', gap:10 }}>
              <div style={{ width:32, height:32, borderRadius:9, background:'linear-gradient(135deg,#2563eb,#7c3aed)', display:'flex', alignItems:'center', justifyContent:'center', boxShadow:'0 2px 8px rgba(124,58,237,.35)' }}>
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                  <rect x="2" y="2" width="5" height="5" rx="1.5" fill="white" opacity=".95"/>
                  <rect x="9" y="2" width="5" height="5" rx="1.5" fill="white" opacity=".5"/>
                  <rect x="2" y="9" width="5" height="5" rx="1.5" fill="white" opacity=".5"/>
                  <rect x="9" y="9" width="5" height="5" rx="1.5" fill="white" opacity=".95"/>
                </svg>
              </div>
              <div>
                <span style={{ fontSize:15, fontWeight:800, letterSpacing:'-0.03em', lineHeight:1 }}>TaxCore</span>
                {rawData.length > 0 && (
                  <span style={{ marginLeft:8, fontSize:11, fontWeight:600, color:'#94a3b8', background:'#f8fafc', border:'1px solid #e2e8f0', borderRadius:5, padding:'1px 7px', letterSpacing:'.04em', verticalAlign:'middle' }}>
                    {rawData.length} baris
                  </span>
                )}
              </div>
            </div>

            {/* Actions */}
            <div style={{ display:'flex', gap:8, alignItems:'center' }}>
              {rawData.length > 0 && (
                <button className="tc-reset"
                  onClick={() => { setRawData([]); setSearchTerm(''); setFilterPIC('ALL'); setFilterBulan('ALL'); }}
                  style={{ border:'1px solid #e2e8f0', background:'#fff', borderRadius:8, padding:'6px 12px', fontSize:12, fontWeight:700, color:'#64748b', cursor:'pointer', display:'flex', alignItems:'center', gap:5, transition:'all .15s' }}>
                  <X size={12} strokeWidth={2.5}/> Reset
                </button>
              )}
              <label className="tc-upload"
                style={{ display:'flex', alignItems:'center', gap:7, padding:'7px 18px', background:'#2563eb', color:'#fff', borderRadius:9, fontSize:13, fontWeight:700, cursor:'pointer', userSelect:'none', transition:'all .15s', boxShadow:'0 2px 8px rgba(37,99,235,.25)' }}>
                <Upload size={14} strokeWidth={2.5}/>
                {!isMobile && 'Upload Excel'}
                <input type="file" accept=".xlsx,.xls" onChange={(e) => processFile(e.target.files[0])} style={{ display:'none' }}/>
              </label>
            </div>
          </div>
        </header>

        {/* ══ MAIN ══ */}
        <main style={{ padding:`32px ${px} 80px` }}>

          {rawData.length === 0 ? (

            /* ── DROP ZONE ── */
            <label className="tc-drop-zone"
              onDragOver={(e)=>{ e.preventDefault(); setIsDragging(true); }}
              onDragLeave={()=>setIsDragging(false)}
              onDrop={(e)=>{ e.preventDefault(); setIsDragging(false); processFile(e.dataTransfer.files[0]); }}
              style={{
                display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center',
                minHeight:460, background: isDragging ? '#eff6ff' : '#fff',
                border:`2px dashed ${isDragging ? '#93c5fd' : '#e2e8f0'}`,
                borderRadius:20, cursor:'pointer', padding:40, textAlign:'center',
              }}
            >
              <div style={{ width:72, height:72, borderRadius:20, background:'linear-gradient(135deg,#eff6ff,#dbeafe)', border:'1px solid #bfdbfe', display:'flex', alignItems:'center', justifyContent:'center', marginBottom:20 }}>
                <FileSpreadsheet size={32} strokeWidth={1.4} color="#2563eb"/>
              </div>
              <p style={{ fontSize:20, fontWeight:800, margin:'0 0 8px', letterSpacing:'-0.03em' }}>Drop file Excel di sini</p>
              <p style={{ fontSize:13, color:'#64748b', margin:'0 0 28px', lineHeight:1.7 }}>
                Atau klik untuk memilih file <strong style={{color:'#0f172a'}}>.xlsx</strong> / <strong style={{color:'#0f172a'}}>.xls</strong>
              </p>
              <span style={{ display:'inline-flex', alignItems:'center', gap:7, background:'#2563eb', color:'#fff', borderRadius:10, padding:'10px 24px', fontSize:13, fontWeight:700, boxShadow:'0 4px 14px rgba(37,99,235,.3)' }}>
                <Upload size={14}/> Pilih File
              </span>
              <input type="file" accept=".xlsx,.xls" onChange={(e)=>processFile(e.target.files[0])} style={{ display:'none' }}/>
            </label>

          ) : (<>

            {/* ── STAT CARDS ── */}
            <div style={{ display:'grid', gridTemplateColumns: isMobile ? 'repeat(2,1fr)' : 'repeat(5,1fr)', gap: isMobile ? 10 : 16, marginBottom:28 }}>
              {STATS.map(({ key, label, short, Icon, color, dark, bg, border }) => (
                <div key={key} className="tc-card"
                  style={{ background:bg, border:`1px solid ${border}`, borderRadius:16, padding: isMobile ? '14px 14px 16px' : '18px 20px 20px', boxShadow:'0 2px 8px rgba(0,0,0,.05)', position:'relative', overflow:'hidden' }}>
                  <div style={{ position:'absolute', right:-16, top:-16, width:72, height:72, borderRadius:'50%', background:`${color}18` }}/>
                  <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:14, position:'relative' }}>
                    <span style={{ fontSize:10, fontWeight:800, color:dark, textTransform:'uppercase', letterSpacing:'.07em', opacity:.7 }}>{isMobile ? short : label}</span>
                    <div style={{ width:30, height:30, borderRadius:8, background:'rgba(255,255,255,.7)', display:'flex', alignItems:'center', justifyContent:'center', color }}>
                      <Icon size={14} strokeWidth={2.5}/>
                    </div>
                  </div>
                  <p style={{ margin:0, fontSize: isMobile ? 14 : 20, fontWeight:800, color:dark, letterSpacing:'-0.03em', fontVariantNumeric:'tabular-nums', position:'relative' }}>
                    {isMobile ? toIDRShort(total[key]) : toIDR(total[key])}
                  </p>
                </div>
              ))}

              {/* Selisih PPN card */}
              {(() => {
                const s = total.selisih;
                const kurangBayar = s > 0;
                const color  = kurangBayar ? '#dc2626' : '#059669';
                const dark   = kurangBayar ? '#991b1b' : '#065f46';
                const bg     = kurangBayar ? 'linear-gradient(135deg,#fef2f2,#fecaca)' : 'linear-gradient(135deg,#ecfdf5,#a7f3d0)';
                const border = kurangBayar ? '#fca5a5' : '#6ee7b7';
                const label  = kurangBayar ? 'Kurang Bayar' : s < 0 ? 'Lebih Bayar' : 'Selisih PPN';
                const Icon   = kurangBayar ? ArrowUpRight : ArrowDownRight;
                return (
                  <div className="tc-card"
                    style={{ background:bg, border:`1px solid ${border}`, borderRadius:16, padding: isMobile ? '14px 14px 16px' : '18px 20px 20px', boxShadow:'0 2px 8px rgba(0,0,0,.05)', position:'relative', overflow:'hidden', gridColumn: isMobile ? 'span 2' : 'auto' }}>
                    <div style={{ position:'absolute', right:-16, top:-16, width:72, height:72, borderRadius:'50%', background:`${color}18` }}/>
                    <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:14, position:'relative' }}>
                      <span style={{ fontSize:10, fontWeight:800, color:dark, textTransform:'uppercase', letterSpacing:'.07em', opacity:.7 }}>{label}</span>
                      <div style={{ width:30, height:30, borderRadius:8, background:'rgba(255,255,255,.7)', display:'flex', alignItems:'center', justifyContent:'center', color }}>
                        <Icon size={14} strokeWidth={2.5}/>
                      </div>
                    </div>
                    <p style={{ margin:'0 0 4px', fontSize: isMobile ? 14 : 20, fontWeight:800, color:dark, letterSpacing:'-0.03em', fontVariantNumeric:'tabular-nums', position:'relative' }}>
                      {isMobile ? toIDRShort(Math.abs(s)) : toIDR(Math.abs(s))}
                    </p>
                    <p style={{ margin:0, fontSize:10, fontWeight:700, color:dark, opacity:.6, position:'relative' }}>
                      {kurangBayar ? 'PPN Keluaran lebih besar' : s < 0 ? 'PPN Masukan lebih besar' : 'Seimbang'}
                    </p>
                  </div>
                );
              })()}
            </div>

            {/* ── TOOLBAR ── */}
            <div style={{
              display:'flex', flexDirection: isMobile ? 'column' : 'row',
              gap:10, marginBottom:14,
              alignItems: isMobile ? 'stretch' : 'center',
              justifyContent:'space-between',
            }}>
              {/* LEFT group */}
              <div style={{ display:'flex', flexWrap:'wrap', gap:8, alignItems:'center' }}>
                {/* PIC segmented */}
                <div style={{ display:'flex', background:'#e8edf3', borderRadius:10, padding:3, gap:2 }}>
                  {['ALL','JIS','DAIVA'].map((p) => (
                    <button key={p} className="tc-seg-btn" onClick={()=>setFilterPIC(p)} style={{
                      padding:'5px 16px', borderRadius:8, border:'none',
                      background: filterPIC===p ? '#fff' : 'transparent',
                      fontSize:12, fontWeight:700, letterSpacing:'.04em',
                      color: filterPIC===p ? '#0f172a' : '#94a3b8',
                      cursor:'pointer', boxShadow: filterPIC===p ? '0 1px 6px rgba(0,0,0,.10)' : 'none',
                    }}>{p}</button>
                  ))}
                </div>

                {/* Bulan select */}
                <div style={{ position:'relative', display:'flex', alignItems:'center' }}>
                  <CalendarDays size={13} style={{ position:'absolute', left:10, color: filterBulan!=='ALL' ? '#2563eb' : '#94a3b8', pointerEvents:'none', zIndex:1 }}/>
                  <select className="tc-select" value={filterBulan} onChange={(e)=>setFilterBulan(e.target.value)}
                    style={{
                      appearance:'none', WebkitAppearance:'none',
                      paddingLeft:30, paddingRight:28, paddingTop:7, paddingBottom:7,
                      fontSize:12, fontWeight:700,
                      border:`1px solid ${filterBulan!=='ALL' ? '#93c5fd' : '#e2e8f0'}`,
                      borderRadius:9,
                      background: filterBulan!=='ALL' ? '#eff6ff' : '#fff',
                      color: filterBulan!=='ALL' ? '#2563eb' : '#64748b',
                      cursor:'pointer', transition:'all .15s',
                    }}>
                    <option value="ALL">Semua Bulan</option>
                    {availableBulan.map((b) => <option key={b} value={b}>{b}</option>)}
                  </select>
                  <svg width="11" height="11" viewBox="0 0 11 11" fill="none"
                    style={{ position:'absolute', right:9, pointerEvents:'none', color: filterBulan!=='ALL' ? '#2563eb' : '#94a3b8' }}>
                    <path d="M1.5 3.5l4 4 4-4" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </div>

                {/* active bulan chip */}
                {filterBulan !== 'ALL' && (
                  <button onClick={()=>setFilterBulan('ALL')}
                    style={{ display:'flex', alignItems:'center', gap:4, background:'#dbeafe', border:'1px solid #bfdbfe', borderRadius:7, padding:'5px 10px', fontSize:11, fontWeight:700, color:'#1d4ed8', cursor:'pointer' }}>
                    {filterBulan} <X size={10} strokeWidth={2.5}/>
                  </button>
                )}
              </div>

              {/* RIGHT: search */}
              <div style={{ position:'relative', width: isMobile ? '100%' : 270 }}>
                <Search size={13} strokeWidth={2} style={{ position:'absolute', left:11, top:'50%', transform:'translateY(-50%)', color:'#94a3b8', pointerEvents:'none' }}/>
                <input className="tc-search" type="text" placeholder="Cari user atau perusahaan…" value={searchTerm}
                  onChange={(e)=>setSearchTerm(e.target.value)}
                  style={{ width:'100%', paddingLeft:33, paddingRight: searchTerm ? 33 : 12, paddingTop:8, paddingBottom:8, fontSize:13, fontWeight:500, border:'1px solid #e2e8f0', borderRadius:9, background:'#fff', color:'#0f172a', transition:'border .15s,box-shadow .15s' }}
                />
                {searchTerm && (
                  <button onClick={()=>setSearchTerm('')} style={{ position:'absolute', right:8, top:'50%', transform:'translateY(-50%)', background:'none', border:'none', cursor:'pointer', color:'#94a3b8', display:'flex', padding:2 }}>
                    <X size={13}/>
                  </button>
                )}
              </div>
            </div>

            {/* ── ACTIVE FILTERS ROW ── */}
            <div style={{ display:'flex', alignItems:'center', gap:6, marginBottom:12, minHeight:22 }}>
              <span style={{ fontSize:12, color:'#94a3b8', fontWeight:600 }}>{filteredData.length} transaksi</span>
              {filterPIC !== 'ALL' && <PICBadge pic={filterPIC}/>}
              {filterBulan !== 'ALL' && (
                <span style={{ fontSize:11, fontWeight:700, color:'#2563eb', background:'#eff6ff', border:'1px solid #bfdbfe', borderRadius:5, padding:'2px 8px' }}>
                  {filterBulan}
                </span>
              )}
              {searchTerm && (
                <span style={{ fontSize:11, color:'#475569', background:'#f1f5f9', border:'1px solid #e2e8f0', borderRadius:5, padding:'2px 8px', fontStyle:'italic' }}>
                  "{searchTerm}"
                </span>
              )}
            </div>

            {/* ── TABLE / MOBILE CARDS ── */}
            {isMobile
              ? <MobileCards data={filteredData}/>
              : (
                <div style={{ background:'#fff', border:'1px solid #e2e8f0', borderRadius:16, overflow:'hidden', boxShadow:'0 2px 12px rgba(0,0,0,.06)' }}>
                  <div style={{ overflowX:'auto' }}>
                    <table style={{ width:'100%', borderCollapse:'collapse', fontSize:13 }}>
                      <thead>
                        <tr style={{ borderBottom:'2px solid #f1f5f9' }}>
                          <Th left>Tanggal</Th>
                          <Th left>User / Customer</Th>
                          {STATS.map(s => <Th key={s.key} right accent={s.color}>{s.label}</Th>)}
                          <Th right accent="#64748b">Selisih PPN</Th>
                          <Th center>PIC</Th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredData.map((row, i) => {
                          const selisih = row.ppnK - row.ppnM;
                          const kurang = selisih > 0;
                          return (
                          <tr key={i} className="tc-row">
                            <td style={{ padding:'14px 18px', verticalAlign:'middle', borderBottom:'1px solid #f8fafc', background:'inherit' }}>
                              <div style={{ fontWeight:700, color:'#1e293b', fontSize:13, letterSpacing:'-0.01em' }}>{row.tanggal||'—'}</div>
                              <div style={{ fontSize:10, color:'#94a3b8', fontWeight:600, marginTop:3, textTransform:'capitalize', letterSpacing:'.02em' }}>{row.bulan} {row.tahun}</div>
                            </td>
                            <td style={{ padding:'14px 18px', verticalAlign:'middle', borderBottom:'1px solid #f8fafc', background:'inherit', maxWidth:240 }}>
                              <div style={{ fontWeight:700, color:'#0f172a', whiteSpace:'nowrap', overflow:'hidden', textOverflow:'ellipsis', fontSize:13 }}>{row.user||'—'}</div>
                            </td>
                            {STATS.map(s => (
                              <td key={s.key} style={{ padding:'14px 18px', textAlign:'right', verticalAlign:'middle', borderBottom:'1px solid #f8fafc', background:'inherit', fontVariantNumeric:'tabular-nums', fontSize:12, fontWeight:700, color: row[s.key] ? s.color : '#e2e8f0', whiteSpace:'nowrap', letterSpacing:'-0.01em' }}>
                                {row[s.key] ? toIDR(row[s.key]) : '—'}
                              </td>
                            ))}
                            <td style={{ padding:'14px 18px', textAlign:'right', verticalAlign:'middle', borderBottom:'1px solid #f8fafc', background:'inherit', whiteSpace:'nowrap' }}>
                              {selisih !== 0 ? (
                                <span style={{ display:'inline-flex', alignItems:'center', gap:4, padding:'3px 8px', borderRadius:6, fontSize:11, fontWeight:800, fontVariantNumeric:'tabular-nums', background: kurang ? '#fef2f2' : '#ecfdf5', color: kurang ? '#dc2626' : '#059669', border:`1px solid ${kurang ? '#fca5a5' : '#6ee7b7'}` }}>
                                  {kurang ? <ArrowUpRight size={11}/> : <ArrowDownRight size={11}/>}
                                  {toIDR(Math.abs(selisih))}
                                </span>
                              ) : <span style={{ color:'#e2e8f0' }}>—</span>}
                            </td>
                            <td style={{ padding:'14px 18px', textAlign:'center', verticalAlign:'middle', borderBottom:'1px solid #f8fafc', background:'inherit' }}>
                              <PICBadge pic={row.pic}/>
                            </td>
                          </tr>
                          );
                        })}
                        {filteredData.length === 0 && (
                          <tr>
                            <td colSpan={8} style={{ padding:'48px 0', textAlign:'center', color:'#94a3b8', fontSize:13, fontWeight:600 }}>
                              Tidak ada data yang cocok.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )
            }
          </>)}
        </main>
      </div>
    </>
  );
}

/* ─── sub-components ─── */

function Th({ children, left, right, center, accent, sticky }) {
  return (
    <th style={{
      padding:'11px 18px', fontSize:10, fontWeight:800,
      textAlign: center ? 'center' : right ? 'right' : 'left',
      color: accent || '#94a3b8',
      textTransform:'uppercase', letterSpacing:'.07em',
      whiteSpace:'nowrap', userSelect:'none',
      background:'#fafbfc',
    }}>
      {children}
    </th>
  );
}

function PICBadge({ pic }) {
  const jis = pic === 'JIS';
  return (
    <span style={{
      display:'inline-block', padding:'3px 10px', borderRadius:6,
      fontSize:10, fontWeight:800, letterSpacing:'.08em', textTransform:'uppercase',
      background: jis ? '#eff6ff' : '#fff7ed',
      color:      jis ? '#1d4ed8' : '#c2410c',
      border:    `1px solid ${jis ? '#bfdbfe' : '#fed7aa'}`,
    }}>{pic}</span>
  );
}

function MobileCards({ data }) {
  if (!data.length) return (
    <div style={{ textAlign:'center', padding:'60px 0', color:'#94a3b8', fontSize:13, fontWeight:600 }}>
      Tidak ada data yang cocok.
    </div>
  );
  return (
    <div style={{ display:'flex', flexDirection:'column', gap:10 }}>
      {data.map((row, i) => (
        <div key={i} style={{ background:'#fff', border:'1px solid #e2e8f0', borderRadius:14, padding:'14px 16px', boxShadow:'0 1px 4px rgba(0,0,0,.04)' }}>
          <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:12 }}>
            <div>
              <div style={{ fontWeight:800, fontSize:14, color:'#0f172a', marginBottom:3 }}>{row.user||'—'}</div>
              <div style={{ fontSize:11, color:'#94a3b8', fontWeight:500 }}>{row.tanggal} · <span style={{ textTransform:'capitalize' }}>{row.bulan} {row.tahun}</span></div>
            </div>
            <PICBadge pic={row.pic}/>
          </div>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:8 }}>
            {STATS.map(({ key, short, color, bg, border, dark }) => (
              <div key={key} style={{ background: row[key] ? bg : '#f8fafc', borderRadius:10, padding:'10px 12px', border:`1px solid ${row[key] ? border : '#f1f5f9'}` }}>
                <div style={{ fontSize:9, fontWeight:800, color: row[key] ? dark : '#94a3b8', textTransform:'uppercase', letterSpacing:'.07em', marginBottom:4, opacity: row[key] ? .7 : 1 }}>
                  {short}
                </div>
                <div style={{ fontSize:12, fontWeight:800, color: row[key] ? dark : '#cbd5e1', fontVariantNumeric:'tabular-nums', letterSpacing:'-0.01em' }}>
                  {row[key] ? toIDRShort(row[key]) : '—'}
                </div>
              </div>
            ))}
            {/* Selisih PPN */}
            {(() => {
              const s = row.ppnK - row.ppnM;
              const kurang = s > 0;
              const color = kurang ? '#dc2626' : '#059669';
              const dark  = kurang ? '#991b1b' : '#065f46';
              const bg    = kurang ? 'linear-gradient(135deg,#fef2f2,#fecaca)' : 'linear-gradient(135deg,#ecfdf5,#a7f3d0)';
              const border= kurang ? '#fca5a5' : '#6ee7b7';
              return (
                <div style={{ background: s !== 0 ? bg : '#f8fafc', borderRadius:10, padding:'10px 12px', border:`1px solid ${s !== 0 ? border : '#f1f5f9'}`, gridColumn:'span 2' }}>
                  <div style={{ fontSize:9, fontWeight:800, color: s !== 0 ? dark : '#94a3b8', textTransform:'uppercase', letterSpacing:'.07em', marginBottom:4, opacity: s !== 0 ? .7 : 1 }}>
                    {kurang ? 'Kurang Bayar' : s < 0 ? 'Lebih Bayar' : 'Selisih PPN'}
                  </div>
                  <div style={{ display:'flex', alignItems:'center', gap:4 }}>
                    {s !== 0 && (kurang ? <ArrowUpRight size={12} color={color}/> : <ArrowDownRight size={12} color={color}/>)}
                    <span style={{ fontSize:12, fontWeight:800, color: s !== 0 ? dark : '#cbd5e1', fontVariantNumeric:'tabular-nums', letterSpacing:'-0.01em' }}>
                      {s !== 0 ? toIDRShort(Math.abs(s)) : '—'}
                    </span>
                  </div>
                </div>
              );
            })()}
          </div>
        </div>
      ))}
    </div>
  );
}