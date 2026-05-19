import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  FileSpreadsheet, Search, TrendingUp, TrendingDown,
  Upload, ArrowUpRight, ArrowDownRight, X,
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
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
    return `${d.getDate()} ${months[d.getMonth()]}`;
  }
  return v;
};
const toIDR = (n) =>
  new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(n || 0);
const toIDRShort = (n) => {
  if (!n) return 'Rp 0';
  if (n >= 1_000_000_000) return `Rp ${(n / 1_000_000_000).toFixed(1)}M`;
  if (n >= 1_000_000) return `Rp ${(n / 1_000_000).toFixed(1)}jt`;
  return toIDR(n);
};

/* ─── accent map ─── */
const STATS = [
  { key: 'fktK', label: 'Faktur Keluaran', Icon: TrendingUp, color: '#2563eb', bg: '#eff6ff', border: '#bfdbfe' },
  { key: 'fktM', label: 'Faktur Masukan', Icon: TrendingDown, color: '#059669', bg: '#ecfdf5', border: '#a7f3d0' },
  { key: 'ppnK', label: 'PPN Keluaran', Icon: ArrowUpRight, color: '#7c3aed', bg: '#f5f3ff', border: '#ddd6fe' },
  { key: 'ppnM', label: 'PPN Masukan', Icon: ArrowDownRight, color: '#d97706', bg: '#fffbeb', border: '#fde68a' },
];

/* ═══════════════════════════════════
   MAIN COMPONENT
═══════════════════════════════════ */
export default function TaxCoreDashboard() {
  const { isMobile, isTablet } = useBreakpoint();
  const [rawData, setRawData] = useState([]);
  const [filterPIC, setFilterPIC] = useState('ALL');
  const [searchTerm, setSearchTerm] = useState('');
  const [isDragging, setIsDragging] = useState(false);

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
            ppnK: cleanNum(nr.ppnkeluaran), ppnM: cleanNum(nr.ppnmasukan),
            pic: nr.pic ? String(nr.pic).trim().toUpperCase() : null,
          };
        })
        .filter((item) => item.pic === 'JIS' || item.pic === 'DAIVA');
      setRawData(cleaned);
    };
    reader.readAsBinaryString(file);
  };

  const filteredData = useMemo(() =>
    rawData.filter((item) =>
      (filterPIC === 'ALL' || item.pic === filterPIC) &&
      item.user.toLowerCase().includes(searchTerm.toLowerCase())
    ), [filterPIC, rawData, searchTerm]);

  const total = useMemo(() =>
    filteredData.reduce(
      (acc, cur) => ({ fktK: acc.fktK + cur.fktK, fktM: acc.fktM + cur.fktM, ppnK: acc.ppnK + cur.ppnK, ppnM: acc.ppnM + cur.ppnM }),
      { fktK: 0, fktM: 0, ppnK: 0, ppnM: 0 }
    ), [filteredData]);

  const gutter = isMobile ? 16 : isTablet ? 24 : 40;

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
        *{box-sizing:border-box}
        .tc-row:hover{background:#f8faff!important}
        .tc-card{transition:transform .18s,box-shadow .18s}
        .tc-card:hover{transform:translateY(-2px);box-shadow:0 8px 28px -6px rgba(0,0,0,.13)!important}
        .tc-upload:hover{background:#1d4ed8!important}
        .tc-search:focus{outline:none;border-color:#93c5fd!important;box-shadow:0 0 0 3px #dbeafe!important}
        ::-webkit-scrollbar{height:4px;width:4px}
        ::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:99px}
      `}</style>

      <div style={{ minHeight: '100vh', background: '#f1f5f9', fontFamily: "'Plus Jakarta Sans',sans-serif", color: '#0f172a' }}>

        {/* NAV */}
        <header style={{ background: '#fff', borderBottom: '1px solid #e2e8f0', position: 'sticky', top: 0, zIndex: 20 }}>
          <div style={{ maxWidth: 1280, margin: '0 auto', padding: `0 ${gutter}px`, height: 56, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>

            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              {/* logo mark */}
              <div style={{ width: 30, height: 30, borderRadius: 8, background: 'linear-gradient(135deg,#2563eb 0%,#7c3aed 100%)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                <svg width="15" height="15" viewBox="0 0 15 15" fill="none">
                  <rect x="1.5" y="1.5" width="5" height="5" rx="1.2" fill="white" opacity=".9" />
                  <rect x="8.5" y="1.5" width="5" height="5" rx="1.2" fill="white" opacity=".5" />
                  <rect x="1.5" y="8.5" width="5" height="5" rx="1.2" fill="white" opacity=".5" />
                  <rect x="8.5" y="8.5" width="5" height="5" rx="1.2" fill="white" opacity=".9" />
                </svg>
              </div>
              <span style={{ fontSize: 15, fontWeight: 800, letterSpacing: '-0.03em' }}>Bacadata</span>
              <span style={{ fontSize: 10, fontWeight: 700, color: '#94a3b8', background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 5, padding: '1px 7px', letterSpacing: '.05em' }}>JIS</span>
            </div>

            <div style={{ display: 'flex', gap: 8 }}>
              {rawData.length > 0 && (
                <button onClick={() => { setRawData([]); setSearchTerm(''); setFilterPIC('ALL'); }}
                  style={{ border: '1px solid #e2e8f0', background: '#fff', borderRadius: 7, padding: '5px 12px', fontSize: 12, fontWeight: 600, color: '#64748b', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4 }}>
                  <X size={11} /> Reset
                </button>
              )}
              <label className="tc-upload" style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '6px 16px', background: '#2563eb', color: '#fff', borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: 'pointer', userSelect: 'none', transition: 'background .15s' }}>
                <Upload size={14} strokeWidth={2.5} />
                {!isMobile && 'Upload Excel'}
                <input type="file" accept=".xlsx,.xls" onChange={(e) => processFile(e.target.files[0])} style={{ display: 'none' }} />
              </label>
            </div>
          </div>
        </header>

        {/* MAIN */}
        <main style={{ maxWidth: 1280, margin: '0 auto', padding: `28px ${gutter}px 80px` }}>

          {rawData.length === 0 ? (
            /* DROP ZONE */
            <label
              onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={(e) => { e.preventDefault(); setIsDragging(false); processFile(e.dataTransfer.files[0]); }}
              style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 440, background: isDragging ? '#eff6ff' : '#fff', border: `2px dashed ${isDragging ? '#93c5fd' : '#e2e8f0'}`, borderRadius: 20, cursor: 'pointer', padding: 40, textAlign: 'center', transition: 'background .15s,border-color .15s' }}
            >
              <div style={{ width: 68, height: 68, borderRadius: 18, background: '#f8fafc', border: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 20 }}>
                <FileSpreadsheet size={30} strokeWidth={1.4} color="#94a3b8" />
              </div>
              <p style={{ fontSize: 18, fontWeight: 800, margin: '0 0 6px', letterSpacing: '-0.03em' }}>Drop file Excel di sini</p>
              <p style={{ fontSize: 13, color: '#64748b', margin: '0 0 24px', lineHeight: 1.6 }}>Atau klik untuk memilih file <strong>.xlsx</strong> / <strong>.xls</strong></p>
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: '#2563eb', color: '#fff', borderRadius: 9, padding: '9px 22px', fontSize: 13, fontWeight: 700 }}>
                <Upload size={14} /> Pilih File
              </span>
              <input type="file" accept=".xlsx,.xls" onChange={(e) => processFile(e.target.files[0])} style={{ display: 'none' }} />
            </label>

          ) : (<>

            {/* STAT CARDS */}
            <div style={{ display: 'grid', gridTemplateColumns: isMobile ? 'repeat(2,1fr)' : 'repeat(4,1fr)', gap: isMobile ? 10 : 14, marginBottom: 24 }}>
              {STATS.map(({ key, label, Icon, color, bg, border }) => (
                <div key={key} className="tc-card" style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 14, padding: isMobile ? '14px 14px' : '18px 20px', boxShadow: '0 1px 4px rgba(0,0,0,.04)' }}>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 14 }}>
                    <span style={{ fontSize: 10, fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '.07em' }}>{label}</span>
                    <div style={{ width: 28, height: 28, borderRadius: 7, background: bg, border: `1px solid ${border}`, display: 'flex', alignItems: 'center', justifyContent: 'center', color, flexShrink: 0 }}>
                      <Icon size={13} strokeWidth={2.5} />
                    </div>
                  </div>
                  <p style={{ margin: 0, fontSize: isMobile ? 15 : 19, fontWeight: 800, color, letterSpacing: '-0.03em', fontVariantNumeric: 'tabular-nums' }}>
                    {isMobile ? toIDRShort(total[key]) : toIDR(total[key])}
                  </p>
                </div>
              ))}
            </div>

            {/* TOOLBAR */}
            <div style={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', gap: 10, marginBottom: 12, alignItems: isMobile ? 'stretch' : 'center', justifyContent: 'space-between' }}>
              {/* segmented */}
              <div style={{ display: 'flex', background: '#e2e8f0', borderRadius: 10, padding: 3, gap: 2, alignSelf: 'flex-start' }}>
                {['ALL', 'JIS', 'DAIVA'].map((p) => (
                  <button key={p} onClick={() => setFilterPIC(p)} style={{
                    padding: '5px 18px', borderRadius: 8, border: 'none',
                    background: filterPIC === p ? '#fff' : 'transparent',
                    fontSize: 12, fontWeight: 700, letterSpacing: '.04em',
                    color: filterPIC === p ? '#0f172a' : '#64748b',
                    cursor: 'pointer', boxShadow: filterPIC === p ? '0 1px 4px rgba(0,0,0,.10)' : 'none',
                    transition: 'all .15s',
                  }}>{p}</button>
                ))}
              </div>
              {/* search */}
              <div style={{ position: 'relative', width: isMobile ? '100%' : 280 }}>
                <Search size={13} strokeWidth={2} style={{ position: 'absolute', left: 11, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8', pointerEvents: 'none' }} />
                <input className="tc-search" type="text" placeholder="Cari user atau perusahaan…" value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  style={{ width: '100%', paddingLeft: 32, paddingRight: searchTerm ? 32 : 12, paddingTop: 8, paddingBottom: 8, fontSize: 13, border: '1px solid #e2e8f0', borderRadius: 8, background: '#fff', color: '#0f172a', transition: 'border .15s,box-shadow .15s' }}
                />
                {searchTerm && (
                  <button onClick={() => setSearchTerm('')} style={{ position: 'absolute', right: 8, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8', display: 'flex', padding: 2 }}>
                    <X size={13} />
                  </button>
                )}
              </div>
            </div>

            {/* PILLS */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 10 }}>
              <span style={{ fontSize: 12, color: '#94a3b8', fontWeight: 600 }}>{filteredData.length} transaksi</span>
              {filterPIC !== 'ALL' && <PICBadge pic={filterPIC} />}
              {searchTerm && <span style={{ fontSize: 11, color: '#64748b', background: '#f1f5f9', border: '1px solid #e2e8f0', borderRadius: 5, padding: '2px 8px' }}>"{searchTerm}"</span>}
            </div>

            {/* TABLE / MOBILE CARDS */}
            {isMobile
              ? <MobileCards data={filteredData} />
              : (
                <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 14, overflow: 'hidden', boxShadow: '0 1px 6px rgba(0,0,0,.05)' }}>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                      <thead>
                        <tr style={{ background: '#fafafa', borderBottom: '1px solid #f1f5f9' }}>
                          <Th>Tanggal</Th>
                          <Th>User / Customer</Th>
                          {STATS.map(s => <Th key={s.key} right accent={s.color}>{s.label}</Th>)}
                          <Th center>PIC</Th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredData.map((row, i) => (
                          <tr key={i} className="tc-row" style={{ borderBottom: '1px solid #f8fafc' }}>
                            <td style={{ padding: '13px 16px', verticalAlign: 'middle' }}>
                              <div style={{ fontWeight: 700, color: '#1e293b', fontSize: 13 }}>{row.tanggal || '—'}</div>
                              <div style={{ fontSize: 11, color: '#94a3b8', fontWeight: 500, marginTop: 2, textTransform: 'capitalize' }}>{row.bulan} {row.tahun}</div>
                            </td>
                            <td style={{ padding: '13px 16px', verticalAlign: 'middle', maxWidth: 220 }}>
                              <div style={{ fontWeight: 700, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{row.user || '—'}</div>
                            </td>
                            {STATS.map(s => (
                              <td key={s.key} style={{ padding: '13px 16px', textAlign: 'right', verticalAlign: 'middle', fontVariantNumeric: 'tabular-nums', fontSize: 12, fontWeight: 600, color: row[s.key] ? s.color : '#e2e8f0', whiteSpace: 'nowrap' }}>
                                {row[s.key] ? toIDR(row[s.key]) : '—'}
                              </td>
                            ))}
                            <td style={{ padding: '13px 16px', textAlign: 'center', verticalAlign: 'middle' }}>
                              <PICBadge pic={row.pic} />
                            </td>
                          </tr>
                        ))}
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

function Th({ children, right, center, accent }) {
  return (
    <th style={{ padding: '10px 16px', fontSize: 11, fontWeight: 700, textAlign: center ? 'center' : right ? 'right' : 'left', color: accent || '#94a3b8', textTransform: 'uppercase', letterSpacing: '.06em', whiteSpace: 'nowrap', userSelect: 'none' }}>
      {children}
    </th>
  );
}

function PICBadge({ pic }) {
  const jis = pic === 'JIS';
  return (
    <span style={{ display: 'inline-block', padding: '3px 9px', borderRadius: 5, fontSize: 10, fontWeight: 800, letterSpacing: '.08em', textTransform: 'uppercase', background: jis ? '#eff6ff' : '#fff7ed', color: jis ? '#1d4ed8' : '#c2410c', border: `1px solid ${jis ? '#bfdbfe' : '#fed7aa'}` }}>
      {pic}
    </span>
  );
}

function MobileCards({ data }) {
  if (!data.length) return <p style={{ textAlign: 'center', color: '#94a3b8', padding: '40px 0', fontSize: 13 }}>Tidak ada data.</p>;
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
      {data.map((row, i) => (
        <div key={i} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 14, padding: '14px 16px', boxShadow: '0 1px 4px rgba(0,0,0,.04)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 12 }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14, color: '#0f172a', marginBottom: 2 }}>{row.user || '—'}</div>
              <div style={{ fontSize: 11, color: '#94a3b8', fontWeight: 500 }}>{row.tanggal} · <span style={{ textTransform: 'capitalize' }}>{row.bulan} {row.tahun}</span></div>
            </div>
            <PICBadge pic={row.pic} />
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
            {STATS.map(({ key, label, color, bg, border }) => (
              <div key={key} style={{ background: row[key] ? bg : '#f8fafc', borderRadius: 9, padding: '9px 11px', border: `1px solid ${row[key] ? border : '#f1f5f9'}` }}>
                <div style={{ fontSize: 10, fontWeight: 700, color: row[key] ? color : '#94a3b8', textTransform: 'uppercase', letterSpacing: '.05em', marginBottom: 3 }}>
                  {label.replace(' Keluaran', '↑').replace(' Masukan', '↓')}
                </div>
                <div style={{ fontSize: 12, fontWeight: 800, color: row[key] ? color : '#cbd5e1', fontVariantNumeric: 'tabular-nums' }}>
                  {row[key] ? toIDRShort(row[key]) : '—'}
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}