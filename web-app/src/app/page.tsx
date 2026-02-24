'use client';
import { useState, useCallback, useRef, useEffect, DragEvent } from 'react';
import { parseExcelFile, exportResults, exportErrors, downloadTemplateLuong1, downloadTemplateLuong2, downloadTemplateGiaHdnd } from '@/lib/excelExport';
import { processLuong1, processLuong2, detectNameColumn, type OutputRow, type Row } from '@/lib/processor';

// ─── Types ────────────────────────────────────────────────────────────────────
type Mode = 'luong1' | 'luong2';
interface RefState { rows: Row[] | null; name: string; source: 'auto' | 'upload' | 'none' }
interface Stats { total: number; matched: number; unmatched: number; ambiguous: number }

// ─── Reference file definitions ───────────────────────────────────────────────
const REF_DEFS = {
    giaHdnd: { label: 'GIA_HDND (NQ HĐND)', icon: '📊', path: '/data/GIA_HDND.xlsx', required: true as const, desc: 'Giá dịch vụ tỉnh/thành' },
    dvktMax: { label: 'DVKT_GIA_MAX', icon: '💹', path: '/data/DVKT_GIA_MAX.xlsx', required: false as const, desc: 'Giá tối đa BYT (TT 23/2024)' },
    quyTrinh: { label: 'QUY_TRINH_DVKT_BYT', icon: '📋', path: '/data/QUY_TRINH_DVKT_BYT.xlsx', required: false as const, desc: 'Danh mục quy trình BYT' },
};

// ─── Auto-fetch helper ────────────────────────────────────────────────────────
async function fetchRefFile(path: string): Promise<Row[] | null> {
    try {
        const res = await fetch(`${path}?v=${Date.now()}`);
        if (!res.ok) return null;
        const blob = await res.blob();
        const file = new File([blob], path.split('/').pop()!);
        return await parseExcelFile(file);
    } catch {
        return null;
    }
}

// ─── Reference file column validation ────────────────────────────────────────
type RefKey = keyof typeof REF_DEFS;

function validateRefFile(key: RefKey, rows: Row[]): string | null {
    if (!rows.length) return '⚠️ File trống hoặc không đọc được.';
    const cols = Object.keys(rows[0]).map(c => c.trim());
    const has = (patterns: string[]) =>
        patterns.some(p => cols.some(c => c.toLowerCase().includes(p.toLowerCase())));

    if (key === 'giaHdnd') {
        // Must have: service name col + at least one price col
        const hasName = has(['Tên dịch vụ kỹ thuật', 'Tên DVKT', 'TEN_DVKT', 'tên']);
        const hasPrice = has(['Mức giá', 'DON_GIA', 'Giá', 'Đơn giá']);
        if (!hasName) return `❌ GIA_HDND thiếu cột Tên dịch vụ.\nCác cột hiện có: ${cols.slice(0, 8).join(', ')}`;
        if (!hasPrice) return `⚠️ GIA_HDND có vẻ thiếu cột giá (“Mức giá” / “DON_GIA”).\nCác cột hiện có: ${cols.slice(0, 8).join(', ')}`;
    }

    if (key === 'dvktMax') {
        const hasName = has(['TEN_DVKT_GIA', 'TEN_DVKT_PHEDUYET', 'Tên DVKT', 'TEN_DVKT']);
        if (!hasName) return `❌ DVKT_GIA_MAX thiếu cột TEN_DVKT_GIA hoặc TEN_DVKT_PHEDUYET.\nCác cột hiện có: ${cols.slice(0, 8).join(', ')}`;
    }

    if (key === 'quyTrinh') {
        const hasName = has(['TEN_DICH_VU', 'Tên dịch vụ', 'Tên kỹ thuật']);
        const hasQt = has(['QUY_TRINH', 'Quy trình']);
        if (!hasName) return `❌ QUY_TRINH_DVKT_BYT thiếu cột TEN_DICH_VU.\nCác cột hiện có: ${cols.slice(0, 8).join(', ')}`;
        if (!hasQt) return `⚠️ QUY_TRINH_DVKT_BYT có vẻ thiếu cột QUY_TRINH.\nCác cột hiện có: ${cols.slice(0, 8).join(', ')}`;
    }

    return null; // valid
}

// ─── Sub-components ───────────────────────────────────────────────────────────
function Dropzone({ label, fileName, onFile }: { label: string; fileName?: string; onFile: (f: File) => void }) {
    const [drag, setDrag] = useState(false);
    const handle = (e: DragEvent, over: boolean) => { e.preventDefault(); setDrag(over); };
    const drop = (e: DragEvent) => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); };
    return (
        <div
            className={`dropzone${drag ? ' hover' : ''}${fileName ? ' loaded' : ''}`}
            onDragEnter={e => handle(e, true)} onDragOver={e => handle(e, true)}
            onDragLeave={e => handle(e, false)} onDrop={drop}
        >
            <input type="file" accept=".xlsx,.xls,.csv" onChange={e => { if (e.target.files?.[0]) { onFile(e.target.files[0]); e.target.value = ''; } }} />
            <div className="dz-icon">{fileName ? '✅' : '📂'}</div>
            {fileName
                ? <div className="dz-name">📄 {fileName}</div>
                : <><div className="dz-label">{label}</div><div className="dz-hint">Kéo thả hoặc click · .xlsx / .xls</div></>}
        </div>
    );
}

interface RefRowProps {
    def: typeof REF_DEFS[keyof typeof REF_DEFS];
    state: RefState;
    visible: boolean;
    onUpload: (f: File) => void;
    refError?: string;
}
function RefFileRow({ def, state, visible, onUpload, refError }: RefRowProps) {
    const inputRef = useRef<HTMLInputElement>(null);
    if (!visible) return null;
    const statusClass = state.rows ? 'ok' : state.source === 'none' ? 'empty' : 'loading';
    const statusText = state.rows
        ? `✓ ${state.rows.length.toLocaleString()} dòng · ${state.source === 'auto' ? 'Hệ thống' : state.name}`
        : state.source === 'none' ? 'Chưa có file' : '➳ Đang tải...';
    return (
        <div className={`ref-row${state.rows && !refError ? ' loaded' : ''}${refError ? ' error' : ''}`}
            style={refError ? { borderColor: 'rgba(240,106,106,.35)', background: 'rgba(240,106,106,.04)' } : {}}>
            <div className="ref-icon">{def.icon}</div>
            <div className="ref-info">
                <div className="ref-name">{def.label} <span className={`ref-badge ${def.required ? 'required' : 'optional'}`}>{def.required ? 'BẮT BUỘC' : 'TÙY CHỌN'}</span></div>
                {refError
                    ? <div className="ref-error">{refError}</div>
                    : <div className={`ref-status ${statusClass}`}>{statusText}</div>
                }
            </div>
            <button className="ref-upload" onClick={() => inputRef.current?.click()}>
                {state.rows ? '🔄' : '📤'}
                <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
                    onChange={e => { if (e.target.files?.[0]) { onUpload(e.target.files[0]); e.target.value = ''; } }} />
            </button>
        </div>
    );
}

// ─── Main Page ────────────────────────────────────────────────────────────────
export default function Home() {
    const [mode, setMode] = useState<Mode>('luong2');
    const [inputFile, setInputFile] = useState<string | null>(null);
    const [inputRows, setInputRows] = useState<Row[] | null>(null);
    const [inputInfo, setInputInfo] = useState<{ cols: string[]; rows: number; nameCol: string } | null>(null);
    const [refs, setRefs] = useState<Record<keyof typeof REF_DEFS, RefState>>({
        giaHdnd: { rows: null, name: '', source: 'none' },
        dvktMax: { rows: null, name: '', source: 'none' },
        quyTrinh: { rows: null, name: '', source: 'none' },
    });
    const [refErrors, setRefErrors] = useState<Partial<Record<RefKey, string>>>({});
    const [threshold, setThreshold] = useState(80);
    const [processing, setProcessing] = useState(false);
    const [progress, setProgress] = useState({ label: '', pct: 0 });
    const [results, setResults] = useState<OutputRow[] | null>(null);
    const [stats, setStats] = useState<Stats | null>(null);
    const [error, setError] = useState<string | null>(null);
    const [success, setSuccess] = useState<string | null>(null);

    // ── Auto-load reference files from /public/data/ on mount ──
    useEffect(() => {
        const load = async () => {
            const keys = Object.keys(REF_DEFS) as Array<keyof typeof REF_DEFS>;
            for (const key of keys) {
                setRefs(prev => ({ ...prev, [key]: { ...prev[key], source: 'loading' as unknown as 'auto' } }));
                const rows = await fetchRefFile(REF_DEFS[key].path);
                if (rows) {
                    setRefs(prev => ({ ...prev, [key]: { rows, name: REF_DEFS[key].path, source: 'auto' } }));
                } else {
                    setRefs(prev => ({ ...prev, [key]: { rows: null, name: '', source: 'none' } }));
                }
            }
        };
        load();
    }, []);

    const handleInput = useCallback(async (f: File) => {
        setInputFile(f.name); setError(null); setResults(null); setStats(null); setSuccess(null); setInputInfo(null);
        try {
            const rows = await parseExcelFile(f);
            if (!rows.length) { setError('❌ File trống hoặc không đọc được.'); return; }
            // Use the same detectNameColumn as processor.ts — handles all underscore/space/diacritic variants
            const nameCol = detectNameColumn(rows);
            if (!nameCol) {
                const allCols = Object.keys(rows[0] ?? {});
                setError(`❌ File đầu vào thiếu cột Tên dịch vụ.\nCác cột hiện có: ${allCols.slice(0, 10).join(', ')}`);
                return;
            }
            const allCols = Object.keys(rows[0] ?? {});
            setInputInfo({ cols: allCols.slice(0, 12), rows: rows.length, nameCol });
            setInputRows(rows);
        } catch { setError('Không đọc được file đầu vào. Kiểm tra định dạng .xlsx/.xls.'); }
    }, []);


    const handleRef = useCallback(async (key: keyof typeof REF_DEFS, f: File) => {
        setError(null);
        setRefErrors(prev => ({ ...prev, [key]: undefined }));
        try {
            const rows = await parseExcelFile(f);
            const err = validateRefFile(key as RefKey, rows);
            if (err) {
                setRefErrors(prev => ({ ...prev, [key]: err }));
                return;  // don't overwrite good existing ref data
            }
            setRefs(prev => ({ ...prev, [key]: { rows, name: f.name, source: 'upload' } }));
            setSuccess(`✅ ${REF_DEFS[key as RefKey].label}: đã tải ${rows.length.toLocaleString()} dòng`);
        } catch { setRefErrors(prev => ({ ...prev, [key]: `Lỗi đọc file: ${f.name} — kiểm tra .xlsx/.xls` })); }
    }, []);

    const canProcess = useCallback(() => {
        if (!inputRows) return false;
        if (mode === 'luong1') return !!refs.quyTrinh.rows;
        return !!refs.giaHdnd.rows;
    }, [inputRows, mode, refs]);

    const handleProcess = useCallback(async () => {
        if (!inputRows || processing) return;
        setProcessing(true); setError(null); setResults(null); setStats(null); setSuccess(null);
        setProgress({ label: 'Đang khởi tạo...', pct: 0 });
        await new Promise(r => setTimeout(r, 30));
        try {
            // Throttle: only update React state when % integer changes to avoid thousands of re-renders
            let lastPct = -1;
            const opts = {
                threshold,
                onProgress: (label: string, pct: number) => {
                    if (pct !== lastPct) {
                        lastPct = pct;
                        setProgress({ label, pct });
                    }
                },
            };
            const result = mode === 'luong1'
                ? await processLuong1({ inputRows, refQuyTrinhRows: refs.quyTrinh.rows!, options: opts })
                : await processLuong2({
                    inputRows,
                    refGiaHdndRows: refs.giaHdnd.rows!,
                    refMaxRows: refs.dvktMax.rows ?? [],
                    refQuyTrinhRows: refs.quyTrinh.rows ?? undefined,
                    options: opts,
                });
            if (result.success) {
                setResults(result.data);
                setStats(result.stats);
                const rate = ((result.stats.matched / Math.max(result.stats.total, 1)) * 100).toFixed(1);
                setSuccess(`Hoàn tất! Khớp ${result.stats.matched}/${result.stats.total} (${rate}%)`);
            } else {
                setError(result.message);
            }
        } catch (e) { setError(`Lỗi: ${e instanceof Error ? e.message : String(e)}`); }
        finally { setProcessing(false); setProgress({ label: '', pct: 0 }); }
    }, [inputRows, mode, refs, threshold, processing]);

    const handleExport = useCallback(() => {
        if (!results) return;
        const base = (inputFile ?? 'ket_qua').replace(/\.[^.]+$/, '');
        exportResults(results, `${base}_${mode === 'luong1' ? 'luong_1' : 'luong_2'}.xlsx`);
    }, [results, inputFile, mode]);

    const handleExportErrors = useCallback(() => {
        if (!results) return;
        const base = (inputFile ?? 'ket_qua').replace(/\.[^.]+$/, '');
        const count = exportErrors(results, `${base}_loi_canh_bao.xlsx`);
        if (count === 0) setSuccess('Không có dòng lỗi — tất cả đều khớp!');
    }, [results, inputFile]);

    const COLS = mode === 'luong1'
        ? ['STT', 'MA_DICH_VU', 'TEN_DICH_VU', 'DON_GIA', 'QUY_TRINH', 'CSKCB_CLS']
        : ['STT', 'MA_TUONG_DUONG', 'TEN_DVKT_GIA', 'TEN_DVKT_PHEDUYET', 'DON_GIA', 'PHAN_LOAI_PTTT', 'CANH_BAO'];

    return (
        <div className="app">
            {/* ── Top Bar ── */}
            <header className="topbar">
                <div>
                    <div className="logo">⚕ Chuẩn hóa Danh mục DVKT</div>
                </div>
                <div className="topbar-right">
                    <div className="mode-toggle">
                        <button className={`mode-btn${mode === 'luong1' ? ' active' : ''}`}
                            onClick={() => { setMode('luong1'); setResults(null); setStats(null); }}>
                            📋 Luồng 1 — QUY TRÌNH
                        </button>
                        <button className={`mode-btn${mode === 'luong2' ? ' active' : ''}`}
                            onClick={() => { setMode('luong2'); setResults(null); setStats(null); }}>
                            💰 Luồng 2 — GIA HDND
                        </button>
                    </div>
                    <div className="version-badge">v2.0</div>
                </div>
            </header>

            <div className="body">
                {/* ──────────── SIDEBAR ──────────── */}
                <aside className="sidebar">
                    {/* Step 1: Input File */}
                    <div className="card">
                        <div className="step-header">
                            <div className={`step-num${inputRows ? ' done' : ''}`}>{inputRows ? '✓' : '1'}</div>
                            <div>
                                <div className="step-title">File đầu vào</div>
                                <div className="step-desc">Danh sách DVKT cần chuẩn hóa</div>
                            </div>
                        </div>
                        <Dropzone label="Chọn file Excel đầu vào" fileName={inputFile ?? undefined} onFile={handleInput} />
                        <button
                            onClick={() => mode === 'luong1' ? downloadTemplateLuong1() : downloadTemplateLuong2()}
                            style={{
                                marginTop: 6, background: 'none', border: 'none',
                                color: 'var(--accent)', cursor: 'pointer', fontSize: 12,
                                padding: '2px 0', textDecoration: 'underline', textUnderlineOffset: 3,
                            }}
                        >
                            📋 Tải file mẫu ({mode === 'luong1' ? 'Luồng 1' : 'Luồng 2 GIA​_HDND'})
                        </button>
                        {inputInfo && (
                            <>
                                <div className="chip-strip">
                                    {inputInfo.cols.map(c => (
                                        <span key={c} className={`chip${c === inputInfo.nameCol ? ' ok' : ' dim'}`}>
                                            {c === inputInfo.nameCol ? '✓ ' : ''}{c}
                                        </span>
                                    ))}
                                </div>
                                <div style={{ fontSize: 10, color: 'var(--text3)', marginTop: 4 }}>
                                    📊 {inputInfo.rows.toLocaleString()} dòng · Cột tên: <strong style={{ color: 'var(--green)' }}>{inputInfo.nameCol}</strong>
                                </div>
                            </>
                        )}
                    </div>

                    {/* Step 2: Reference Files */}
                    <div className="card">
                        <div className="step-header">
                            <div className={`step-num${canProcess() ? ' done' : ''}`}>2</div>
                            <div>
                                <div className="step-title">File tham chiếu</div>
                                <div className="step-desc">Tự động tải từ hệ thống · có thể ghi đè</div>
                            </div>
                        </div>
                        <RefFileRow def={REF_DEFS.giaHdnd} state={refs.giaHdnd} visible={mode === 'luong2'} onUpload={f => handleRef('giaHdnd', f)} refError={refErrors.giaHdnd} />
                        {mode === 'luong2' && (
                            <button onClick={downloadTemplateGiaHdnd} style={{
                                background: 'none', border: 'none', color: 'var(--accent)',
                                cursor: 'pointer', fontSize: 11, padding: '0 0 4px 36px',
                                textDecoration: 'underline', textUnderlineOffset: 3,
                            }}>📥 Tải file mẫu GIA_HDND (⭐ Mã + Tên + Giá)</button>
                        )}
                        <RefFileRow def={REF_DEFS.dvktMax} state={refs.dvktMax} visible={mode === 'luong2'} onUpload={f => handleRef('dvktMax', f)} refError={refErrors.dvktMax} />
                        <RefFileRow def={REF_DEFS.quyTrinh} state={refs.quyTrinh} visible={true} onUpload={f => handleRef('quyTrinh', f)} refError={refErrors.quyTrinh} />
                    </div>

                    {/* Step 3: Options + Process */}
                    <div className="card">
                        <div className="step-header">
                            <div className="step-num">3</div>
                            <div>
                                <div className="step-title">Tùy chọn & Xử lý</div>
                                <div className="step-desc">Điều chỉnh ngưỡng khớp tên</div>
                            </div>
                        </div>
                        <div className="section-lbl">Ngưỡng khớp tên (Fuzzy Threshold)</div>
                        <div className="slider-row">
                            <span className="slider-label">{threshold < 70 ? 'Lỏng' : threshold >= 90 ? 'Chặt' : 'Cân bằng'}</span>
                            <input type="range" min={50} max={99} value={threshold}
                                onChange={e => setThreshold(Number(e.target.value))} />
                            <span className="slider-val">{threshold}%</span>
                        </div>

                        <button className={`process-btn${processing ? ' running' : ''}`} disabled={!canProcess() || processing} onClick={handleProcess}>
                            {processing ? <><div className="spinner" /> Đang xử lý...</> : '▶ BẮT ĐẦU XỬ LÝ'}
                        </button>

                        {error && <div className="alert error"><button className="alert-close" onClick={() => setError(null)}>✕</button>⛔ {error}</div>}
                        {success && !error && <div className="alert success"><button className="alert-close" onClick={() => setSuccess(null)}>✕</button>✅ {success}</div>}
                    </div>
                </aside>

                {/* ──────────── MAIN ──────────── */}
                <main className="main">
                    {/* Stats */}
                    {stats && (
                        <div className="stats-grid">
                            {[
                                { val: stats.total, lbl: 'Tổng bản ghi', cls: 'blue' },
                                { val: stats.matched, lbl: 'Khớp thành công', cls: 'green' },
                                { val: stats.unmatched, lbl: 'Không khớp', cls: 'red' },
                                { val: stats.ambiguous, lbl: 'Cảnh báo', cls: 'yellow' },
                            ].map(s => (
                                <div className="stat-card" key={s.lbl}>
                                    <div className={`stat-val ${s.cls}`}>{s.val.toLocaleString()}</div>
                                    <div className="stat-lbl">{s.lbl}</div>
                                    {s.lbl === 'Khớp thành công' && (
                                        <div className="match-rate">
                                            {((s.val / Math.max(stats.total, 1)) * 100).toFixed(1)}% tỷ lệ
                                        </div>
                                    )}
                                </div>
                            ))}
                        </div>
                    )}

                    {/* Results Table */}
                    <div className="results-panel">
                        <div className="results-header">
                            <div className="results-title">
                                Kết quả xử lý
                                {results && <span className="results-count">({results.length.toLocaleString()} dòng)</span>}
                            </div>
                            {results && (
                                <div style={{ display: 'flex', gap: 8 }}>
                                    <button className="export-btn" onClick={handleExport}>⬇ Tải Excel</button>
                                    <button className="export-btn" onClick={handleExportErrors}
                                        style={{ borderColor: 'rgba(240,106,106,.35)', background: 'rgba(240,106,106,.08)', color: 'var(--red)' }}>
                                        ⚠ Xuất lỗi ({results.filter(r => !r.TEN_DVKT_PHEDUYET || r._highlight).length})
                                    </button>
                                </div>
                            )}
                        </div>

                        {!results ? (
                            <div className="empty">
                                <div className="empty-icon">📊</div>
                                <h3>Chưa có kết quả</h3>
                                <p>Tải file đầu vào và nhấn Bắt Đầu Xử Lý</p>
                            </div>
                        ) : (
                            <>
                                <div className="table-wrap">
                                    <table className="results-table">
                                        <thead>
                                            <tr>{COLS.map(c => <th key={c}>{c}</th>)}</tr>
                                        </thead>
                                        <tbody>
                                            {results.slice(0, 500).map((row, i) => (
                                                <tr key={i} className={row._highlight ? `hl-${row._highlight}` : ''}>
                                                    {COLS.map(col => {
                                                        const v = String(row[col as keyof OutputRow] ?? '');
                                                        let cls = '';
                                                        if (col === 'DON_GIA') cls = v ? 'td-price' : 'td-missing';
                                                        else if (col === 'CANH_BAO') cls = v ? 'td-warn' : '';
                                                        else if (col === 'MA_TUONG_DUONG') cls = v ? 'td-code' : 'td-missing';
                                                        return <td key={col} className={cls} title={v}>{v || (col === 'MA_TUONG_DUONG' || col === 'DON_GIA' ? '—' : v)}</td>;
                                                    })}
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                                {results.length > 500 && (
                                    <div className="legend">
                                        <span style={{ fontSize: 11, color: 'var(--text3)' }}>
                                            Hiển thị 500/{results.length} dòng — Tải Excel để xem đầy đủ
                                        </span>
                                    </div>
                                )}
                                <div className="legend">
                                    <div className="legend-item"><div className="legend-dot" style={{ background: 'rgba(245,166,35,.5)' }} />Sai mã (fallback tên)</div>
                                    <div className="legend-item"><div className="legend-dot" style={{ background: 'rgba(240,106,106,.5)' }} />Nhiều kết quả tương tự</div>
                                </div>
                            </>
                        )}
                    </div>
                </main>
            </div>

            {/* ─── Processing Popup Overlay ─── */}
            {processing && (
                <div className="progress-overlay">
                    <div className="progress-popup">
                        <div className="progress-popup-icon">⚙️</div>
                        <div className="progress-popup-title">Đang xử lý dữ liệu...</div>
                        <div className="progress-popup-file">📄 {inputFile ?? ''}</div>
                        <div className="progress-popup-bar">
                            <div className="progress-popup-fill" style={{ width: `${Math.max(progress.pct, 3)}%` }} />
                        </div>
                        <div className="progress-popup-meta">
                            <div className="progress-popup-pct">{progress.pct === 0 ? '···' : `${progress.pct}%`}</div>
                            <div className="progress-popup-lbl">{progress.label}</div>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
