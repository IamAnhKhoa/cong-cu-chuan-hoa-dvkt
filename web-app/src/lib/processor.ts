/**
 * processor.ts — Async TypeScript port of processor.py
 * Logic is matched exactly to the Python version (Luồng 1 + Luồng 2).
 */

// eslint-disable-next-line @typescript-eslint/no-require-imports
const fuzzball = require('fuzzball');

export type Row = Record<string, string | number | null | undefined>;

export interface ProcessingOptions {
    threshold?: number;
    onProgress?: (msg: string, pct: number) => void;
}

export interface ProcessingResult {
    success: boolean;
    message: string;
    data: OutputRow[];
    stats: { total: number; matched: number; unmatched: number; ambiguous: number };
}

export interface OutputRow {
    STT: number;
    // ── Luồng 1 columns (Python: process_luong1) ─────────────────────────────
    MA_DICH_VU?: string;          // Code from QUY_TRINH ref
    TEN_DICH_VU?: string;         // Standardized name from QUY_TRINH ref
    // ── Luồng 2 columns (Python: process_gia_hdnd_file) ─────────────────────
    MA_TUONG_DUONG?: string;
    TEN_DVKT_PHEDUYET?: string;
    TEN_DVKT_GIA?: string;
    PHAN_LOAI_PTTT?: string;
    // ── Shared columns ────────────────────────────────────────────────────────
    DON_GIA?: string | number;
    GHI_CHU?: string;
    QUYET_DINH?: string;
    QUY_TRINH?: string;
    TU_NGAY?: string;
    DEN_NGAY?: string;
    CSKCB_CGKT?: string;
    CSKCB_CLS?: string;           // Luồng 1: warning goes here; Luồng 2: from QUY_TRINH lookup
    CANH_BAO?: string;            // Luồng 2 only
    // Internal markers — stripped before export
    _highlight?: 'yellow' | null;
    _matchedName?: string;
    _score?: number;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

export function getColVal(row: Row, col: string): string {
    if (!col) return '';
    const v = row[col];
    if (v === null || v === undefined) return '';
    return String(v).trim();
}

/** Yield to event loop — keeps the UI responsive during large processing */
const yield_ = () => new Promise<void>(r => setTimeout(r, 0));

// ─── Column Detection (mirrors Python detect_code_column / possible_headers) ───

const CODE_COLUMN_NAMES = [
    'MA_DICH_VU', 'MA_TUONG_DUONG', 'MA_KY_THUAT', 'MA_DVKT',
    'mã dịch vụ', 'mã kỹ thuật', 'mã dvkt', 'mã dịch vụ kỹ thuật',
    'ma dich vu', 'ma ky thuat', 'ma dvkt', 'mã tương đương',
    'ma tuong duong', 'service code', 'code', 'ma',
];

const CODE_VALUE_REGEX = /^\d{2,}\.\d{4}\.\d{4}$|^\w{2,}\.\d+/;

export function detectCodeColumn(rows: Row[]): string | null {
    if (!rows?.length) return null;
    const cols = Object.keys(rows[0]);
    const lower = CODE_COLUMN_NAMES.map(c => c.toLowerCase());

    for (const col of cols) {
        const c = col.trim().toLowerCase();
        if (lower.includes(c)) return col;
    }
    for (const col of cols) {
        const c = col.trim().toLowerCase();
        for (const pat of lower) {
            if (c.includes(pat) || pat.includes(c)) return col;
        }
    }
    // Auto-detect by value pattern
    for (const col of cols.slice(0, 6)) {
        const sample = rows.slice(0, 20).map(r => String(r[col] ?? '').trim()).filter(Boolean);
        const hits = sample.filter(v => CODE_VALUE_REGEX.test(v));
        if (hits.length >= Math.max(1, Math.floor(sample.length * 0.3))) return col;
    }
    return null;
}

/** Mirror Python's possible_headers list exactly */
export function detectNameColumn(rows: Row[]): string | null {
    if (!rows?.length) return null;
    const cols = Object.keys(rows[0]);
    const exactPatterns = [
        'tên kỹ thuật', 'tên dịch vụ', 'tên dvkt', 'ten dich vu', 'ten ky thuat', 'tên',
        'ten_dvkt_pheduyet', 'ten_dvkt_gia', 'tên dịch vụ kỹ thuật', 'ten dich vu ky thuat',
        'ten_dich_vu', 'tên dịch vụ kỹ thuật',
    ];
    for (const col of cols) {
        if (exactPatterns.includes(col.trim().toLowerCase())) return col;
    }
    for (const col of cols) {
        const c = col.trim().toLowerCase();
        if (c.includes('ten_dvkt') || c.includes('dvkt_pheduyet')) return col;
    }
    for (const col of cols) {
        const c = col.trim().toLowerCase();
        if ((c.includes('tên') || c.includes('ten')) &&
            (c.includes('thuật') || c.includes('dịch vụ') || c.includes('dvkt') || c.includes('ky thuat'))) {
            return col;
        }
    }
    return null;
}

function detectColumn(rows: Row[], candidates: string[]): string | null {
    if (!rows?.length) return null;
    const cols = Object.keys(rows[0]);
    const lower = candidates.map(c => c.toLowerCase());
    for (const col of cols) {
        if (lower.includes(col.trim().toLowerCase())) return col;
    }
    for (const col of cols) {
        const c = col.trim().toLowerCase();
        for (const p of lower) {
            if (c.includes(p) || p.includes(c)) return col;
        }
    }
    return null;
}

// ─── Fuzzy Matching (mirrors Python find_best_match exactly) ─────────────────

export interface MatchResult {
    name: string;
    score: number;
    ambiguous?: boolean;
    choices?: string[];  // all same-score options when ambiguous
}

/**
 * Python equivalent:
 *   results = process.extract(query, choices, scorer=token_sort_ratio, limit=5)
 *   high_score = [(n, s) for n, s in results if s >= 95]
 *   if len(high_score) > 1: return ('AMBIGUOUS', [n for n, s in high_score])
 */
export function findBestMatch(query: string, choices: string[], threshold = 80): MatchResult | null {
    if (!query || !choices.length) return null;

    const queryNorm = query.trim().toLowerCase();

    // STEP 1: Exact match (Python checks exact first)
    const exactMatches = choices.filter(c => c.trim().toLowerCase() === queryNorm);
    if (exactMatches.length > 1) return { name: exactMatches[0], score: 100, ambiguous: true, choices: exactMatches };
    if (exactMatches.length === 1) return { name: exactMatches[0], score: 100 };

    // STEP 2: fuzzball.extract — limit=5, cutoff=threshold
    const results: Array<[string, number]> = fuzzball.extract(query, choices, {
        scorer: fuzzball.token_sort_ratio,
        limit: 5,
        cutoff: threshold,
    });
    if (!results || results.length === 0) return null;

    // AMBIGUOUS: multiple results with score >= 95 (same as Python)
    const highScore = results.filter(([, s]) => s >= 95);
    if (highScore.length > 1) {
        return {
            name: highScore[0][0],
            score: highScore[0][1],
            ambiguous: true,
            choices: highScore.map(([n]) => n),
        };
    }

    const [bestName, bestScore] = results[0];
    if (bestScore >= threshold) return { name: bestName, score: bestScore };
    return null;
}

// ─── Pre-built indexes for O(1) code lookup ───────────────────────────────────

interface RefIndex {
    nameToRow: Map<string, Row>;
    codeToRow: Map<string, Row>;
    nameList: string[];
}

function buildIndex(rows: Row[], nameCol: string, codeCol: string | null): RefIndex {
    const nameToRow = new Map<string, Row>();
    const codeToRow = new Map<string, Row>();
    const nameList: string[] = [];
    for (const row of rows) {
        const name = getColVal(row, nameCol);
        if (name && !nameToRow.has(name)) { nameToRow.set(name, row); nameList.push(name); }
        if (codeCol) {
            const code = getColVal(row, codeCol);
            if (code && !codeToRow.has(code)) codeToRow.set(code, row);
        }
    }
    return { nameToRow, codeToRow, nameList };
}

// ─── Luồng 1: QUY_TRINH matching ─────────────────────────────────────────────

export interface ProcessLuong1Params {
    inputRows: Row[];
    refQuyTrinhRows: Row[];
    options?: ProcessingOptions;
}

export async function processLuong1({ inputRows, refQuyTrinhRows, options }: ProcessLuong1Params): Promise<ProcessingResult> {
    const threshold = options?.threshold ?? 80;
    const onProgress = options?.onProgress;

    // Detect reference columns (Python: prefer TEN_DICH_VU, then detectNameColumn)
    let refNameCol: string | null = null;
    for (const col of Object.keys(refQuyTrinhRows[0] ?? {})) {
        if (col.includes('TEN_DICH_VU') || (col.toLowerCase().includes('tên') && col.toLowerCase().includes('vụ'))) {
            refNameCol = col; break;
        }
    }
    if (!refNameCol) refNameCol = detectNameColumn(refQuyTrinhRows);

    let refCodeCol: string | null = null;
    for (const col of Object.keys(refQuyTrinhRows[0] ?? {})) {
        if (col.includes('MA_DICH_VU') || (col.toLowerCase().includes('mã') && col.toLowerCase().includes('vụ'))) {
            refCodeCol = col; break;
        }
    }

    if (!refNameCol) return {
        success: false,
        message: 'Không tìm thấy cột tên trong file QUY_TRINH.',
        data: [], stats: { total: 0, matched: 0, unmatched: 0, ambiguous: 0 },
    };

    const refQuyTrinhCol = detectColumn(refQuyTrinhRows, ['QUY_TRINH', 'quy_trinh', 'quy trinh']);
    const inputNameCol = detectNameColumn(inputRows);
    const inputCodeCol = detectCodeColumn(inputRows);

    if (!inputNameCol) return {
        success: false,
        message: `Không tìm thấy cột tên dịch vụ. Các cột: ${Object.keys(inputRows[0] ?? {}).join(', ')}`,
        data: [], stats: { total: 0, matched: 0, unmatched: 0, ambiguous: 0 },
    };

    const refIdx = buildIndex(refQuyTrinhRows, refNameCol, refCodeCol);
    const stats = { total: 0, matched: 0, unmatched: 0, ambiguous: 0 };
    const results: OutputRow[] = [];
    let stt = 1;

    for (let i = 0; i < inputRows.length; i++) {
        const row = inputRows[i];
        const rawName = getColVal(row, inputNameCol);
        if (!rawName) continue;
        stats.total++;

        let matchedByCode = false;
        let matchResult: MatchResult | null = null;
        let inputCodeValue = '';
        let codeNotFound = false;

        // Priority 1: code matching (same as Python)
        if (inputCodeCol) {
            inputCodeValue = getColVal(row, inputCodeCol);
            if (inputCodeValue) {
                const refRow = refIdx.codeToRow.get(inputCodeValue);
                if (refRow) {
                    matchedByCode = true;
                    matchResult = { name: getColVal(refRow, refNameCol!), score: 100 };
                } else {
                    codeNotFound = true;
                }
            }
        }

        // Priority 2: fuzzy name (only if code not matched)
        if (!matchedByCode) matchResult = findBestMatch(rawName, refIdx.nameList, threshold);

        let quyTrinh = '', matchedCode = '', donGia = '';
        let warning = '';
        let matchedNameDisplay = 'KHÔNG TÌM THẤY';
        let tenDichVu = rawName; // Default: use input name (Python: unmatched case)

        if (matchResult) {
            if (matchResult.ambiguous) {
                // Python: "⚠️ CÓ N DỊCH VỤ TƯƠNG TỰ - XEM KỸ" in CSKCB_CLS (last column)
                const n = matchResult.choices?.length ?? 2;
                warning = `⚠️ CÓ ${n} DỊCH VỤ TƯƠNG TỰ - XEM KỸ`;
                // Fill data from first match (Python does this)
                const firstRefRow = refIdx.nameToRow.get(matchResult.name);
                if (firstRefRow) {
                    if (refQuyTrinhCol) quyTrinh = getColVal(firstRefRow, refQuyTrinhCol);
                    if (refCodeCol) matchedCode = getColVal(firstRefRow, refCodeCol);
                    donGia = getColVal(firstRefRow, 'DON_GIA');
                    tenDichVu = matchResult.name; // Use matched ref name
                }
                matchedNameDisplay = `${matchResult.name} (⚠️ ${n} tương tự)`;
                stats.ambiguous++;
            } else {
                const refRow = refIdx.nameToRow.get(matchResult.name);
                if (refRow) {
                    if (refQuyTrinhCol) quyTrinh = getColVal(refRow, refQuyTrinhCol);
                    if (refCodeCol) matchedCode = getColVal(refRow, refCodeCol);
                    donGia = getColVal(refRow, 'DON_GIA');
                    tenDichVu = matchResult.name; // Use ref standardized name
                }
                matchedNameDisplay = matchResult.name;
                stats.matched++;
            }
        } else {
            stats.unmatched++;
        }

        // Code mismatch warning (Luồng 1 also tracks this)
        if (codeNotFound && matchedNameDisplay !== 'KHÔNG TÌM THẤY') {
            const w = `⚠️ Sai mã | Mã nhập: ${inputCodeValue}${matchedCode ? ` | Mã theo tên: ${matchedCode}` : ''}`;
            warning = warning ? `${warning} | ${w}` : w;
        }

        results.push({
            STT: stt++,
            // ── Luồng 1 exact Python columns ─
            MA_DICH_VU: matchedCode,                   // from QUY_TRINH ref (MA_DICH_VU)
            TEN_DICH_VU: tenDichVu,                    // ref standardized name OR input name
            DON_GIA: donGia,                           // from QUY_TRINH ref (usually empty)
            QUY_TRINH: quyTrinh,
            CSKCB_CGKT: '',                            // Python: always '' for Luồng 1
            CSKCB_CLS: warning,                        // Python: warning goes HERE (last col)
            // Fill Luồng 2 fields as empty (for TypeScript type compatibility)
            MA_TUONG_DUONG: '',
            TEN_DVKT_PHEDUYET: '',
            TEN_DVKT_GIA: '',
            PHAN_LOAI_PTTT: '',
            GHI_CHU: '',
            QUYET_DINH: '',
            TU_NGAY: '',
            DEN_NGAY: '',
            CANH_BAO: '',
            // Python: yellow if ⚠️ in last column (CSKCB_CLS)
            _highlight: warning ? 'yellow' : null,
            _matchedName: matchedNameDisplay,
            _score: matchResult?.score,
        });

        // Per-row progress (Python does this every row; we yield every 80 for performance)
        const pct = Math.round(((i + 1) / inputRows.length) * 100);
        if (warning) {
            onProgress?.(`Đã xử lý ${i + 1}/${inputRows.length}: Cảnh báo - ${rawName.slice(0, 50)}...`, pct);
        } else if (matchResult) {
            onProgress?.(`Đã xử lý ${i + 1}/${inputRows.length}: Khớp - ${rawName.slice(0, 50)}...`, pct);
        } else {
            onProgress?.(`Đã xử lý ${i + 1}/${inputRows.length}: Không khớp - ${rawName.slice(0, 50)}...`, pct);
        }

        // Yield every 10 rows OR on last row so React can re-render progress popup
        if (i % 10 === 9 || i === inputRows.length - 1) await yield_();
    }

    return { success: true, message: 'Luồng 1 hoàn tất.', data: results, stats };
}

// ─── Luồng 2: GIA_HDND matching ──────────────────────────────────────────────

export interface ProcessLuong2Params {
    inputRows: Row[];
    refGiaHdndRows: Row[];
    refMaxRows: Row[];
    refQuyTrinhRows?: Row[];
    options?: ProcessingOptions;
}

export async function processLuong2({ inputRows, refGiaHdndRows, refMaxRows, refQuyTrinhRows, options }: ProcessLuong2Params): Promise<ProcessingResult> {
    const threshold = options?.threshold ?? 80;
    const onProgress = options?.onProgress;

    // ── Detect GIA_HDND columns (Python: prefers "Tên dịch vụ kỹ thuật" or TT23+Tên)
    let refGiaNameCol: string | null = null;
    for (const col of Object.keys(refGiaHdndRows[0] ?? {})) {
        if (col.includes('Tên dịch vụ kỹ thuật') || (col.includes('TT23') && col.includes('Tên'))) {
            refGiaNameCol = col; break;
        }
    }
    if (!refGiaNameCol) refGiaNameCol = detectNameColumn(refGiaHdndRows);
    if (!refGiaNameCol) return { success: false, message: 'Không tìm thấy cột tên trong GIA_HDND.', data: [], stats: { total: 0, matched: 0, unmatched: 0, ambiguous: 0 } };

    // Python: detect_code_column with preference for MA_TUONG_DUONG, 'mã tương đương'
    const refGiaCodeCol = detectColumn(refGiaHdndRows, ['Mã tương đương', 'MA_TUONG_DUONG', 'MA_DICH_VU', 'MA_DVKT']);
    const refGiaChapterCol = detectColumn(refGiaHdndRows, ['Tên chương', 'TEN_CHUONG', 'Chương', 'chương', 'tên chương']);

    // ── Detect DVKT_GIA_MAX columns (Python: prefer TEN_DVKT_GIA then TEN_DVKT_PHEDUYET)
    let refMaxNameCol: string | null = null;
    for (const col of Object.keys(refMaxRows[0] ?? {})) {
        if (col.includes('TEN_DVKT_GIA') || (col.includes('Tên') && col.includes('Gia'))) {
            refMaxNameCol = col; break;
        }
    }
    // Also find TEN_DVKT_PHEDUYET column
    let refMaxPheduyetCol: string | null = null;
    for (const col of Object.keys(refMaxRows[0] ?? {})) {
        if (col.includes('TEN_DVKT_PHEDUYET') || col.trim().toLowerCase() === 'ten_dvkt_pheduyet') {
            refMaxPheduyetCol = col; break;
        }
    }
    // Find code column in MAX (Python: 'MA' + 'DICH_VU'/'DVKT'/'TUONG_DUONG')
    let refMaxCodeCol: string | null = null;
    for (const col of Object.keys(refMaxRows[0] ?? {})) {
        if (col.includes('MA') && (col.includes('DICH_VU') || col.includes('DVKT') || col.includes('TUONG_DUONG'))) {
            refMaxCodeCol = col; break;
        }
    }

    // ── Detect input columns
    const inputNameCol = detectNameColumn(inputRows);
    const inputCodeCol = detectCodeColumn(inputRows);
    const inputChapterCol = detectColumn(inputRows, ['Tên chương', 'TEN_CHUONG', 'Chương', 'chương', 'tên chương']);

    if (!inputNameCol) return {
        success: false,
        message: `Không tìm thấy cột tên. Các cột: ${Object.keys(inputRows[0] ?? {}).join(', ')}`,
        data: [], stats: { total: 0, matched: 0, unmatched: 0, ambiguous: 0 },
    };

    // ── Build lookup indexes
    const giaIdx = buildIndex(refGiaHdndRows, refGiaNameCol, refGiaCodeCol);
    const maxIdx = refMaxNameCol ? buildIndex(refMaxRows, refMaxNameCol, refMaxCodeCol) : null;

    // Chapter cache for chapter-filtered matching
    const chapterCache = new Map<string, string[]>();
    if (refGiaChapterCol) {
        for (const row of refGiaHdndRows) {
            const ch = getColVal(row, refGiaChapterCol).toLowerCase().trim();
            const name = getColVal(row, refGiaNameCol);
            if (ch && name) {
                if (!chapterCache.has(ch)) chapterCache.set(ch, []);
                chapterCache.get(ch)!.push(name);
            }
        }
    }

    // ── QUY_TRINH reference (Python: TEN_DICH_VU col, MA_DICH_VU for code)
    let qtNameCol: string | null = null;
    let qtIdx: RefIndex | null = null;
    if (refQuyTrinhRows?.length) {
        for (const col of Object.keys(refQuyTrinhRows[0])) {
            if (col.includes('TEN_DICH_VU') || (col.toLowerCase().includes('tên') && col.toLowerCase().includes('vụ'))) {
                qtNameCol = col; break;
            }
        }
        if (qtNameCol) qtIdx = buildIndex(refQuyTrinhRows, qtNameCol, 'MA_DICH_VU');
    }

    const stats = { total: 0, matched: 0, unmatched: 0, ambiguous: 0 };
    const results: OutputRow[] = [];
    let stt = 1;

    for (let i = 0; i < inputRows.length; i++) {
        const row = inputRows[i];
        const rawName = getColVal(row, inputNameCol);
        if (!rawName) continue;
        stats.total++;

        const chapterName = inputChapterCol ? getColVal(row, inputChapterCol).toLowerCase().trim() : '';

        let matchedByCodeGia = false;
        let matchedByCodeMax = false;
        let matchResultGia: MatchResult | null = null;
        let matchResultMax: MatchResult | null = null;
        let inputCodeValue = '';
        let codeNotFoundInRef = false;

        // ── Priority 1: Code matching
        if (inputCodeCol) {
            inputCodeValue = getColVal(row, inputCodeCol);
            if (inputCodeValue) {
                const giaRow = giaIdx.codeToRow.get(inputCodeValue);
                if (giaRow) { matchedByCodeGia = true; matchResultGia = { name: getColVal(giaRow, refGiaNameCol), score: 100 }; }

                if (maxIdx) {
                    const maxRow = maxIdx.codeToRow.get(inputCodeValue);
                    if (maxRow) { matchedByCodeMax = true; matchResultMax = { name: getColVal(maxRow, refMaxNameCol!), score: 100 }; }
                }

                if (!matchedByCodeGia && !matchedByCodeMax) codeNotFoundInRef = true;
            }
        }

        // ── Priority 2: Fuzzy name matching (skip only if code matched for that file)
        if (!matchedByCodeGia) {
            const searchNames = chapterName && chapterCache.has(chapterName)
                ? chapterCache.get(chapterName)!
                : giaIdx.nameList;
            matchResultGia = findBestMatch(rawName, searchNames.length > 0 ? searchNames : giaIdx.nameList, threshold);
        }
        if (!matchedByCodeMax && maxIdx) {
            matchResultMax = findBestMatch(rawName, maxIdx.nameList, threshold);
        }

        // ── Build output row (same field names as Python)
        const output: OutputRow = {
            STT: stt++,
            MA_TUONG_DUONG: '', TEN_DVKT_PHEDUYET: '', TEN_DVKT_GIA: '',
            PHAN_LOAI_PTTT: '', DON_GIA: '', GHI_CHU: '', QUYET_DINH: '',
            QUY_TRINH: '', TU_NGAY: '', DEN_NGAY: '', CSKCB_CGKT: '', CSKCB_CLS: '',
            CANH_BAO: '', _highlight: null,
        };

        let matchedNameDisplay = 'KHÔNG TÌM THẤY';
        let scoreDisplay = 0;
        let warning = '';

        const isAmbiguousGia = matchResultGia?.ambiguous ?? false;
        const isAmbiguousMax = matchResultMax?.ambiguous ?? false;

        if (isAmbiguousGia || isAmbiguousMax) {
            // ── Ambiguous: collect all unique choices, use first match for data
            const allAmbiguous: string[] = [];
            const seen = new Set<string>();
            for (const name of [...(matchResultGia?.choices ?? []), ...(matchResultMax?.choices ?? [])]) {
                if (!seen.has(name)) { seen.add(name); allAmbiguous.push(name); }
            }

            // Python: "⚠️ CÓ N DỊCH VỤ TƯƠNG TỰ - XEM KỸ"
            warning = `⚠️ CÓ ${allAmbiguous.length} DỊCH VỤ TƯƠNG TỰ - XEM KỸ`;

            // Still fill data from first ambiguous MAX match (Python does this)
            if (isAmbiguousMax && matchResultMax?.choices?.length) {
                const firstMax = matchResultMax.choices[0];
                const maxRow = maxIdx?.nameToRow.get(firstMax);
                if (maxRow) {
                    output.TEN_DVKT_GIA = firstMax;
                    output.TEN_DVKT_PHEDUYET = refMaxPheduyetCol ? getColVal(maxRow, refMaxPheduyetCol) : getColVal(maxRow, 'TEN_DVKT_PHEDUYET');
                    output.MA_TUONG_DUONG = getColVal(maxRow, 'MA_TUONG_DUONG');
                    output.PHAN_LOAI_PTTT = getColVal(maxRow, 'PHAN_LOAI_PTTT');
                    output.GHI_CHU = getColVal(maxRow, 'GHI_CHU');
                }
            }
            // Fill price from GIA (highest priority for price)
            if (isAmbiguousGia && matchResultGia?.choices?.length) {
                const firstGia = matchResultGia.choices[0];
                const giaRow = giaIdx.nameToRow.get(firstGia);
                if (giaRow) {
                    for (const pc of ['Mức giá', 'DON_GIA', 'Giá', 'Đơn giá', 'don_gia', 'muc_gia', 'Mức Giá']) {
                        const v = getColVal(giaRow, pc);
                        if (v) { output.DON_GIA = v; break; }
                    }
                    output.QUYET_DINH = getColVal(giaRow, 'Quyết định') || getColVal(giaRow, 'QUYET_DINH');
                    if (!output.TEN_DVKT_GIA) output.TEN_DVKT_GIA = firstGia;
                    if (!output.TEN_DVKT_PHEDUYET) output.TEN_DVKT_PHEDUYET = getColVal(giaRow, 'Tên chương theo TT 23/2024');
                }
            }
            matchedNameDisplay = `${allAmbiguous[0]} (⚠️ ${allAmbiguous.length} tương tự)`;
            scoreDisplay = 99;
            stats.ambiguous++;

        } else {
            // ── Step 1: Fill from GIA_HDND (highest priority for price — same as Python)
            if (matchResultGia) {
                const giaRow = giaIdx.nameToRow.get(matchResultGia.name);
                if (giaRow) {
                    for (const pc of ['Mức giá', 'DON_GIA', 'Giá', 'Đơn giá', 'don_gia', 'muc_gia', 'Mức Giá']) {
                        const v = getColVal(giaRow, pc);
                        if (v) { output.DON_GIA = v; break; }
                    }
                    output.QUYET_DINH = getColVal(giaRow, 'Quyết định') || getColVal(giaRow, 'QUYET_DINH');
                    output.MA_TUONG_DUONG = getColVal(giaRow, 'Mã tương đương') || getColVal(giaRow, 'MA_TUONG_DUONG');
                    output.GHI_CHU = getColVal(giaRow, 'Ghi chú') || getColVal(giaRow, 'GHI_CHU');
                    output.TEN_DVKT_GIA = matchResultGia.name;
                    output.TEN_DVKT_PHEDUYET = getColVal(giaRow, 'Tên chương theo TT 23/2024');
                    matchedNameDisplay = matchResultGia.name;
                    scoreDisplay = matchResultGia.score;
                }
            }

            // ── Step 2: Fill from DVKT_GIA_MAX (supplementary, lower priority)
            // TEN_DVKT_PHEDUYET ALWAYS from MAX (user requirement from Python comments)
            if (matchResultMax && maxIdx) {
                const maxRow = maxIdx.nameToRow.get(matchResultMax.name);
                if (maxRow) {
                    if (!output.MA_TUONG_DUONG) output.MA_TUONG_DUONG = getColVal(maxRow, 'MA_TUONG_DUONG');
                    const pheduyet = refMaxPheduyetCol ? getColVal(maxRow, refMaxPheduyetCol) : getColVal(maxRow, 'TEN_DVKT_PHEDUYET');
                    output.TEN_DVKT_PHEDUYET = pheduyet || output.TEN_DVKT_PHEDUYET;
                    if (!output.TEN_DVKT_GIA || output.TEN_DVKT_GIA === rawName) output.TEN_DVKT_GIA = matchResultMax.name;
                    output.PHAN_LOAI_PTTT = getColVal(maxRow, 'PHAN_LOAI_PTTT');
                    if (!output.GHI_CHU) output.GHI_CHU = getColVal(maxRow, 'GHI_CHU');
                    // MAX price: ONLY fallback when GIA_HDND has no price (same as Python)
                    if (!output.DON_GIA) {
                        output.DON_GIA = getColVal(maxRow, 'DON_GIA') || getColVal(maxRow, 'GIÁ_MAX');
                    }
                    if (matchedNameDisplay === 'KHÔNG TÌM THẤY') {
                        matchedNameDisplay = matchResultMax.name;
                        scoreDisplay = matchResultMax.score;
                    }
                }
            }

            // Stats (Python: matched if display != 'KHÔNG TÌM THẤY', else unmatched)
            if (matchedNameDisplay !== 'KHÔNG TÌM THẤY') stats.matched++;
            else stats.unmatched++;
        }

        // ── Step 3: QUY_TRINH lookup (direct by MA_DICH_VU, then fuzzy by name)
        if (qtIdx && qtNameCol) {
            let qtRow: Row | undefined;
            // Python: direct lookup by MA_DICH_VU in QUY_TRINH
            if (output.MA_TUONG_DUONG) qtRow = qtIdx.codeToRow.get(output.MA_TUONG_DUONG);
            // Fallback: name-based (Python uses matched_name_display or service_name)
            if (!qtRow && matchedNameDisplay !== 'KHÔNG TÌM THẤY') {
                const nameToLookup = matchedNameDisplay.includes('(⚠️') ? rawName : matchedNameDisplay;
                const qtMatch = findBestMatch(nameToLookup, qtIdx.nameList, threshold);
                if (qtMatch && !qtMatch.ambiguous) qtRow = qtIdx.nameToRow.get(qtMatch.name);
            }
            if (qtRow) {
                output.QUY_TRINH = getColVal(qtRow, 'QUY_TRINH');
                if (!output.CSKCB_CGKT) output.CSKCB_CGKT = getColVal(qtRow, 'CSKCB_CGKT');
                if (!output.CSKCB_CLS) output.CSKCB_CLS = getColVal(qtRow, 'CSKCB_CLS');
            }
        }

        // ── Step 4: Code mismatch warning (Python: "⚠️ Sai mã so với file HDND | Mã nhập: X | Mã theo tên: Y")
        if (codeNotFoundInRef) {
            const matchedCode = output.MA_TUONG_DUONG;
            const w = `⚠️ Sai mã so với file HDND | Mã nhập: ${inputCodeValue}${matchedCode ? ` | Mã theo tên: ${matchedCode}` : ''}`;
            warning = warning ? `${warning} | ${w}` : w;
        }

        if (!output.TEN_DVKT_GIA) output.TEN_DVKT_GIA = rawName;
        output.CANH_BAO = warning;
        // Python: yellow if '⚠️' in CANH_BAO (last column)
        output._highlight = warning ? 'yellow' : null;
        output._matchedName = matchedNameDisplay;
        output._score = scoreDisplay;

        results.push(output);

        // Per-row progress callback (Python does this each row)
        const pct = Math.round(((i + 1) / inputRows.length) * 100);
        if (warning) {
            onProgress?.(`Đã xử lý ${i + 1}/${inputRows.length}: Cảnh báo - ${rawName.slice(0, 50)}...`, pct);
        } else if (matchedNameDisplay !== 'KHÔNG TÌM THẤY') {
            onProgress?.(`Đã xử lý ${i + 1}/${inputRows.length}: Khớp - ${rawName.slice(0, 50)}...`, pct);
        } else {
            onProgress?.(`Đã xử lý ${i + 1}/${inputRows.length}: Không khớp - ${rawName.slice(0, 50)}...`, pct);
        }

        // Yield every 10 rows OR on last row so React can re-render progress popup
        if (i % 10 === 9 || i === inputRows.length - 1) await yield_();
    }

    return { success: true, message: 'Luồng 2 hoàn tất.', data: results, stats };
}
