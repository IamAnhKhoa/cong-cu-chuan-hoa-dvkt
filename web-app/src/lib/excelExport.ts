/**
 * excelExport.ts
 * Excel read/write utilities. Yellow/orange highlighting matches Python openpyxl output.
 */
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import type { OutputRow, Row } from './processor';

// ─── Column order matches Python output exactly ───────────────────────────────
// Python process_luong1: STT, MA_DICH_VU, TEN_DICH_VU, DON_GIA, QUY_TRINH, CSKCB_CGKT, CSKCB_CLS
const LUONG1_COLS: (keyof OutputRow)[] = [
    'STT', 'MA_DICH_VU', 'TEN_DICH_VU', 'DON_GIA', 'QUY_TRINH', 'CSKCB_CGKT', 'CSKCB_CLS',
];
// Python process_gia_hdnd_file: 14 cols (already verified matching Python)
const LUONG2_COLS: (keyof OutputRow)[] = [
    'STT', 'MA_TUONG_DUONG', 'TEN_DVKT_PHEDUYET', 'TEN_DVKT_GIA',
    'PHAN_LOAI_PTTT', 'DON_GIA', 'GHI_CHU', 'QUYET_DINH',
    'QUY_TRINH', 'TU_NGAY', 'DEN_NGAY', 'CSKCB_CGKT', 'CSKCB_CLS', 'CANH_BAO',
];

const COL_WIDTHS: Record<string, number> = {
    STT: 6,
    // Luồng 1 specific
    MA_DICH_VU: 18, TEN_DICH_VU: 55,
    // Luồng 2 specific
    MA_TUONG_DUONG: 18, TEN_DVKT_PHEDUYET: 55, TEN_DVKT_GIA: 55,
    PHAN_LOAI_PTTT: 14, DON_GIA: 14, GHI_CHU: 22, QUYET_DINH: 16,
    QUY_TRINH: 10, TU_NGAY: 12, DEN_NGAY: 12,
    CSKCB_CGKT: 16, CSKCB_CLS: 65, CANH_BAO: 65,
};

// SheetJS fill helper — matches openpyxl PatternFill('solid', fgColor=…)
function solidFill(rgb: string) {
    return { patternType: 'solid', fgColor: { rgb }, bgColor: { indexed: 64 } };
}

// Python uses ONLY yellow (FFFF00) for ALL warning rows (ambiguous + code mismatch)
const YELLOW = solidFill('FFFF00');
const HEADER_FILL = solidFill('1A3A5C');
const ERROR_HEADER_FILL = solidFill('C0392B');

// ─── Parse uploaded file into row objects ─────────────────────────────────────
export async function parseExcelFile(file: File, sheetIndex = 0): Promise<Row[]> {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target?.result as ArrayBuffer);
                const wb = XLSX.read(data, { type: 'array' });
                const ws = wb.Sheets[wb.SheetNames[sheetIndex]];
                resolve(XLSX.utils.sheet_to_json<Row>(ws, { defval: '' }));
            } catch (err) { reject(err); }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// ─── Build ordered export rows (strips internal fields) ───────────────────────
function buildExportRows(data: OutputRow[], cols: (keyof OutputRow)[]) {
    return data.map(row => {
        const out: Record<string, unknown> = {};
        for (const col of cols) {
            out[col as string] = row[col] ?? '';
        }
        return out;
    });
}

// ─── Apply styles to all cells in a row ───────────────────────────────────────
function applyRowFill(ws: XLSX.WorkSheet, rowIdx: number, colCount: number, fill: object) {
    for (let c = 0; c <= colCount; c++) {
        const addr = XLSX.utils.encode_cell({ r: rowIdx, c });
        if (!ws[addr]) ws[addr] = { t: 's', v: '' };
        ws[addr].s = { fill };
    }
}

// ─── Apply header styles ────────────────────────────────────────────────────────
function applyHeaderStyle(ws: XLSX.WorkSheet, colCount: number, fill: object) {
    for (let c = 0; c <= colCount; c++) {
        const addr = XLSX.utils.encode_cell({ r: 0, c });
        if (ws[addr]) {
            ws[addr].s = {
                fill,
                font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
            };
        }
    }
}

// ─── Write workbook to blob and trigger download ────────────────────────────────
function saveWorkbook(wb: XLSX.WorkBook, filename: string) {
    const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
    saveAs(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
}

// ─── Main export — all rows, yellow/orange highlights ─────────────────────────
export function exportResults(data: OutputRow[], filename: string): void {
    const cols = LUONG2_COLS; // same columns for both luong1 and luong2
    const exportData = buildExportRows(data, cols);

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);

    // Column widths
    ws['!cols'] = cols.map(c => ({ wch: COL_WIDTHS[c as string] ?? 16 }));

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const colCount = range.e.c;

    // Apply row highlights — yellow for all warning rows (same as Python)
    for (let r = 1; r <= range.e.r; r++) {
        const row = data[r - 1];
        if (!row?._highlight) continue;
        applyRowFill(ws, r, colCount, YELLOW);
    }

    // Apply header style
    applyHeaderStyle(ws, colCount, HEADER_FILL);

    // Freeze top row
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };

    XLSX.utils.book_append_sheet(wb, ws, 'Kết quả');
    saveWorkbook(wb, filename);
}

// ─── Error export — only unmatched / warning rows ────────────────────────────
export function exportErrors(data: OutputRow[], filename: string): number {
    const errorRows = data.filter(row =>
        !row.TEN_DVKT_PHEDUYET || !!row._highlight
    );

    if (errorRows.length === 0) return 0;

    const cols = LUONG2_COLS;
    const exportData = errorRows.map(row => {
        const out: Record<string, unknown> = {};
        for (const col of cols) out[col as string] = row[col] ?? '';
        // Derive reason from CANH_BAO content (same as Python warning text)
        let lyDo = '';
        const canh_bao = row.CANH_BAO ?? '';
        if (canh_bao.includes('TƯƠNG TỰ')) lyDo = '⚠ Nhiều kết quả tương tự (ambiguous)';
        else if (canh_bao.includes('Sai mã')) lyDo = '⚠ Sai mã — fallback theo tên';
        else if (!row.TEN_DVKT_PHEDUYET) lyDo = '❌ Không tìm thấy kết quả phù hợp';
        else if (canh_bao) lyDo = `⚠ ${canh_bao}`;
        out['LY_DO_LOI'] = lyDo;
        return out;
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);

    ws['!cols'] = [...cols.map(c => ({ wch: COL_WIDTHS[c as string] ?? 16 })), { wch: 50 }];

    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const colCount = range.e.c;

    for (let r = 1; r <= range.e.r; r++) {
        const row = errorRows[r - 1];
        if (!row) continue;
        applyRowFill(ws, r, colCount, YELLOW);
    }

    applyHeaderStyle(ws, colCount, ERROR_HEADER_FILL);
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };

    XLSX.utils.book_append_sheet(wb, ws, 'Lỗi & Cảnh báo');
    saveWorkbook(wb, filename);
    return errorRows.length;
}

// ─── Template downloads (exact column format Python expects) ─────────────────

/** Luồng 2 input template: Mã tương đương + Tên chương + Tên dịch vụ */
export function downloadTemplateLuong2(): void {
    const headers = ['STT', 'Mã tương đương', 'Tên chương', 'Tên dịch vụ'];
    const sampleRows = [
        [1, '18.0030.0001', 'Chương I: Thăm dò chức năng', 'Siêu âm tim qua thành ngực'],
        [2, '09.0012.0003', 'Chương II: Chẩn đoán hình ảnh', 'Chụp X-quang phổi thẳng'],
        [3, '', 'Chương III: Xét nghiệm', 'Xét nghiệm công thức máu'],
        [4, '17.0055.0002', '', 'Nội soi dạ dày tá tràng'],
        [5, '', '', 'Phẫu thuật cắt ruột thừa nội soi'],
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);

    ws['!cols'] = [{ wch: 5 }, { wch: 18 }, { wch: 38 }, { wch: 55 }];
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };

    // Light blue header to distinguish from output
    const TMPL_FILL = solidFill('2E75B6');
    for (let c = 0; c < headers.length; c++) {
        const addr = XLSX.utils.encode_cell({ r: 0, c });
        if (ws[addr]) ws[addr].s = {
            fill: TMPL_FILL,
            font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
            alignment: { horizontal: 'center', vertical: 'center' },
        };
    }
    // Light grey for optional columns header note
    for (let r = 1; r <= sampleRows.length; r++) {
        for (let c = 0; c < headers.length; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            if (ws[addr]) ws[addr].s = { alignment: { vertical: 'center', wrapText: false } };
        }
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Mẫu nhập liệu');

    // Instruction sheet
    const instrData = [
        ['CỘT', 'BẮT BUỘC?', 'GHI CHÚ'],
        ['STT', 'Không', 'Số thứ tự (tùy chọn)'],
        ['Mã tương đương', 'Không (nhưng nên có)', 'Mã XX.XXXX.XXXX — dùng để khớp chính xác 100%. Không có mã → fuzzy matching theo tên.'],
        ['Tên chương', 'Không', 'Tên chương dịch vụ — giúp thu hẹp phạm vi tìm kiếm, tăng độ chính xác.'],
        ['Tên dịch vụ', '✅ BẮT BUỘC', 'Tên dịch vụ kỹ thuật cần tra cứu. Cột có thể đặt tên: "Tên kỹ thuật", "Tên dịch vụ", "Tên DVKT", v.v.'],
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(instrData);
    ws2['!cols'] = [{ wch: 20 }, { wch: 18 }, { wch: 80 }];
    const HDR_FILL = solidFill('1A3A5C');
    for (let c = 0; c < 3; c++) {
        const addr = XLSX.utils.encode_cell({ r: 0, c });
        if (ws2[addr]) ws2[addr].s = {
            fill: HDR_FILL,
            font: { bold: true, color: { rgb: 'FFFFFF' } },
            alignment: { horizontal: 'center' },
        };
    }
    XLSX.utils.book_append_sheet(wb, ws2, 'Hướng dẫn');

    saveWorkbook(wb, 'MAU_NHAP_LIEU_LUONG2_GIA_HDND.xlsx');
}

/** Luồng 1 input template: MA_DICH_VU + Tên dịch vụ */
export function downloadTemplateLuong1(): void {
    const headers = ['STT', 'MA_DICH_VU', 'Tên dịch vụ'];
    const sampleRows = [
        [1, '18.0030.0001', 'Siêu âm tim qua thành ngực'],
        [2, '09.0012.0003', 'Chụp X-quang phổi thẳng'],
        [3, '', 'Xét nghiệm công thức máu'],
        [4, '17.0055.0002', 'Nội soi dạ dày tá tràng'],
        [5, '', 'Phẫu thuật cắt ruột thừa nội soi'],
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = [{ wch: 5 }, { wch: 18 }, { wch: 55 }];
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };

    const TMPL_FILL = solidFill('2E75B6');
    for (let c = 0; c < headers.length; c++) {
        const addr = XLSX.utils.encode_cell({ r: 0, c });
        if (ws[addr]) ws[addr].s = {
            fill: TMPL_FILL,
            font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
            alignment: { horizontal: 'center' },
        };
    }
    XLSX.utils.book_append_sheet(wb, ws, 'Mẫu nhập liệu');
    saveWorkbook(wb, 'MAU_NHAP_LIEU_LUONG1_QUY_TRINH.xlsx');
}

/** GIA_HDND reference file template: Mã tương đương (priority code match) + name + price */
export function downloadTemplateGiaHdnd(): void {
    const headers = [
        'Mã tương đương',             // ⭐ Priority: code-based exact match
        'Tên dịch vụ kỹ thuật',       // ✅ Required: fuzzy name matching fallback
        'Mức giá',                     // ✅ Required: price output
        'Quyết định',                  // Optional: decision ref number
        'Tên chương theo TT 23/2024',  // Optional: chapter name (narrows search scope)
        'Ghi chú',                     // Optional: notes
    ];
    const sampleRows = [
        ['18.0030.0001', 'Siêu âm tim qua thành ngực', 649000, 'QĐ 3974/QĐ-BYT', 'Chương I: Thăm dò chức năng', ''],
        ['09.0012.0003', 'Chụp X-quang phổi thẳng (1 phim)', 53000, 'QĐ 3974/QĐ-BYT', 'Chương II: Chẩn đoán hình ảnh', ''],
        ['', 'Xét nghiệm công thức máu', 21000, '', 'Chương III: Xét nghiệm', 'Không có mã → khớp theo tên'],
        ['17.0055.0002', 'Nội soi dạ dày tá tràng', 350000, 'QĐ 3974/QĐ-BYT', 'Chương IV: Nội soi tiêu hóa', ''],
        ['01.0001.0001', 'Khám bệnh (nội khoa)', 38000, 'QĐ 3974/QĐ-BYT', 'Chương V: Khám bệnh', ''],
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = [{ wch: 18 }, { wch: 55 }, { wch: 12 }, { wch: 20 }, { wch: 38 }, { wch: 30 }];
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };

    // Header styling: green for priority code col, blue for required, grey for optional
    const PRIORITY_FILL = solidFill('375623'); // dark green
    const REQUIRED_FILL = solidFill('1F4E79'); // dark blue
    const OPTIONAL_FILL = solidFill('595959'); // grey
    const headerFills = [PRIORITY_FILL, REQUIRED_FILL, REQUIRED_FILL, OPTIONAL_FILL, OPTIONAL_FILL, OPTIONAL_FILL];

    for (let c = 0; c < headers.length; c++) {
        const addr = XLSX.utils.encode_cell({ r: 0, c });
        if (ws[addr]) ws[addr].s = {
            fill: headerFills[c],
            font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
            alignment: { horizontal: 'center', wrapText: false },
        };
    }

    // Row data alignment
    for (let r = 1; r <= sampleRows.length; r++) {
        for (let c = 0; c < headers.length; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            if (ws[addr]) ws[addr].s = { alignment: { vertical: 'center' } };
        }
    }

    XLSX.utils.book_append_sheet(wb, ws, 'GIA_HDND');

    // Instruction sheet
    const instrData = [
        ['CỘT', 'LOẠI', 'MÔ TẢ'],
        ['Mã tương đương', '⭐ ƯU TIÊN', 'Mã kỹ thuật dạng XX.XXXX.XXXX — nếu có sẽ khớp chính xác 100%, bỏ qua fuzzy matching. Tên cột chấp nhận: Mã tương đương, MA_TUONG_DUONG, MA_DICH_VU, MA_DVKT.'],
        ['Tên dịch vụ kỹ thuật', '✅ BẮT BUỘC', 'Tên chuẩn hóa của dịch vụ. Dùng để fuzzy matching khi không có mã. Tên cột linh hoạt: "Tên dịch vụ kỹ thuật", "Tên DVKT", v.v.'],
        ['Mức giá', '✅ BẮT BUỘC', 'Đơn giá dịch vụ (số). Tên cột chấp nhận: Mức giá, DON_GIA, Giá, Đơn giá.'],
        ['Quyết định', 'Tùy chọn', 'Số quyết định ban hành giá. Sẽ điền vào cột QUYET_DINH trong kết quả.'],
        ['Tên chương theo TT 23/2024', 'Tùy chọn', 'Thu hẹp phạm vi tìm kiếm → tăng độ chính xác khi file đầu vào có cột "Tên chương".'],
        ['Ghi chú', 'Tùy chọn', 'Ghi chú thêm. Điền vào cột GHI_CHU trong kết quả.'],
        ['', '', ''],
        ['LƯU Ý', '', 'Không xóa dòng tiêu đề. Có thể thêm cột khác nhưng không đổi tên các cột trên.'],
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(instrData);
    ws2['!cols'] = [{ wch: 28 }, { wch: 14 }, { wch: 90 }];
    const HDR_FILL = solidFill('1A3A5C');
    for (let c = 0; c < 3; c++) {
        const addr = XLSX.utils.encode_cell({ r: 0, c });
        if (ws2[addr]) ws2[addr].s = {
            fill: HDR_FILL, font: { bold: true, color: { rgb: 'FFFFFF' } }, alignment: { horizontal: 'center' },
        };
    }
    XLSX.utils.book_append_sheet(wb, ws2, 'Hướng dẫn');

    saveWorkbook(wb, 'MAU_REF_GIA_HDND.xlsx');
}

