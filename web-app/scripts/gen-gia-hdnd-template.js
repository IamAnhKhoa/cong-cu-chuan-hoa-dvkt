/**
 * Generates a sample/template GIA_HDND.xlsx file and places it
 * next to the real reference file in public/data/.
 * Run: node scripts/gen-gia-hdnd-template.js
 */
const XLSX = require('xlsx');
const path = require('path');

const OUT = path.join(__dirname, '../public/data/MAU_GIA_HDND.xlsx');

// ─── Data ────────────────────────────────────────────────────────────────────
const headers = ['Mã tương đương', 'Tên dịch vụ kỹ thuật', 'Mức giá'];

const rows = [
    ['18.0030.0001', 'Siêu âm tim qua thành ngực', 649000],
    ['18.0030.0002', 'Siêu âm tim có cản quang', 1240000],
    ['09.0012.0003', 'Chụp X-quang phổi thẳng (1 phim)', 53000],
    ['09.0012.0004', 'Chụp X-quang phổi nghiêng (1 phim)', 53000],
    ['13.0001.0001', 'Xét nghiệm công thức máu', 21000],
    ['13.0001.0002', 'Xét nghiệm sinh hóa máu (Glucose)', 15000],
    ['17.0055.0002', 'Nội soi dạ dày tá tràng', 350000],
    ['17.0055.0003', 'Nội soi đại tràng', 420000],
    ['01.0001.0001', 'Khám bệnh (nội khoa)', 38000],
    ['', 'Ví dụ không có mã — khớp theo tên', 50000],
];

// ─── Build workbook ───────────────────────────────────────────────────────────
const wb = XLSX.utils.book_new();
const data = [headers, ...rows];
const ws = XLSX.utils.aoa_to_sheet(data);

// Column widths
ws['!cols'] = [{ wch: 18 }, { wch: 55 }, { wch: 14 }];
ws['!freeze'] = { xSplit: 0, ySplit: 1 };

// Header style: green=code(priority), blue=name(required), orange=price(required)
const solidFill = (rgb) => ({ patternType: 'solid', fgColor: { rgb }, bgColor: { rgb: 'FFFFFF' } });
const fills = [solidFill('375623'), solidFill('1F4E79'), solidFill('7B3F00')];
const white = { rgb: 'FFFFFF' };

for (let c = 0; c < headers.length; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    ws[addr].s = {
        fill: fills[c],
        font: { bold: true, color: white, sz: 11 },
        alignment: { horizontal: 'center' },
    };
}

XLSX.utils.book_append_sheet(wb, ws, 'GIA_HDND');

// Instruction sheet
const instrRows = [
    ['CỘT', 'BẮT BUỘC?', 'MÔ TẢ'],
    ['Mã tương đương', '⭐ Ưu tiên', 'Mã kỹ thuật XX.XXXX.XXXX — khớp chính xác 100%. Để trống nếu không có → tự chuyển sang khớp theo tên.'],
    ['Tên dịch vụ kỹ thuật', '✅ Bắt buộc', 'Tên chuẩn của dịch vụ. Được dùng để fuzzy matching. Tên cột linh hoạt.'],
    ['Mức giá', '✅ Bắt buộc', 'Đơn giá (số). Chấp nhận cột tên: Mức giá, DON_GIA, Giá, Đơn giá, muc_gia.'],
    ['', '', ''],
    ['Lưu ý', '', 'Không xóa dòng tiêu đề. Có thể thêm cột Quyết định, Tên chương, Ghi chú để điền thêm vào kết quả.'],
    ['Tên file khi thay', '', 'Đặt tên GIA_HDND.xlsx và copy vào thư mục public/data/ rồi deploy lại.'],
];
const ws2 = XLSX.utils.aoa_to_sheet(instrRows);
ws2['!cols'] = [{ wch: 22 }, { wch: 14 }, { wch: 90 }];
const hdrFill = solidFill('1A3A5C');
for (let c = 0; c < 3; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    ws2[addr].s = { fill: hdrFill, font: { bold: true, color: white }, alignment: { horizontal: 'center' } };
}
XLSX.utils.book_append_sheet(wb, ws2, 'Hướng dẫn');

// ─── Write ────────────────────────────────────────────────────────────────────
XLSX.writeFile(wb, OUT, { bookType: 'xlsx', type: 'file', cellStyles: true });
console.log('✅ Created:', OUT);
