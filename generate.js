#!/usr/bin/env node
/**
 * Generate the Kompit Futsal team roster Excel template.
 *
 * Editable ranges (everything else is sheet-protected):
 *   - TEAM: D1:D4 (identitas) + G1:G2 (Campus League) + F8:I9 (Partisipasi & warna)
 *   - PA / PI: A3:Q16 (Kompit) + R3:AA16 (Campus League)
 *   - Wilayah: fully locked (dropdown source only)
 *
 * Campus League (CL) additions are grouped under "CL_*" constants and
 * "applyCl*" functions so Kompit's own columns stay untouched. CL cells share
 * Kompit's styling.
 *
 * Usage:
 *   npm install
 *   npm run generate
 *   node generate.js --out ./output/Template-Futsal-Kompit.xlsx
 *   node generate.js --password mypass
 */

'use strict';

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const REGIONS = require('./data/indonesia-regions.json');

const SHEET_PROTECT_PASSWORD = 'kompit';
const DATA_START_ROW = 3;
const DATA_END_ROW = 16;
const CURRENT_YEAR = new Date().getFullYear();
const TAHUN_AWAL_MIN = 2023;

const COLORS = [
  'PUTIH',
  'HITAM',
  'KUNING',
  'BIRU',
  'ORANGE',
  'HIJAU',
  'UNGU',
  'COKELAT',
  'ABU-ABU',
  'PINK',
];

const TIPE_OPTIONS = ['Athlete', 'Official', 'Coach', 'Manager'];
const POSISI_OPTIONS = ['Goalkeeper', 'Pivot', 'Flank', 'Anchor', '-'];

// ── Campus League additions ────────────────────────────────────────────────
// Last column owned by Kompit on PA/PI (Q). CL columns start right after it.
const KOMPIT_LAST_COL = 17;

const CL_SPORT_OPTIONS = ['Futsal', 'Basketball', 'Badminton'];
const CL_BPJSTK_LENGTH = 11;

/**
 * Extra roster columns CL needs to build User / UserAttribute records.
 * Appended after Kompit's columns in this order (R..AA). `sample` holds the
 * example values for the Athlete (row 3) and Official (row 4) sample rows.
 *
 * Optional `clubYear: true` enforces format "<club>,<tahun>" when filled.
 */
const CL_ROSTER_COLUMNS = [
  { header: 'Ukuran Baju', width: 16, sample: { PA: ['L', 'XL'], PI: ['M', 'S'] } },
  { header: 'Ukuran Celana', width: 16, sample: { PA: ['32', '34'], PI: ['28', '27'] } },
  { header: 'Ukuran Sepatu', width: 16, sample: { PA: [42, 43], PI: [38, 37] } },
  {
    header: 'Merk dan Tipe HP',
    width: 24,
    sample: { PA: ['Samsung Galaxy A54', 'iPhone 13'], PI: ['iPhone 14', 'Xiaomi Redmi Note 12'] },
  },
  { header: 'Nama Bank', width: 18, sample: { PA: ['BCA', 'Mandiri'], PI: ['BRI', 'BNI'] } },
  {
    header: 'Merk dan Tipe Kendaraan',
    width: 26,
    sample: { PA: ['Honda Beat', 'Toyota Avanza'], PI: ['Yamaha Mio', 'Honda Scoopy'] },
  },
  {
    header: 'Merk dan Tipe Laptop',
    width: 26,
    sample: { PA: ['ASUS VivoBook', 'Lenovo IdeaPad'], PI: ['MacBook Air', 'Acer Aspire'] },
  },
  {
    header: 'Nomor Kepesertaan BPJSTK',
    width: 26,
    sample: { PA: ['12345678901', '10987654321'], PI: ['11223344556', '16543210987'] },
    text: true, // keep leading zeros
  },
  {
    header: 'Asal Club (Pro/Non Pro) Sebelumnya\ndan tahun bergabung',
    width: 35,
    sample: { PA: ['Persib,2020', 'Arema,2019'], PI: ['Persib,2020', 'Persebaya,2021'] },
    text: true,
    wrapHeader: true,
    clubYear: true,
  },
  {
    header: 'Asal Club (Pro/Non Pro) saat ini\ndan tahun bergabung',
    width: 31,
    sample: { PA: ['Persija,2025', 'Bali United,2024'], PI: ['Persija,2025', 'PSIS,2023'] },
    text: true,
    wrapHeader: true,
    clubYear: true,
  },
];
const CL_LAST_COL = KOMPIT_LAST_COL + CL_ROSTER_COLUMNS.length;

/** Instructions appended to "Petunjuk Pengisian". */
const CL_PETUNJUK_ROWS = [
  ['Ukuran Baju / Ukuran Celana', 'Opsional. Boleh teks atau angka (contoh: XL, 32).'],
  ['Ukuran Sepatu', 'Opsional. Angka 30–50 (contoh: 42).'],
  ['Merk dan Tipe HP', 'Opsional. Contoh: Samsung Galaxy A54.'],
  ['Nama Bank', 'Opsional. Nama bank rekening peserta (contoh: BCA).'],
  ['Merk dan Tipe Kendaraan', 'Opsional. Contoh: Honda Beat.'],
  ['Merk dan Tipe Laptop', 'Opsional. Contoh: ASUS VivoBook.'],
  [
    'Nomor Kepesertaan BPJSTK',
    `Nomor kepesertaan BPJS Ketenagakerjaan, ${CL_BPJSTK_LENGTH} digit angka. Diisi untuk Athlete dan Official yang sudah terdaftar.`,
  ],
  [
    'Asal Club (Pro/Non Pro) Sebelumnya dan tahun bergabung',
    'Opsional. Format: <nama club>,<tahun> (contoh: Persib,2020). Hanya satu koma; tahun 4 digit.',
  ],
  [
    'Asal Club (Pro/Non Pro) saat ini dan tahun bergabung',
    'Opsional. Format: <nama club>,<tahun> (contoh: Persija,2025). Hanya satu koma; tahun 4 digit.',
  ],
  [
    'Cabang Olahraga (sheet TEAM, G1)',
    `Wajib. Cabang olahraga tim pada file ini (dropdown: ${CL_SPORT_OPTIONS.join(', ')}). Dipakai untuk membentuk kategori tim.`,
  ],
  [
    'Wilayah (sheet TEAM, G2)',
    'Wilayah/region pertandingan tim (contoh: Yogyakarta). Jika kosong, ditentukan saat file diunggah di CMS Campus League.',
  ],
  [
    'Kategori Tim (sheet TEAM, kolom E)',
    'Terisi otomatis dari Cabang Olahraga + jenis tim Putra/Putri (contoh: Futsal Putra). Tidak perlu diisi manual.',
  ],
  [
    'Partisipasi (sheet TEAM, kolom F)',
    'Opsional. TRUE jika tim ikut berpartisipasi, FALSE jika tidak. Default TRUE.',
  ],
  [
    'Catatan Campus League',
    'Setiap tim (sheet PA / PI) wajib memiliki minimal satu baris dengan Tipe = Manager. Baris contoh (Rizky Pratama, Budi Santoso, Hendra Gunawan, Agus Firmansyah / Aulia Rahma, Maya Kusumawati, Rina Wulandari, Fitri Handayani) harus dihapus atau ditimpa sebelum diunggah.',
  ],
];

const THIN_BORDER = {
  left: { style: 'thin', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } },
  top: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
};

const FILL_HEADER_GRAY = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFD3D3D3' },
  bgColor: { argb: 'FFD3D3D3' },
};

/** Auto/locked TEAM section (banner + formula columns). */
const FILL_AUTO_GRAY = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFCCCCCC' },
  bgColor: { argb: 'FFCCCCCC' },
};

function parseArgs(argv) {
  const args = { out: null, password: SHEET_PROTECT_PASSWORD };
  for (let i = 2; i < argv.length; i++) {
    if (argv[i] === '--out' && argv[i + 1]) {
      args.out = argv[++i];
    } else if (argv[i] === '--password' && argv[i + 1]) {
      args.password = argv[++i];
    } else if (argv[i] === '--help' || argv[i] === '-h') {
      args.help = true;
    }
  }
  return args;
}

function requiredHeader(text) {
  // Required column header with red asterisk
  return {
    richText: [
      {
        font: { name: 'Cambria', bold: true, color: { theme: 1 }, size: 11 },
        text,
      },
      {
        font: { name: 'Cambria', bold: true, color: { argb: 'FFFF0000' }, size: 11 },
        text: '*',
      },
    ],
  };
}

function lockCell(cell, locked) {
  // Sheet protection uses cell.protection.locked (default true when protected)
  cell.protection = { locked, hidden: false };
}

/** Lock every cell in a rectangular used range. */
function lockAllUsed(ws, maxRow, maxCol) {
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      lockCell(ws.getCell(r, c), true);
    }
  }
}

function addDataValidation(ws, address, validation) {
  // Shared defaults for all data validations on the workbook
  ws.dataValidations.add(address, {
    allowBlank: true,
    showErrorMessage: true,
    showInputMessage: true,
    // Reject invalid input in Excel / LibreOffice / Google Sheets
    errorStyle: 'stop',
    ...validation,
  });
}

/**
 * Returns an Excel formula that is TRUE when the cell (after removing spaces
 * and '+') contains digits 0-9 only. More reliable than ISNUMBER(VALUE(...)).
 */
function formulaDigitsOnlyPhone(cell) {
  const cleaned = `SUBSTITUTE(SUBSTITUTE(${cell}," ",""),"+","")`;
  return `LEN(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(${cleaned},"0",""),"1",""),"2",""),"3",""),"4",""),"5",""),"6",""),"7",""),"8",""),"9",""))=0`;
}

/** Indonesian mobile/WhatsApp validation: 08… / 62… / +62…, digits only. */
function formulaWhatsApp(cell) {
  const clean = `SUBSTITUTE(${cell}," ","")`;
  const digitsOnly = formulaDigitsOnlyPhone(cell);
  return `OR(${cell}="",AND(${digitsOnly},OR(AND(LEFT(${clean},2)="08",LEN(${clean})>=10,LEN(${clean})<=13),AND(LEFT(${clean},3)="+62",LEN(SUBSTITUTE(${clean},"+",""))>=11,LEN(SUBSTITUTE(${clean},"+",""))<=14),AND(LEFT(${clean},2)="62",LEN(${clean})>=11,LEN(${clean})<=14))))`;
}

// Populated by initRegionListFormulas()
let PROVINSI_LIST_FORMULA = null;
/** Full city list (birth place dropdown on PA/PI). */
let KOTA_LIST_FORMULA = null;
/**
 * Province-dependent city list for TEAM!D2.
 * Points at Wilayah!D which is filled by a FILTER formula keyed off TEAM!D3.
 */
let KOTA_CASCADING_FORMULA = null;

function initRegionListFormulas() {
  // Build sorted province/city lists and the Excel range formulas used by dropdowns
  const provinces = Object.keys(REGIONS).sort((a, b) => a.localeCompare(b, 'id'));
  const allCities = [
    ...new Set(provinces.flatMap((p) => REGIONS[p])),
  ].sort((a, b) => a.localeCompare(b, 'id'));

  const maxPerProv = Math.max(1, ...provinces.map((p) => REGIONS[p].length));
  const cityPairCount = provinces.reduce((n, p) => n + REGIONS[p].length, 0);
  const dataLastRow = cityPairCount + 1;

  PROVINSI_LIST_FORMULA = `Wilayah!$F$2:$F$${provinces.length + 1}`;
  KOTA_LIST_FORMULA = `Wilayah!$G$2:$G$${allCities.length + 1}`;
  // Cascading cities from FILTER output (cross-sheet list DV works in Google Sheets)
  KOTA_CASCADING_FORMULA = `Wilayah!$D$2:$D$${maxPerProv + 5}`;

  return { provinces, allCities, maxPerProv, dataLastRow };
}

function buildRefSheet(wb) {
  // Wilayah sheet: province/city pairs + FILTER helper + lookup columns
  const { provinces, allCities, maxPerProv, dataLastRow } = initRegionListFormulas();

  const ws = wb.addWorksheet('Wilayah', {
    properties: { defaultRowHeight: 15 },
    views: [{ state: 'normal', showGridLines: true }],
  });
  ws.properties.tabColor = { argb: 'FF808080' };

  ws.getCell('A1').value = 'Provinsi';
  ws.getCell('B1').value = 'Kota';
  ws.getCell('A1').font = { bold: true, name: 'Calibri' };
  ws.getCell('B1').font = { bold: true, name: 'Calibri' };

  let row = 2;
  // Flat Province|City table used as FILTER input
  provinces.forEach((p) => {
    REGIONS[p].forEach((city) => {
      ws.getCell(`A${row}`).value = p;
      ws.getCell(`B${row}`).value = city;
      row += 1;
    });
  });

  // Column D: FILTER cities matching TEAM!D3 (feeds cascading city dropdown)
  ws.getCell('D1').value = 'KotaSesuaiProvinsi';
  ws.getCell('D1').font = { bold: true, name: 'Calibri' };
  ws.getCell('D2').value = {
    formula: `IFERROR(FILTER(B2:B${dataLastRow},A2:A${dataLastRow}=TEAM!D3),"")`,
  };
  ws.getCell('E1').value = '← otomatis dari Provinsi (TEAM!D3). Jangan diubah.';
  ws.getCell('E1').font = {
    italic: true,
    color: { argb: 'FF666666' },
    name: 'Calibri',
    size: 9,
  };

  ws.getCell('F1').value = 'DaftarProvinsi';
  ws.getCell('F1').font = { bold: true, name: 'Calibri' };
  provinces.forEach((p, i) => {
    ws.getCell(`F${i + 2}`).value = p;
  });

  ws.getCell('G1').value = 'SemuaKota';
  ws.getCell('G1').font = { bold: true, name: 'Calibri' };
  allCities.forEach((c, i) => {
    ws.getCell(`G${i + 2}`).value = c;
  });

  ws.getColumn(1).width = 26;
  ws.getColumn(2).width = 22;
  ws.getColumn(4).width = 22;
  ws.getColumn(5).width = 48;
  ws.getColumn(6).width = 26;
  ws.getColumn(7).width = 22;

  const usedLastRow = Math.max(
    dataLastRow,
    allCities.length,
    provinces.length,
    maxPerProv + 5
  ) + 1;
  // Lock a generous range so nothing on Wilayah is editable under sheet protection
  lockAllUsed(ws, Math.max(usedLastRow, 500), 10);

  return { provinces, allCities, dataLastRow, maxPerProv };
}

function buildPetunjukSheet(wb) {
  // Instructions sheet (Indonesian copy for end users; fully locked)
  const ws = wb.addWorksheet('Petunjuk Pengisian', {
    properties: { defaultRowHeight: 15, defaultColWidth: 14.43 },
  });

  ws.getColumn(1).width = 25;
  ws.getColumn(2).width = 113;

  const rows = [
    ['Kolom', 'Aturan & Format'],
    [
      'Sheet TEAM',
      "Wajib diisi. Yang bisa diedit: D1–D4 (identitas), G1–G2 (Campus League), F8–I9 (Partisipasi & warna kostum). Sel abu-abu (No., Nama Tim, Singkatan, Kategori Tim) terisi otomatis dari rumus — jangan diubah. Provinsi & Kota wajib dari dropdown.",
    ],
    [
      'Nama Sheet',
      'Format: [SINGKATAN] PA (untuk tim Putra) atau [SINGKATAN] PI (untuk tim Putri). Jika tanpa inisial maka dianggap tim Putra. Contoh: UGM PA.',
    ],
    ['Tipe*', 'Pilih "Athlete", "Official", "Coach", atau "Manager" (dropdown). Wajib diisi.'],
    ['Nama*', 'Nama lengkap. Wajib diisi.'],
    [
      'Foto',
      'URL foto peserta (http:// atau https://). Opsional.',
    ],
    [
      'Email*',
      'Alamat email yang valid (contoh: rizky.pratama@student.ugm.ac.id). Wajib diisi dan tidak boleh duplikat.',
    ],
    [
      'No. WhatsApp*',
      'Nomor HP/WhatsApp Indonesia. Hanya angka (huruf ditolak). Format: 08xxxxxxxxxx, 62xxxxxxxxxx, atau +62xxxxxxxxxx. Wajib diisi.',
    ],
    [
      'No. Punggung*',
      'Angka 0–99 (maksimal 2 digit). Wajib diisi untuk Athlete. Tidak boleh duplikat dengan baris lain di sheet yang sama.',
    ],
    ['Posisi*', 'Posisi dalam tim (Goalkeeper, Pivot, Flank, Anchor, -). Wajib diisi untuk Athlete.'],
    ['Tempat Lahir*', 'Pilih kota dari dropdown. Wajib diisi untuk Athlete.'],
    [
      'Tanggal Lahir*',
      'Wajib menggunakan format DD-MM-YYYY atau DD/MM/YYYY (contoh: 12-01-2005 atau 12/01/2005). Wajib diisi untuk Athlete. Usia Athlete maksimal 24 tahun.',
    ],
    [
      'NIM*',
      'Nomor Induk Mahasiswa Indonesia: 8–18 digit angka saja (tanpa spasi/huruf). Wajib diisi untuk Athlete.',
    ],
    ['Jurusan*', 'Jurusan kuliah. Wajib diisi untuk Athlete.'],
    [
      'Tahun Awal*',
      `Tahun masuk kuliah. Minimal ${TAHUN_AWAL_MIN}, maksimal tahun berjalan (${CURRENT_YEAR}). Wajib diisi untuk Athlete.`,
    ],
    [
      'IPK*',
      'Format desimal dengan titik, contoh: 3.5 atau 3.50. Minimal 2.25 dan maksimal 4.00. Wajib diisi untuk Athlete jika Tahun Awal bukan tahun saat ini.',
    ],
    ['Berat Badan*', 'Hanya angka dalam satuan kg. Wajib diisi untuk Athlete.'],
    ['Tinggi Badan*', 'Hanya angka dalam satuan cm. Wajib diisi untuk Athlete.'],
    [
      'Instagram*',
      'Username Instagram. Hanya huruf, angka, titik (.), dan underscore (_). Maksimal 30 karakter. Wajib diisi untuk Athlete.',
    ],
    [
      'TikTok*',
      'Username TikTok. Hanya huruf, angka, titik (.), dan underscore (_). Maksimal 24 karakter. Wajib diisi untuk Athlete.',
    ],
    [
      'Proteksi Sheet',
      'Sheet terkunci. TEAM: D1–D4, G1–G2, dan F8–I9 bisa diedit. Sel abu-abu di tabel tim terkunci karena terisi rumus. PA & PI: hanya A3–AA16 bisa diedit. Sheet Wilayah terkunci penuh (sumber dropdown, jangan diubah).',
    ],
    ...CL_PETUNJUK_ROWS,
  ];

  rows.forEach((row, idx) => {
    const r = idx + 1;
    const a = ws.getCell(`A${r}`);
    const b = ws.getCell(`B${r}`);
    a.value = row[0];
    b.value = row[1];
    a.border = { ...THIN_BORDER };
    b.border = { ...THIN_BORDER };
    a.alignment = { vertical: 'top', wrapText: true };
    b.alignment = { vertical: 'top', wrapText: true };

    if (r === 1) {
      a.font = { bold: true, size: 12, name: 'Cambria', color: { theme: 1 } };
      b.font = { bold: true, size: 12, name: 'Cambria', color: { theme: 1 } };
      a.fill = FILL_HEADER_GRAY;
      b.fill = FILL_HEADER_GRAY;
    } else {
      a.font = { size: 11, name: 'Calibri', color: { theme: 1 } };
      b.font = { size: 11, name: 'Calibri', color: { theme: 1 } };
    }
    lockCell(a, true);
    lockCell(b, true);
  });

  return ws;
}

function buildTeamSheet(wb) {
  // TEAM sheet: university identity + derived Putra/Putri rows
  initRegionListFormulas();

  const ws = wb.addWorksheet('TEAM', {
    properties: {
      defaultRowHeight: 15,
      defaultColWidth: 14.43,
      outlineProperties: { summaryBelow: false, summaryRight: false },
    },
  });

  ws.getColumn(1).width = 4.43;
  ws.getColumn(2).width = 7.14;
  ws.getColumn(3).width = 29.71;
  ws.getColumn(4).width = 34.71;
  ws.getColumn(5).width = 17.57;
  ws.getColumn(6).width = 14;
  ws.getColumn(7).width = 19.43;
  ws.getColumn(8).width = 19.43;
  ws.getColumn(9).width = 19.43;

  const labels = [
    [1, 'Nama Universitas :', 'Universitas Gadjah Mada'],
    [2, 'Kota:', 'Yogyakarta'],
    [3, 'Provinsi:', 'DI Yogyakarta'],
    [4, 'Singkatan:', 'UGM'],
  ];

  labels.forEach(([row, label, value]) => {
    const c = ws.getCell(`C${row}`);
    const d = ws.getCell(`D${row}`);
    c.value = label;
    d.value = value;
    c.font = { name: 'Calibri', color: { theme: 1 } };
    d.font = { name: 'Calibri', color: { argb: 'FF000000' } };
    lockCell(c, true);
    lockCell(d, false);
  });

  // Banner: auto section notice (merged, locked, gray)
  ws.mergeCells('B6:E6');
  const banner = ws.getCell('B6');
  banner.value = 'BAGIAN INI TIDAK PERLU DIISI KARENA OTOMATIS';
  banner.font = { name: 'Calibri', color: { theme: 1 } };
  banner.fill = FILL_AUTO_GRAY;
  banner.border = { ...THIN_BORDER };
  banner.alignment = { horizontal: 'center' };
  lockCell(banner, true);
  for (const addr of ['C6', 'D6', 'E6']) {
    const cell = ws.getCell(addr);
    cell.border = { ...THIN_BORDER };
    lockCell(cell, true);
  }

  // Team table header row (row 7)
  // Auto cols B–E + Partisipasi/Kostum Ketiga: CCCCCC; editable warna G–H: D3D3D3
  const headers = [
    ['B7', 'No.', FILL_AUTO_GRAY, 'Cambria'],
    ['C7', 'Nama Tim', FILL_AUTO_GRAY, 'Cambria'],
    ['D7', 'Singkatan', FILL_AUTO_GRAY, 'Cambria'],
    ['E7', 'Kategori Tim', FILL_AUTO_GRAY, 'Cambria'],
    ['F7', 'Partisipasi', FILL_AUTO_GRAY, 'Calibri'],
    ['G7', 'Warna Kandang', FILL_HEADER_GRAY, 'Cambria'],
    ['H7', 'Warna Tandang', FILL_HEADER_GRAY, 'Cambria'],
    ['I7', 'Kostum Ketiga', FILL_AUTO_GRAY, 'Calibri'],
  ];
  headers.forEach(([addr, text, fill, fontName]) => {
    const cell = ws.getCell(addr);
    cell.value = text;
    cell.font = { bold: true, name: fontName, color: { theme: 1 } };
    cell.fill = fill;
    cell.border = { ...THIN_BORDER };
    cell.alignment = { horizontal: 'center' };
    lockCell(cell, true);
  });

  // Men's team row (Putra) — row 8
  ws.getCell('B8').value = 1;
  ws.getCell('C8').value = { formula: 'IF($D$1="", "", $D$1 & " Putra")' };
  ws.getCell('D8').value = { formula: 'IF($D$4="", "", $D$4 & "PA")' };
  ws.getCell('E8').value = {
    formula: 'IF($G$1="","",$G$1&" "&IF(RIGHT(D8,2)="PI","Putri","Putra"))',
  };
  ws.getCell('F8').value = true;
  ws.getCell('G8').value = 'PUTIH';
  ws.getCell('H8').value = 'HITAM';
  ws.getCell('I8').value = 'KUNING';

  // Women's team row (Putri) — row 9
  ws.getCell('B9').value = 2;
  ws.getCell('C9').value = { formula: 'IF($D$1="", "", $D$1 & " Putri")' };
  ws.getCell('D9').value = { formula: 'IF($D$4="", "", $D$4 & "PI")' };
  ws.getCell('E9').value = {
    formula: 'IF($G$1="","",$G$1&" "&IF(RIGHT(D9,2)="PI","Putri","Putra"))',
  };
  ws.getCell('F9').value = true;
  ws.getCell('G9').value = 'BIRU';
  ws.getCell('H9').value = 'ORANGE';
  ws.getCell('I9').value = 'HIJAU';

  // B–E: auto/formula (gray + locked). F–I: editable (no gray fill).
  for (const row of [8, 9]) {
    for (const col of ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']) {
      const cell = ws.getCell(`${col}${row}`);
      cell.border = { ...THIN_BORDER };
      cell.font = { name: 'Calibri', color: { argb: 'FF000000' } };
      if (col === 'B') cell.alignment = { horizontal: 'right' };

      const isAutoCol = col === 'B' || col === 'C' || col === 'D' || col === 'E';
      if (isAutoCol) {
        cell.fill = FILL_AUTO_GRAY;
        lockCell(cell, true);
      } else {
        lockCell(cell, false);
      }
    }
  }

  // Province list from Wilayah!F; city list from Wilayah!D (FILTER by D3)
  addDataValidation(ws, 'D3', {
    type: 'list',
    formulae: [PROVINSI_LIST_FORMULA],
    promptTitle: 'Provinsi',
    prompt: 'Pilih provinsi Indonesia',
    errorTitle: 'Provinsi tidak valid',
    error: 'Pilih provinsi dari daftar.',
  });

  addDataValidation(ws, 'D2', {
    type: 'list',
    formulae: [KOTA_CASCADING_FORMULA],
    promptTitle: 'Kota',
    prompt: 'Pilih kota sesuai provinsi di D3',
    errorTitle: 'Kota tidak valid',
    error:
      'Pilih kota yang sesuai dengan provinsi terpilih. Ganti Provinsi dulu, lalu pilih ulang Kota.',
  });

  addDataValidation(ws, 'D4', {
    type: 'custom',
    formulae: [
      'AND(LEN(D4)>=2,LEN(D4)<=10,EXACT(D4,UPPER(D4)),ISERROR(FIND(" ",D4)))',
    ],
    promptTitle: 'Singkatan',
    prompt: '2–10 karakter, huruf kapital tanpa spasi (contoh: UGM)',
    errorTitle: 'Singkatan tidak valid',
    error: 'Gunakan 2–10 karakter huruf kapital/angka tanpa spasi.',
  });

  for (const addr of ['F8', 'F9']) {
    addDataValidation(ws, addr, {
      type: 'list',
      formulae: ['"TRUE,FALSE"'],
      promptTitle: 'Partisipasi',
      prompt: 'TRUE jika ikut, FALSE jika tidak',
      errorTitle: 'Partisipasi tidak valid',
      error: 'Pilih TRUE atau FALSE.',
    });
  }

  const colorList = `"${COLORS.join(',')}"`;
  for (const addr of ['G8', 'G9', 'H8', 'H9', 'I8', 'I9']) {
    addDataValidation(ws, addr, {
      type: 'list',
      formulae: [colorList],
    });
  }

  applyClTeamFields(ws);

  return ws;
}

/**
 * Campus League fields on TEAM:
 *   - F1:G2  Cabang Olahraga / Wilayah (G1:G2 editable)
 * Kategori Tim lives in the main team table (E8:E9, formula-driven).
 */
function applyClTeamFields(ws) {
  ws.getColumn(6).width = Math.max(ws.getColumn(6).width || 0, 20);
  ws.getColumn(7).width = Math.max(ws.getColumn(7).width || 0, 19.43);

  const fields = [
    [1, 'Cabang Olahraga:', CL_SPORT_OPTIONS[0]],
    [2, 'Wilayah:', 'Yogyakarta'],
  ];
  fields.forEach(([row, label, value]) => {
    const f = ws.getCell(`F${row}`);
    const g = ws.getCell(`G${row}`);
    f.value = label;
    g.value = value;
    f.font = { name: 'Calibri', color: { theme: 1 } };
    g.font = { name: 'Calibri', color: { argb: 'FF000000' } };
    lockCell(f, true);
    lockCell(g, false);
  });

  addDataValidation(ws, 'G1', {
    type: 'list',
    formulae: [`"${CL_SPORT_OPTIONS.join(',')}"`],
    promptTitle: 'Cabang Olahraga',
    prompt: 'Pilih cabang olahraga tim pada file ini',
    errorTitle: 'Cabang olahraga tidak valid',
    error: 'Pilih dari dropdown.',
  });
}

function applyRosterHeaders(ws) {
  // PA/PI header row (row 2); locked
  const headers = [
    ['A2', requiredHeader('Tipe')],
    ['B2', requiredHeader('Nama')],
    ['C2', 'Foto'],
    ['D2', requiredHeader('Email')],
    ['E2', requiredHeader('No. WhatsApp')],
    ['F2', requiredHeader('No. Punggung')],
    ['G2', requiredHeader('Posisi')],
    ['H2', requiredHeader('Tempat Lahir')],
    ['I2', requiredHeader('Tanggal Lahir')],
    ['J2', requiredHeader('NIM')],
    ['K2', requiredHeader('Jurusan')],
    ['L2', requiredHeader('Tahun Awal')],
    ['M2', requiredHeader('IPK')],
    ['N2', requiredHeader('Berat Badan')],
    ['O2', requiredHeader('Tinggi Badan (cm)')],
    ['P2', requiredHeader('Instagram')],
    ['Q2', requiredHeader('TikTok')],
  ];

  headers.forEach(([addr, value]) => {
    const cell = ws.getCell(addr);
    cell.value = value;
    if (addr === 'C2') {
      cell.font = { name: 'Cambria', bold: true, color: { theme: 1 }, size: 11 };
    }
    cell.border = { ...THIN_BORDER };
    lockCell(cell, true);
  });

  for (let c = 1; c <= 17; c++) {
    ws.getColumn(c).width = c === 3 ? 30.29 : 20;
  }
}

function photoHyperlink(slug) {
  return {
    text: `https://drive.google.com/file/d/${slug}/view`,
    hyperlink: `https://www.example.com/foto/${slug}`,
  };
}

/** Realistic Indonesian sample rosters for Putra (PA) and Putri (PI). */
const SAMPLE_ROSTERS = {
  PA: [
    {
      row: 3,
      values: {
        A: 'Athlete',
        B: 'Rizky Pratama',
        C: photoHyperlink('rizky-pratama'),
        D: 'rizky.pratama@student.ugm.ac.id',
        E: '081278345621',
        F: 1,
        G: 'Goalkeeper',
        H: 'Yogyakarta',
        I: '14-03-2004',
        J: '2351507001',
        K: 'Ilmu Keolahragaan',
        L: 2023,
        M: 3.42,
        N: 72,
        O: 178,
        P: 'rizkypratama',
        Q: 'rizkypratama',
      },
    },
    {
      row: 4,
      values: {
        A: 'Official',
        B: 'Budi Santoso',
        C: photoHyperlink('budi-santoso'),
        D: 'budi.santoso@mail.ugm.ac.id',
        E: '081356792084',
        F: 99,
        G: '-',
        H: 'Sleman',
        I: '22-08-2003',
        J: '2251504012',
        K: 'Manajemen',
        L: 2023,
        M: 3.28,
        N: 68,
        O: 172,
        P: 'budisantoso',
        Q: 'budisantoso',
      },
    },
    {
      row: 5,
      values: {
        A: 'Coach',
        B: 'Hendra Gunawan',
        C: photoHyperlink('hendra-gunawan'),
        G: '-',
      },
    },
    {
      row: 6,
      values: {
        A: 'Manager',
        B: 'Agus Firmansyah',
        C: photoHyperlink('agus-firmansyah'),
        G: '-',
      },
    },
  ],
  PI: [
    {
      row: 3,
      values: {
        A: 'Athlete',
        B: 'Aulia Rahma',
        C: photoHyperlink('aulia-rahma'),
        D: 'aulia.rahma@student.ugm.ac.id',
        E: '081245678913',
        F: 10,
        G: 'Flank',
        H: 'Bantul',
        I: '05-11-2005',
        J: '2451508033',
        K: 'Kedokteran',
        L: 2024,
        M: 3.67,
        N: 55,
        O: 163,
        P: 'auliarahma',
        Q: 'auliarahma',
      },
    },
    {
      row: 4,
      values: {
        A: 'Official',
        B: 'Maya Kusumawati',
        C: photoHyperlink('maya-kusumawati'),
        D: 'maya.kusumawati@mail.ugm.ac.id',
        E: '081389201456',
        F: 99,
        G: '-',
        H: 'Yogyakarta',
        I: '18-06-2004',
        J: '2351502041',
        K: 'Psikologi',
        L: 2023,
        M: 3.51,
        N: 52,
        O: 160,
        P: 'mayakusuma',
        Q: 'mayakusuma',
      },
    },
    {
      row: 5,
      values: {
        A: 'Coach',
        B: 'Rina Wulandari',
        C: photoHyperlink('rina-wulandari'),
        G: '-',
      },
    },
    {
      row: 6,
      values: {
        A: 'Manager',
        B: 'Fitri Handayani',
        C: photoHyperlink('fitri-handayani'),
        G: '-',
      },
    },
  ],
};

function applySampleRows(ws, sheetName) {
  // Example roster rows so users see expected formats
  const samples = SAMPLE_ROSTERS[sheetName] || SAMPLE_ROSTERS.PA;

  samples.forEach(({ row, values }) => {
    Object.entries(values).forEach(([col, value]) => {
      const cell = ws.getCell(`${col}${row}`);
      cell.value = value;
      if (col === 'C') {
        cell.font = { underline: true, color: { argb: 'FF0000FF' } };
      } else if (col === 'I') {
        cell.numFmt = '@';
        cell.font = { size: 11, name: 'Calibri', color: { theme: 1 } };
      } else {
        cell.font = { name: 'Calibri', color: { theme: 1 } };
      }
      if (col === 'M') {
        cell.numFmt = '0.0';
      }
    });
  });

  // Borders + unlock A3:Q16; force text format for phone / date / student ID
  for (let r = DATA_START_ROW; r <= DATA_END_ROW; r++) {
    for (let c = 1; c <= 17; c++) {
      const cell = ws.getCell(r, c);
      cell.border = { ...THIN_BORDER };
      lockCell(cell, false);
      if (c === 5) cell.numFmt = '@'; // WhatsApp as text
      if (c === 9) cell.numFmt = '@'; // Birth date as text
      if (c === 10) cell.numFmt = '@'; // Student ID (NIM) as text
      if (c === 13) cell.numFmt = '0.0'; // GPA (IPK)
    }
  }
}

function applyRosterValidations(ws) {
  // Data validations for editable roster rows A3:Q16
  const tipeList = `"${TIPE_OPTIONS.join(',')}"`;
  const posisiList = `"${POSISI_OPTIONS.join(',')}"`;

  for (let r = DATA_START_ROW; r <= DATA_END_ROW; r++) {
    // Column A: role type
    addDataValidation(ws, `A${r}`, {
      type: 'list',
      formulae: [tipeList],
      promptTitle: 'Tipe',
      prompt: 'Athlete / Official / Coach / Manager',
      errorTitle: 'Tipe tidak valid',
      error: 'Pilih dari dropdown.',
    });

    // Column C: photo URL
    addDataValidation(ws, `C${r}`, {
      type: 'custom',
      formulae: [
        `OR(C${r}="",LEFT(C${r},7)="http://",LEFT(C${r},8)="https://")`,
      ],
      promptTitle: 'Foto',
      prompt: 'URL gambar (http:// atau https://)',
      errorTitle: 'URL tidak valid',
      error: 'Foto harus berupa URL http:// atau https://',
    });

    // Column D: email (Excel-compatible; avoid Google-only ISEMAIL)
    addDataValidation(ws, `D${r}`, {
      type: 'custom',
      formulae: [
        `OR(D${r}="",AND(ISNUMBER(FIND("@",D${r})),ISNUMBER(FIND(".",D${r},FIND("@",D${r}))),ISERROR(FIND(" ",D${r})),COUNTIF($D$${DATA_START_ROW}:$D$${DATA_END_ROW},D${r})=1))`,
      ],
      promptTitle: 'Email',
      prompt: 'Email valid dan unik di sheet ini',
      errorTitle: 'Email tidak valid',
      error: 'Email harus valid, tanpa spasi, dan tidak duplikat.',
    });

    // Column E: Indonesian WhatsApp / phone — digits only
    addDataValidation(ws, `E${r}`, {
      type: 'custom',
      formulae: [formulaWhatsApp(`E${r}`)],
      allowBlank: true,
      promptTitle: 'No. WhatsApp',
      prompt: '08… / 62… / +62… (hanya angka)',
      errorTitle: 'Nomor HP tidak valid',
      error:
        'Hanya angka. Format: 08xxxxxxxxxx, 62xxxxxxxxxx, atau +62xxxxxxxxxx. Huruf tidak diperbolehkan.',
    });

    // Column F: jersey number 0–99, integer, unique in sheet
    addDataValidation(ws, `F${r}`, {
      type: 'custom',
      formulae: [
        `OR(F${r}="",AND(ISNUMBER(F${r}),F${r}=INT(F${r}),F${r}>=0,F${r}<=99,COUNTIF($F$${DATA_START_ROW}:$F$${DATA_END_ROW},F${r})=1))`,
      ],
      promptTitle: 'No. Punggung',
      prompt: 'Angka 0–99, unik di sheet',
      errorTitle: 'No. Punggung tidak valid',
      error: 'Harus angka 0–99 (maks 2 digit) dan tidak boleh duplikat.',
    });

    // Column G: position
    addDataValidation(ws, `G${r}`, {
      type: 'list',
      formulae: [posisiList],
      promptTitle: 'Posisi',
      prompt: 'Pilih posisi futsal',
      errorTitle: 'Posisi tidak valid',
      error: 'Pilih dari dropdown.',
    });

    // Column H: birth city (Indonesia)
    addDataValidation(ws, `H${r}`, {
      type: 'list',
      formulae: [KOTA_LIST_FORMULA],
      promptTitle: 'Tempat Lahir',
      prompt: 'Pilih kota di Indonesia',
      errorTitle: 'Kota tidak valid',
      error: 'Pilih kota dari daftar.',
    });

    // Column I: birth date as text DD-MM-YYYY or DD/MM/YYYY
    addDataValidation(ws, `I${r}`, {
      type: 'custom',
      formulae: [
        `OR(I${r}="",AND(OR(AND(LEN(I${r})=10,MID(I${r},3,1)="-",MID(I${r},6,1)="-"),AND(LEN(I${r})=10,MID(I${r},3,1)="/",MID(I${r},6,1)="/")),ISNUMBER(VALUE(LEFT(I${r},2))),ISNUMBER(VALUE(MID(I${r},4,2))),ISNUMBER(VALUE(RIGHT(I${r},4))),VALUE(LEFT(I${r},2))>=1,VALUE(LEFT(I${r},2))<=31,VALUE(MID(I${r},4,2))>=1,VALUE(MID(I${r},4,2))<=12,VALUE(RIGHT(I${r},4))>=1980,VALUE(RIGHT(I${r},4))<=YEAR(TODAY())))`,
      ],
      promptTitle: 'Tanggal Lahir',
      prompt: 'DD-MM-YYYY atau DD/MM/YYYY',
      errorTitle: 'Tanggal tidak valid',
      error: 'Gunakan format DD-MM-YYYY atau DD/MM/YYYY.',
    });

    // Column J: Indonesian student ID (NIM), 8–18 digits
    addDataValidation(ws, `J${r}`, {
      type: 'custom',
      formulae: [
        `OR(J${r}="",AND(LEN(TRIM(J${r}))>=8,LEN(TRIM(J${r}))<=18,ISNUMBER(VALUE(J${r})),ISERROR(FIND(" ",J${r})),ISERROR(FIND(".",J${r})),ISERROR(FIND(",",J${r})),ISERROR(FIND("-",J${r})),ISERROR(FIND("e",LOWER(J${r}))),ISERROR(FIND("+",J${r}))))`,
      ],
      promptTitle: 'NIM',
      prompt: '8–18 digit angka (NIM Indonesia)',
      errorTitle: 'NIM tidak valid',
      error: 'NIM harus 8–18 digit angka tanpa huruf/spasi.',
    });

    // Column L: enrollment year 2023–current year
    addDataValidation(ws, `L${r}`, {
      type: 'whole',
      operator: 'between',
      formulae: [TAHUN_AWAL_MIN, CURRENT_YEAR],
      promptTitle: 'Tahun Awal',
      prompt: `${TAHUN_AWAL_MIN}–${CURRENT_YEAR}`,
      errorTitle: 'Tahun Awal tidak valid',
      error: `Tahun awal harus antara ${TAHUN_AWAL_MIN} dan ${CURRENT_YEAR}.`,
    });

    // Column M: GPA (IPK) 2.25–4.00 (e.g. 3.5)
    addDataValidation(ws, `M${r}`, {
      type: 'decimal',
      operator: 'between',
      formulae: [2.25, 4],
      promptTitle: 'IPK',
      prompt: 'Contoh: 3.5 (min 2.25, max 4.00)',
      errorTitle: 'IPK tidak valid',
      error: 'IPK harus desimal 2.25–4.00 (contoh: 3.5).',
    });

    // Column N: weight (kg)
    addDataValidation(ws, `N${r}`, {
      type: 'whole',
      operator: 'between',
      formulae: [30, 200],
      promptTitle: 'Berat Badan',
      prompt: 'kg (30–200)',
      errorTitle: 'Berat tidak valid',
      error: 'Berat badan harus angka 30–200 kg.',
    });

    // Column O: height (cm)
    addDataValidation(ws, `O${r}`, {
      type: 'whole',
      operator: 'between',
      formulae: [100, 250],
      promptTitle: 'Tinggi Badan',
      prompt: 'cm (100–250)',
      errorTitle: 'Tinggi tidak valid',
      error: 'Tinggi badan harus angka 100–250 cm.',
    });

    // Column P: Instagram username
    addDataValidation(ws, `P${r}`, {
      type: 'custom',
      formulae: [
        `OR(P${r}="",AND(LEN(P${r})>=1,LEN(P${r})<=30,ISERROR(FIND(" ",P${r})),ISERROR(FIND("@",P${r}))))`,
      ],
      promptTitle: 'Instagram',
      prompt: 'Username tanpa @, maks 30 karakter',
      errorTitle: 'Instagram tidak valid',
      error: 'Username tanpa spasi/@, maksimal 30 karakter.',
    });

    // Column Q: TikTok username
    addDataValidation(ws, `Q${r}`, {
      type: 'custom',
      formulae: [
        `OR(Q${r}="",AND(LEN(Q${r})>=1,LEN(Q${r})<=24,ISERROR(FIND(" ",Q${r})),ISERROR(FIND("@",Q${r}))))`,
      ],
      promptTitle: 'TikTok',
      prompt: 'Username tanpa @, maks 24 karakter',
      errorTitle: 'TikTok tidak valid',
      error: 'Username tanpa spasi/@, maksimal 24 karakter.',
    });
  }
}

function buildRosterSheet(wb, name) {
  // Build a PA or PI roster sheet with sample data + validations
  const ws = wb.addWorksheet(name, {
    properties: { defaultRowHeight: 15, defaultColWidth: 14.43 },
  });

  applyRosterHeaders(ws);
  applySampleRows(ws, name);
  applyRosterValidations(ws);
  applyClRosterColumns(ws, name);

  // Keep header / empty row locked; only A3:AA16 is editable
  for (let c = 1; c <= CL_LAST_COL; c++) {
    lockCell(ws.getCell(1, c), true);
    lockCell(ws.getCell(2, c), true);
  }

  return ws;
}

/**
 * Optional club+year format: "<club>,<tahun>" (blank allowed).
 * Exactly one comma, non-empty club name, 4-digit year in a sensible range.
 */
function formulaClubYear(cell) {
  const club = `TRIM(LEFT(${cell},FIND(",",${cell})-1))`;
  const year = `TRIM(MID(${cell},FIND(",",${cell})+1,99))`;
  return `OR(${cell}="",AND(ISNUMBER(FIND(",",${cell})),FIND(",",${cell})>1,LEN(${cell})-LEN(SUBSTITUTE(${cell},",",""))=1,LEN(${club})>=1,LEN(${year})=4,ISNUMBER(VALUE(${year})),VALUE(${year})>=1950,VALUE(${year})<=${CURRENT_YEAR + 1}))`;
}

/**
 * Campus League roster columns (R..AA): header, widths, sample values,
 * borders/unlock for the data rows and validations. Mirrors what Kompit's
 * applyRosterHeaders / applySampleRows / applyRosterValidations do for A..Q.
 */
function applyClRosterColumns(ws, sheetName) {
  const gender = SAMPLE_ROSTERS[sheetName] ? sheetName : 'PA';
  let needsWrappedHeader = false;

  CL_ROSTER_COLUMNS.forEach((col, i) => {
    const c = KOMPIT_LAST_COL + 1 + i;
    const colLetter = ws.getColumn(c).letter;
    const sample = col.sample[gender] || col.sample.PA;

    const header = ws.getCell(2, c);
    header.value = col.header;
    header.font = { name: 'Cambria', bold: true, color: { theme: 1 }, size: 11 };
    header.border = { ...THIN_BORDER };
    if (col.wrapHeader) {
      header.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
      needsWrappedHeader = true;
    }
    ws.getColumn(c).width = col.width;

    for (let r = DATA_START_ROW; r <= DATA_END_ROW; r++) {
      const cell = ws.getCell(r, c);
      cell.border = { ...THIN_BORDER };
      cell.font = { name: 'Calibri', color: { theme: 1 } };
      if (col.text) cell.numFmt = '@';
      lockCell(cell, false);
    }

    // Sample values on the Athlete (row 3) and Official (row 4) example rows
    ws.getCell(DATA_START_ROW, c).value = sample[0];
    ws.getCell(DATA_START_ROW + 1, c).value = sample[1];

    if (col.header === 'Ukuran Sepatu') {
      for (let r = DATA_START_ROW; r <= DATA_END_ROW; r++) {
        addDataValidation(ws, `${colLetter}${r}`, {
          type: 'whole',
          operator: 'between',
          formulae: [30, 50],
          promptTitle: 'Ukuran Sepatu',
          prompt: 'Angka 30–50 (contoh: 42)',
          errorTitle: 'Ukuran sepatu tidak valid',
          error: 'Ukuran sepatu harus angka 30–50.',
        });
      }
    }

    if (col.header === 'Nomor Kepesertaan BPJSTK') {
      for (let r = DATA_START_ROW; r <= DATA_END_ROW; r++) {
        const ref = `${colLetter}${r}`;
        addDataValidation(ws, ref, {
          type: 'custom',
          formulae: [
            `OR(${ref}="",AND(LEN(${ref})=${CL_BPJSTK_LENGTH},ISNUMBER(VALUE(${ref})),ISERROR(FIND(" ",${ref})),ISERROR(FIND(".",${ref})),ISERROR(FIND(",",${ref})),ISERROR(FIND("-",${ref}))))`,
          ],
          promptTitle: 'Nomor Kepesertaan BPJSTK',
          prompt: `${CL_BPJSTK_LENGTH} digit angka`,
          errorTitle: 'Nomor BPJSTK tidak valid',
          error: `Nomor kepesertaan BPJSTK harus ${CL_BPJSTK_LENGTH} digit angka tanpa spasi/huruf.`,
        });
      }
    }

    if (col.clubYear) {
      for (let r = DATA_START_ROW; r <= DATA_END_ROW; r++) {
        const ref = `${colLetter}${r}`;
        addDataValidation(ws, ref, {
          type: 'custom',
          formulae: [formulaClubYear(ref)],
          promptTitle: 'Asal Club',
          prompt: 'Opsional. Format: nama club,tahun (contoh: Persib,2020)',
          errorTitle: 'Format club tidak valid',
          error: `Gunakan format <club>,<tahun> (contoh: Persib,2020). Tahun 1950–${CURRENT_YEAR + 1}. Boleh dikosongkan.`,
        });
      }
    }
  });

  if (needsWrappedHeader) {
    ws.getRow(2).height = 36;
  }
}

async function protectSheets(wb, password) {
  // Protect all sheets; unlock rules are set per-cell above
  const opts = {
    selectLockedCells: true,
    selectUnlockedCells: true,
    formatCells: false,
    formatColumns: false,
    formatRows: false,
    insertColumns: false,
    insertRows: false,
    insertHyperlinks: false,
    deleteColumns: false,
    deleteRows: false,
    sort: false,
    autoFilter: false,
    pivotTables: false,
  };

  for (const ws of wb.worksheets) {
    if (ws.name === 'Wilayah') {
      // Dropdown source: fully locked — no unlocked cells to edit
      await ws.protect(password, {
        ...opts,
        selectLockedCells: true,
        selectUnlockedCells: false,
      });
      continue;
    }
    await ws.protect(password, opts);
  }
}

async function generate(outPath, password) {
  // Assemble workbook, protect sheets, write .xlsx
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Kompit';
  wb.created = new Date();
  wb.properties.date1904 = false;

  initRegionListFormulas();

  // Sheet order: instructions → TEAM → PA → PI → Wilayah (dropdown source)
  buildPetunjukSheet(wb);
  buildTeamSheet(wb);
  buildRosterSheet(wb, 'PA');
  buildRosterSheet(wb, 'PI');
  buildRefSheet(wb);

  wb.views = [{ x: 0, y: 0, width: 12000, height: 15000, firstSheet: 0, activeTab: 1 }];

  await protectSheets(wb, password);

  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  await wb.xlsx.writeFile(outPath);
  return outPath;
}

async function main() {
  const args = parseArgs(process.argv);
  if (args.help) {
    console.log(`Usage: node generate.js [--out path.xlsx] [--password secret]

Editable:
  TEAM  D1:D4, G1:G2, F8:I9
  PA/PI A3:AA16
  Wilayah (locked)

Default password: ${SHEET_PROTECT_PASSWORD}`);
    process.exit(0);
  }

  const outPath =
    args.out ||
    path.join(__dirname, 'output', 'Template Futsal Kompit - CL.xlsx');

  const file = await generate(outPath, args.password);
  console.log(`Generated: ${file}`);
  console.log(`Sheet password: ${args.password}`);
  console.log('Editable: TEAM!D1:D4 + G1:G2 + F8:I9 , PA/PI!A3:AA16 (Wilayah locked)');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
