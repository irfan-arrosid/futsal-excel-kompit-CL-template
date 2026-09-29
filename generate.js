#!/usr/bin/env node
/**
 * Generate the Kompit Futsal team roster Excel template.
 *
 * Editable ranges (everything else is sheet-protected):
 *   - TEAM: D1:D4 (identitas) + G1:G2 (Campus League) + F8:I9 (Ikut/Tidak Ikut & warna)
 *   - PA / PI: sectioned roster (Athlete / Official / Coach / Manager), each with own header
 *     Athlete 20 + Official 3 + Coach 2 + Manager 1 example rows; Kompit A..R + CL S..AD
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
/** Sorted PT names from data/daftar-kampus.xlsx (see indonesia-universities.json). */
const UNIVERSITIES = require('./data/indonesia-universities.json');

const SHEET_PROTECT_PASSWORD = 'kompit';
/**
 * PA/PI layout: each role has its own header row + data block,
 * with 1 blank spacer row between sections.
 *   Athlete  header@2  data 3–22  (20)
 *   Official header@24 data 25–27 (3)
 *   Coach    header@29 data 30–31 (2)
 *   Manager  header@33 data 34    (1)
 */
const ROSTER_SECTIONS = [
  { tipe: 'Athlete', headerRow: 2, dataStart: 3, count: 20 },
  { tipe: 'Official', headerRow: 24, dataStart: 25, count: 3 },
  { tipe: 'Coach', headerRow: 29, dataStart: 30, count: 2 },
  { tipe: 'Manager', headerRow: 33, dataStart: 34, count: 1 },
];
const SAMPLE_COUNTS = { Athlete: 20, Official: 3, Coach: 2, Manager: 1 };
const CURRENT_YEAR = new Date().getFullYear();
const TAHUN_AWAL_MIN = 2023;

function rosterDataRows() {
  const rows = [];
  ROSTER_SECTIONS.forEach((s) => {
    for (let r = s.dataStart; r < s.dataStart + s.count; r++) rows.push(r);
  });
  return rows;
}

function rosterLastRow() {
  const last = ROSTER_SECTIONS[ROSTER_SECTIONS.length - 1];
  return last.dataStart + last.count - 1;
}

function rosterHeaderRows() {
  return ROSTER_SECTIONS.map((s) => s.headerRow);
}

/** Blank spacer rows between sections (1 row before each section after the first). */
function rosterSpacerRows() {
  return ROSTER_SECTIONS.slice(1).map((s) => s.headerRow - 1);
}

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

const POSISI_OPTIONS = ['Goalkeeper', 'Pivot', 'Flank', 'Anchor', '-'];

// ── Campus League additions ────────────────────────────────────────────────
// Last column owned by Kompit on PA/PI (R). CL columns start right after it.
const KOMPIT_LAST_COL = 18;

const CL_SPORT_OPTIONS = ['Futsal', 'Basketball', 'Badminton'];
const CL_BPJSTK_LENGTH = 11;

/**
 * Extra roster columns CL needs to build User / UserAttribute records.
 * Appended after Kompit's columns in this order (S..AD).
 *
 * Optional `clubYear: true` enforces format "<club>,<tahun>" when filled.
 * Optional `requiredForAthlete: true` adds red asterisk + validation when Tipe=Athlete.
 */
const CL_ROSTER_COLUMNS = [
  { header: 'Ukuran Baju', width: 16 },
  { header: 'Ukuran Celana', width: 16 },
  { header: 'Ukuran Sepatu', width: 16 },
  { header: 'Merk dan Tipe HP', width: 24 },
  { header: 'Nama Bank', width: 18 },
  { header: 'Merk dan Tipe Kendaraan', width: 26 },
  { header: 'Merk dan Tipe Laptop', width: 26 },
  { header: 'Nomor Kepesertaan BPJSTK', width: 26, text: true },
  {
    header: 'Asal Club (Pro/Non Pro) Sebelumnya\ndan tahun bergabung',
    width: 35,
    text: true,
    wrapHeader: true,
    clubYear: true,
  },
  {
    header: 'Asal Club (Pro/Non Pro) saat ini\ndan tahun bergabung',
    width: 31,
    text: true,
    wrapHeader: true,
    clubYear: true,
  },
  { header: 'Asal SMP', width: 28, text: true, requiredForAthlete: true },
  { header: 'Asal SMA', width: 28, text: true, requiredForAthlete: true },
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
  ['Asal SMP*', 'Wajib untuk Athlete. Nama sekolah menengah pertama (teks bebas).'],
  ['Asal SMA*', 'Wajib untuk Athlete. Nama sekolah menengah atas (teks bebas).'],
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
    'Ikut/Tidak Ikut (sheet TEAM, kolom F)',
    'Opsional. TRUE jika tim ikut, FALSE jika tidak ikut. Default TRUE.',
  ],
  [
    'Catatan Campus League',
    `Setiap tim (sheet PA / PI) wajib memiliki minimal satu baris Manager. Sheet PA/PI dipisah per seksi (Athlete / Official / Coach / Manager), masing-masing dengan header sendiri. Baris contoh (${SAMPLE_COUNTS.Athlete} Athlete, ${SAMPLE_COUNTS.Official} Official, ${SAMPLE_COUNTS.Coach} Coach, ${SAMPLE_COUNTS.Manager} Manager) harus dihapus atau ditimpa sebelum diunggah.`,
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
 * Points at Wilayah!D filled by legacy INDEX/MATCH + COUNTIF rank (no FILTER),
 * so cascading works in Excel desktop, LibreOffice, Excel web, and Google Sheets.
 */
let KOTA_CASCADING_FORMULA = null;
/** Full university list for TEAM!D1 (daftar-kampus.xlsx). */
let UNIVERSITAS_LIST_FORMULA = null;

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
  // Cascading cities: static range on Wilayah!D (INDEX/MATCH output of ranked pairs)
  KOTA_CASCADING_FORMULA = `Wilayah!$D$2:$D$${maxPerProv + 1}`;
  UNIVERSITAS_LIST_FORMULA = `Universitas!$A$2:$A$${UNIVERSITIES.length + 1}`;

  return { provinces, allCities, maxPerProv, dataLastRow };
}

function buildRefSheet(wb) {
  // Wilayah sheet: province/city pairs + legacy cascading helper + lookup columns
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
  // Flat Province|City table used by cascading helper formulas
  provinces.forEach((p) => {
    REGIONS[p].forEach((city) => {
      ws.getCell(`A${row}`).value = p;
      ws.getCell(`B${row}`).value = city;
      row += 1;
    });
  });

  // Column C: rank matching cities for TEAM!D3 (1,2,3… on matching rows only).
  // Uses IF + COUNTIF — works in Excel desktop, LibreOffice, Sheets, Excel web.
  ws.getCell('C1').value = 'UrutanKota';
  ws.getCell('C1').font = { bold: true, name: 'Calibri' };
  for (let r = 2; r <= dataLastRow; r++) {
    ws.getCell(`C${r}`).value = {
      formula: `IF($A${r}=TEAM!$D$3,COUNTIF($A$2:$A${r},TEAM!$D$3),"")`,
    };
  }

  // Column D: compact list of cities for the selected province (feeds TEAM!D2 DV)
  ws.getCell('D1').value = 'KotaSesuaiProvinsi';
  ws.getCell('D1').font = { bold: true, name: 'Calibri' };
  for (let i = 0; i < maxPerProv; i++) {
    const r = i + 2;
    ws.getCell(`D${r}`).value = {
      formula: `IFERROR(INDEX($B$2:$B$${dataLastRow},MATCH(${i + 1},$C$2:$C$${dataLastRow},0)),"")`,
    };
  }
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
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 22;
  ws.getColumn(5).width = 48;
  ws.getColumn(6).width = 26;
  ws.getColumn(7).width = 22;

  const usedLastRow = Math.max(
    dataLastRow,
    allCities.length,
    provinces.length,
    maxPerProv + 1
  ) + 1;
  // Lock a generous range so nothing on Wilayah is editable under sheet protection
  lockAllUsed(ws, Math.max(usedLastRow, 500), 10);

  return { provinces, allCities, dataLastRow, maxPerProv };
}

/**
 * Locked dropdown source: PT names from daftar-kampus.xlsx
 * (data/indonesia-universities.json). Sheet is hidden + fully protected.
 */
function buildUniversitiesSheet(wb) {
  const ws = wb.addWorksheet('Universitas', {
    properties: { defaultRowHeight: 15 },
    views: [{ state: 'normal', showGridLines: true }],
    state: 'hidden',
  });
  ws.properties.tabColor = { argb: 'FF808080' };

  const header = ws.getCell('A1');
  header.value = 'Nama Perguruan Tinggi — jangan diubah';
  header.font = { bold: true, name: 'Calibri' };
  lockCell(header, true);
  ws.getColumn(1).width = 72;

  UNIVERSITIES.forEach((name, i) => {
    const cell = ws.getCell(i + 2, 1);
    cell.value = name;
    cell.font = { name: 'Calibri' };
    lockCell(cell, true);
  });

  // Lock a generous range so empty cells stay non-editable under sheet protection
  lockAllUsed(ws, Math.max(UNIVERSITIES.length + 2, 500), 3);
  return ws;
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
      "Wajib diisi. Yang bisa diedit: D1–D4 (identitas), G1–G2 (Campus League), F8–I9 (Ikut/Tidak Ikut & warna kostum). Nama Universitas (D1) wajib dipilih dari dropdown daftar kampus. Sel abu-abu (No., Nama Tim, Singkatan, Kategori Tim) terisi otomatis dari rumus — jangan diubah. Provinsi & Kota wajib dari dropdown.",
    ],
    [
      'Nama Universitas (D1)',
      `Wajib. Pilih dari dropdown daftar perguruan tinggi (${UNIVERSITIES.length.toLocaleString('id-ID')} entri dari daftar kampus).`,
    ],
    [
      'Nama Sheet',
      'Format: [SINGKATAN] PA (untuk tim Putra) atau [SINGKATAN] PI (untuk tim Putri). Jika tanpa inisial maka dianggap tim Putra. Contoh: UGM PA.',
    ],
    [
      'Tipe*',
      'Sudah terisi sesuai seksi (Athlete / Official / Coach / Manager) dan terkunci. Isi data di bawah header seksi yang sesuai.',
    ],
    [
      'Layout PA / PI',
      `Dipisah per seksi dengan header sendiri: Athlete (20 baris), Official (3), Coach (2), Manager (1). Baris contoh harus diganti sebelum diunggah.`,
    ],
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
    ['Fakultas*', 'Fakultas kuliah. Wajib diisi untuk Athlete.'],
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
      'Sheet terkunci. TEAM: D1–D4, G1–G2, dan F8–I9 bisa diedit. Sel abu-abu di tabel tim terkunci karena terisi rumus. PA & PI: hanya A3–AD32 bisa diedit. Sheet Wilayah & Universitas terkunci penuh (sumber dropdown; Universitas disembunyikan). Jangan diubah.',
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
  ws.getColumn(4).width = 42;
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

  addDataValidation(ws, 'D1', {
    type: 'list',
    formulae: [UNIVERSITAS_LIST_FORMULA],
    promptTitle: 'Nama Universitas',
    prompt: 'Pilih perguruan tinggi dari daftar kampus',
    errorTitle: 'Universitas tidak valid',
    error: 'Pilih nama universitas dari dropdown (daftar kampus).',
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
  // Auto cols B–E + Ikut/Tidak Ikut/Kostum Ketiga: CCCCCC; editable warna G–H: D3D3D3
  const headers = [
    ['B7', 'No.', FILL_AUTO_GRAY, 'Cambria'],
    ['C7', 'Nama Tim', FILL_AUTO_GRAY, 'Cambria'],
    ['D7', 'Singkatan', FILL_AUTO_GRAY, 'Cambria'],
    ['E7', 'Kategori Tim', FILL_AUTO_GRAY, 'Cambria'],
    ['F7', 'Ikut/Tidak Ikut', FILL_AUTO_GRAY, 'Calibri'],
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

  // Province list from Wilayah!F; city list from Wilayah!D (INDEX/MATCH by D3)
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
      promptTitle: 'Ikut/Tidak Ikut',
      prompt: 'TRUE jika ikut, FALSE jika tidak ikut',
      errorTitle: 'Nilai tidak valid',
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
  // Kompit headers per section; red * only on Athlete (Email keeps * on all sections)
  const plainHeader = (text) => text;
  const athleteDefs = [
    ['A', requiredHeader('Tipe')],
    ['B', requiredHeader('Nama')],
    ['C', 'Foto'],
    ['D', requiredHeader('Email')],
    ['E', requiredHeader('No. WhatsApp')],
    ['F', requiredHeader('No. Punggung')],
    ['G', requiredHeader('Posisi')],
    ['H', requiredHeader('Tempat Lahir')],
    ['I', requiredHeader('Tanggal Lahir')],
    ['J', requiredHeader('NIM')],
    ['K', requiredHeader('Fakultas')],
    ['L', requiredHeader('Jurusan')],
    ['M', requiredHeader('Tahun Awal')],
    ['N', requiredHeader('IPK')],
    ['O', requiredHeader('Berat Badan')],
    ['P', requiredHeader('Tinggi Badan (cm)')],
    ['Q', requiredHeader('Instagram')],
    ['R', requiredHeader('TikTok')],
  ];
  const staffDefs = [
    ['A', plainHeader('Tipe')],
    ['B', plainHeader('Nama')],
    ['C', 'Foto'],
    ['D', requiredHeader('Email')],
    ['E', plainHeader('No. WhatsApp')],
    ['F', plainHeader('No. Punggung')],
    ['G', plainHeader('Posisi')],
    ['H', plainHeader('Tempat Lahir')],
    ['I', plainHeader('Tanggal Lahir')],
    ['J', plainHeader('NIM')],
    ['K', plainHeader('Fakultas')],
    ['L', plainHeader('Jurusan')],
    ['M', plainHeader('Tahun Awal')],
    ['N', plainHeader('IPK')],
    ['O', plainHeader('Berat Badan')],
    ['P', plainHeader('Tinggi Badan (cm)')],
    ['Q', plainHeader('Instagram')],
    ['R', plainHeader('TikTok')],
  ];

  ROSTER_SECTIONS.forEach((section) => {
    const defs = section.tipe === 'Athlete' ? athleteDefs : staffDefs;
    defs.forEach(([col, value]) => {
      const cell = ws.getCell(`${col}${section.headerRow}`);
      cell.value = value;
      if (typeof value === 'string') {
        cell.font = { name: 'Cambria', bold: true, color: { theme: 1 }, size: 11 };
      }
      cell.border = { ...THIN_BORDER };
      lockCell(cell, true);
    });
  });

  for (let c = 1; c <= KOMPIT_LAST_COL; c++) {
    ws.getColumn(c).width = c === 3 ? 30.29 : 20;
  }
}

function photoHyperlink(slug) {
  return {
    text: `https://drive.google.com/file/d/${slug}/view`,
    hyperlink: `https://www.example.com/foto/${slug}`,
  };
}

function slugifyName(name) {
  return name
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
}

function pick(arr, i) {
  return arr[i % arr.length];
}

/** Build dummy roster per section: 20 Athlete + 3 Official + 2 Coach + 1 Manager. */
function buildSampleRoster(gender) {
  const isPA = gender === 'PA';
  const athleteNames = isPA
    ? [
        'Rizky Pratama',
        'Dimas Nugroho',
        'Andi Saputra',
        'Fajar Maulana',
        'Bayu Setiawan',
        'Aditya Wijaya',
        'Rafi Alfarizi',
        'Yoga Kurniawan',
        'Gilang Ramadhan',
        'Eko Prasetyo',
        'Arif Rahman',
        'Hafiz Nurcahyo',
        'Ilham Fadilah',
        'Joko Santoso',
        'Kevin Apriliansyah',
        'Lutfi Hakim',
        'Muhammad Irfan',
        'Naufal Akbar',
        'Putra Mahendra',
        'Reza Firmansyah',
      ]
    : [
        'Aulia Rahma',
        'Siti Nurhaliza',
        'Dewi Lestari',
        'Putri Ayuningtyas',
        'Nadia Safitri',
        'Intan Permata',
        'Citra Melati',
        'Dinda Kartika',
        'Farah Anindya',
        'Gita Prameswari',
        'Hana Kusuma',
        'Indah Puspita',
        'Jasmine Aurelia',
        'Kirana Putri',
        'Larasati Dewi',
        'Mira Anggraini',
        'Nabila Zahra',
        'Olivia Maharani',
        'Putri Wulandari',
        'Rani Febrianti',
      ];
  const officialNames = isPA
    ? ['Budi Santoso', 'Dedi Kurniawan', 'Eko Wahyudi']
    : ['Maya Kusumawati', 'Sari Melati', 'Wulan Cahya'];
  const coachNames = isPA ? ['Hendra Gunawan', 'Slamet Riyadi'] : ['Rina Wulandari', 'Susi Susanti'];
  const managerNames = isPA ? ['Agus Firmansyah'] : ['Fitri Handayani'];

  const positions = ['Goalkeeper', 'Pivot', 'Flank', 'Anchor'];
  const cities = isPA
    ? ['Yogyakarta', 'Sleman', 'Bantul', 'Kulon Progo', 'Gunung Kidul']
    : ['Bantul', 'Yogyakarta', 'Sleman', 'Wonosari', 'Wates'];
  const faculties = isPA
    ? [
        'Fakultas Ilmu Keolahragaan',
        'Fakultas Ekonomika dan Bisnis',
        'Fakultas Teknik',
        'Fakultas Ilmu Sosial',
        'Fakultas MIPA',
      ]
    : [
        'Fakultas Kedokteran',
        'Fakultas Psikologi',
        'Fakultas Ilmu Budaya',
        'Fakultas Farmasi',
        'Fakultas Hukum',
      ];
  const majors = isPA
    ? ['Ilmu Keolahragaan', 'Manajemen', 'Teknik Informatika', 'Akuntansi', 'Ilmu Komunikasi']
    : ['Kedokteran', 'Psikologi', 'Pendidikan Bahasa', 'Farmasi', 'Ilmu Hukum'];
  const baju = isPA ? ['M', 'L', 'XL', 'L', 'XXL'] : ['S', 'M', 'L', 'M', 'S'];
  const celana = isPA ? ['30', '32', '34', '33', '36'] : ['26', '27', '28', '29', '30'];
  const sepatu = isPA ? [40, 41, 42, 43, 44] : [36, 37, 38, 39, 40];
  const phones = isPA
    ? ['Samsung Galaxy A54', 'iPhone 13', 'Xiaomi Redmi Note 12', 'OPPO A78', 'Vivo Y36']
    : ['iPhone 14', 'Samsung Galaxy A34', 'Xiaomi 13T', 'OPPO Reno10', 'Realme C55'];
  const banks = isPA ? ['BCA', 'Mandiri', 'BRI', 'BNI', 'BTN'] : ['BRI', 'BNI', 'BCA', 'Mandiri', 'CIMB'];
  const kendaraan = isPA
    ? ['Honda Beat', 'Yamaha NMAX', 'Honda Vario', 'Toyota Avanza', 'Suzuki Carry']
    : ['Yamaha Mio', 'Honda Scoopy', 'Honda Beat', 'Yamaha Fino', 'Honda Genio'];
  const laptops = isPA
    ? ['ASUS VivoBook', 'Lenovo IdeaPad', 'Acer Aspire', 'HP Pavilion', 'Dell Inspiron']
    : ['MacBook Air', 'ASUS Zenbook', 'Lenovo Yoga', 'Acer Swift', 'HP Envy'];
  const clubsPrev = ['Persib,2020', 'Arema,2019', 'Persebaya,2021', 'PSIS,2018', 'Bali United,2022'];
  const clubsNow = ['Persija,2025', 'PSIM,2024', 'PSS,2023', 'Persis,2025', 'Madura United,2024'];
  const smp = isPA
    ? [
        'SMP Negeri 1 Yogyakarta',
        'SMP Negeri 5 Sleman',
        'SMP Negeri 2 Bantul',
        'SMP Muhammadiyah 3 Yogya',
        'SMP Negeri 1 Wates',
      ]
    : [
        'SMP Negeri 2 Bantul',
        'SMP Muhammadiyah 1 Yogya',
        'SMP Negeri 4 Sleman',
        'SMP Negeri 1 Yogyakarta',
        'SMP Pangudi Luhur',
      ];
  const sma = isPA
    ? [
        'SMA Negeri 1 Yogyakarta',
        'SMA Negeri 3 Sleman',
        'SMA Negeri 1 Bantul',
        'SMA Negeri 8 Yogyakarta',
        'MAN 1 Yogyakarta',
      ]
    : [
        'SMA Negeri 1 Bantul',
        'SMA Negeri 8 Yogyakarta',
        'SMA Negeri 1 Sleman',
        'SMA Negeri 3 Yogyakarta',
        'MAN 2 Bantul',
      ];

  const sectionByTipe = Object.fromEntries(ROSTER_SECTIONS.map((s) => [s.tipe, s]));
  const rows = [];
  let personIdx = 0;

  const pushFull = (tipe, name, jersey, posisi, row) => {
    const i = personIdx++;
    const slug = slugifyName(name);
    const emailLocal = slug.replace(/-/g, '.');
    const nimYear = 23 + (i % 3);
    const values = {
      A: tipe,
      B: name,
      C: photoHyperlink(slug),
      D: `${emailLocal}@student.ugm.ac.id`,
      E: `0812${String(70000000 + i * 137).slice(0, 8)}`,
      F: jersey,
      G: posisi,
      H: pick(cities, i),
      I: `${String(10 + (i % 18)).padStart(2, '0')}-${String(1 + (i % 12)).padStart(2, '0')}-200${4 + (i % 3)}`,
      J: `${nimYear}5150${String(7000 + i).padStart(4, '0')}`,
      K: pick(faculties, i),
      L: pick(majors, i),
      M: 2023 + (i % 2),
      N: Number((2.8 + (i % 12) * 0.1).toFixed(1)),
      O: isPA ? 65 + (i % 15) : 48 + (i % 12),
      P: isPA ? 168 + (i % 14) : 155 + (i % 12),
      Q: slug.replace(/-/g, '').slice(0, 30),
      R: slug.replace(/-/g, '').slice(0, 24),
      S: pick(baju, i),
      T: pick(celana, i),
      U: pick(sepatu, i),
      V: pick(phones, i),
      W: pick(banks, i),
      X: pick(kendaraan, i),
      Y: pick(laptops, i),
      Z: String(10000000000 + i * 111).slice(0, 11),
      AA: pick(clubsPrev, i),
      AB: pick(clubsNow, i),
    };
    if (tipe === 'Athlete') {
      values.AC = pick(smp, i);
      values.AD = pick(sma, i);
    }
    rows.push({ row, values });
  };

  const pushStaff = (tipe, name, row) => {
    personIdx++;
    rows.push({
      row,
      values: {
        A: tipe,
        B: name,
        C: photoHyperlink(slugifyName(name)),
        G: '-',
      },
    });
  };

  const ath = sectionByTipe.Athlete;
  athleteNames.slice(0, SAMPLE_COUNTS.Athlete).forEach((name, i) => {
    pushFull('Athlete', name, i + 1, pick(positions, i), ath.dataStart + i);
  });
  const off = sectionByTipe.Official;
  officialNames.slice(0, SAMPLE_COUNTS.Official).forEach((name, i) => {
    pushFull('Official', name, 97 + i, '-', off.dataStart + i);
  });
  const coach = sectionByTipe.Coach;
  coachNames.slice(0, SAMPLE_COUNTS.Coach).forEach((name, i) => {
    pushStaff('Coach', name, coach.dataStart + i);
  });
  const mgr = sectionByTipe.Manager;
  managerNames.slice(0, SAMPLE_COUNTS.Manager).forEach((name, i) => {
    pushStaff('Manager', name, mgr.dataStart + i);
  });

  return rows;
}

const SAMPLE_ROSTERS = {
  PA: buildSampleRoster('PA'),
  PI: buildSampleRoster('PI'),
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
      } else if (
        col === 'I' ||
        col === 'E' ||
        col === 'J' ||
        col === 'Z' ||
        col === 'AA' ||
        col === 'AB' ||
        col === 'AC' ||
        col === 'AD'
      ) {
        cell.numFmt = '@';
        cell.font = { size: 11, name: 'Calibri', color: { theme: 1 } };
      } else {
        cell.font = { name: 'Calibri', color: { theme: 1 } };
      }
      if (col === 'N') {
        cell.numFmt = '0.0';
      }
    });
  });

  // Borders + unlock data rows only (skip section headers); Tipe (col A) stays locked
  rosterDataRows().forEach((r) => {
    for (let c = 1; c <= KOMPIT_LAST_COL; c++) {
      const cell = ws.getCell(r, c);
      cell.border = { ...THIN_BORDER };
      if (c === 1) {
        lockCell(cell, true);
      } else {
        lockCell(cell, false);
      }
      if (c === 5) cell.numFmt = '@'; // WhatsApp as text
      if (c === 9) cell.numFmt = '@'; // Birth date as text
      if (c === 10) cell.numFmt = '@'; // Student ID (NIM) as text
      if (c === 14) cell.numFmt = '0.0'; // GPA (IPK)
    }
  });
}

function applyRosterValidations(ws) {
  // Validations only on section data rows (not header rows between sections)
  const posisiList = `"${POSISI_OPTIONS.join(',')}"`;
  const uniqFirst = ROSTER_SECTIONS[0].dataStart;
  const uniqLast = rosterLastRow();

  rosterDataRows().forEach((r) => {
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
        `OR(D${r}="",AND(ISNUMBER(FIND("@",D${r})),ISNUMBER(FIND(".",D${r},FIND("@",D${r}))),ISERROR(FIND(" ",D${r})),COUNTIF($D$${uniqFirst}:$D$${uniqLast},D${r})=1))`,
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
        `OR(F${r}="",AND(ISNUMBER(F${r}),F${r}=INT(F${r}),F${r}>=0,F${r}<=99,COUNTIF($F$${uniqFirst}:$F$${uniqLast},F${r})=1))`,
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

    // Column K: Fakultas (required for Athlete)
    addDataValidation(ws, `K${r}`, {
      type: 'custom',
      formulae: [`OR(A${r}<>"Athlete",LEN(TRIM(K${r}))>=1)`],
      promptTitle: 'Fakultas',
      prompt: 'Wajib diisi jika Tipe = Athlete',
      errorTitle: 'Fakultas wajib',
      error: 'Fakultas wajib diisi untuk Athlete.',
    });

    // Column M: enrollment year 2023–current year
    addDataValidation(ws, `M${r}`, {
      type: 'whole',
      operator: 'between',
      formulae: [TAHUN_AWAL_MIN, CURRENT_YEAR],
      promptTitle: 'Tahun Awal',
      prompt: `${TAHUN_AWAL_MIN}–${CURRENT_YEAR}`,
      errorTitle: 'Tahun Awal tidak valid',
      error: `Tahun awal harus antara ${TAHUN_AWAL_MIN} dan ${CURRENT_YEAR}.`,
    });

    // Column N: GPA (IPK) 2.25–4.00 (e.g. 3.5)
    addDataValidation(ws, `N${r}`, {
      type: 'decimal',
      operator: 'between',
      formulae: [2.25, 4],
      promptTitle: 'IPK',
      prompt: 'Contoh: 3.5 (min 2.25, max 4.00)',
      errorTitle: 'IPK tidak valid',
      error: 'IPK harus desimal 2.25–4.00 (contoh: 3.5).',
    });

    // Column O: weight (kg)
    addDataValidation(ws, `O${r}`, {
      type: 'whole',
      operator: 'between',
      formulae: [30, 200],
      promptTitle: 'Berat Badan',
      prompt: 'kg (30–200)',
      errorTitle: 'Berat tidak valid',
      error: 'Berat badan harus angka 30–200 kg.',
    });

    // Column P: height (cm)
    addDataValidation(ws, `P${r}`, {
      type: 'whole',
      operator: 'between',
      formulae: [100, 250],
      promptTitle: 'Tinggi Badan',
      prompt: 'cm (100–250)',
      errorTitle: 'Tinggi tidak valid',
      error: 'Tinggi badan harus angka 100–250 cm.',
    });

    // Column Q: Instagram username
    addDataValidation(ws, `Q${r}`, {
      type: 'custom',
      formulae: [
        `OR(Q${r}="",AND(LEN(Q${r})>=1,LEN(Q${r})<=30,ISERROR(FIND(" ",Q${r})),ISERROR(FIND("@",Q${r}))))`,
      ],
      promptTitle: 'Instagram',
      prompt: 'Username tanpa @, maks 30 karakter',
      errorTitle: 'Instagram tidak valid',
      error: 'Username tanpa spasi/@, maksimal 30 karakter.',
    });

    // Column R: TikTok username
    addDataValidation(ws, `R${r}`, {
      type: 'custom',
      formulae: [
        `OR(R${r}="",AND(LEN(R${r})>=1,LEN(R${r})<=24,ISERROR(FIND(" ",R${r})),ISERROR(FIND("@",R${r}))))`,
      ],
      promptTitle: 'TikTok',
      prompt: 'Username tanpa @, maks 24 karakter',
      errorTitle: 'TikTok tidak valid',
      error: 'Username tanpa spasi/@, maksimal 24 karakter.',
    });
  });
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

  // Lock row 1 + section headers + spacer rows; data rows unlocked in apply* above
  for (let c = 1; c <= CL_LAST_COL; c++) {
    lockCell(ws.getCell(1, c), true);
  }
  rosterHeaderRows().forEach((hr) => {
    for (let c = 1; c <= CL_LAST_COL; c++) {
      lockCell(ws.getCell(hr, c), true);
    }
  });
  rosterSpacerRows().forEach((sr) => {
    for (let c = 1; c <= CL_LAST_COL; c++) {
      lockCell(ws.getCell(sr, c), true);
    }
  });

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
 * Campus League roster columns (S..AD): header, widths, sample values,
 * borders/unlock for the data rows and validations. Mirrors what Kompit's
 * applyRosterHeaders / applySampleRows / applyRosterValidations do for A..R.
 */
function applyClRosterColumns(ws, sheetName) {
  let needsWrappedHeader = false;
  const dataRows = rosterDataRows();

  CL_ROSTER_COLUMNS.forEach((col, i) => {
    const c = KOMPIT_LAST_COL + 1 + i;
    const colLetter = ws.getColumn(c).letter;

    ROSTER_SECTIONS.forEach((section) => {
      const header = ws.getCell(section.headerRow, c);
      const showRequired =
        col.requiredForAthlete && section.tipe === 'Athlete';
      header.value = showRequired ? requiredHeader(col.header) : col.header;
      if (!showRequired) {
        header.font = { name: 'Cambria', bold: true, color: { theme: 1 }, size: 11 };
      }
      header.border = { ...THIN_BORDER };
      lockCell(header, true);
      if (col.wrapHeader) {
        header.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
        needsWrappedHeader = true;
      }
    });
    ws.getColumn(c).width = col.width;

    dataRows.forEach((r) => {
      const cell = ws.getCell(r, c);
      cell.border = { ...THIN_BORDER };
      cell.font = { name: 'Calibri', color: { theme: 1 } };
      if (col.text) cell.numFmt = '@';
      lockCell(cell, false);
    });

    // Sample CL values are filled by applySampleRows (S..AD on example rows).

    if (col.header === 'Ukuran Sepatu') {
      dataRows.forEach((r) => {
        addDataValidation(ws, `${colLetter}${r}`, {
          type: 'whole',
          operator: 'between',
          formulae: [30, 50],
          promptTitle: 'Ukuran Sepatu',
          prompt: 'Angka 30–50 (contoh: 42)',
          errorTitle: 'Ukuran sepatu tidak valid',
          error: 'Ukuran sepatu harus angka 30–50.',
        });
      });
    }

    if (col.header === 'Nomor Kepesertaan BPJSTK') {
      dataRows.forEach((r) => {
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
      });
    }

    if (col.clubYear) {
      dataRows.forEach((r) => {
        const ref = `${colLetter}${r}`;
        addDataValidation(ws, ref, {
          type: 'custom',
          formulae: [formulaClubYear(ref)],
          promptTitle: 'Asal Club',
          prompt: 'Opsional. Format: nama club,tahun (contoh: Persib,2020)',
          errorTitle: 'Format club tidak valid',
          error: `Gunakan format <club>,<tahun> (contoh: Persib,2020). Tahun 1950–${CURRENT_YEAR + 1}. Boleh dikosongkan.`,
        });
      });
    }

    if (col.requiredForAthlete) {
      dataRows.forEach((r) => {
        const ref = `${colLetter}${r}`;
        // Required only when Tipe = Athlete; other roles / empty rows may stay blank
        addDataValidation(ws, ref, {
          type: 'custom',
          formulae: [`OR(A${r}<>"Athlete",LEN(TRIM(${ref}))>=1)`],
          promptTitle: col.header,
          prompt: 'Wajib diisi jika Tipe = Athlete',
          errorTitle: `${col.header} wajib`,
          error: `${col.header} wajib diisi untuk Athlete.`,
        });
      });
    }
  });

  if (needsWrappedHeader) {
    rosterHeaderRows().forEach((hr) => {
      ws.getRow(hr).height = 36;
    });
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

  const sourceSheetOpts = {
    ...opts,
    selectLockedCells: true,
    selectUnlockedCells: false,
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
    if (ws.name === 'Wilayah' || ws.name === 'Universitas') {
      // Dropdown sources: every cell locked, no unlocked cells to edit
      await ws.protect(password, sourceSheetOpts);
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

  // Sheet order: instructions → TEAM → PA → PI → Wilayah / Universitas (dropdown sources)
  buildPetunjukSheet(wb);
  buildTeamSheet(wb);
  buildRosterSheet(wb, 'PA');
  buildRosterSheet(wb, 'PI');
  buildRefSheet(wb);
  buildUniversitiesSheet(wb);

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
  PA/PI section data rows (Athlete/Official/Coach/Manager)
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
  console.log(
    'Editable: TEAM!D1:D4 + G1:G2 + F8:I9 , PA/PI section data rows (Wilayah locked)',
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
