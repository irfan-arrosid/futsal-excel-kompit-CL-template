# Futsal Excel Kompit CL Template

Node.js generator for the Kompit Futsal CL team roster Excel template.

## Features

- Sheet protection with limited editable ranges
  - **TEAM**: `D1:D4` (identitas) + `G1:G2` (Campus League) + `F8:I9` (Partisipasi & warna kostum)
  - **PA / PI**: `A3:Q16` (Kompit) + `R3:AA16` (Campus League)
  - **Wilayah**: fully locked (dropdown source)
- Indonesian province & cascading city dropdowns (via `FILTER` helper on `Wilayah`)
- Hardened roster validations (WhatsApp phone, jersey number, NIM, IPK, enrollment year, etc.)
- Campus League columns (grouped under `CL_*` / `applyCl*` in `generate.js`):
  - **PA / PI** `R:AA` — Ukuran Baju/Celana/Sepatu, Merk HP/Kendaraan/Laptop, Nama Bank, BPJSTK, Asal Club sebelumnya & saat ini (`<club>,<tahun>`, opsional)
  - **TEAM** `F1:G2` — Cabang Olahraga (dropdown) & Wilayah; tabel tim `B6:I9` — kolom abu-abu otomatis (Nama Tim, Singkatan, Kategori Tim) + Partisipasi & warna kostum

## Requirements

- Node.js >= 18

## Usage

```bash
npm install
npm run generate
```

Output: `output/Template Futsal Kompit - CL.xlsx`

```bash
node generate.js --out /path/to/file.xlsx
node generate.js --password mysecret
```

Default sheet-unprotect password: `kompit`

## Project layout

| Path | Description |
|------|-------------|
| `generate.js` | Template generator (ExcelJS) |
| `data/indonesia-regions.json` | Province → city map |
| `output/` | Generated `.xlsx` (gitignored) |

## License

Private / internal Kompit use.
