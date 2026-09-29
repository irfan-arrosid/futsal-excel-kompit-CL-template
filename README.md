# Futsal Excel Kompit CL Template

Node.js generator for the Kompit Futsal CL team roster Excel template.

## Features

- Sheet protection with limited editable ranges
  - **TEAM**: `D1:D4` (identitas) + `G1:G2` (Campus League) + `F8:I9` (Ikut/Tidak Ikut & warna kostum)
  - **PA / PI**: sectioned roster — Athlete / Official / Coach / Manager each with own header; dummy 20 / 3 / 2 / 1; Kompit `A:R` + Campus League `S:AD`
  - **Wilayah**: fully locked (dropdown source)
- Indonesian province & cascading city dropdowns (COUNTIF rank + INDEX/MATCH on `Wilayah`)
- Nama Universitas dropdown dari `data/daftar-kampus.xlsx` (`Universitas` sheet / `data/indonesia-universities.json`)
- Hardened roster validations (WhatsApp phone, jersey number, NIM, IPK, enrollment year, etc.)
- Campus League columns (grouped under `CL_*` / `applyCl*` in `generate.js`):
  - **PA / PI** `S:AD` — Ukuran Baju/Celana/Sepatu, Merk HP/Kendaraan/Laptop, Nama Bank, BPJSTK, Asal Club sebelumnya & saat ini (`<club>,<tahun>`, opsional), Asal SMP*/SMA* (wajib Athlete)
  - **TEAM** `F1:G2` — Cabang Olahraga (dropdown) & Wilayah; tabel tim `B6:I9` — kolom abu-abu otomatis (Nama Tim, Singkatan, Kategori Tim) + Ikut/Tidak Ikut & warna kostum
  - **Wilayah** / **Universitas** — locked dropdown source sheets

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

Refresh university list from daftar kampus Excel:

```bash
npm run fetch:universities
# or
node scripts/fetch-universities.js --from ./data/daftar-kampus.xlsx
npm run generate
```

Default sheet-unprotect password: `kompit`

## Project layout

| Path | Description |
|------|-------------|
| `generate.js` | Template generator (ExcelJS) |
| `data/indonesia-regions.json` | Province → city map |
| `data/daftar-kampus.xlsx` | Source list of universities (Nama, Singkatan, NPSN) |
| `data/indonesia-universities.json` | University names used by TEAM!D1 dropdown |
| `scripts/fetch-universities.js` | Rebuild university JSON from daftar-kampus.xlsx |
| `output/` | Generated `.xlsx` (gitignored) |

## License

Private / internal Kompit use.
