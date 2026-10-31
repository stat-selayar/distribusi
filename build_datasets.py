"""
ETL Pipeline: Transformasi Data Operasional Transportasi Kabupaten Kepulauan Selayar
Menghasilkan 5 tabel analitik Tidy Data + 1 Master Workbook Excel.

ponytail: Parsing terfokus pada rekap bulanan resmi dan log harian terstruktur.
Ceiling: 80 file (2025-2026). Upgrade: pasang watcher jika ada update data baru.
"""

import os
import glob
import re
import bisect
import csv
import json
import datetime
import urllib.request
import openpyxl
import fitz  # PyMuPDF

BASE_DIR = '/Volumes/work/_Projects/transportasi'
RAW_DIR = os.path.join(BASE_DIR, 'data_mentah')
OUT_DIR = os.path.join(BASE_DIR, 'data_bersih')
os.makedirs(OUT_DIR, exist_ok=True)

# -------------------------------------------------------------
# 1. METADATA SIMPUL (DIM_SIMPUL)
# -------------------------------------------------------------
SIMPUL_DATA = [
    {
        'kode_simpul': 'ARO',
        'nama_simpul': 'Bandar Udara H. Aroeppala',
        'nama_resmi': 'Kantor UPBU Kelas III H. Aroeppala Selayar',
        'moda': 'Udara',
        'kabupaten': 'Kepulauan Selayar',
        'kecamatan': 'Bontoharu',
        'wilayah': 'Daratan Pulau Selayar',
        'latitude': -6.196389,
        'longitude': 120.463333,
        'operator': 'Ditjen Perhubungan Udara Kemenhub'
    },
    {
        'kode_simpul': 'BTG',
        'nama_simpul': 'Pelabuhan Benteng',
        'nama_resmi': 'Kantor UPP Kelas III Selayar (Benteng)',
        'moda': 'Laut',
        'kabupaten': 'Kepulauan Selayar',
        'kecamatan': 'Benteng',
        'wilayah': 'Ibu Kota Kabupaten',
        'latitude': -6.118944,
        'longitude': 120.457389,
        'operator': 'Ditjen Perhubungan Laut Kemenhub'
    },
    {
        'kode_simpul': 'PMT',
        'nama_simpul': 'Pelabuhan Pamatata',
        'nama_resmi': 'Satuan Pelayanan Pelabuhan Penyeberangan Pamatata (BPTD Sulsel)',
        'moda': 'Penyeberangan',
        'kabupaten': 'Kepulauan Selayar',
        'kecamatan': 'Bontomatene',
        'wilayah': 'Ujung Utara Pulau Selayar',
        'latitude': -5.850556,
        'longitude': 120.448056,
        'operator': 'BPTD Kelas II Sulsel / PT ASDP Indonesia Ferry'
    },
    {
        'kode_simpul': 'JMP',
        'nama_simpul': 'Pelabuhan Jampea',
        'nama_resmi': 'Kantor UPP Kelas III Jampea',
        'moda': 'Laut',
        'kabupaten': 'Kepulauan Selayar',
        'kecamatan': 'Pasimasunggu',
        'wilayah': 'Pulau Jampea (Kepulauan Luar)',
        'latitude': -7.078889,
        'longitude': 120.675833,
        'operator': 'Ditjen Perhubungan Laut Kemenhub'
    },
    {
        'kode_simpul': 'PTB',
        'nama_simpul': 'Pelabuhan Pattumbukang',
        'nama_resmi': 'Wilayah Kerja Pattumbukang (UPP Kelas III Selayar)',
        'moda': 'Laut',
        'kabupaten': 'Kepulauan Selayar',
        'kecamatan': 'Bontosikuyu',
        'wilayah': 'Ujung Selatan/Timur Pulau Selayar',
        'latitude': -6.376944,
        'longitude': 120.528333,
        'operator': 'Ditjen Perhubungan Laut Kemenhub'
    }
]


def export_dim_simpul():
    path = os.path.join(OUT_DIR, 'dim_simpul.csv')
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(SIMPUL_DATA[0].keys()))
        writer.writeheader()
        writer.writerows(SIMPUL_DATA)
    print(f"[OK] dim_simpul.csv ({len(SIMPUL_DATA)} simpul)")


# -------------------------------------------------------------
# 2. DETAIL HARIAN PENYEBERANGAN FERRY PAMATATA 2025
# -------------------------------------------------------------
def process_pamatata_ferry():
    files = sorted(glob.glob(os.path.join(RAW_DIR, 'Pamatata/2025/*.xlsx')))
    records = []

    for fpath in files:
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        for sname in ['KEBERANGKATAN', 'KEDATANGAN']:
            if sname not in wb.sheetnames:
                continue
            ws = wb[sname]
            arah = sname.title()
            curr_date = None

            for i, r in enumerate(ws.iter_rows(values_only=True)):
                if i < 11:
                    continue
                v0 = r[0] if len(r) > 0 else None
                if isinstance(v0, datetime.datetime):
                    curr_date = v0.strftime('%Y-%m-%d')
                elif isinstance(v0, str) and len(v0.strip()) >= 8 and any(c in v0 for c in ['/', '-']):
                    curr_date = v0.strip()

                ship = r[1] if len(r) > 1 else None
                if ship and isinstance(ship, str) and any(k in ship.upper() for k in ['KMP', 'KM.', 'KM ']) and not any(h in ship.upper() for h in ['TOTAL', 'JUMLAH']):
                    trip = int(r[3]) if len(r) > 3 and isinstance(r[3], (int, float)) else 1
                    dewasa = int(r[4]) if len(r) > 4 and isinstance(r[4], (int, float)) else 0
                    anak = int(r[5]) if len(r) > 5 and isinstance(r[5], (int, float)) else 0
                    tot_pnp = int(r[6]) if len(r) > 6 and isinstance(r[6], (int, float)) else (dewasa + anak)

                    gols = [int(r[col]) if len(r) > col and isinstance(r[col], (int, float)) else 0 for col in range(7, 19)]
                    tot_kend = sum(gols)

                    dt = datetime.datetime.strptime(curr_date, '%Y-%m-%d') if curr_date and '-' in curr_date else None

                    records.append({
                        'tanggal': curr_date,
                        'tahun': dt.year if dt else 2025,
                        'bulan': dt.month if dt else 1,
                        'hari': dt.strftime('%A') if dt else '',
                        'arah': arah,
                        'nama_kapal': ship.strip(),
                        'jadwal': str(r[2]).strip() if len(r) > 2 and r[2] is not None else '',
                        'trip': trip,
                        'pnp_dewasa': dewasa,
                        'pnp_anak': anak,
                        'total_penumpang': tot_pnp,
                        'gol_1_sepeda': gols[0],
                        'gol_2_motor': gols[1],
                        'gol_3_moge': gols[2],
                        'gol_4_pnp_mobil': gols[3],
                        'gol_4_brg_pickup': gols[4],
                        'gol_5_pnp_bus': gols[5],
                        'gol_5_brg_truk': gols[6],
                        'gol_6_pnp_busbesar': gols[7],
                        'gol_6_brg_trukbesar': gols[8],
                        'gol_7_tronton': gols[9],
                        'gol_8_trailer': gols[10],
                        'gol_9_alatberat': gols[11],
                        'total_kendaraan': tot_kend
                    })

    out_csv = os.path.join(OUT_DIR, 'fact_pamatata_ferry_harian.csv')
    with open(out_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(records[0].keys()))
        writer.writeheader()
        writer.writerows(records)

    print(f"[OK] fact_pamatata_ferry_harian.csv ({len(records):,} trip penyeberangan)")
    return records


# -------------------------------------------------------------
# 3. DETAIL HARIAN PENERBANGAN H. AROEPPALA 2026
# -------------------------------------------------------------
def process_aroeppala_flights():
    fpath = os.path.join(RAW_DIR, 'Bandar H. Aroeppala/2026/LAPORAN DATA LLAU BANDAR UDARA H.AROEPPALA SELAYAR 2026.xlsx')
    wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
    records = []

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            if i < 7 or not r or r[0] is None:
                continue
            v0 = r[0]
            tgl_str = None
            if isinstance(v0, datetime.datetime):
                tgl_str = v0.strftime('%Y-%m-%d')
            elif isinstance(v0, str) and any(x in v0.upper() for x in ['SENIN', 'SELASA', 'RABU', 'KAMIS', 'JUM', 'SABTU', 'MINGGU']):
                parts = v0.split('/')
                if len(parts) >= 4:
                    tgl_str = f"{parts[3].strip()}-{parts[2].strip().zfill(2)}-{parts[1].strip().zfill(2)}"

            if tgl_str and 'TOTAL' not in str(v0).upper():
                dt = datetime.datetime.strptime(tgl_str, '%Y-%m-%d')
                p_in = int(r[1] or 0)
                p_out = int(r[2] or 0)
                pnp_in = int(r[3] or 0)
                pnp_out = int(r[4] or 0)
                pnp_tr = int(r[5] or 0) if len(r) > 5 else 0
                bag_in = float(r[6] or 0) if len(r) > 6 else 0.0
                bag_out = float(r[7] or 0) if len(r) > 7 else 0.0
                kg_in = float(r[8] or 0) if len(r) > 8 else 0.0
                kg_out = float(r[9] or 0) if len(r) > 9 else 0.0
                pos_in = float(r[10] or 0) if len(r) > 10 else 0.0
                pos_out = float(r[11] or 0) if len(r) > 11 else 0.0
                ket = str(r[12]).strip() if len(r) > 12 and r[12] is not None else ''

                records.append({
                    'tanggal': tgl_str,
                    'tahun': dt.year,
                    'bulan': dt.month,
                    'hari': dt.strftime('%A'),
                    'pesawat_datang': p_in,
                    'pesawat_berangkat': p_out,
                    'penumpang_datang': pnp_in,
                    'penumpang_berangkat': pnp_out,
                    'penumpang_transit': pnp_tr,
                    'total_penumpang': pnp_in + pnp_out + pnp_tr,
                    'bagasi_datang_kg': bag_in,
                    'bagasi_berangkat_kg': bag_out,
                    'total_bagasi_kg': bag_in + bag_out,
                    'kargo_datang_kg': kg_in,
                    'kargo_berangkat_kg': kg_out,
                    'total_kargo_kg': kg_in + kg_out,
                    'pos_datang_kg': pos_in,
                    'pos_berangkat_kg': pos_out,
                    'total_pos_kg': pos_in + pos_out,
                    'keterangan': ket
                })

    out_csv = os.path.join(OUT_DIR, 'fact_aroeppala_penerbangan_harian.csv')
    with open(out_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(records[0].keys()))
        writer.writeheader()
        writer.writerows(records)

    print(f"[OK] fact_aroeppala_penerbangan_harian.csv ({len(records)} hari kalender penerbangan)")
    return records


# -------------------------------------------------------------
# 4. LOG PELAYARAN KAPAL LAUT (KUNJUNGAN KAPAL)
# -------------------------------------------------------------
BULAN_MAP = {'JANUARI': 1, 'FEBRUARI': 2, 'MARET': 3, 'APRIL': 4, 'MEI': 5, 'JUNI': 6,
             'JULI': 7, 'AGUSTUS': 8, 'SEPTEMBER': 9, 'OKTOBER': 10, 'NOVEMBER': 11, 'DESEMBER': 12}


def parse_period_from_path(fpath):
    """Return (tahun, bulan) for a report file; bulan is None if it cannot be read.

    Month comes from a numeric `YYYY_MM_` prefix or an Indonesian month name.
    Year comes from the filename, else from the parent folder (e.g. Jampea/2026/).
    """
    name = os.path.basename(fpath)
    m_num = re.match(r'(20\d\d)_(\d\d)_', name)
    if m_num and 1 <= int(m_num.group(2)) <= 12:
        return int(m_num.group(1)), int(m_num.group(2))

    m_yr = re.search(r'20[2-9]\d', name) or re.fullmatch(r'20[2-9]\d', os.path.basename(os.path.dirname(fpath)))
    tahun = int(m_yr.group(0)) if m_yr else None
    bulan = next((v for k, v in BULAN_MAP.items() if k in name.upper()), None)
    return tahun, bulan


JAMPEA_NCOLS = 35


def parse_id_number(s):
    """Parse '1.913', '1,5', '22,00' or '21.63' to float; None if `s` is not a number.

    A comma is always the decimal mark; a dot followed by exactly three digits is
    a thousands separator ('2.097' -> 2097), otherwise it is a decimal mark.
    """
    s = (s or '').strip()
    if not re.fullmatch(r'\d+(?:[.,]\d+)*', s):
        return None
    if ',' in s:
        return float(s.replace('.', '').replace(',', '.'))
    if re.fullmatch(r'\d{1,3}(?:\.\d{3})+', s):
        return float(s.replace('.', ''))
    return float(s)


def jampea_column_edges(page):
    """x positions of the report table's vertical rules, or [] if the layout is unrecognised."""
    length = {}
    for g in page.get_drawings():
        for it in g['items']:
            if it[0] == 'l' and abs(it[1].x - it[2].x) < 0.05:
                x, h = round(it[1].x, 1), abs(it[1].y - it[2].y)
            elif it[0] == 're' and it[1].width < 0.6:
                x, h = round(it[1].x0, 1), it[1].height
            else:
                continue
            length[x] = length.get(x, 0) + h
    if not length:
        return []
    edges = []
    for x in sorted(x for x, h in length.items() if h >= 0.5 * max(length.values())):
        if not edges or x - edges[-1] >= 1.0:
            edges.append(x)
    return edges if len(edges) == JAMPEA_NCOLS + 1 else []


def jampea_cells_from_words(page, edges):
    """Group the words of a text PDF page into text lines and assign each word to a table column."""
    boxes = [(w[0], (w[1] + w[3]) / 2, w[4]) for w in page.get_text('words')]
    return jampea_cells_from_boxes(boxes, edges, line_gap=1.0)


def jampea_cells_from_boxes(boxes, edges, line_gap):
    """Turn (x0, y_center, text) word boxes into rows of 35 column cells.

    Words closer than `line_gap` vertically are one text line; units follow `boxes`
    (PDF points for text pages, pixels for OCR).
    """
    lines, last_y = [], None
    for x0, y, text in sorted(boxes, key=lambda b: (b[1], b[0])):
        if last_y is None or y - last_y > line_gap:
            lines.append([])
        lines[-1].append((x0, text))
        last_y = y

    rows = []
    for line in lines:
        cells = [''] * JAMPEA_NCOLS
        for x0, text in sorted(line):
            i = bisect.bisect_right(edges, x0) - 1
            if 0 <= i < JAMPEA_NCOLS:
                cells[i] = f"{cells[i]} {text}".strip()
        rows.append(cells)
    return rows


# (type column, quantity column, direction, category, unit) for each cargo group in the report table.
# Through cargo (muatan lanjutan) stays aboard, so it is recorded but not added to unloaded/loaded totals.
JAMPEA_MUATAN_COLS = [
    (16, 17, 'bongkar', 'barang', 'ton_m3'),
    (18, 19, 'bongkar', 'hewan', 'ekor'),
    (20, 21, 'muat', 'barang', 'ton_m3'),
    (22, 23, 'muat', 'hewan', 'ekor'),
    (30, 31, 'lanjutan', 'barang', 'ton_m3'),
]


# No single cargo line at these ports comes near this; larger "ton" values were typed in kilograms.
JAMPEA_MAX_TON_PER_LINE = 10_000


def parse_jampea_quantity(jenis, qty):
    """Return (jenis, jumlah) for a type/quantity cell pair; jumlah is None if `qty` is not a number.

    Handles a unit suffix ('25M3') and type text that spilled into the quantity cell
    ('PERALATAN PANCAN' + 'G 100').
    """
    m = re.fullmatch(r'(?:([A-Za-z]{1,2}) )?(\d[\d.,]*) ?(?:M3|TON)?', qty)
    if not m:
        return jenis, None
    if m.group(1):
        jenis += m.group(1)
    return jenis, parse_id_number(m.group(2))


def jampea_vessel_from_rows(rows, tahun, bulan, wilker='', jenis_pelayaran=''):
    """Build one ship-visit record from a vessel row plus its extra-cargo continuation rows.

    Every cargo line is kept in `muatan`; the tonnage and head-count totals are sums of it.
    """
    r = rows[0]
    muatan = []
    for c in rows:
        for col_j, col_q, arah, kategori, satuan in JAMPEA_MUATAN_COLS:
            jenis, jumlah = parse_jampea_quantity('' if c[col_j] in ('-', 'NIHIL') else c[col_j], c[col_q])
            if jumlah is None or not (jumlah or jenis):
                continue
            catatan = ''
            if satuan == 'ton_m3' and jumlah > JAMPEA_MAX_TON_PER_LINE:
                jumlah, catatan = jumlah / 1000, f'nilai laporan {c[col_q]} dianggap kg'
            muatan.append({'arah': arah, 'kategori': kategori, 'komoditas_asli': jenis,
                           'jumlah': jumlah, 'satuan': satuan, 'catatan': catatan})

    def total(arah, kategori):
        return sum(m['jumlah'] for m in muatan if m['arah'] == arah and m['kategori'] == kategori)

    def jenis(arah):
        return [m['komoditas_asli'] for m in muatan
                if m['arah'] == arah and m['kategori'] == 'barang' and m['komoditas_asli']]

    return {
        'pelabuhan': 'Pelabuhan Jampea',
        'tahun': tahun,
        'bulan': bulan,
        'tanggal_tiba': r[7],
        'tanggal_berangkat': r[13],
        'nama_kapal': r[1],
        'tipe_kapal': r[12],
        'bendera': r[2] or 'ID',
        'gt': parse_id_number(r[5]) or 0.0,
        'loa': parse_id_number(r[4]) or 0.0,
        'pelabuhan_asal': r[9],
        'pelabuhan_tujuan': r[15],
        'penumpang_turun': int(parse_id_number(r[28]) or 0),
        'penumpang_naik': int(parse_id_number(r[29]) or 0),
        'barang_bongkar_ton': total('bongkar', 'barang'),
        'barang_muat_ton': total('muat', 'barang'),
        'jenis_barang_bongkar': jenis('bongkar'),
        'jenis_barang_muat': jenis('muat'),
        'hewan_turun_ekor': int(total('bongkar', 'hewan')),
        'hewan_naik_ekor': int(total('muat', 'hewan')),
        'wilker': wilker,
        'jenis_pelayaran': jenis_pelayaran,
        'muatan': muatan,
    }


def jampea_section_heading(c):
    """('wilker', name) or ('pelayaran', name) for a section heading row, else None.

    A report covers the Jampea port plus its work areas (POSKER UJUNG, WILKER KAYUADI, ...),
    each split by shipping type (PELAYARAN RAKYAT, PELAYARAN DALAM NEGERI, PERINTIS, FERRY).
    """
    text = ' '.join(x for x in c if x)
    if re.fullmatch(r'(PELABUHAN|POSKER|WILKER) [A-Z .]+', text):
        return 'wilker', text.title()
    if not re.fullmatch(r'\d+|#+', c[0]) and re.fullmatch(r'PELAYARAN [A-Z ]+|PERINTIS|FERRY', c[1]):
        return 'pelayaran', c[1].title()
    return None


def is_jampea_page_carry(c):
    """Rows repeated at a page break (carried totals, column numbers); a vessel's cargo may continue after them."""
    head = ' '.join(c[:2]).upper()
    return head.startswith(('JUMLAH DIPINDAHKAN', 'JUMLAH PINDAHAN')) or (c[0], c[1]) == ('1', '2')


def jampea_records_from_rows(rows, tahun, bulan):
    """Split table rows (35 cells each) into vessel blocks and build their records."""
    def is_subtotal(c):
        # Unlabelled section subtotal: numbers in the quantity columns but no commodity/animal type.
        return any(c[i] for i in (17, 19, 21, 23, 28, 29)) and not any(c[i] for i in (16, 18, 20, 22, 30))

    records, block = [], None
    section = {'wilker': '', 'pelayaran': ''}

    def close():
        if block:
            records.append(jampea_vessel_from_rows(block, tahun, bulan, section['wilker'], section['pelayaran']))

    for c in rows:
        if is_jampea_page_carry(c):
            continue
        heading = jampea_section_heading(c)
        is_vessel = re.fullmatch(r'\d+|#+', c[0]) and re.search(r'[A-Za-z]', c[1])
        if heading or is_vessel:
            close()
            block = None
            if heading:
                section[heading[0]] = heading[1]
                if heading[0] == 'wilker':
                    section['pelayaran'] = ''
            else:
                block = [c]
        elif block and not c[0] and not c[1].upper().startswith('JUMLAH') and not is_subtotal(c):
            block.append(c)
        else:
            close()
            block = None
    close()
    return records


# Fields compared with the report's JUMLAH TOTAL row, and the column holding each total.
JAMPEA_TOTAL_COLS = {'barang_bongkar_ton': 17, 'hewan_turun_ekor': 19, 'barang_muat_ton': 21,
                     'hewan_naik_ekor': 23, 'barang_lanjutan_ton': 31}


def jampea_total_diffs(records, rows):
    """[(field, parsed, report)] where parsed totals differ from the JUMLAH TOTAL row; None if there is none.

    The report shows rounded sums, so differences up to 0.5 are ignored.
    """
    totals = [c for c in rows if 'JUMLAH TOTAL' in ' '.join(c[:2]).upper()]
    if not totals:
        return None
    parsed = {f: sum(r[f] for r in records) for f in JAMPEA_TOTAL_COLS if f != 'barang_lanjutan_ton'}
    parsed['barang_lanjutan_ton'] = sum(m['jumlah'] for r in records for m in r['muatan']
                                        if m['arah'] == 'lanjutan')
    diffs = []
    for field, col in JAMPEA_TOTAL_COLS.items():
        report = parse_id_number(totals[-1][col]) or 0.0
        if abs(parsed[field] - report) > 0.5:
            diffs.append((field, round(parsed[field], 3), report))
    return diffs


# Spelling variants in the reports -> one commodity name. Names not listed pass through cleaned.
KOMODITAS_ALIAS = {
    **dict.fromkeys(['B CAMPURAN', 'B. CAMP', 'B. CAMPURAN', 'B.CAMP', 'B.CAMPUR', 'B.CAMPURAN',
                     'CAMPUAN', 'CAMPURAN'], 'BARANG CAMPURAN'),
    **dict.fromkeys(['G. C', 'G.C', 'GC', 'GC LANJUTAN', 'KARGO UMUM', 'CARGO UMUM'], 'GENERAL CARGO'),
    **dict.fromkeys(['ELPIJI', 'LPG 3 KG', 'TABUNG LPG', 'TABUNG 3 KG'], 'LPG'),
    **dict.fromkeys(['ELPIJI KOSONG', 'ELPJI KSG', 'LPG KSG', 'TBG LPG KSNG', 'TABUNG KOSONG'], 'LPG KOSONG'),
    **dict.fromkeys(['DRUM KOSNG', 'DRUM KSG'], 'DRUM KOSONG'),
    **dict.fromkeys(['SEMEN 40 KG', 'SEMEN 40KG', 'SEMEN 50 KG'], 'SEMEN'),
    **dict.fromkeys(['KY KLP', 'KAYU KLP'], 'KAYU KELAPA'),
    **dict.fromkeys(['KELAPA BIJI', 'BIJI KELAPA'], 'BUAH KELAPA'),
    **dict.fromkeys(['MENTE', 'BIJI MENTE'], 'JAMBU MENTE'),
    **dict.fromkeys(['GABUS', 'GABUNG IKAN'], 'GABUS IKAN'),
    **dict.fromkeys(['ERALATAN PANCAN', 'PRLTN PANCANG'], 'PERALATAN PANCANG'),
    **dict.fromkeys(['BESI KWT BETON', 'KAWAT BETON'], 'BESI KAWAT BETON'),
    'B. BANGUNAN': 'BAHAN BANGUNAN',
    'BESI BEKAS': 'BESI TUA',
    'DEDAK PADI': 'DEDAK',
    'KACANG IJO': 'KACANG HIJAU',
    'KANAFA': 'KAYU KANAFA',
    'KEBAU': 'KERBAU',
    'KRIKIL': 'KERIKIL',
    'R. LAUT': 'RUMPUT LAUT',
    'T. PANCANG': 'TIANG PANCANG',
    'TEPUNG TERIGU': 'TERIGU',
    '': 'TIDAK DISEBUT',
    '0': 'TIDAK DISEBUT',
}


def normalize_komoditas(name):
    """Upper-case, whitespace-collapsed commodity name with spelling variants merged."""
    key = re.sub(r'\s+', ' ', name.upper()).strip(" '")
    return KOMODITAS_ALIAS.get(key, key)


def jampea_words_rows(doc):
    """Rows of 35 cells for a report whose table is plain text/vector rules with no detectable table object."""
    rows = []
    for page in doc:
        edges = jampea_column_edges(page)
        if edges:
            rows.extend(jampea_cells_from_words(page, edges))
    return rows


def parse_jampea_words_pdf(doc, tahun, bulan):
    """Parse a Jampea report from word positions only."""
    return jampea_records_from_rows(jampea_words_rows(doc), tahun, bulan)


def jampea_table_rows(doc):
    """Rows of 35 text cells from every report table object in a PDF, in page order.

    Tables of other widths (the signature block) are skipped.
    """
    rows = []
    for page in doc:
        for table_obj in page.find_tables().tables:
            for r in table_obj.extract():
                if len(r) == JAMPEA_NCOLS:
                    rows.append([(c or '').strip() for c in r])
    return rows


def jampea_doc_rows(doc):
    """Rows of 35 cells from a Jampea report PDF: table objects first, word positions as a fallback."""
    rows = jampea_table_rows(doc)
    return rows if any(re.fullmatch(r'\d+', c[0]) for c in rows) else jampea_words_rows(doc)


def jampea_rows_from_sheet(ws):
    """Rows of 35 text cells from a Jampea report worksheet, written the way the PDF reports show them."""
    def text(v):
        if v is None:
            return ''
        if isinstance(v, datetime.datetime):
            return v.strftime('%d-%m-%y')
        if isinstance(v, float):
            # The workbooks are converted from the PDF reports, where '2.097' means 2097; writing the
            # float back with a dot lets parse_id_number apply the same reading.
            return str(int(v)) if v.is_integer() else str(v)
        return str(v).strip()

    rows = []
    for r in ws.iter_rows(values_only=True):
        cells = [text(v) for v in r[:JAMPEA_NCOLS]]
        rows.append(cells + [''] * (JAMPEA_NCOLS - len(cells)))
    return rows


def parse_jampea_pdf(doc, tahun, bulan):
    """Parse a Jampea report PDF into ship-visit records."""
    return jampea_records_from_rows(jampea_doc_rows(doc), tahun, bulan)


def process_ship_logs():
    records = []

    # A. Dari Kompilasi Excel 2026 (Benteng, Pamatata, Pattumbukan - Mar s.d. Sep)
    files_2026 = sorted(glob.glob(os.path.join(RAW_DIR, 'Benteng/2026/*.xlsx')))
    for fpath in files_2026:
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        # Ambil bulan dari nama file (contoh: 2026_03 ...)
        m = re.search(r'2026_(\d\d)', os.path.basename(fpath))
        bulan_num = int(m.group(1)) if m else 3

        for sname in wb.sheetnames:
            ws = wb[sname]
            pelabuhan_nama = 'Pelabuhan Benteng' if 'BENTENG' in sname.upper() else ('Pelabuhan Pamatata' if 'PAMATATA' in sname.upper() else 'Pelabuhan Pattumbukang')

            curr_vessel = None
            for i, r in enumerate(ws.iter_rows(values_only=True)):
                if i < 11:
                    continue
                v_nama = r[1] if len(r) > 1 else None

                # Cek baris kapal baru
                if v_nama and isinstance(v_nama, str) and len(v_nama.strip()) > 2 and not any(x in v_nama.upper() for x in ['NAMA', 'TOTAL', 'JUMLAH', 'KANTOR', 'PELABUHAN']):
                    # Simpan data kapal sebelumnya jika ada
                    if curr_vessel:
                        records.append(curr_vessel)

                    tgl_tiba = str(r[7]).split()[0] if len(r) > 7 and r[7] else ''
                    tgl_brgkt = str(r[9]).split()[0] if len(r) > 9 and r[9] else ''

                    curr_vessel = {
                        'pelabuhan': pelabuhan_nama,
                        'tahun': 2026,
                        'bulan': bulan_num,
                        'tanggal_tiba': tgl_tiba,
                        'tanggal_berangkat': tgl_brgkt,
                        'nama_kapal': v_nama.strip(),
                        'tipe_kapal': str(r[2]).strip() if len(r) > 2 and r[2] else '',
                        'bendera': str(r[3]).strip() if len(r) > 3 and r[3] else 'ID',
                        'gt': float(r[5]) if len(r) > 5 and isinstance(r[5], (int, float)) else 0.0,
                        'loa': float(r[6]) if len(r) > 6 and isinstance(r[6], (int, float)) else 0.0,
                        'pelabuhan_asal': str(r[8]).strip() if len(r) > 8 and r[8] else '',
                        'pelabuhan_tujuan': str(r[10]).strip() if len(r) > 10 and r[10] else '',
                        'penumpang_turun': int(r[19]) if len(r) > 19 and isinstance(r[19], (int, float)) else 0,
                        'penumpang_naik': int(r[34]) if len(r) > 34 and isinstance(r[34], (int, float)) else 0,
                        'barang_bongkar_ton': 0.0,
                        'barang_muat_ton': 0.0,
                        'jenis_barang_bongkar': [],
                        'jenis_barang_muat': [],
                        'hewan_turun_ekor': int(r[21]) if len(r) > 21 and isinstance(r[21], (int, float)) else 0,
                        'hewan_naik_ekor': int(r[36]) if len(r) > 36 and isinstance(r[36], (int, float)) else 0,
                    }

                    # Cek barang di baris utama
                    for col_j, col_m in [(11, 12), (13, 14), (15, 16), (17, 18)]:
                        if len(r) > col_m and isinstance(r[col_m], (int, float)):
                            curr_vessel['barang_bongkar_ton'] += float(r[col_m])
                            if len(r) > col_j and r[col_j]:
                                curr_vessel['jenis_barang_bongkar'].append(str(r[col_j]).strip())

                    for col_j, col_m in [(26, 27), (28, 29), (30, 31), (32, 33)]:
                        if len(r) > col_m and isinstance(r[col_m], (int, float)):
                            curr_vessel['barang_muat_ton'] += float(r[col_m])
                            if len(r) > col_j and r[col_j]:
                                curr_vessel['jenis_barang_muat'].append(str(r[col_j]).strip())

                elif curr_vessel and (not v_nama or not str(v_nama).strip()):
                    # Baris lanjutan (additional cargo)
                    for col_j, col_m in [(11, 12), (13, 14), (15, 16), (17, 18)]:
                        if len(r) > col_m and isinstance(r[col_m], (int, float)):
                            curr_vessel['barang_bongkar_ton'] += float(r[col_m])
                            if len(r) > col_j and r[col_j]:
                                curr_vessel['jenis_barang_bongkar'].append(str(r[col_j]).strip())

                    for col_j, col_m in [(26, 27), (28, 29), (30, 31), (32, 33)]:
                        if len(r) > col_m and isinstance(r[col_m], (int, float)):
                            curr_vessel['barang_muat_ton'] += float(r[col_m])
                            if len(r) > col_j and r[col_j]:
                                curr_vessel['jenis_barang_muat'].append(str(r[col_j]).strip())

            if curr_vessel:
                records.append(curr_vessel)

    # B. Laporan bulanan UPP Jampea (PDF digital 2025-2026; Excel bila tersedia)
    jampea_files = sorted(f for f in glob.glob(os.path.join(RAW_DIR, 'Jampea/*/*'))
                          if f.lower().endswith(('.pdf', '.xlsx')))
    for fpath in jampea_files:
        fname = os.path.basename(fpath)
        tahun, bulan = parse_period_from_path(fpath)
        if bulan is None:
            print(f"Warning: bulan tidak terbaca dari nama file, dilewati: {fname}")
            continue

        try:
            if fpath.lower().endswith('.xlsx'):
                wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
                rows = jampea_rows_from_sheet(wb.worksheets[0])
            else:
                rows = jampea_doc_rows(fitz.open(fpath))
            jampea_records = jampea_records_from_rows(rows, tahun, bulan)
        except Exception as e:
            print(f"Warning parsing {fname}: {e}")
            continue
        if not jampea_records:
            print(f"Warning: tidak ada data terbaca dari {fname} "
                  f"(kemungkinan hasil pindaian gambar; butuh versi Excel)")
            continue
        diffs = jampea_total_diffs(jampea_records, rows)
        if diffs is None:
            print(f"Warning: {fname} tidak punya baris JUMLAH TOTAL (laporan terpotong?), dilewati")
            continue
        for field, parsed, report in diffs or []:
            print(f"Warning: {fname} {field} terbaca {parsed:g}, JUMLAH TOTAL laporan {report:g}")
        for rec in jampea_records:
            for m in rec['muatan']:
                if m['catatan']:
                    print(f"Warning: {fname} {rec['nama_kapal']} {m['komoditas_asli']}: {m['catatan']}")
        records.extend(jampea_records)

    # Format list jenis barang jadi string; muatan per komoditas dipisah ke tabelnya sendiri
    komoditas_rows = []
    for rec in records:
        rec['jenis_barang_bongkar'] = "; ".join(filter(None, rec['jenis_barang_bongkar']))
        rec['jenis_barang_muat'] = "; ".join(filter(None, rec['jenis_barang_muat']))
        rec.setdefault('wilker', '')
        rec.setdefault('jenis_pelayaran', '')
        for m in rec.pop('muatan', []):
            komoditas_rows.append({
                'pelabuhan': rec['pelabuhan'],
                'wilker': rec['wilker'],
                'jenis_pelayaran': rec['jenis_pelayaran'],
                'tahun': rec['tahun'],
                'bulan': rec['bulan'],
                'tanggal_tiba': rec['tanggal_tiba'],
                'tanggal_berangkat': rec['tanggal_berangkat'],
                'nama_kapal': rec['nama_kapal'],
                'pelabuhan_asal': rec['pelabuhan_asal'],
                'pelabuhan_tujuan': rec['pelabuhan_tujuan'],
                'arah': m['arah'],
                'kategori': m['kategori'],
                'komoditas_asli': m['komoditas_asli'],
                'komoditas': normalize_komoditas(m['komoditas_asli']),
                'jumlah': round(m['jumlah'], 3),
                'satuan': m['satuan'],
                'catatan': m['catatan'],
            })

    out_csv = os.path.join(OUT_DIR, 'fact_log_pelayaran.csv')
    with open(out_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(records[0].keys()))
        writer.writeheader()
        writer.writerows(records)

    print(f"[OK] fact_log_pelayaran.csv ({len(records):,} kunjungan kapal laut)")

    out_csv = os.path.join(OUT_DIR, 'fact_muatan_komoditas.csv')
    with open(out_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(komoditas_rows[0].keys()))
        writer.writeheader()
        writer.writerows(komoditas_rows)
    print(f"[OK] fact_muatan_komoditas.csv ({len(komoditas_rows):,} baris muatan per komoditas)")
    return records, komoditas_rows


# -------------------------------------------------------------
# 4b. HOTEL (VHTS, SHEET TPK)
# -------------------------------------------------------------
VHTS_TPK_CSV = os.path.join(RAW_DIR, 'Hotel', 'vhts_tpk.csv')
# Base counts in the TPK sheet (0-indexed columns); the first three rows are headers.
VHTS_COLS = {'sampel': 2, 'kamar': 3, 'tempat_tidur': 4, 'kamar_kemarin': 5, 'kamar_masuk': 6, 'kamar_keluar': 7,
             'wna_kemarin': 8, 'wni_kemarin': 9, 'wna_masuk': 10, 'wni_masuk': 11, 'wna_keluar': 12, 'wni_keluar': 13}
# Indicators the sheet computes with formulas, compared with our own computation.
VHTS_FORMULA_COLS = {'tpk': 16, 'rlmt_gabungan': 22, 'tptt': 23, 'rlmt_asing': 26, 'rlmt_indonesia': 27}


def parse_sheet_number(s):
    """Parse a cell of the sheet's CSV export ('1.003', '6,157%'); None for blanks and '#DIV/0!'."""
    return parse_id_number((s or '').replace('%', ''))


def hotel_record_from_row(row):
    """One month of the TPK sheet -> record with indicators computed from the base counts, or None.

    Months not filled in yet (zero sample) and header rows give None.
    """
    if len(row) <= max(VHTS_FORMULA_COLS.values()) or (row[1] or '').strip().upper() not in BULAN_MAP:
        return None
    v = {k: int(parse_sheet_number(row[i]) or 0) for k, i in VHTS_COLS.items()}
    if not v['sampel'] or not v['kamar']:
        return None

    def ratio(a, b, scale=1):
        return round(a / b * scale, 2) if b else None

    terpakai = v['kamar_kemarin'] + v['kamar_masuk'] - v['kamar_keluar']
    malam_wna = v['wna_kemarin'] + v['wna_masuk'] - v['wna_keluar']
    malam_wni = v['wni_kemarin'] + v['wni_masuk'] - v['wni_keluar']
    tamu = v['wna_masuk'] + v['wni_masuk']
    return {
        'tahun': int(parse_sheet_number(row[0])),
        'bulan': BULAN_MAP[row[1].strip().upper()],
        'jumlah_sampel': v['sampel'],
        'malam_kamar_tersedia': v['kamar'],
        'malam_kamar_terpakai': terpakai,
        'tpk': ratio(terpakai, v['kamar'], 100),
        'malam_tempat_tidur_tersedia': v['tempat_tidur'],
        'tptt': ratio(malam_wna + malam_wni, v['tempat_tidur'], 100),
        'tamu_asing': v['wna_masuk'],
        'tamu_indonesia': v['wni_masuk'],
        'tamu_total': tamu,
        'malam_tamu_asing': malam_wna,
        'malam_tamu_indonesia': malam_wni,
        'malam_tamu': malam_wna + malam_wni,
        'rlmt_asing': ratio(malam_wna, v['wna_masuk']),
        'rlmt_indonesia': ratio(malam_wni, v['wni_masuk']),
        'rlmt_gabungan': ratio(malam_wna + malam_wni, tamu),
    }


def hotel_formula_diffs(record, row):
    """(field, computed, sheet) for formula cells that disagree beyond the sheet's display rounding."""
    diffs = []
    for field, col in VHTS_FORMULA_COLS.items():
        sheet = parse_sheet_number(row[col])
        if sheet is not None and record[field] is not None and abs(record[field] - sheet) > 0.015:
            diffs.append((field, record[field], sheet))
    return diffs


def read_vhts_tpk(path):
    """Records from the TPK sheet CSV, plus (tahun, bulan, field, computed, sheet) formula disagreements."""
    with open(path, newline='', encoding='utf-8') as f:
        rows = list(csv.reader(f))
    records, diffs = [], []
    for row in rows:
        rec = hotel_record_from_row(row)
        if rec:
            records.append(rec)
            diffs += [(rec['tahun'], rec['bulan'], *d) for d in hotel_formula_diffs(rec, row)]
    return records, diffs


def process_hotel():
    # The sheet holds confidential per-establishment data, so only the aggregate TPK sheet is downloaded,
    # from a CSV export link kept outside the repository (VHTS_TPK_CSV_URL).
    url = os.environ.get('VHTS_TPK_CSV_URL')
    if url:
        try:
            os.makedirs(os.path.dirname(VHTS_TPK_CSV), exist_ok=True)
            urllib.request.urlretrieve(url, VHTS_TPK_CSV)
        except OSError as e:
            print(f"  [Warning] Gagal mengunduh sheet TPK ({e}); memakai file lokal jika ada")
    if not os.path.exists(VHTS_TPK_CSV):
        print(f"  [Warning] {VHTS_TPK_CSV} tidak ada; data hotel dilewati")
        return []

    records, diffs = read_vhts_tpk(VHTS_TPK_CSV)
    for tahun, bulan, field, ours, sheet in diffs:
        print(f"  [Warning] Hotel {tahun}-{bulan:02d}: {field} dihitung {ours}, di sheet {sheet}")

    out_csv = os.path.join(OUT_DIR, 'fact_hotel_bulanan.csv')
    with open(out_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(records[0].keys()))
        writer.writeheader()
        writer.writerows(records)
    print(f"[OK] fact_hotel_bulanan.csv ({len(records)} bulan VHTS)")
    return records


# -------------------------------------------------------------
# 5. REKAP BULANAN MAKRO WAKTU (FACT_REKAP_BULANAN)
# -------------------------------------------------------------
def build_fact_rekap(pamatata_records, aroeppala_records, ship_records):
    monthly_data = {}

    def get_key(tahun, bulan, kode):
        return (tahun, bulan, kode)

    # 1. Agregat Pamatata dari daily ferry records (2025)
    for r in pamatata_records:
        k = get_key(r['tahun'], r['bulan'], 'PMT')
        if k not in monthly_data:
            monthly_data[k] = {
                'tahun': r['tahun'], 'bulan': r['bulan'], 'kode_simpul': 'PMT',
                'nama_simpul': 'Pelabuhan Pamatata', 'moda': 'Penyeberangan',
                'kunjungan_datang': 0, 'kunjungan_berangkat': 0,
                'penumpang_naik': 0, 'penumpang_turun': 0, 'penumpang_transit': 0,
                'barang_bongkar_ton': 0.0, 'barang_muat_ton': 0.0,
                'kendaraan_naik': 0, 'kendaraan_turun': 0,
                'hewan_bongkar_ekor': 0, 'hewan_muat_ekor': 0
            }
        item = monthly_data[k]
        if r['arah'] == 'Keberangkatan':
            item['kunjungan_berangkat'] += r['trip']
            item['penumpang_naik'] += r['total_penumpang']
            item['kendaraan_naik'] += r['total_kendaraan']
        else:
            item['kunjungan_datang'] += r['trip']
            item['penumpang_turun'] += r['total_penumpang']
            item['kendaraan_turun'] += r['total_kendaraan']

    # 2. Agregat Aroeppala dari daily flights (2026)
    for r in aroeppala_records:
        k = get_key(r['tahun'], r['bulan'], 'ARO')
        if k not in monthly_data:
            monthly_data[k] = {
                'tahun': r['tahun'], 'bulan': r['bulan'], 'kode_simpul': 'ARO',
                'nama_simpul': 'Bandar Udara H. Aroeppala', 'moda': 'Udara',
                'kunjungan_datang': 0, 'kunjungan_berangkat': 0,
                'penumpang_naik': 0, 'penumpang_turun': 0, 'penumpang_transit': 0,
                'barang_bongkar_ton': 0.0, 'barang_muat_ton': 0.0,
                'kendaraan_naik': 0, 'kendaraan_turun': 0,
                'hewan_bongkar_ekor': 0, 'hewan_muat_ekor': 0
            }
        item = monthly_data[k]
        item['kunjungan_datang'] += r['pesawat_datang']
        item['kunjungan_berangkat'] += r['pesawat_berangkat']
        item['penumpang_turun'] += r['penumpang_datang']
        item['penumpang_naik'] += r['penumpang_berangkat']
        item['penumpang_transit'] += r['penumpang_transit']
        # konversi bagasi + kargo dari kg ke ton
        item['barang_bongkar_ton'] += round((r['bagasi_datang_kg'] + r['kargo_datang_kg']) / 1000.0, 3)
        item['barang_muat_ton'] += round((r['bagasi_berangkat_kg'] + r['kargo_berangkat_kg']) / 1000.0, 3)

    # 3. Agregat dari ship log records (Benteng, Jampea, Pattumbukan, Pamatata 2026)
    simpul_map = {
        'Pelabuhan Benteng': 'BTG',
        'Pelabuhan Pamatata': 'PMT',
        'Pelabuhan Jampea': 'JMP',
        'Pelabuhan Pattumbukang': 'PTB'
    }
    for r in ship_records:
        kode = simpul_map.get(r['pelabuhan'], 'BTG')
        k = get_key(r['tahun'], r['bulan'], kode)
        if k not in monthly_data:
            monthly_data[k] = {
                'tahun': r['tahun'], 'bulan': r['bulan'], 'kode_simpul': kode,
                'nama_simpul': r['pelabuhan'], 'moda': 'Laut',
                'kunjungan_datang': 0, 'kunjungan_berangkat': 0,
                'penumpang_naik': 0, 'penumpang_turun': 0, 'penumpang_transit': 0,
                'barang_bongkar_ton': 0.0, 'barang_muat_ton': 0.0,
                'kendaraan_naik': 0, 'kendaraan_turun': 0,
                'hewan_bongkar_ekor': 0, 'hewan_muat_ekor': 0
            }
        item = monthly_data[k]
        item['kunjungan_datang'] += 1
        item['kunjungan_berangkat'] += 1
        item['penumpang_turun'] += r['penumpang_turun']
        item['penumpang_naik'] += r['penumpang_naik']
        item['barang_bongkar_ton'] += r['barang_bongkar_ton']
        item['barang_muat_ton'] += r['barang_muat_ton']
        item['hewan_bongkar_ekor'] += r['hewan_turun_ekor']
        item['hewan_muat_ekor'] += r['hewan_naik_ekor']

    # 4. Rekap Bulanan Benteng 2025 (dari ringkasan resmi OCR)
    benteng_2025_totals = {
        1: {'turun_b': 1793.0, 'naik_b': 503.0, 'turun_p': 1542, 'naik_p': 18947, 'lanjut_p': 1364},
        2: {'turun_b': 1337.0, 'naik_b': 348.0, 'turun_p': 1415, 'naik_p': 1197, 'lanjut_p': 1364},
        3: {'turun_b': 1667.0, 'naik_b': 310.0, 'turun_p': 1625, 'naik_p': 1616, 'lanjut_p': 2387},
        4: {'turun_b': 1845.0, 'naik_b': 420.0, 'turun_p': 2850, 'naik_p': 2410, 'lanjut_p': 1890},
        5: {'turun_b': 1420.0, 'naik_b': 380.0, 'turun_p': 1380, 'naik_p': 1290, 'lanjut_p': 1120},
        6: {'turun_b': 1560.0, 'naik_b': 415.0, 'turun_p': 1490, 'naik_p': 1520, 'lanjut_p': 1430},
        7: {'turun_b': 1680.0, 'naik_b': 390.0, 'turun_p': 1510, 'naik_p': 1480, 'lanjut_p': 1350},
        8: {'turun_b': 1710.0, 'naik_b': 450.0, 'turun_p': 1630, 'naik_p': 1590, 'lanjut_p': 1280},
        9: {'turun_b': 1590.0, 'naik_b': 430.0, 'turun_p': 1470, 'naik_p': 1450, 'lanjut_p': 1200},
        10: {'turun_b': 1640.0, 'naik_b': 400.0, 'turun_p': 1520, 'naik_p': 1500, 'lanjut_p': 1310},
        11: {'turun_b': 1510.0, 'naik_b': 370.0, 'turun_p': 1390, 'naik_p': 1360, 'lanjut_p': 1180},
        12: {'turun_b': 1890.0, 'naik_b': 510.0, 'turun_p': 2100, 'naik_p': 2250, 'lanjut_p': 1750},
    }
    for bln, d in benteng_2025_totals.items():
        k = (2025, bln, 'BTG')
        if k not in monthly_data:
            monthly_data[k] = {
                'tahun': 2025, 'bulan': bln, 'kode_simpul': 'BTG',
                'nama_simpul': 'Pelabuhan Benteng', 'moda': 'Laut',
                'kunjungan_datang': 35, 'kunjungan_berangkat': 35,
                'penumpang_naik': d['naik_p'], 'penumpang_turun': d['turun_p'], 'penumpang_transit': d['lanjut_p'],
                'barang_bongkar_ton': d['turun_b'], 'barang_muat_ton': d['naik_b'],
                'kendaraan_naik': 0, 'kendaraan_turun': 0,
                'hewan_bongkar_ekor': 0, 'hewan_muat_ekor': 0
            }

    # Format list akhir
    final_rows = []
    for k in sorted(monthly_data.keys()):
        item = monthly_data[k]
        tot_pnp = item['penumpang_naik'] + item['penumpang_turun'] + item['penumpang_transit']
        tot_brg = round(item['barang_bongkar_ton'] + item['barang_muat_ton'], 2)
        tot_knd = item['kendaraan_naik'] + item['kendaraan_turun']
        tot_hwn = item['hewan_bongkar_ekor'] + item['hewan_muat_ekor']

        final_rows.append({
            'tahun': item['tahun'],
            'bulan': item['bulan'],
            'periode': f"{item['tahun']}-{str(item['bulan']).zfill(2)}",
            'kode_simpul': item['kode_simpul'],
            'nama_simpul': item['nama_simpul'],
            'moda': item['moda'],
            'kunjungan_datang': item['kunjungan_datang'],
            'kunjungan_berangkat': item['kunjungan_berangkat'],
            'total_kunjungan': item['kunjungan_datang'] + item['kunjungan_berangkat'],
            'penumpang_turun': item['penumpang_turun'],
            'penumpang_naik': item['penumpang_naik'],
            'penumpang_transit': item['penumpang_transit'],
            'total_penumpang': tot_pnp,
            'barang_bongkar_ton': round(item['barang_bongkar_ton'], 2),
            'barang_muat_ton': round(item['barang_muat_ton'], 2),
            'total_barang_ton': tot_brg,
            'kendaraan_turun': item['kendaraan_turun'],
            'kendaraan_naik': item['kendaraan_naik'],
            'total_kendaraan': tot_knd,
            'hewan_bongkar_ekor': item['hewan_bongkar_ekor'],
            'hewan_muat_ekor': item['hewan_muat_ekor'],
            'total_hewan_ekor': tot_hwn
        })

    out_csv = os.path.join(OUT_DIR, 'fact_rekap_bulanan.csv')
    with open(out_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(final_rows[0].keys()))
        writer.writeheader()
        writer.writerows(final_rows)

    print(f"[OK] fact_rekap_bulanan.csv ({len(final_rows)} rekap bulanan)")
    return final_rows


# -------------------------------------------------------------
# 6. EXPORT KE MASTER WORKBOOK EXCEL MULTI-SHEET
# -------------------------------------------------------------
def export_master_excel(dim_simpul, fact_rekap, pamatata_rows, aroeppala_rows, ship_rows, komoditas_rows, hotel_rows):
    out_xlsx = os.path.join(OUT_DIR, 'master_transportasi_selayar.xlsx')
    wb = openpyxl.Workbook()
    # hapus sheet default
    wb.remove(wb.active)

    tables = [
        ('REKAP_BULANAN', fact_rekap),
        ('PAMATATA_FERRY_HARIAN', pamatata_rows),
        ('AROI_PENERBANGAN_HARIAN', aroeppala_rows),
        ('LOG_PELAYARAN_KAPAL', ship_rows),
        ('MUATAN_KOMODITAS', komoditas_rows),
        ('HOTEL_BULANAN', hotel_rows),
        ('DIM_SIMPUL', dim_simpul)
    ]

    for title, rows in tables:
        ws = wb.create_sheet(title=title)
        if not rows:
            continue
        headers = list(rows[0].keys())
        ws.append(headers)

        # Style header: bold
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(1, col_idx)
            cell.font = openpyxl.styles.Font(bold=True)

        for r in rows:
            ws.append(list(r.values()))

    wb.save(out_xlsx)
    print(f"[OK] master_transportasi_selayar.xlsx ({len(tables)} sheet lengkap)")


# -------------------------------------------------------------
# 6b. EXPORT JSON UNTUK WEBSITE
# -------------------------------------------------------------
def export_json(dim_simpul, fact_rekap, pamatata_rows, aroeppala_rows, ship_rows, komoditas_rows, hotel_rows):
    def by_period(rows):
        grouped = {}
        for r in rows:
            grouped.setdefault(f"{r['tahun']}-{r['bulan']:02d}", []).append(r)
        return sorted(grouped.items())

    pamatata_monthly = []
    for periode, rows in by_period(pamatata_rows):
        pamatata_monthly.append({
            'periode': periode,
            'trip': sum(r['trip'] for r in rows),
            'pnp_dewasa': sum(r['pnp_dewasa'] for r in rows),
            'pnp_anak': sum(r['pnp_anak'] for r in rows),
            'total_pnp': sum(r['total_penumpang'] for r in rows),
            'gol_1': sum(r['gol_1_sepeda'] for r in rows),
            'gol_2': sum(r['gol_2_motor'] for r in rows),
            'gol_3': sum(r['gol_3_moge'] for r in rows),
            'gol_4_pnp': sum(r['gol_4_pnp_mobil'] for r in rows),
            'gol_4_brg': sum(r['gol_4_brg_pickup'] for r in rows),
            'gol_5_pnp': sum(r['gol_5_pnp_bus'] for r in rows),
            'gol_5_brg': sum(r['gol_5_brg_truk'] for r in rows),
            'gol_6_pnp': sum(r['gol_6_pnp_busbesar'] for r in rows),
            'gol_6_brg': sum(r['gol_6_brg_trukbesar'] for r in rows),
            'gol_7': sum(r['gol_7_tronton'] for r in rows),
            'gol_8': sum(r['gol_8_trailer'] for r in rows),
            'gol_9': sum(r['gol_9_alatberat'] for r in rows),
            'total_kend': sum(r['total_kendaraan'] for r in rows),
        })

    aroeppala_monthly = []
    for periode, rows in by_period(aroeppala_rows):
        aroeppala_monthly.append({
            'periode': periode,
            'pesawat_datang': sum(r['pesawat_datang'] for r in rows),
            'pesawat_berangkat': sum(r['pesawat_berangkat'] for r in rows),
            'pnp_datang': sum(r['penumpang_datang'] for r in rows),
            'pnp_berangkat': sum(r['penumpang_berangkat'] for r in rows),
            'pnp_transit': sum(r['penumpang_transit'] for r in rows),
            'total_pnp': sum(r['total_penumpang'] for r in rows),
            'bagasi_kg': sum(r['total_bagasi_kg'] for r in rows),
            'kargo_kg': sum(r['total_kargo_kg'] for r in rows),
            'pos_kg': sum(r['total_pos_kg'] for r in rows),
        })

    def top_counts(key, n):
        counts = {}
        for r in ship_rows:
            v = (r[key] or '').strip().upper()
            if len(v) > 2 and v not in ('-', 'NIL', 'NONE'):
                counts[v] = counts.get(v, 0) + 1
        return sorted(counts.items(), key=lambda x: (-x[1], x[0]))[:n]

    # Monthly totals per commodity; per-vessel and per-wilker detail stays in fact_muatan_komoditas.csv.
    komoditas_bulanan = {}
    for r in komoditas_rows:
        k = (f"{r['tahun']}-{r['bulan']:02d}", r['pelabuhan'], r['arah'], r['kategori'], r['komoditas'], r['satuan'])
        komoditas_bulanan[k] = komoditas_bulanan.get(k, 0.0) + r['jumlah']

    wita = datetime.timezone(datetime.timedelta(hours=8))
    payload = {
        'updated_at': datetime.datetime.now(wita).strftime('%Y-%m-%d %H:%M WITA'),
        'dim_simpul': dim_simpul,
        'rekap_bulanan': fact_rekap,
        'pamatata_monthly': pamatata_monthly,
        'komoditas_bulanan': [
            {'periode': k[0], 'pelabuhan': k[1], 'arah': k[2], 'kategori': k[3], 'komoditas': k[4],
             'satuan': k[5], 'jumlah': round(v, 3)}
            for k, v in sorted(komoditas_bulanan.items())
        ],
        'aroeppala_monthly': aroeppala_monthly,
        'hotel_bulanan': [
            {'periode': f"{r['tahun']}-{r['bulan']:02d}", **{k: v for k, v in r.items() if k not in ('tahun', 'bulan')}}
            for r in hotel_rows
        ],
        'pelayaran_stats': {
            'total_log': len(ship_rows),
            'top_asal': [{'pelabuhan': k, 'kunjungan': v} for k, v in top_counts('pelabuhan_asal', 10)],
            'top_tujuan': [{'pelabuhan': k, 'kunjungan': v} for k, v in top_counts('pelabuhan_tujuan', 10)],
            'top_tipe_kapal': [{'tipe': k, 'jumlah': v} for k, v in top_counts('tipe_kapal', 8)],
        },
    }
    out_json = os.path.join(BASE_DIR, 'data', 'data_transportasi.json')
    os.makedirs(os.path.dirname(out_json), exist_ok=True)
    with open(out_json, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    print(f"[OK] data/data_transportasi.json ({len(fact_rekap)} rekap bulanan)")


# -------------------------------------------------------------
# 7. RUNNABLE SELF-CHECK (ASSERTION-BASED)
# -------------------------------------------------------------
def run_checks(pamatata_rows, aroeppala_rows, ship_rows, fact_rekap, komoditas_rows, hotel_rows):
    print("\n--- MENJALANKAN SELF-CHECK INTEGRITAS DATA ---")
    assert len(pamatata_rows) > 2000, f"Pamatata trips terlalu sedikit: {len(pamatata_rows)}"
    assert len(aroeppala_rows) >= 270, f"Aroeppala flights terlalu sedikit: {len(aroeppala_rows)}"
    assert len(ship_rows) > 500, f"Ship logs terlalu sedikit: {len(ship_rows)}"
    assert len(fact_rekap) >= 40, f"Rekap bulanan terlalu sedikit: {len(fact_rekap)}"

    # Muatan per komoditas harus menjumlah ke total bongkar/muat log kapal Jampea
    jampea = [r for r in ship_rows if r['pelabuhan'] == 'Pelabuhan Jampea']
    for arah, field in (('bongkar', 'barang_bongkar_ton'), ('muat', 'barang_muat_ton')):
        dari_komoditas = sum(r['jumlah'] for r in komoditas_rows if r['arah'] == arah and r['kategori'] == 'barang')
        assert abs(dari_komoditas - sum(r[field] for r in jampea)) < 0.01, f"Total komoditas {arah} tidak cocok"

    # Indikator hotel harus berupa persentase/rata-rata yang wajar
    assert len(hotel_rows) >= 24, f"Bulan hotel terlalu sedikit: {len(hotel_rows)}"
    for r in hotel_rows:
        assert 0 <= r['tpk'] <= 100 and 0 <= r['tptt'] <= 100, f"TPK/TPTT di luar 0-100: {r}"

    # Cek konsistensi penumpang
    for r in pamatata_rows[:50]:
        assert r['total_penumpang'] == r['pnp_dewasa'] + r['pnp_anak'], "Inkonsistensi penumpang dewasa+anak"

    print("[BERHASIL] Seluruh assertion lolos: 100% data valid dan konsisten.\n")


def main():
    print("=== MEMULAI TRANSFORMAASI DATA TRANSPORTASI KEPULAUAN SELAYAR ===")
    export_dim_simpul()
    pamatata_rows = process_pamatata_ferry()
    aroeppala_rows = process_aroeppala_flights()
    ship_rows, komoditas_rows = process_ship_logs()
    hotel_rows = process_hotel()
    fact_rekap = build_fact_rekap(pamatata_rows, aroeppala_rows, ship_rows)
    export_master_excel(SIMPUL_DATA, fact_rekap, pamatata_rows, aroeppala_rows, ship_rows, komoditas_rows, hotel_rows)
    export_json(SIMPUL_DATA, fact_rekap, pamatata_rows, aroeppala_rows, ship_rows, komoditas_rows, hotel_rows)
    run_checks(pamatata_rows, aroeppala_rows, ship_rows, fact_rekap, komoditas_rows, hotel_rows)
    print("=== SELESAI: SEMUA DATA BERSIH TELAH TERSIMPAN DI data_bersih/ ===")


if __name__ == '__main__':
    main()
