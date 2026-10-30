"""
Generate Standalone Interactive HTML Dashboard
Menggunakan data bersih CSV dan Chart.js (CDN) untuk visualisasi instan.
"""

import csv
import json
import os

BASE_DIR = '/Volumes/work/_Projects/transportasi'
OUT_DIR = os.path.join(BASE_DIR, 'data_bersih')

# 1. Load Rekap Bulanan
with open(os.path.join(OUT_DIR, 'fact_rekap_bulanan.csv')) as f:
    rekap = list(csv.DictReader(f))

# 2. Load Pamatata Harian
with open(os.path.join(OUT_DIR, 'fact_pamatata_ferry_harian.csv')) as f:
    pamatata = list(csv.DictReader(f))

# 3. Load Aroeppala Harian
with open(os.path.join(OUT_DIR, 'fact_aroeppala_penerbangan_harian.csv')) as f:
    aroeppala = list(csv.DictReader(f))

# 4. Load Pelayaran
with open(os.path.join(OUT_DIR, 'fact_log_pelayaran.csv')) as f:
    ships = list(csv.DictReader(f))

# --- AGGREGASI UNTUK VISUALISASI ---
# A. Tren Penumpang per Moda / Simpul per Bulan (2025)
pnp_trends = {}
for r in rekap:
    per = r['periode']
    simpul = r['nama_simpul']
    pnp = int(r['total_penumpang'])
    if per not in pnp_trends:
        pnp_trends[per] = {'periode': per, 'Pamatata (Ferry)': 0, 'Benteng (Laut)': 0, 'Aroeppala (Udara)': 0, 'Jampea (Laut)': 0, 'Pattumbukang (Laut)': 0}
    if 'Pamatata' in simpul:
        pnp_trends[per]['Pamatata (Ferry)'] += pnp
    elif 'Benteng' in simpul:
        pnp_trends[per]['Benteng (Laut)'] += pnp
    elif 'Aroeppala' in simpul:
        pnp_trends[per]['Aroeppala (Udara)'] += pnp
    elif 'Jampea' in simpul:
        pnp_trends[per]['Jampea (Laut)'] += pnp
    elif 'Pattumbukang' in simpul:
        pnp_trends[per]['Pattumbukang (Laut)'] += pnp

sorted_periods = sorted(pnp_trends.keys())
trend_labels = sorted_periods
pamatata_series = [pnp_trends[p]['Pamatata (Ferry)'] for p in sorted_periods]
benteng_series = [pnp_trends[p]['Benteng (Laut)'] for p in sorted_periods]
aroeppala_series = [pnp_trends[p]['Aroeppala (Udara)'] for p in sorted_periods]

# B. Distribusi Kendaraan Ferry Pamatata 2025
kend_dist = {
    'Sepeda Motor (Gol II)': sum(int(r['gol_2_motor']) for r in pamatata),
    'Mobil Penumpang (Gol IV PNP)': sum(int(r['gol_4_pnp_mobil']) for r in pamatata),
    'Pick Up Barang (Gol IV BRG)': sum(int(r['gol_4_brg_pickup']) for r in pamatata),
    'Truk Sedang (Gol V BRG)': sum(int(r['gol_5_brg_truk']) for r in pamatata),
    'Bus Sedang/Besar (Gol V-VI PNP)': sum(int(r['gol_5_pnp_bus']) + int(r['gol_6_pnp_busbesar']) for r in pamatata),
    'Truk Besar / Tronton (Gol VI-VIII)': sum(int(r['gol_6_brg_trukbesar']) + int(r['gol_7_tronton']) + int(r['gol_8_trailer']) for r in pamatata),
}

# C. Top Pelabuhan Asal Kapal (Konektivitas Selayar)
asal_counts = {}
for s in ships:
    asal = s['pelabuhan_asal'].strip().upper()
    if asal and len(asal) > 2 and asal not in ['-', 'NIL', 'NONE']:
        asal_counts[asal] = asal_counts.get(asal, 0) + 1

top_asal = sorted(asal_counts.items(), key=lambda x: x[1], reverse=True)[:8]

# D. Ringkasan KPI
total_pnp_all = sum(int(r['total_penumpang']) for r in rekap)
total_trip_all = sum(int(r['total_kunjungan']) for r in rekap)
total_brg_all = sum(float(r['total_barang_ton']) for r in rekap)
total_knd_all = sum(int(r['total_kendaraan']) for r in rekap)

html_content = f"""<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Analisis Transportasi Kabupaten Kepulauan Selayar</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {{
            --bg: #0f172a;
            --card-bg: #1e293b;
            --card-border: #334155;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --primary: #38bdf8;
            --secondary: #818cf8;
            --accent: #34d399;
            --warning: #fbbf24;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg);
            color: var(--text-main);
            margin: 0;
            padding: 24px;
        }}
        .header {{
            margin-bottom: 24px;
            border-bottom: 1px solid var(--card-border);
            padding-bottom: 16px;
        }}
        .header h1 {{
            margin: 0 0 8px 0;
            font-size: 26px;
            font-weight: 700;
            color: var(--primary);
        }}
        .header p {{
            margin: 0;
            color: var(--text-muted);
            font-size: 14px;
        }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }}
        .kpi-card {{
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 12px;
            padding: 18px;
        }}
        .kpi-label {{
            font-size: 13px;
            color: var(--text-muted);
            margin-bottom: 6px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .kpi-value {{
            font-size: 28px;
            font-weight: 700;
            color: var(--text-main);
        }}
        .kpi-sub {{
            font-size: 12px;
            color: var(--accent);
            margin-top: 4px;
        }}
        .charts-grid {{
            display: grid;
            grid-template-columns: 2fr 1fr;
            gap: 20px;
            margin-bottom: 24px;
        }}
        @media (max-width: 900px) {{
            .charts-grid {{ grid-template-columns: 1fr; }}
        }}
        .chart-card {{
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 12px;
            padding: 20px;
        }}
        .chart-title {{
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 16px;
            color: var(--primary);
        }}
        .table-card {{
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 12px;
            padding: 20px;
            overflow-x: auto;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            text-align: left;
        }}
        th {{
            background: #0f172a;
            color: var(--text-muted);
            padding: 10px 14px;
            border-bottom: 1px solid var(--card-border);
        }}
        td {{
            padding: 10px 14px;
            border-bottom: 1px solid #243248;
        }}
        tr:hover {{
            background: #243248;
        }}
        .badge {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
        }}
        .badge-laut {{ background: #0369a1; color: #e0f2fe; }}
        .badge-ferry {{ background: #15803d; color: #dcfce7; }}
        .badge-udara {{ background: #7e22ce; color: #f3e8ff; }}
    </style>
</head>
<body>

    <div class="header">
        <h1>Dashboard Analisis Operasional Transportasi</h1>
        <p>Wilayah Kerja Kabupaten Kepulauan Selayar (Tahun 2025 – 2026) | Simpul: Pamatata, Benteng, Jampea, Pattumbukang & Bandara H. Aroeppala</p>
    </div>

    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-label">Total Mobilitas Penumpang</div>
            <div class="kpi-value">{total_pnp_all:,}</div>
            <div class="kpi-sub">Laut, Ferry ASDP & Udara</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Total Muatan Logistik</div>
            <div class="kpi-value">{total_brg_all:,.1f} <span style="font-size:16px;">Ton</span></div>
            <div class="kpi-sub">Bongkar & Muat Komoditas</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Kendaraan Penyeberangan</div>
            <div class="kpi-value">{total_knd_all:,}</div>
            <div class="kpi-sub">Lintasan Bira – Pamatata</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Pergerakan Armada</div>
            <div class="kpi-value">{total_trip_all:,}</div>
            <div class="kpi-sub">Trip Ferry, Kapal Laut & Flight</div>
        </div>
    </div>

    <div class="charts-grid">
        <div class="chart-card">
            <div class="chart-title">Tren Bulanan Penumpang Berdasarkan Simpul Utama</div>
            <canvas id="pnpTrendChart" height="130"></canvas>
        </div>
        <div class="chart-card">
            <div class="chart-title">Distribusi Jenis Kendaraan Ferry Pamatata</div>
            <canvas id="kendPieChart" height="200"></canvas>
        </div>
    </div>

    <div class="charts-grid">
        <div class="chart-card">
            <div class="chart-title">Top Asal Konektivitas Pelayaran Kapal Laut Selayar</div>
            <canvas id="asalBarChart" height="120"></canvas>
        </div>
        <div class="chart-card">
            <div class="chart-title">Profil Simpul Transportasi Kepulauan Selayar</div>
            <div style="font-size: 13px; line-height: 1.6; color: var(--text-muted);">
                <p>• <b style="color:var(--text-main);">Pelabuhan Pamatata (ASDP):</b> Jalur urat nadi utama penghubung daratan Sulawesi (Bira) ke Selayar.</p>
                <p>• <b style="color:var(--text-main);">Pelabuhan Benteng:</b> Simpul logistik ibukota, melayani kapal perintis, Pelni, dan kapal layar motor (KLM).</p>
                <p>• <b style="color:var(--text-main);">Pelabuhan Jampea & Pattumbukang:</b> Penghubung antar pulau wilayah kepulauan luar dan jalur timur.</p>
                <p>• <b style="color:var(--text-main);">Bandara H. Aroeppala:</b> Akses udara cepat Makassar – Selayar (Wings Air / Perintis).</p>
            </div>
        </div>
    </div>

    <div class="table-card">
        <div class="chart-title">Ringkasan Rekap Bulanan Terakhir (Sample View)</div>
        <table>
            <thead>
                <tr>
                    <th>Periode</th>
                    <th>Simpul</th>
                    <th>Moda</th>
                    <th>Trip / Kunjungan</th>
                    <th>Penumpang Turun</th>
                    <th>Penumpang Naik</th>
                    <th>Total Penumpang</th>
                    <th>Barang (Ton)</th>
                    <th>Kendaraan</th>
                </tr>
            </thead>
            <tbody>
"""

# Sample 10 rows dari rekap bulanan
for r in rekap[-10:]:
    badge_cls = 'badge-ferry' if r['moda'] == 'Penyeberangan' else ('badge-udara' if r['moda'] == 'Udara' else 'badge-laut')
    html_content += f"""                <tr>
                    <td><b>{r['periode']}</b></td>
                    <td>{r['nama_simpul']}</td>
                    <td><span class="badge {badge_cls}">{r['moda']}</span></td>
                    <td>{int(r['total_kunjungan']):,}</td>
                    <td>{int(r['penumpang_turun']):,}</td>
                    <td>{int(r['penumpang_naik']):,}</td>
                    <td><b>{int(r['total_penumpang']):,}</b></td>
                    <td>{float(r['total_barang_ton']):,.1f}</td>
                    <td>{int(r['total_kendaraan']):,}</td>
                </tr>
"""

html_content += f"""            </tbody>
        </table>
    </div>

    <script>
        // 1. Line Chart: Tren Penumpang
        new Chart(document.getElementById('pnpTrendChart'), {{
            type: 'line',
            data: {{
                labels: {json.dumps(trend_labels)},
                datasets: [
                    {{
                        label: 'Pelabuhan Pamatata (Ferry)',
                        data: {json.dumps(pamatata_series)},
                        borderColor: '#38bdf8',
                        backgroundColor: 'rgba(56, 189, 248, 0.1)',
                        fill: true,
                        tension: 0.3
                    }},
                    {{
                        label: 'Pelabuhan Benteng (Laut)',
                        data: {json.dumps(benteng_series)},
                        borderColor: '#34d399',
                        backgroundColor: 'transparent',
                        tension: 0.3
                    }},
                    {{
                        label: 'Bandara H. Aroeppala (Udara)',
                        data: {json.dumps(aroeppala_series)},
                        borderColor: '#c084fc',
                        backgroundColor: 'transparent',
                        tension: 0.3
                    }}
                ]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    legend: {{ labels: {{ color: '#94a3b8' }} }}
                }},
                scales: {{
                    x: {{ ticks: {{ color: '#94a3b8' }}, grid: {{ color: '#334155' }} }},
                    y: {{ ticks: {{ color: '#94a3b8' }}, grid: {{ color: '#334155' }} }}
                }}
            }}
        }});

        // 2. Doughnut Chart: Kendaraan Ferry Pamatata
        new Chart(document.getElementById('kendPieChart'), {{
            type: 'doughnut',
            data: {{
                labels: {json.dumps(list(kend_dist.keys()))},
                datasets: [{{
                    data: {json.dumps(list(kend_dist.values()))},
                    backgroundColor: ['#38bdf8', '#818cf8', '#34d399', '#fbbf24', '#f87171', '#c084fc']
                }}]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    legend: {{ position: 'bottom', labels: {{ color: '#94a3b8', boxWidth: 12 }} }}
                }}
            }}
        }});

        // 3. Bar Chart: Konektivitas Asal Kapal
        new Chart(document.getElementById('asalBarChart'), {{
            type: 'bar',
            data: {{
                labels: {json.dumps([x[0] for x in top_asal])},
                datasets: [{{
                    label: 'Frekuensi Kunjungan Kapal Masuk',
                    data: {json.dumps([x[1] for x in top_asal])},
                    backgroundColor: '#818cf8'
                }}]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    legend: {{ labels: {{ color: '#94a3b8' }} }}
                }},
                scales: {{
                    x: {{ ticks: {{ color: '#94a3b8' }}, grid: {{ display: false }} }},
                    y: {{ ticks: {{ color: '#94a3b8' }}, grid: {{ color: '#334155' }} }}
                }}
            }}
        }});
    </script>
</body>
</html>
"""

dashboard_path = os.path.join(BASE_DIR, 'dashboard_transportasi.html')
with open(dashboard_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"[OK] Dashboard HTML berhasil dibuat: {dashboard_path}")
