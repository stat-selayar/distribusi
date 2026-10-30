# Distribusi Selayar: Dashboard Statistik Transportasi & Pariwisata

Dashboard interaktif dan repositori data terbuka untuk pemantauan dan analisis statistik distribusi, transportasi, dan pariwisata di **Kabupaten Kepulauan Selayar**, Sulawesi Selatan.

🌐 **Live Dashboard**: [https://stat-selayar.github.io/distribusi/](https://stat-selayar.github.io/distribusi/)

---

## 📌 Cakupan Simpul & Moda Transportasi

| Kode | Simpul Transportasi | Moda | Cakupan Layanan |
| :--- | :--- | :--- | :--- |
| **PMT** | **Pelabuhan Pamatata** | Penyeberangan (ASDP) | Lintasan utama penyeberangan Bira – Pamatata (penumpang & 12 golongan kendaraan). |
| **BTG** | **Pelabuhan Benteng** | Laut (KUPP Selayar) | Logistik ibukota kabupaten, kapal perintis, Pelni, dan kapal layar motor (KLM). |
| **ARO** | **Bandar Udara H. Aroeppala** | Udara | Penerbangan domestik/perintis rute Makassar (UPG) – Selayar (KSR). |
| **JMP** | **Pelabuhan Jampea** | Laut (UPP Jampea) | Simpul konektivitas pulau luar (Pulau Jampea) ke Selayar, Bulukumba, dan Flores. |
| **PTB** | **Pelabuhan Pattumbukang** | Laut (Wilker UPP) | Simpul penyeberangan dan pelayaran wilayah timur & selatan pulau. |

---

## 📂 Struktur Dataset (`data_bersih/` & `data/`)

Dataset disusun dengan prinsip **Tidy Data** (1 baris = 1 observasi) yang siap digunakan di Excel, Power BI, Tableau, Python, dan R:

1. **`master_transportasi_selayar.xlsx`** (Workbook Excel kompilasi 5 sheet dengan header tebal dan siap pakai).
2. **`fact_rekap_bulanan.csv`** (Tabel fakta bulanan per moda/simpul: total trip, penumpang naik/turun/transit, tonase bongkar/muat barang, total kendaraan).
3. **`fact_pamatata_ferry_harian.csv`** (2.058 trip harian kapal ferry ASDP: penumpang dewasa/anak dan rincian kendaraan Golongan I s.d. IX).
4. **`fact_aroeppala_penerbangan_harian.csv`** (273 hari kalender penerbangan: pesawat, penumpang datang/berangkat, bagasi, kargo, dan pos).
5. **`fact_log_pelayaran.csv`** (4.300+ record kunjungan kapal laut: nama kapal, GT, asal/tujuan, muatan komoditas, dan ternak).
6. **`dim_simpul.csv`** (Metadata geospasial simpul transportasi).
7. **`data/data_transportasi.json`** (Data ringkas yang dibaca langsung oleh `index.html`; dihasilkan otomatis oleh `build_datasets.py`).

---

## 🔄 Cara Memperbarui Data (*Updating the Data*)

Untuk memperbarui data saat ada file laporan baru:
1. Simpan file laporan bulanan (Excel/PDF) ke folder `data_mentah/<Simpul>/<Tahun>/`.
2. Jalankan pipeline transformasi:
   ```bash
   python3 build_datasets.py
   ```
   Perintah ini memperbarui `data_bersih/` dan `data/data_transportasi.json`, sehingga website (`index.html`) otomatis menampilkan data terbaru. `generate_dashboard.py` hanya membuat `dashboard_transportasi.html` (versi statis lama) dan tidak diperlukan untuk website.
   Butuh paket `openpyxl` dan `PyMuPDF`. Tes: `python3 -m unittest tests.test_build_datasets`.
3. Commit dan push ke GitHub:
   ```bash
   git add .
   git commit -m "Update data transportasi periode [Bulan Tahun]"
   git push origin main
   ```
   *Dashboard GitHub Pages akan otomatis terbarui secara instan.*

---

## 🛠️ Lisensi & Sumber Data
- **Sumber Data Mentah**: Laporan Resmi SIMOPPEL Ditjen Perhubungan Laut, BPTD Kelas II Sulsel (Satpel Pamatata), dan Kantor UPBU Kelas III H. Aroeppala Selayar.
- **Pengelola**: Badan Pusat Statistik (BPS) Kabupaten Kepulauan Selayar.
