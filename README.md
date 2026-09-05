# Import Data Master (xlsx & Google Sheets)

Modul shared untuk import Chart of Accounts (CoA), kontak partner, jurnal akuntansi, saldo awal kas & bank, aset tetap & depresiasi, saldo awal hutang & piutang, struktur analitik, serta penyesuaian nomenklatur laporan keuangan dari satu file Excel atau Google Sheets.

- **Versi**: `19.0.1.1.0`
- **Dependencies**: `accountant`, `mail` (Odoo Enterprise), Python: `openpyxl`, `requests`
- **Fitur Bawaan**: Relabel `account_type` ke Bahasa Indonesia secara otomatis (`data/ir.model.fields.selection.csv`).
- **Template Google Sheet**: [Template Import Data Master (Google Sheets)](https://docs.google.com/spreadsheets/d/1Hs-XjWxnb8qFXmuTXrZHzpJnQw4aXy_LuqDtuFQzAGY/edit?usp=sharing)

---

## Fitur Utama

1. **Manajemen Riwayat Import (`sf.import.history`)**:
   - Setiap proses import tercatat otomatis dalam riwayat resmi (`IMP/YYYY/MM/XXXX`).
   - Menyimpan tanggal eksekusi, user yang menjalankan, file `.xlsx` atau URL Google Sheet, dan status akhir (*Draft*, *Sedang Diproses*, *Selesai*, *Peringatan / Duplikat*, *Gagal*).
   - Dilengkapi kartu ringkasan visual (*summary card*) yang menampilkan statistik jumlah baris sukses, peringatan, skip, dan error.

2. **Logging Rinci per Record (`sf.import.history.line`)**:
   - Pencatatan log komprehensif untuk **seluruh sheet**: `company`, `res.partner`, `account.account`, `account.journal`, `account.asset`, `vendor_bill`, `customer_invoice`, `account.analytic.*`, dan `account.report.*`.
   - **Form Sheet Khusus**: Setiap baris log dapat diklik untuk melihat detail nomor baris spreadsheet, identifier/kode, nama record, statusbadge, dan pesan sistem lengkap.

3. **Tampilan Form Multi-Tab per Sheet**:
   - Baris log tidak dicampur aduk dalam satu daftar panjang, melainkan **dikelompokkan ke tab khusus**:
     - 📑 **Tab Akun (COA)**: Log bagan akun dan saldo awal neraca.
     - 📑 **Tab Jurnal**: Log konfigurasi buku jurnal.
     - 📑 **Tab Kas & Bank**: Log mutasi saldo awal rekening kas dan bank.
     - 📑 **Tab Aset Tetap**: Log pendaftaran aset tetap dan depresiasi.
     - 📑 **Tab Kontak / Partner**: Log sinkronisasi partner / kontak.
     - 📑 **Tab Saldo Awal Hutang / Piutang**: Log transaksi saldo awal faktur vendor dan customer.
     - 📑 **Tab Laporan & Analitik**: Log profil perusahaan, rencana analitik, akun analitik, dan baris laporan keuangan.
     - 📑 **Tab Semua Log Record**: Tampilan global seluruh baris dari semua sheet.
     - 📑 **Tab Ringkasan Statistik**: Rangkuman angka kalkulasi metrik import.
   - *Kondisional*: Tab otomatis tersembunyi jika sheet tersebut tidak terdapat pada file yang diunggah.

4. **Analisis Matriks & Visualisasi (Pivot & Graph View)**:
   - **Pivot View**: Matriks interaktif dengan sumbu baris (*Row*) per nama sheet, sumbu kolom (*Column*) per status eksekusi (*Sukses*, *Peringatan*, *Dilewati*, *Gagal*), dan ukuran cacah baris (*Count*).
   - **Graph View**: Grafik batang (*bar chart*) komposisi hasil import per sheet.
   - **Smart Buttons**: Akses instan di header form untuk *Log Baris*, *Analisis Pivot*, *Peringatan*, dan *Gagal*.

5. **Deteksi Otomatis Duplikat & Proteksi Integritas**:
   - Deteksi otomatis kode ganda (*duplicate code*) pada sheet `account.account`. Baris duplikat ditandai badge kuning `Peringatan` dengan catatan detail referensi baris awal.
   - Pencegahan penimpaan saldo awal jika sudah terdapat transaksi operasional manual yang diinput oleh pengguna.

6. **Chatter & Jejak Audit (`mail.thread`, `mail.activity.mixin`)**:
   - Dilengkapi widget **Chatter** di bagian bawah form view untuk mengirim pesan, log internal note, dan aktivitas.
   - **Otomatis Posting Log Eksekusi**: Sistem secara otomatis mencatat pesan saat proses import dimulai, ringkasan saat selesai (jumlah baris sukses, peringatan, skip, error), serta detail kegagalan jika terjadi error.
   - **Field Tracking**: Perubahan status, user, tanggal, dan metrik total baris tercatat pada log jejak audit.

7. **Arsitektur Bersih (`tools/import_engine.py`)**:
   - Seluruh logika bisnis import dimodularisasi ke dalam folder `tools/import_engine.py`.
   - Root `hooks.py` disediakan sebagai *backward-compatibility wrapper* sehingga modul klien lama (`sf_lazis_unisia_konfig`, `sf_sma_uii_konfig`, dll.) tetap kompatibel 100% tanpa breaking changes.

---

## Cara Pakai

### 1. Antarmuka Web (Menu Import Data Master)
Buka menu **Accounting / Invoicing → Configuration → Import Data Master** (atau via **Settings / Technical → Import Data Master**):
1. Klik tombol **Baru**.
2. Pilih sumber data: **Google Sheets URL** (pastikan akses publik *Anyone with the link*) atau **Upload File (.xlsx)**.
3. Klik tombol **Jalankan Import**.
4. Sistem akan memproses seluruh sheet, menampilkan status progress, summary card, serta membagi hasil log ke tab-tab per sheet.

### 2. Bundled di Modul Client (Otomatis saat Install Modul)
Taruh file `.xlsx` master data di folder `data/` modul client, lalu panggil pada `post_init_hook`:

```python
# hooks.py (modul client)
import os
from odoo.addons.odoo_sf_import_data.hooks import import_bundled_data

def post_init_hook(env):
    import_bundled_data(env, os.path.dirname(__file__))
```

```python
# __manifest__.py (modul client)
{
    ...
    "depends": ["accountant", "odoo_sf_import_data"],
    "post_init_hook": "post_init_hook",
}
```

### 3. Eksekusi via Odoo Shell (Docker)
```bash
cat <<'EOF' | docker compose exec -T odoo odoo shell -c /etc/odoo/odoo.conf --db_host=psqlx -r odoo -w <POSTGRES_PASSWORD> -d <db> --stop-after-init
# Opsi A: Eksekusi dengan pencatatan Riwayat & Logging UI
history = env["sf.import.history"].create({
    "source_type": "google_sheet",
    "google_sheet_url": "https://docs.google.com/spreadsheets/d/.../edit?usp=sharing",
})
history.action_run_import()
env.cr.commit()
print(f"Selesai: {history.name}, Total baris: {history.total_rows}")

# Opsi B: Eksekusi langsung dari folder modul client
from odoo.addons.odoo_sf_import_data.hooks import import_bundled_data
import_bundled_data(env, "/mnt/odoo_addon_shared/<nama_module_client>")
env.cr.commit()
EOF
```

---

## Urutan & Aturan Sheet

Modul ini mendukung **dual format**: format nama sheet singkat (Google Sheet template) maupun format nama lengkap/klasik (.xlsx). Keduanya dapat digunakan secara bergantian.

| No | Tipe / Model Data | Nama Singkat (Google Sheet) | Nama Lengkap / Klasik | Wajib / Opsional |
|:---:|:---|:---:|:---|:---:|
| 1 | Konfigurasi Perusahaan | `c` | `company` | **Wajib** (memuat `OPENING_BALANCE_DATE`) |
| 2 | Kontak / Partner | `r.p` | `res.partner` | Opsional |
| 3 | Chart of Accounts | `a.a` | `account.account` | **Wajib** |
| 4 | Jurnal Akuntansi & Kas/Bank | `a.j` | `account.journal` | **Wajib** |
| 5 | Aset Tetap & Depresiasi | `a.as` | `account.asset` | Opsional |
| 6 | Saldo Awal Hutang (Vendor) | `v.b` | `vendor_bill` | Opsional |
| 7 | Saldo Awal Piutang (Customer) | `c.i` | `customer_invoice` | Opsional |
| 8 | Rencana Analitik | `a.an.p` | `account.analytic.plan` | Opsional |
| 9 | Akun Analitik | `a.an.a` | `account.analytic.account` | Opsional |
| 10 | Rename Laporan Keuangan | `a.r` | `account.report` | Opsional |
| 11 | Rename Baris Laporan | `a.r.l` | `account.report.line` | Opsional |
| — | Petunjuk (Informasi) | `p` | `petunjuk` | Informasi |


Urutan eksekusi import otomatis dijalankan secara sekuensial:
`c` (`company`) → `r.p` (`res.partner`) → `a.a` (`account.account`) → `a.j` (`account.journal`) → `a.as` (`account.asset`) → `v.b` (`vendor_bill`) → `c.i` (`customer_invoice`) → `a.an.p` (`account.analytic.plan`) → `a.an.a` (`account.analytic.account`) → `a.r` (`account.report`) → `a.r.l` (`account.report.line`)

---

### 1. Sheet `company` / `c` (Key-Value)
Menyimpan konfigurasi umum perusahaan dalam format pasangan kunci-nilai (Kolom A: Key, Kolom B: Value):
- **`OPENING_BALANCE_DATE`** *(Wajib)*: Tanggal cutover saldo awal (mis. `2026-06-30`). Tahun buku Odoo (`account_opening_date`) otomatis menjadi `H+1` (`2026-07-01`).
- **`logo`**: URL gambar logo (otomatis di-download & di-encode ke base64).
- **`analytic_accounting`** & **`budget_management`**: Toggle fitur accounting (`TRUE`/`FALSE`).
- Profil perusahaan: `name`, `street`, `street2`, `city`, `zip`, `phone`, `email`, `website`, `report_footer`.

---

### 2. Sheet `res.partner` / `r.p` (Kontak / Partner)
Master data kontak pelanggan, pemasok, donatur, muzakki, amil, atau karyawan:
- **Kolom**: `id`, `name`, `email`, `phone`, `is_company`, `street`, `city`, `state`, `country_id`, `ref`.
- **`is_company`**: Diisi `TRUE` untuk institusi/badan usaha, `FALSE` untuk individu.
- **`state`** & **`country_id`**: Nama provinsi (mis. `D.I. Yogyakarta`) dan negara (mis. `Indonesia`).

---

### 3. Sheet `account.account` / `a.a` (Chart of Accounts & Saldo Awal)
Bagan akun (COA) beserta nilai saldo awal neraca:
- **Kolom**: `id`, `code`, `name`, `account_type`, `opening_debit`, `opening_credit`, `active`.
- **`account_type`**: Label Bahasa Indonesia (misal `Aktiva Lancar`, `Ekuitas`) atau kode teknis internal Odoo (misal `asset_current`, `equity`).
- **`active`**: `TRUE`/`FALSE` (opsional, default `TRUE`).

> [!IMPORTANT]
> **Aturan Akun Khusus (Jangan Diisi di Sheet Ini):**
> - **Kas & Bank**: Kosongkan `opening_debit`/`opening_credit` di sheet ini. Saldo kas/bank diisi via kolom `opening_balance` di sheet `account.journal` / `a.j`.
> - **Aset Tetap & Akumulasi Penyusutan**: Kosongkan jika menggunakan sheet `account.asset` / `a.as` agar nilai tidak tercatat ganda.
> - **Liquidity Transfer**: Dihitung otomatis dari akumulasi saldo kas/bank.
> - **Laba Ditahan / Penyeimbang**: Dihitung otomatis oleh Odoo untuk menyeimbangkan total Debit dan Kredit.

#### Panduan Posisi Saldo Normal (Debit / Kredit) per `account_type`

| Kategori | Tipe Akun (Bahasa Indonesia) | Technical Type (`account_type`) | Posisi Saldo Normal | Keterangan & Catatan Khusus |
| :--- | :--- | :--- | :--- | :---: | :--- |
| **Aset** | Bank dan Tunai | `asset_cash` | **DEBIT** | Kosongkan di sheet ini, diisi via `opening_balance` di sheet `account.journal` / `a.j`. |
| **Aset** | Piutang | `asset_receivable` | **DEBIT** | Disarankan lewat sheet `customer_invoice` / `c.i` agar ada rincian per partner. |
| **Aset** | Cadangan Piutang Tak Tertagih *(Contra-Asset)* | `asset_receivable` / `asset_current` | **KREDIT** | Penyisihan piutang ragu-ragu. Masukkan di `opening_credit` (positif). |
| **Aset** | Aktiva Lancar | `asset_current` | **DEBIT** | Persediaan (*inventory*), uang muka, perlengkapan, dll. |
| **Aset** | Aktiva Tidak Lancar | `asset_non_current` | **DEBIT** | Investasi jangka panjang, piutang jangka panjang. |
| **Aset** | Prabayar | `asset_prepayments` | **DEBIT** | Biaya dibayar dimuka (*prepaid expenses*), sewa dibayar dimuka. |
| **Aset** | Aktiva Tetap | `asset_fixed` | **DEBIT** | Nilai perolehan aset tetap (tanah, bangunan, kendaraan, peralatan). |
| **Aset** | Akumulasi Penyusutan *(Contra-Asset)* | `asset_fixed` / `asset_non_current` | **KREDIT** | **PENTING**: Masukkan di kolom `opening_credit` (positif), **jangan** angka minus di debit! |
| **Liabilitas** | Utang | `liability_payable` | **KREDIT** | Disarankan lewat sheet `vendor_bill` / `v.b` agar ada rincian per partner. |
| **Liabilitas** | Kartu Kredit | `liability_credit_card` | **KREDIT** | Kewajiban kartu kredit korporasi. |
| **Liabilitas** | Pasiva Terkini | `liability_current` | **KREDIT** | Utang lancar, utang gaji, utang pajak, pendapatan diterima dimuka. |
| **Liabilitas** | Hutang Tidak Lancar | `liability_non_current` | **KREDIT** | Utang bank jangka panjang, obligasi, liabilitas sewa pembiayaan. |
| **Ekuitas** | Ekuitas | `equity` | **KREDIT** | Modal disetor, saldo dana zakat/infaq/amil, aset bersih yayasan (surplus). |
| **Ekuitas** | Prive / Penarikan Modal *(Contra-Equity)* | `equity` | **DEBIT** | Pengambilan pribadi pemilik/dividen. Masukkan di kolom `opening_debit` (positif). |
| **Ekuitas** | Akumulasi Defisit / Rugi Tahun Lalu | `equity` | **DEBIT** | Jika historis saldo dana/ekuitas bernilai negatif/defisit. |
| **Ekuitas** | Penghasilan Tahun Terkini | `equity_unaffected` | **KREDIT** | Laba Ditahan / Penyeimbang otomatis Odoo (sebaiknya kosongkan). |
| **Laba Rugi** | Penghasilan | `income` | **KREDIT** | Kosongkan jika cutover awal tahun. Jika tengah tahun: isi di `opening_credit`. |
| **Laba Rugi** | Penghasilan Lainnya | `income_other` | **KREDIT** | Pendapatan bunga bank, keuntungan selisih kurs. |
| **Laba Rugi** | Pengeluaran | `expense` | **DEBIT** | Kosongkan jika cutover awal tahun. Jika tengah tahun: isi di `opening_debit`. |
| **Laba Rugi** | Pengeluaran Lainnya | `expense_other` | **DEBIT** | Beban bunga, kerugian selisih kurs, biaya admin bank. |
| **Laba Rugi** | Penyusutan | `expense_depreciation` | **DEBIT** | Beban penyusutan / depresiasi periode berjalan. |
| **Laba Rugi** | Biaya Pendapatan | `expense_direct_cost` | **DEBIT** | Harga Pokok Penjualan (HPP) / Beban program langsung. |
| **Lainnya** | Off-Balance Sheet | `off_balance` | — | Akun komitmen / kontinjensi di luar neraca. |

---

### 4. Sheet `account.journal` / `a.j` & Saldo Awal Kas/Bank
Buku jurnal operasional Odoo serta penentuan saldo awal kas & rekening bank:
- **Kolom**: `id`, `sequence`, `name`, `type`, `code`, `default_account_id`, `Bank Feed`, **`opening_balance`** *(opsional)*.
- **`type`**: `Bank`, `Kas`, `Penjualan`, `Pembelian`, atau `Lain-lain`.
- **`opening_balance`**: Saldo awal rekening bank/kas (angka positif untuk saldo normal debit, angka negatif untuk cerukan/overdraft kredit).
  - Sistem otomatis membuat record transaksi mutasi (*Bank Statement Line*) yang ter-posting langsung di dashboard jurnal terkait.
  - Akun lawan otomatis diarahkan ke akun **Liquidity Transfer** penyeimbang.
  - Rekening kas/bank yang bernilai 0 / kosong tidak akan dibuatkan baris transaksi dummy.


---

### 5. Sheet `account.asset` / `a.as` (Aset Tetap & Depresiasi)
Master aset tetap untuk manajemen depresiasi otomatis Odoo Enterprise:
- **Kolom**: `id`, `name`, `original_value`, `acquisition_date`, `model_id`, `already_depreciated_amount_import` *(opsional)*.
- **Opening Balance Otomatis**: Total `original_value` dan akumulasi depresiasi otomatis mengisi debit/kredit akun terkait pada opening move.
- **`already_depreciated_amount_import`**: Jika kosong, otomatis dihitung dari `acquisition_date` hingga `OPENING_BALANCE_DATE`. Isi manual hanya untuk override kalkulasi.

#### Rumus & Logika Perhitungan Akumulasi Penyusutan (Prorata Odoo 19)
Kalkulasi otomatis (*fallback*) menggunakan metode **Garis Lurus (*Straight Line*)** dengan standar Odoo 19 Enterprise berbasis **Prorata Temporis (*Constant Periods*)**:

```text
Akumulasi Depresiasi = min(original_value, (original_value / method_number) * Periode Berlalu)
```

Di mana **Periode Berlalu** (dalam satuan bulan) dihitung dari:
1. **Prorata Bulan Pertama (Bulan Perolehan)**:
   - `Prorata Awal = (Hari Aktif di Bulan Beli) / (Total Hari Kalender Bulan Beli)`
   - `Prorata Awal = (days_in_month - acquisition_date.day + 1) / days_in_month`
2. **Bulan Penuh Antara**:
   - `Bulan Penuh = (balance_date.year - acquisition_date.year) * 12 + (balance_date.month - acquisition_date.month - 1)`
3. **Prorata Bulan Cut-off**:
   - `Prorata Akhir = balance_date.day / days_in_month` *(bernilai `1.0` jika tanggal cut-off adalah akhir bulan)*
4. **Total Periode Berlalu**:
   - `Periode Berlalu = Prorata Awal + Bulan Penuh + Prorata Akhir`

> **Rumus Formula Spreadsheet / Excel (Kalkulasi Manual)**:
> ```excel
> =IF(
>   acquisition_date > balance_date,
>   0,
>   MIN(
>     original_value,
>     (original_value / method_number) * (
>       DATEDIF(EOMONTH(acquisition_date, 0), EOMONTH(balance_date, 0), "M") +
>       (EOMONTH(acquisition_date, 0) - acquisition_date + 1) / DAY(EOMONTH(acquisition_date, 0))
>     )
>   )
> )
> ```

---

### 6. Sheet `vendor_bill` / `v.b` & `customer_invoice` / `c.i` (Rincian Utang & Piutang)
Digunakan jika saldo awal piutang atau utang ingin dipecah per partner / faktur belum lunas:
- **Kolom**: `id`, `Reference`, `date`, `journal`, `line_ids/account`, `line_ids/debit`, `line_ids/credit`, `line_ids/name`, `line_ids/partner`, `line_ids/date_maturity`.
- **Fungsi**: Membentuk journal entry saldo awal per vendor/customer sehingga umur piutang (*Aged Receivable*) dan umur utang (*Aged Payable*) tercatat akurat per rekanan.
- Jika detail per partner tidak diperlukan, saldo awal piutang dan utang cukup diisi gelondongan pada sheet `account.account` / `a.a`.

---

### 7. Sheet `account.analytic.plan` / `a.an.p` (Rencana Analitik)
Kategori / dimensi pelaporan analitik (misal: Program Dakwah, Departemen, Proyek):
- **Kolom**: `id`, `name`, `sequence`, `default_applicability`, `color`, `description` *(opsional)*, `parent_id` *(opsional)*.
- **`default_applicability`**: Pilihan `Optional`, `Mandatory`, atau `Unavailable`.

---

### 8. Sheet `account.analytic.account` / `a.an.a` (Akun Analitik)
Pusat biaya (*cost center*) atau akun unit kerja di bawah rencana analitik:
- **Kolom**: `id`, `name`, `plan_id`, `code` *(opsional)*, `active` *(opsional)*, `partner_id` *(opsional)*.
- **`plan_id`**: Nama atau XML ID dari Rencana Analitik terkait.

---

### 9. Sheet `account.report` / `a.r` (Rename Laporan Keuangan)
Menyesuaikan nama laporan resmi Odoo ke nomenklatur lembaga:
- **Kolom**: `id`, `name`.
- **Fungsi**: Mengubah nama laporan keuangan (`account.report`), judul window action (`ir.actions.client`), serta menu sidebar terkait (`ir.ui.menu`) di menu *Accounting → Reporting* (misal: `account_reports.balance_sheet` ➔ `Laporan Posisi Keuangan`, `account_reports.profit_and_loss` ➔ `Laporan Perubahan Dana`).

---

### 10. Sheet `account.report.line` / `a.r.l` (Rename Baris/Hirarki Laporan)
Menyesuaikan nama baris laporan keuangan:
- **Kolom**: `id`, `name`, `code` *(opsional)*.
- **Fungsi**: Mengubah label/nama baris dan sub-grup hirarki laporan keuangan (`account.report.line`), misalnya mengubah baris `EQUITY (& EARNINGS)` menjadi `SALDO DANA`.

---

## Konfigurasi Flags (`tools/import_engine.py`)

- `VALIDATE_IMPORTED_ASSETS` (default `False`): Set `True` jika ingin asset langsung tervalidasi setelah di-import.
- `POST_OPENING_MOVES` (default `False`): Set `True` jika ingin opening move partner langsung ter-post.

