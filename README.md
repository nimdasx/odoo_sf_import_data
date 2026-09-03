# Import Data Master (xlsx)

Modul shared untuk import Chart of Accounts (CoA), partner, jurnal, kas/bank, asset tetap, dan opening balance dari satu file Excel / Google Sheets.

- **Dependencies**: `accountant` (Odoo Enterprise), Python: `openpyxl`, `requests`
- **Fitur Bawaan**: Relabel `account_type` ke Bahasa Indonesia secara otomatis (`data/ir.model.fields.selection.csv`).
- **Template Google Sheet**: [Template Import Data Master (Google Sheets)](https://docs.google.com/spreadsheets/d/1Hs-XjWxnb8qFXmuTXrZHzpJnQw4aXy_LuqDtuFQzAGY/edit?usp=sharing)

---

## Cara Pakai

### 1. Bundled di Modul Client (Otomatis saat Install)
Taruh satu file `.xlsx` di folder `data/` modul client, lalu panggil di `post_init_hook`:

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

### 2. Manual via UI Wizard
Buka menu **Settings → Import Data Master** (Hak akses: Administrator):
- Upload file `.xlsx`, **atau**
- Masukkan URL Google Sheets publik (*Anyone with the link*).
- Duplikasi template master data berikut (*File → Make a copy*): [Template Google Sheets](https://docs.google.com/spreadsheets/d/1Hs-XjWxnb8qFXmuTXrZHzpJnQw4aXy_LuqDtuFQzAGY/edit?usp=sharing).

### 3. Re-run via Odoo Shell (Docker)
```bash
cat <<'EOF' | docker compose exec -T odoo odoo shell -c /etc/odoo/odoo.conf --db_host=psqlx -r odoo -w <POSTGRES_PASSWORD> -d <db> --stop-after-init
from odoo.addons.odoo_sf_import_data.hooks import import_bundled_data
import_bundled_data(env, "/mnt/extra-addons/<nama_module_client>")
env.cr.commit()
print("DONE")
EOF
```

---

## Urutan & Aturan Sheet

Urutan import dijalankan secara sekuensial:
`company` → `res.partner` → `account.account` → `account.journal` → `account.asset` → `kas_bank (opsional)` → `vendor_bill` → `customer_invoice` → `account.analytic.plan` → `account.analytic.account` → `account.report` → `account.report.line`

---

### 1. Sheet `company` (Key-Value)
Menyimpan konfigurasi umum perusahaan dalam format pasangan kunci-nilai (Kolom A: Key, Kolom B: Value):
- **`OPENING_BALANCE_DATE`** *(Wajib)*: Tanggal cutover saldo awal (mis. `2026-06-30`). Tahun buku Odoo (`account_opening_date`) otomatis menjadi `H+1` (`2026-07-01`).
- **`logo`**: URL gambar logo (otomatis di-download & di-encode ke base64).
- **`analytic_accounting`** & **`budget_management`**: Toggle fitur accounting (`TRUE`/`FALSE`).
- Profil perusahaan: `name`, `street`, `street2`, `city`, `zip`, `phone`, `email`, `website`, `report_footer`.

---

### 2. Sheet `res.partner` (Kontak / Partner)
Master data kontak pelanggan, pemasok, donatur, muzakki, amil, atau karyawan:
- **Kolom**: `id`, `name`, `email`, `phone`, `is_company`, `street`, `city`, `state`, `country_id`, `ref`.
- **`is_company`**: Diisi `TRUE` untuk institusi/badan usaha, `FALSE` untuk individu.
- **`state`** & **`country_id`**: Nama provinsi (mis. `D.I. Yogyakarta`) dan negara (mis. `Indonesia`).

---

### 3. Sheet `account.account` (Chart of Accounts & Saldo Awal)
Bagan akun (COA) beserta nilai saldo awal neraca:
- **Kolom**: `id`, `code`, `name`, `account_type`, `opening_debit`, `opening_credit`, `active`.
- **`account_type`**: Label Bahasa Indonesia (misal `Aktiva Lancar`, `Ekuitas`) atau kode teknis internal Odoo (misal `asset_current`, `equity`).
- **`active`**: `TRUE`/`FALSE` (opsional, default `TRUE`).

> [!IMPORTANT]
> **Aturan Akun Khusus (Jangan Diisi di Sheet Ini):**
> - **Kas & Bank**: Kosongkan `opening_debit`/`opening_credit` di sheet ini. Saldo kas/bank diisi via kolom `opening_balance` di sheet `account.journal`.
> - **Aset Tetap & Akumulasi Penyusutan**: Kosongkan jika menggunakan sheet `account.asset` agar nilai tidak tercatat ganda.
> - **Liquidity Transfer**: Dihitung otomatis dari akumulasi saldo kas/bank.
> - **Laba Ditahan / Penyeimbang**: Dihitung otomatis oleh Odoo untuk menyeimbangkan total Debit dan Kredit.

#### Panduan Posisi Saldo Normal (Debit / Kredit) per `account_type`

| Kategori | Tipe Akun (Bahasa Indonesia) | Technical Type (`account_type`) | Posisi Saldo Normal | Keterangan & Catatan Khusus |
| :--- | :--- | :--- | :---: | :--- |
| **Aset** | Bank dan Tunai | `asset_cash` | **DEBIT** | Kosongkan di sheet ini, diisi via `opening_balance` di sheet `account.journal`. |
| **Aset** | Piutang | `asset_receivable` | **DEBIT** | Disarankan lewat sheet `customer_invoice` agar ada rincian per partner. |
| **Aset** | Cadangan Piutang Tak Tertagih *(Contra-Asset)* | `asset_receivable` / `asset_current` | **KREDIT** | Penyisihan piutang ragu-ragu. Masukkan di `opening_credit` (positif). |
| **Aset** | Aktiva Lancar | `asset_current` | **DEBIT** | Persediaan (*inventory*), uang muka, perlengkapan, dll. |
| **Aset** | Aktiva Tidak Lancar | `asset_non_current` | **DEBIT** | Investasi jangka panjang, piutang jangka panjang. |
| **Aset** | Prabayar | `asset_prepayments` | **DEBIT** | Biaya dibayar dimuka (*prepaid expenses*), sewa dibayar dimuka. |
| **Aset** | Aktiva Tetap | `asset_fixed` | **DEBIT** | Nilai perolehan aset tetap (tanah, bangunan, kendaraan, peralatan). |
| **Aset** | Akumulasi Penyusutan *(Contra-Asset)* | `asset_fixed` / `asset_non_current` | **KREDIT** | **PENTING**: Masukkan di kolom `opening_credit` (positif), **jangan** angka minus di debit! |
| **Liabilitas** | Utang | `liability_payable` | **KREDIT** | Disarankan lewat sheet `vendor_bill` agar ada rincian per partner. |
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

### 4. Sheet `account.journal` & Saldo Awal Kas/Bank
Buku jurnal operasional Odoo serta penentuan saldo awal kas & rekening bank:
- **Kolom**: `id`, `sequence`, `name`, `type`, `code`, `default_account_id`, `Bank Feed`, **`opening_balance`** *(opsional)*.
- **`type`**: `Bank`, `Kas`, `Penjualan`, `Pembelian`, atau `Lain-lain`.
- **`opening_balance`**: Saldo awal rekening bank/kas (angka positif untuk saldo normal debit, angka negatif untuk cerukan/overdraft kredit).
  - Sistem otomatis membuat record transaksi mutasi (*Bank Statement Line*) yang ter-posting langsung di dashboard jurnal terkait.
  - Akun lawan otomatis diarahkan ke akun **Liquidity Transfer** penyeimbang.
  - Rekening kas/bank yang bernilai 0 / kosong tidak akan dibuatkan baris transaksi dummy.
- *(Opsional / Legacy)*: Sheet `kas_bank` terpisah tetap didukung sebagai cadangan jika kolom `opening_balance` di sheet ini tidak digunakan.

---

### 5. Sheet `account.asset` (Aset Tetap & Depresiasi)
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

### 6. Sheet `vendor_bill` & `customer_invoice` (Rincian Utang & Piutang)
Digunakan jika saldo awal piutang atau utang ingin dipecah per partner / faktur belum lunas:
- **Kolom**: `id`, `Reference`, `date`, `journal`, `line_ids/account`, `line_ids/debit`, `line_ids/credit`, `line_ids/name`, `line_ids/partner`, `line_ids/date_maturity`.
- **Fungsi**: Membentuk journal entry saldo awal per vendor/customer sehingga umur piutang (*Aged Receivable*) dan umur utang (*Aged Payable*) tercatat akurat per rekanan.
- Jika detail per partner tidak diperlukan, saldo awal piutang dan utang cukup diisi gelondongan pada sheet `account.account`.

---

### 7. Sheet `account.analytic.plan` (Rencana Analitik)
Kategori / dimensi pelaporan analitik (misal: Program Dakwah, Departemen, Proyek):
- **Kolom**: `id`, `name`, `sequence`, `default_applicability`, `color`, `description` *(opsional)*, `parent_id` *(opsional)*.
- **`default_applicability`**: Pilihan `Optional`, `Mandatory`, atau `Unavailable`.

---

### 8. Sheet `account.analytic.account` (Akun Analitik)
Pusat biaya (*cost center*) atau akun unit kerja di bawah rencana analitik:
- **Kolom**: `id`, `name`, `plan_id`, `code` *(opsional)*, `active` *(opsional)*, `partner_id` *(opsional)*.
- **`plan_id`**: Nama atau XML ID dari Rencana Analitik terkait.

---

### 9. Sheet `account.report` (Rename Laporan Keuangan)
Menyesuaikan nama laporan resmi Odoo ke nomenklatur lembaga:
- **Kolom**: `id`, `name`.
- **Fungsi**: Mengubah nama laporan keuangan (`account.report`), judul window action (`ir.actions.client`), serta menu sidebar terkait (`ir.ui.menu`) di menu *Accounting → Reporting* (misal: `account_reports.balance_sheet` ➔ `Laporan Posisi Keuangan`, `account_reports.profit_and_loss` ➔ `Laporan Perubahan Dana`).

---

### 10. Sheet `account.report.line` (Rename Baris/Hirarki Laporan)
Menyesuaikan nama baris laporan keuangan:
- **Kolom**: `id`, `name`, `code` *(opsional)*.
- **Fungsi**: Mengubah label/nama baris dan sub-grup hirarki laporan keuangan (`account.report.line`), misalnya mengubah baris `EQUITY (& EARNINGS)` menjadi `SALDO DANA`.

---

## Konfigurasi Flags (`hooks.py`)

- `VALIDATE_IMPORTED_ASSETS` (default `False`): Set `True` jika ingin asset langsung tervalidasi setelah di-import.
- `POST_OPENING_MOVES` (default `False`): Set `True` jika ingin opening move partner langsung ter-post.
