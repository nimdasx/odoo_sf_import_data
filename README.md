# Import Data Master (xlsx)

Modul shared untuk import Chart of Accounts, jurnal, asset, kas/bank, dan opening balance vendor/customer dari satu file Excel — dipakai lintas client, bukan spesifik satu database.

Depends: `accountant` (Enterprise Accounting, untuk `account.asset`). External Python dependency: `openpyxl`, `requests` (untuk download logo, lihat sheet `company` di bawah).

Modul ini juga bawa data bawaan sendiri (`data/ir.model.fields.selection.csv`) yang relabel selection field `account_type` bawaan Odoo ke Bahasa Indonesia (mis. "Bank dan Tunai", "Piutang", dll) — otomatis ter-apply begitu modul ini di-install, tidak perlu setup tambahan.

Contoh file Excel siap pakai (untuk referensi format sheet, atau langsung dites lewat wizard/Opsi 2 di bawah): [`contoh-master-data.xlsx`](contoh-master-data.xlsx).

## Cara Pakai

### Opsi 1 — Bundled di modul client (otomatis saat install)

Taruh **satu** file `.xlsx` di folder `data/` module client, lalu panggil `import_bundled_data` dari `post_init_hook`:

```python
# hooks.py milik module client
import os
from odoo.addons.odoo_sf_import_data.hooks import import_bundled_data

def post_init_hook(env):
    import_bundled_data(env, os.path.dirname(__file__))
```

```python
# __manifest__.py milik module client
{
    ...
    "depends": ["accountant", "odoo_sf_import_data"],
    "post_init_hook": "post_init_hook",
}
```

`import_bundled_data` cari satu-satunya file `.xlsx` di `<module_dir>/data/` — kalau **tidak ada**, skip dengan warning log (bukan blocker install); kalau **lebih dari 1**, raise error (ambiguous, bukan skenario yang aman di-skip diam-diam).

### Opsi 2 — Upload manual lewat UI

Menu **Settings → Import Data Master** (butuh akses Administrator/`base.group_system`). Upload file `.xlsx` sesuai format sheet, **atau** paste link Google Sheet-nya langsung di field terpisah (harus di-share sebagai "Anyone with the link"/public — kalau tidak, download-nya gagal dengan error yang jelas, bukan diam-diam import file kosong/salah). Kalau dua-duanya diisi, file yang dipakai. Klik Import — wizard otomatis tertutup setelah selesai dengan notifikasi sukses.

### Re-run manual tanpa uninstall/install ulang

```bash
cat <<'EOF' | docker compose exec -T odoo /entrypoint.sh odoo shell -c /etc/odoo/odoo.conf -d <db> --stop-after-init
from odoo.addons.odoo_sf_import_data.hooks import import_bundled_data
import os
# ganti path ke folder module client yang punya data/*.xlsx
import_bundled_data(env, "/mnt/extra-addons/<nama_module_client>")
env.cr.commit()
print("DONE")
EOF
```

> Catatan: `post_init_hook` (kalau dipakai lewat Opsi 1) hanya jalan sekali saat **install pertama** module client, tidak ikut jalan lagi saat `-u` update — pakai cara re-run manual di atas kalau perlu jalankan ulang.

## Data Perusahaan (sheet `company`)

Sheet opsional berisi baris key/value, ditulis langsung ke `res.company` (`base.main_company`) — tidak butuh `id`/xml_id karena company sudah pasti ada, cukup `write()`. Sheet ini juga tempat baris `OPENING_BALANCE_DATE` (lihat bagian [Tanggal Opening Balance](#tanggal-opening-balance-sheet-company)).

| key | Keterangan |
|---|---|
| `name`, `street`, `street2`, `city`, `zip`, `phone`, `email`, `website`, `report_footer` | Teks biasa, ditulis apa adanya. Kalau cell-nya ke-format angka oleh Excel/Sheets (mis. `zip`), otomatis di-normalize balik ke text tanpa `.0` di belakang. |
| `state`, `country` | Nama (bukan kode), di-resolve lewat pencarian nama — sama seperti kolom `state`/`country_id` di sheet `res.partner`. |
| `logo` | **URL** gambar (bukan file upload) — otomatis di-`download` (timeout 10 detik) dan di-encode base64 ke `company.logo`. Gagal download (mis. URL mati/timeout) cuma di-log warning, tidak menggagalkan import lain, dan logo lama **tidak dihapus**. |
| `external_report_layout` | xml_id layout laporan (mis. `web.external_layout_folder`). Kosong = pakai default `web.external_layout_folder` (`DEFAULT_EXTERNAL_REPORT_LAYOUT` di `hooks.py`). |
| `analytic_accounting` | `TRUE`/`FALSE` (boleh juga `ya`/`tidak`, `yes`/`no`, `1`/`0`, `aktif`/`nonaktif`, atau checkbox Excel asli) — aktifkan/nonaktifkan fitur **Analytic Accounting**. |
| `budget_management` | Sama format dengan di atas — aktifkan/nonaktifkan **Budget Management** (install modul `account_budget` kalau `TRUE`). |

Baris/kolom yang dikosongkan **tidak** menimpa data yang sudah ada (kecuali sheet-nya sendiri tidak ada sama sekali, dalam hal ini `_import_company` di-skip) — jadi sheet boleh diisi bertahap. Untuk `analytic_accounting`/`budget_management` khususnya: kosong = **tidak diubah** (beda dengan field lain, nilai "off" untuk toggle ini tetap nilai yang valid untuk di-set eksplisit, jadi butuh dibedakan dari "belum diisi").

## Urutan Import

`company` → `res.partner` → `account.account` → `account.journal` → `account.asset` → `kas_bank` → `vendor_bill` → `customer_invoice` (urutan ini sudah tetap di dalam `run_import`, tidak configurable).

Dependency nyata ada 2, keduanya sudah terjaga oleh urutan di atas:
- `account.asset` resolve `journal_id`-nya dengan cari nama jurnal "Aset Tetap" yang **baru dibuat** oleh import `account.journal` (bukan bawaan CoA template) — jadi jurnal wajib di-import sebelum asset.
- `kas_bank` resolve jurnalnya dengan cari nama "Bank"/"Kas" — nama Indonesia ini hasil **relabel** import `account.journal` dari default Odoo "Bank"/"Cash", jadi jurnal juga wajib di-import sebelum kas_bank.

Sheet lain independen satu sama lain (resolve akun/partner via kode/nama yang sudah stabil).

## Tanggal Opening Balance (sheet `company`)

Tanggal cutover opening balance diambil dari sheet `company`, baris `OPENING_BALANCE_DATE` — isi tanggal saldo awal berlaku di situ (mis. "saldo per 30 Juni 2026"). Awal tahun buku (`account_opening_date`) otomatis di-derive sebagai `OPENING_BALANCE_DATE + 1 hari`, mengikuti aturan Odoo bahwa opening move di-tanggal-kan sehari sebelum awal tahun buku. Kalau baris `OPENING_BALANCE_DATE` tidak ketemu, import akan gagal dengan error yang jelas (bukan diam-diam pakai tanggal default).

Kalau hasil derive-nya bukan tanggal 1 (mis. `OPENING_BALANCE_DATE` diisi tanggal 15), import **tidak** gagal — cuma log warning, karena Odoo sendiri tidak mewajibkan tahun buku mulai di tanggal 1 (ada kasus valid seperti migrasi sistem di tengah bulan). Warning ini cuma jaring pengaman untuk typo tanggal, bukan validasi keras.

> Sebelumnya nilai ini dibaca dari baris `OPENING_FISCAL_YEAR_START` (arah sebaliknya) di sheet `petunjuk` — `_read_opening_balance_date` cek sheet `company` dulu baru fallback ke `petunjuk` untuk workbook lama yang belum di-migrasi.

## Opening Balance untuk Asset

`account.asset` **tidak otomatis** menulis nilai perolehan/akumulasi penyusutan ke buku besar — normalnya nilai itu masuk lewat transaksi pembelian aslinya. Import ini **otomatis menghitung dan mengisi** ini sendiri — tidak perlu diisi manual di sheet `account.account`:

| Akun (kolom di sheet `account.asset`) | Sisi | Nilai |
|---|---|---|
| `account_asset_id` (akun aset tetap, mis. "Peralatan Kantor") | `opening_debit` | jumlah `original_value` semua asset yang pakai akun ini |
| `account_depreciation_id` (akun akumulasi penyusutan) | `opening_credit` | jumlah `already_depreciated_amount_import` semua asset yang pakai akun ini |

Kalau beberapa baris asset pakai akun yang sama, nilainya otomatis dijumlahkan jadi satu opening balance per akun. Ini otomatis balance ke retained earnings lewat mekanisme `opening_debit`/`opening_credit` yang sama seperti akun lain.

> Catatan: kalau kolom `opening_debit`/`opening_credit` di sheet `account.account` **juga** diisi manual untuk akun aset/depresiasi yang sama, nilai dari `account.asset` akan menimpanya (last-write-wins, karena import asset jalan setelah import `account.account`) — jadi biarkan kosong di sheet `account.account` untuk akun-akun ini.

### Kalkulasi `already_depreciated_amount_import`

Field ini harus akurat sampai tepat tanggal opening balance — kalau kekecilan, saat asset di-`validate()`/post, Odoo akan generate journal entry penyusutan tambahan yang **ke-backdate ke bulan-bulan sebelum opening balance** (lihat mekanisme `compute_depreciation_board()` di `account_asset.py`: `already_depreciated_amount_import` cuma "memakan" amount dari periode-periode paling awal — kalau nilainya kurang, sisa periode sebelum cutover tetap generate move baru).

Import ini **otomatis menghitung** field ini kalau kolom `already_depreciated_amount_import` di sheet `account.asset` **dibiarkan kosong** — tidak perlu formula Excel:

```
already_depreciated_amount_import = (original_value / method_number) × jumlah_periode_yang_sudah_lewat
```

`jumlah_periode_yang_sudah_lewat` = banyak periode (bulan/tahun, sesuai `method_period`) dari `acquisition_date` sampai tanggal opening balance (dibaca dari sheet `petunjuk`), dihitung penuh (termasuk periode acquisition), dibatasi maksimal `method_number`.

Kalau kolom itu **diisi manual**, nilai dari Excel yang dipakai (override, tidak dihitung ulang) — berguna kalau metode penyusutan real-nya bukan garis lurus murni atau ada penyesuaian khusus.

> Catatan: rumus di atas mengasumsikan `acquisition_date` persis tanggal 1. Kalau tidak, Odoo default pakai prorata `constant_periods` (30 hari/bulan) untuk periode pertama, jadi hasil kalkulasi otomatis bisa sedikit meleset (biasanya bisa diabaikan kalau selisihnya cuma beberapa hari) — isi manual kalau butuh presisi sampai ke rupiah terakhir.

## Opening Balance untuk Kas/Bank (Liquidity Transfer)

Import `kas_bank` juga **otomatis menghitung** total opening balance akun **"Liquidity Transfer"** dari sheet `kas_bank` (jumlah bersih debit-kredit semua baris yang match `journal.default_account_id`) dan langsung menulisnya ke `opening_debit`/`opening_credit` — sama seperti asset, tidak perlu diisi manual di sheet `account.account`. Nilai ini yang nanti direkonsil otomatis lewat `_reconcile_liquidity_transfer` begitu opening move di-post.

> Catatan: jangan isi manual `opening_debit`/`opening_credit` untuk akun "Liquidity Transfer" di sheet `account.account` — akan tertimpa oleh hasil kalkulasi `kas_bank` (import kas_bank jalan setelah import `account.account`), sama seperti kasus asset di atas.

> Catatan tambahan: begitu opening move (`account_opening_move_id`) sudah **di-post**, semua penulisan `opening_debit`/`opening_credit` (dari `account.account`, `account.asset`, maupun `kas_bank`) otomatis **di-skip**, bukan error — jadi aman untuk re-run import kapan saja setelah post, tidak akan merusak data yang sudah di-review.

## Catatan Minor: Akun Retained Earnings

Kalau kolom `opening_debit`/`opening_credit` di sheet `account.account` diisi untuk akun **"Laba/Rugi Belum Dialokasikan"** (akun retained-earnings yang sama dipakai `vendor_bill`/`customer_invoice` untuk auto-balancing), berpotensi terjadi double-balancing (entry jadi "not balanced") — hindari mengisi opening balance manual untuk akun ini.

## Toggle Perilaku

Dua flag di `hooks.py` (default `False`, tinggalkan hasil import sebagai draft untuk direview manual):

- `VALIDATE_IMPORTED_ASSETS` — kalau `True`, asset otomatis di-`validate()` (generate depreciation board + posting) begitu diimport.
- `POST_OPENING_MOVES` — kalau `True`, move opening balance vendor/customer otomatis di-post begitu diimport.

(Tidak berlaku untuk `kas_bank` — `account.bank.statement.line` di Odoo selalu auto-post begitu dibuat, itu perilaku core Odoo.)
