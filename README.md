# Import Data Master (xlsx)

Modul shared untuk import Chart of Accounts, jurnal, asset, kas/bank, dan opening balance vendor/customer dari satu file Excel — dipakai lintas client, bukan spesifik satu database.

Depends: `accountant` (Enterprise Accounting, untuk `account.asset`). External Python dependency: `openpyxl`.

Modul ini juga bawa data bawaan sendiri (`data/ir.model.fields.selection.csv`) yang relabel selection field `account_type` bawaan Odoo ke Bahasa Indonesia (mis. "Bank dan Tunai", "Piutang", dll) — otomatis ter-apply begitu modul ini di-install, tidak perlu setup tambahan.

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

Menu **Settings → Import Data Master** (butuh akses Administrator/`base.group_system`). Upload file `.xlsx` sesuai format sheet, klik Import. Wizard otomatis tertutup setelah selesai dengan notifikasi sukses.

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

## Urutan Import

`res.partner` → `account.account` → `account.journal` → `account.asset` → `kas_bank` → `vendor_bill` → `customer_invoice` (urutan ini sudah tetap di dalam `run_import`, tidak configurable).

Dependency nyata ada 2, keduanya sudah terjaga oleh urutan di atas:
- `account.asset` resolve `journal_id`-nya dengan cari nama jurnal "Aset Tetap" yang **baru dibuat** oleh import `account.journal` (bukan bawaan CoA template) — jadi jurnal wajib di-import sebelum asset.
- `kas_bank` resolve jurnalnya dengan cari nama "Bank"/"Kas" — nama Indonesia ini hasil **relabel** import `account.journal` dari default Odoo "Bank"/"Cash", jadi jurnal juga wajib di-import sebelum kas_bank.

Sheet lain independen satu sama lain (resolve akun/partner via kode/nama yang sudah stabil).

## Tanggal Opening Balance (sheet `petunjuk`)

Tanggal fiskal awal tahun buku diambil dari sheet `petunjuk`, baris `OPENING_FISCAL_YEAR_START` — isi tanggal awal tahun buku di situ. Tanggal opening move sendiri otomatis di-derive sebagai `OPENING_FISCAL_YEAR_START - 1 hari`, mengikuti aturan Odoo bahwa opening move di-tanggal-kan sehari sebelum awal tahun buku. Kalau baris ini tidak ketemu di sheet, import akan gagal dengan error yang jelas (bukan diam-diam pakai tanggal default).

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
