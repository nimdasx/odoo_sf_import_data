# Import Data Master (xlsx)

Modul shared untuk import Chart of Accounts (CoA), partner, jurnal, asset, kas/bank, dan opening balance dari satu file Excel / Google Sheets.

- **Dependencies**: `accountant` (Odoo Enterprise), Python: `openpyxl`, `requests`
- **Fitur Bawaan**: Relabel `account_type` ke Bahasa Indonesia secara otomatis (`data/ir.model.fields.selection.csv`).

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

### 3. Re-run via Odoo Shell
```bash
cat <<'EOF' | docker compose exec -T odoo odoo shell -c /etc/odoo/odoo.conf -d <db> --stop-after-init
from odoo.addons.odoo_sf_import_data.hooks import import_bundled_data
import_bundled_data(env, "/mnt/extra-addons/<nama_module_client>")
env.cr.commit()
print("DONE")
EOF
```

---

## Urutan & Aturan Sheet

Urutan import:
`company` → `res.partner` → `account.account` → `account.journal` → `account.asset` → `kas_bank` → `vendor_bill` → `customer_invoice`

### 1. Sheet `company` (Key-Value)
- **`OPENING_BALANCE_DATE`** *(Wajib)*: Tanggal cutover saldo awal (mis. `2026-06-30`). Tahun buku Odoo (`account_opening_date`) otomatis menjadi `H+1` (`2026-07-01`).
- **`logo`**: URL gambar logo (otomatis di-download & di-encode ke base64).
- **`analytic_accounting`** & **`budget_management`**: Toggle fitur (`TRUE`/`FALSE`).
- Profil perusahaan: `name`, `street`, `city`, `zip`, `phone`, `email`, `website`, dll.

### 2. Sheet `account.asset`
- **Opening Balance Otomatis**: Total `original_value` dan akumulasi depresiasi otomatis mengisi debit/kredit akun terkait pada opening move.
- **`already_depreciated_amount_import`**: Jika kosong, otomatis dihitung dari `acquisition_date` hingga `OPENING_BALANCE_DATE`. Isi manual hanya untuk override kalkulasi.

### 3. Sheet `kas_bank`
- Total saldo otomatis mengisi opening balance akun **"Liquidity Transfer"**.

### 4. Sheet `account.account` (Perhatian Khusus)
> **Penting**: Kosongkan nilai `opening_debit`/`opening_credit` untuk akun-akun berikut agar tidak double-balance atau tertimpa:
> - Akun Aset Tetap & Akumulasi Penyusutan (diisi dari sheet `account.asset`).
> - Akun Liquidity Transfer (diisi dari sheet `kas_bank`).
> - Akun Retained Earnings / Laba Ditahan (dihitung otomatis sebagai penyeimbang saldo).

---

## Konfigurasi Flags (`hooks.py`)

- `VALIDATE_IMPORTED_ASSETS` (default `False`): Set `True` jika ingin asset langsung tervalidasi setelah di-import.
- `POST_OPENING_MOVES` (default `False`): Set `True` jika ingin opening move partner langsung ter-post.
