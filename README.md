# Analisa Harga Satuan — Tim Fasilitas

Aplikasi untuk menyusun **analisa harga satuan pekerjaan infrastruktur** (mulai dari
drainase) berbasis AHSP PUPR. Harga satuan dihitung otomatis dari rincian
A) upah, B) bahan, C) alat + overhead & profit — jadi cukup update harga dasar
sekali, seluruh item langsung terbaru.

## Menjalankan di komputer

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy ke Streamlit Community Cloud

1. Push seluruh folder ini ke repository GitHub.
2. Buka https://share.streamlit.io → **New app** → pilih repo & branch.
3. Main file path: `app.py` → **Deploy**.

## Struktur

| File | Isi |
|------|-----|
| `app.py` | UI Streamlit (3 tab: Buat Analisa, Detail AHSP, Kelola Harga) |
| `logic.py` | Mesin hitung AHSP + ekspor Excel |
| `data/harga_dasar.csv` | Harga upah/bahan/alat (per satuan) |
| `data/ahsp_drainase.csv` | Koefisien AHSP per item pekerjaan |
| `data/edisi.txt` | Label edisi sumber harga yang tampil di aplikasi |

## Update harga

- **Sesekali (harga bahan naik/turun):** edit kolom Harga di tab *Kelola Harga* → Simpan.
- **Edisi AHSP baru (mis. AHSP 2026):** siapkan CSV dengan kolom yang sama
  (template bisa diunduh dari tab *Kelola Harga*), unggah → data lama terganti,
  label edisi ikut berubah.

> Sumber koefisien: AHSP PUPR (SE Binkon No. 182/2025). Harga dasar awal adalah
> harga umum nasional 2025 — sesuaikan dengan harga setempat sebelum dipakai resmi.
