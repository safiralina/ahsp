"""Mesin hitung Analisa Harga Satuan (AHSP) — tim fasilitas."""
from pathlib import Path

import pandas as pd

DATA_DIR = Path(__file__).parent / "data"
HARGA_PATH = DATA_DIR / "harga_dasar.csv"
AHSP_PATH = DATA_DIR / "ahsp_drainase.csv"
EDISI_PATH = DATA_DIR / "edisi.txt"

KOLONA_HARGA = ["jenis", "nama", "satuan", "harga"]
KOLONA_AHSP = [
    "kategori", "kode", "uraian", "satuan",
    "komponen", "jenis_komponen", "satuan_komponen", "koefisien",
]
JENIS_LABEL = {"upah": "A. Upah", "bahan": "B. Bahan", "alat": "C. Alat"}
URUTAN_JENIS = ["upah", "bahan", "alat"]


def edisi_saat_ini() -> str:
    if EDISI_PATH.exists():
        return EDISI_PATH.read_text(encoding="utf-8").strip()
    return "AHSP PUPR 2025 (SE Binkon No. 182/2025)"


def simpan_edisi(label: str) -> None:
    EDISI_PATH.write_text(label.strip(), encoding="utf-8")


def load_harga_dasar() -> pd.DataFrame:
    df = pd.read_csv(HARGA_PATH)
    df.columns = [c.strip().lower() for c in df.columns]
    df["harga"] = pd.to_numeric(df["harga"], errors="coerce").fillna(0.0)
    return df[KOLONA_HARGA].copy()


def simpan_harga_dasar(df: pd.DataFrame) -> None:
    df[KOLONA_HARGA].to_csv(HARGA_PATH, index=False)


def load_ahsp() -> pd.DataFrame:
    df = pd.read_csv(AHSP_PATH)
    df.columns = [c.strip().lower() for c in df.columns]
    df["koefisien"] = pd.to_numeric(df["koefisien"], errors="coerce").fillna(0.0)
    return df[KOLONA_AHSP].copy()


def simpan_ahsp(df: pd.DataFrame) -> None:
    df[KOLONA_AHSP].to_csv(AHSP_PATH, index=False)


def kamus_harga(harga: pd.DataFrame) -> pd.Series:
    """Seri harga dengan kunci (jenis, nama) untuk pencarian."""
    k = harga.copy()
    k["jenis"] = k["jenis"].astype(str).str.strip().str.lower()
    k["nama"] = k["nama"].astype(str).str.strip()
    return k.set_index(["jenis", "nama"])["harga"]


def harga_komponen(k: pd.Series, jenis: str, nama: str) -> float:
    kunci = (str(jenis).strip().lower(), str(nama).strip())
    if kunci in k.index:
        return float(k[kunci])
    return 0.0


def komponen_tanpa_harga(ahsp: pd.DataFrame, harga: pd.DataFrame) -> list[tuple[str, str]]:
    k = kamus_harga(harga)
    hilang = []
    for _, row in ahsp.iterrows():
        kunci = (str(row["jenis_komponen"]).strip().lower(), str(row["komponen"]).strip())
        if kunci not in k.index and kunci not in hilang:
            hilang.append(kunci)
    return hilang


def rincian_item(ahsp: pd.DataFrame, harga: pd.DataFrame, kode: str) -> pd.DataFrame:
    """Rincian komponen satu item AHSP lengkap dengan harga dasar dan subtotal."""
    k = kamus_harga(harga)
    df = ahsp[ahsp["kode"] == kode].copy()
    if df.empty:
        return pd.DataFrame()
    df["harga_dasar"] = [harga_komponen(k, j, n) for j, n in zip(df["jenis_komponen"], df["komponen"])]
    df["subtotal"] = df["koefisien"] * df["harga_dasar"]
    df["urutan"] = df["jenis_komponen"].map({j: i for i, j in enumerate(URUTAN_JENIS)}).fillna(9)
    df = df.sort_values(["urutan", "komponen"]).drop(columns="urutan")
    df.insert(0, "jenis_label", df["jenis_komponen"].map(JENIS_LABEL).fillna(df["jenis_komponen"]))
    return df


def rekap_harga_satuan(
    ahsp: pd.DataFrame,
    harga: pd.DataFrame,
    overhead_pct: float,
    profit_pct: float,
) -> pd.DataFrame:
    """Satu baris per item AHSP: subtotal A/B/C dan harga satuan jadi."""
    k = kamus_harga(harga)
    df = ahsp.copy()
    df["harga_dasar"] = [
        harga_komponen(k, j, n) for j, n in zip(df["jenis_komponen"], df["komponen"])
    ]
    df["subtotal"] = df["koefisien"] * df["harga_dasar"]
    faktor = 1.0 + overhead_pct / 100.0 + profit_pct / 100.0
    rekap = df.pivot_table(
        index=["kategori", "kode", "uraian", "satuan"],
        columns="jenis_komponen",
        values="subtotal",
        aggfunc="sum",
        fill_value=0.0,
    ).reset_index()
    for j in URUTAN_JENIS:
        if j not in rekap.columns:
            rekap[j] = 0.0
    rekap["total_dasar"] = rekap[URUTAN_JENIS].sum(axis=1)
    rekap["harga_satuan"] = rekap["total_dasar"] * faktor
    rekap = rekap.rename(columns={j: JENIS_LABEL[j] for j in URUTAN_JENIS})
    return rekap.sort_values("kode").reset_index(drop=True)


def buat_excel(
    judul: str,
    edisi: str,
    overhead_pct: float,
    profit_pct: float,
    analisa: pd.DataFrame,
    kode_item: list[str],
    ahsp: pd.DataFrame,
    harga: pd.DataFrame,
    buf,
) -> None:
    """Ekspor analisa + rincian AHSP ke file Excel (buf: BytesIO atau path).

    `analisa` berkolom: kode, uraian, qty, satuan, harga_satuan, total.
    """
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    faktor = 1.0 + overhead_pct / 100.0 + profit_pct / 100.0
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        info = pd.DataFrame({
            "Keterangan": [
                "Judul", "Sumber harga", "Overhead", "Profit",
                "Faktor pengali", "Tanggal ekspor",
            ],
            "Nilai": [
                judul, edisi, f"{overhead_pct:g}%", f"{profit_pct:g}%",
                f"{faktor:.4f}", pd.Timestamp.now().strftime("%d-%m-%Y %H:%M"),
            ],
        })
        info.to_excel(writer, sheet_name="Info", index=False)

        kolom_analisa = [
            "No", "Item Pekerjaan", "Jumlah/Qty", "Satuan",
            "Harga Satuan", "Total Harga Satuan",
        ]
        if analisa.empty:
            tabel = pd.DataFrame(columns=kolom_analisa)
        else:
            tabel = pd.DataFrame({
                "No": range(1, len(analisa) + 1),
                "Item Pekerjaan": analisa["uraian"].tolist(),
                "Jumlah/Qty": analisa["qty"].astype(float).tolist(),
                "Satuan": analisa["satuan"].tolist(),
                "Harga Satuan": analisa["harga_satuan"].astype(float).tolist(),
                "Total Harga Satuan": analisa["total"].astype(float).tolist(),
            })
            total_row = {
                "No": "", "Item Pekerjaan": "TOTAL",
                "Jumlah/Qty": float(analisa["qty"].sum()), "Satuan": "",
                "Harga Satuan": "",
                "Total Harga Satuan": float(analisa["total"].sum()),
            }
            tabel = pd.concat([tabel, pd.DataFrame([total_row])], ignore_index=True)
        tabel.to_excel(writer, sheet_name="Analisa Harga Satuan", index=False, startrow=2)
        ws_a = writer.sheets["Analisa Harga Satuan"]
        ws_a["A1"] = judul
        ws_a["A1"].font = Font(bold=True)
        ws_a["A2"] = f"{edisi} — overhead {overhead_pct:g}% + profit {profit_pct:g}%"
        if len(tabel):
            for c in ws_a[ws_a.max_row]:
                c.font = Font(bold=True)

        if kode_item:
            rincian = pd.concat(
                [rincian_item(ahsp, harga, k) for k in kode_item], ignore_index=True
            )
            rincian = rincian.rename(columns={
                "kode": "Kode", "uraian": "Item Pekerjaan", "satuan": "Satuan",
                "jenis_label": "Kelompok", "komponen": "Komponen",
                "satuan_komponen": "Satuan Komponen", "koefisien": "Koefisien",
                "harga_dasar": "Harga Dasar", "subtotal": "Subtotal",
            })
            rincian.to_excel(writer, sheet_name="Rincian AHSP", index=False)

        def rapikan(ws: "object", header_row: int, lebar: list[float]) -> None:
            headers = {c.column: str(c.value) for c in ws[header_row]}
            for i, w in enumerate(lebar, start=1):
                ws.column_dimensions[get_column_letter(i)].width = w
            for c in ws[header_row]:
                c.font = Font(bold=True)
            for row in ws.iter_rows(min_row=header_row + 1):
                for cell in row:
                    h = headers.get(cell.column, "")
                    if not isinstance(cell.value, (int, float)):
                        continue
                    if "Harga" in h or "Subtotal" in h:
                        cell.number_format = "#,##0"
                    elif h == "Jumlah/Qty":
                        cell.number_format = "#,##0.00"

        rapikan(writer.sheets["Info"], 1, [22, 48])
        rapikan(ws_a, 3, [6, 46, 12, 10, 16, 20])
        if kode_item:
            rapikan(writer.sheets["Rincian AHSP"], 1,
                    [8, 42, 8, 12, 24, 16, 10, 14, 14])
