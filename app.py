"""Analisa Harga Satuan — Tim Fasilitas (Streamlit)."""
import io

import pandas as pd
import streamlit as st

import logic

st.set_page_config(
    page_title="Analisa Harga Satuan — Tim Fasilitas",
    page_icon="🧰",
    layout="wide",
)

st.session_state.setdefault("analisa", [])  # [{kode, qty}]
st.session_state.setdefault("judul_analisa", "Analisa Harga Satuan Pekerjaan Drainase")

# ---------- sidebar ----------
with st.sidebar:
    st.title("🧰 Analisa Harga Satuan")
    st.caption(f"Sumber harga: {logic.edisi_saat_ini()}")
    st.divider()
    st.session_state["judul_analisa"] = st.text_input(
        "Judul analisa", st.session_state["judul_analisa"]
    )
    overhead = st.number_input(
        "Overhead (%)", min_value=0.0, max_value=100.0, value=10.0, step=0.5
    )
    profit = st.number_input(
        "Profit (%)", min_value=0.0, max_value=100.0, value=15.0, step=0.5
    )

# ---------- data ----------
harga = logic.load_harga_dasar()
ahsp = logic.load_ahsp()
rekap = logic.rekap_harga_satuan(ahsp, harga, overhead, profit)
rekap_index = rekap.set_index("kode")
hilang = logic.komponen_tanpa_harga(ahsp, harga)

tab_analisa, tab_detail, tab_harga = st.tabs(
    ["📋 Buat Analisa", "🔍 Detail AHSP", "💰 Kelola Harga"]
)

# ---------- tab 1: buat analisa ----------
with tab_analisa:
    kategori = st.selectbox(
        "Kategori pekerjaan",
        sorted(ahsp["kategori"].unique()),
        help="Kategori baru (mis. jalan) otomatis muncul setelah datanya dimuat di tab Kelola Harga.",
    )
    daftar = rekap[rekap["kategori"] == kategori]
    opsi = {
        f"{r.kode} — {r.uraian} ({r.satuan})": r.kode
        for r in daftar.itertuples()
    }
    kiri, kanan = st.columns([4, 1])
    pilihan = kiri.selectbox("Item pekerjaan", list(opsi), index=None,
                             placeholder="Cari / pilih item…")
    if kanan.button("➕ Tambah", width="stretch", disabled=not pilihan):
        st.session_state["analisa"].append({"kode": opsi[pilihan], "qty": 1.0})
        st.rerun()

    if not st.session_state["analisa"]:
        st.info("Belum ada item. Pilih item pekerjaan di atas lalu klik **Tambah**.")
    else:
        baris = []
        for i, item in enumerate(st.session_state["analisa"], start=1):
            r = rekap_index.loc[item["kode"]]
            qty = float(item["qty"])
            baris.append({
                "No": i,
                "Kode": item["kode"],
                "Item Pekerjaan": r["uraian"],
                "Jumlah/Qty": qty,
                "Satuan": r["satuan"],
                "Harga Satuan": float(r["harga_satuan"]),
                "Total Harga Satuan": float(r["harga_satuan"]) * qty,
                "Hapus": False,
            })
        def sinkron():
            ed = st.session_state["editor_analisa"]
            def angka(v):
                return 0.0 if pd.isna(v) else float(v)
            st.session_state["analisa"] = [
                {"kode": r["Kode"], "qty": angka(r["Jumlah/Qty"])}
                for _, r in ed.iterrows() if not r["Hapus"]
            ]

        edited = st.data_editor(
            pd.DataFrame(baris),
            key="editor_analisa",
            on_change=sinkron,
            hide_index=True,
            column_config={
                "No": st.column_config.NumberColumn(disabled=True),
                "Kode": st.column_config.TextColumn(disabled=True),
                "Item Pekerjaan": st.column_config.TextColumn(disabled=True),
                "Jumlah/Qty": st.column_config.NumberColumn(format="%.2f"),
                "Satuan": st.column_config.TextColumn(disabled=True),
                "Harga Satuan": st.column_config.NumberColumn("Harga Satuan (Rp)", format="Rp %d", disabled=True),
                "Total Harga Satuan": st.column_config.NumberColumn("Total Harga Satuan (Rp)", format="Rp %d", disabled=True),
                "Hapus": st.column_config.CheckboxColumn("Hapus"),
            },
            width="stretch",
        )
        def angka(v):
            return 0.0 if pd.isna(v) else float(v)

        baris_aktif = edited.loc[~edited["Hapus"]]
        st.session_state["analisa"] = [
            {"kode": r["Kode"], "qty": angka(r["Jumlah/Qty"])}
            for _, r in baris_aktif.iterrows()
        ]
        total = float(baris_aktif["Total Harga Satuan"].fillna(0).sum())
        st.metric("TOTAL KESELURUHAN", f"Rp {total:,.0f}".replace(",", "."))

        ekspor = baris_aktif.rename(columns={
            "Kode": "kode", "Item Pekerjaan": "uraian", "Jumlah/Qty": "qty",
            "Satuan": "satuan", "Harga Satuan": "harga_satuan",
            "Total Harga Satuan": "total",
        }).drop(columns=["No", "Hapus"])
        buf = io.BytesIO()
        logic.buat_excel(
            st.session_state["judul_analisa"], logic.edisi_saat_ini(),
            overhead, profit, ekspor,
            list(ekspor["kode"]),
            ahsp, harga, buf,
        )
        st.download_button(
            "⬇️ Ekspor Excel",
            data=buf.getvalue(),
            file_name="analisa_harga_satuan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

# ---------- tab 2: detail AHSP ----------
with tab_detail:
    st.caption("Rincian upah, bahan, dan alat untuk satu item — ini sumber harga satuannya.")
    pilih_detail = st.selectbox(
        "Item", list(opsi), index=None, placeholder="Cari / pilih item…"
    )
    if hilang:
        st.warning(
            "Komponen berikut belum punya harga dasar (dihitung Rp 0): "
            + ", ".join(f"{j}/{n}" for j, n in hilang)
            + " — lengkapi di tab Kelola Harga."
        )
    if pilih_detail:
        kode_detail = opsi[pilih_detail]
        r = rekap_index.loc[kode_detail]
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("A. Upah", f"Rp {r['A. Upah']:,.0f}".replace(",", "."))
        m2.metric("B. Bahan", f"Rp {r['B. Bahan']:,.0f}".replace(",", "."))
        m3.metric("C. Alat", f"Rp {r['C. Alat']:,.0f}".replace(",", "."))
        m4.metric("Harga Satuan", f"Rp {r['harga_satuan']:,.0f}".replace(",", "."))
        st.caption(
            f"Total dasar Rp {r['total_dasar']:,.0f} × "
            f"(1 + {overhead:g}% overhead + {profit:g}% profit)"
        )
        rincian = logic.rincian_item(ahsp, harga, kode_detail)
        st.data_editor(
            rincian[["jenis_label", "komponen", "satuan_komponen",
                     "koefisien", "harga_dasar", "subtotal"]],
            hide_index=True,
            disabled=True,
            column_config={
                "jenis_label": "Kelompok",
                "komponen": "Komponen",
                "satuan_komponen": "Satuan",
                "koefisien": st.column_config.NumberColumn(format="%.4f"),
                "harga_dasar": st.column_config.NumberColumn("Harga Dasar (Rp)", format="Rp %d"),
                "subtotal": st.column_config.NumberColumn("Subtotal (Rp)", format="Rp %d"),
            },
            width="stretch",
        )

# ---------- tab 3: kelola harga ----------
with tab_harga:
    st.caption(
        "Ubah harga dasar di tabel, lalu klik **Simpan**. Semua harga satuan "
        "otomatis menghitung ulang."
    )
    edited_harga = st.data_editor(
        harga,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "jenis": st.column_config.SelectboxColumn(
                "Jenis", options=["upah", "bahan", "alat"], required=True
            ),
            "nama": "Nama Komponen",
            "satuan": "Satuan",
            "harga": st.column_config.NumberColumn("Harga (Rp)", format="Rp %d"),
        },
        width="stretch",
        key="editor_harga",
    )
    if st.button("💾 Simpan harga dasar", type="primary"):
        logic.simpan_harga_dasar(edited_harga)
        st.success("Harga dasar tersimpan.")
        st.rerun()

    st.divider()
    st.subheader("Muat data edisi baru")
    st.caption(
        "PUPR menerbitkan AHSP terbaru tiap tahun (saat ini: "
        f"{logic.edisi_saat_ini()}). Saat ada edisi baru, siapkan file CSV/Excel "
        "dengan kolom yang sama seperti di bawah, lalu unggah di sini untuk "
        "mengganti seluruh data."
    )
    dengan_data = st.expander("Contoh struktur kolom")
    dengan_data.write("**harga_dasar**: `jenis, nama, satuan, harga`")
    dengan_data.write(
        "**AHSP**: `kategori, kode, uraian, satuan, komponen, "
        "jenis_komponen, satuan_komponen, koefisien`"
    )
    label_edisi = st.text_input(
        "Label edisi baru", "AHSP PUPR 2026", key="label_edisi"
    )
    up_harga = st.file_uploader(
        "File harga dasar baru", type=["csv", "xlsx"], key="up_harga"
    )
    up_ahsp = st.file_uploader(
        "File data AHSP baru", type=["csv", "xlsx"], key="up_ahsp"
    )
    if st.button("🔄 Muat & ganti data", disabled=not (up_harga or up_ahsp)):
        def baca(berkas):
            if berkas.name.endswith(".csv"):
                return pd.read_csv(berkas)
            return pd.read_excel(berkas)
        try:
            gagal = []
            if up_harga is not None:
                baru = baca(up_harga)
                if list(map(str.lower, baru.columns)) != logic.KOLONA_HARGA:
                    gagal.append("harga dasar (kolom tidak sesuai)")
                elif baru.empty:
                    gagal.append("harga dasar (file kosong)")
                else:
                    logic.simpan_harga_dasar(baru)
            if up_ahsp is not None:
                baru = baca(up_ahsp)
                if list(map(str.lower, baru.columns)) != logic.KOLONA_AHSP:
                    gagal.append("AHSP (kolom tidak sesuai)")
                elif baru.empty:
                    gagal.append("AHSP (file kosong)")
                else:
                    logic.simpan_ahsp(baru)
            if gagal:
                st.error("File ditolak: " + "; ".join(gagal))
            else:
                logic.simpan_edisi(label_edisi)
                st.success("Data berhasil diganti.")
                st.rerun()
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

    st.divider()
    st.download_button(
        "⬇️ Unduh template harga dasar (CSV)",
        data=(harga).to_csv(index=False).encode("utf-8"),
        file_name="template_harga_dasar.csv",
        mime="text/csv",
    )
    st.download_button(
        "⬇️ Unduh template data AHSP (CSV)",
        data=ahsp.to_csv(index=False).encode("utf-8"),
        file_name="template_ahsp.csv",
        mime="text/csv",
    )
