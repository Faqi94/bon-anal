import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from fpdf import FPDF
import os

# --- Konfigurasi Halaman ---
st.set_page_config(page_title="Pro Analitik Kasbon Dashboard", layout="wide")

# --- Fungsi Helper Formatting ---
def format_rupiah(value: float) -> str:
    try:
        return f"Rp {value:,.0f}".replace(",", ".")
    except Exception:
        return "Rp 0"

def format_int(num) -> str:
    """Format integer dengan pemisah ribuan '.' (contoh: 50.579)"""
    try:
        return f"{int(num):,}".replace(",", ".")
    except Exception:
        return "0"

def format_singkat(num: float) -> str:
    """Mengubah angka besar menjadi format pendek (2M, 500jt, 10k)"""
    try:
        n = float(num)
    except Exception:
        return "0"
    if n >= 1_000_000_000:
        return f"{n/1_000_000_000:.1f}M"
    elif n >= 1_000_000:
        return f"{n/1_000_000:.0f}jt"
    elif n >= 1000:
        return f"{n/1000:.0f}k"
    return f"{n:,.0f}".replace(",", ".")

def create_chart_image(fig, filename: str) -> str:
    """Simpan figure ke folder charts dan kembalikan path-nya."""
    charts_dir = "charts"
    os.makedirs(charts_dir, exist_ok=True)
    path = os.path.join(charts_dir, filename)
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)
    return path

# --- Header ---
st.title("🚀 Dashboard Analitik EWA & PPOB")
st.markdown(
    """
    1. Upload file Excel (`.xlsx`) berisi data kasbon.  
    2. Sistem akan otomatis membaca data & menampilkan dashboard.  
    3. Analitik tersedia untuk **gabungan (EWA + PPOB)**, khusus **EWA saja**, dan khusus **PPOB saja**.  
    4. Di akhir, kamu bisa **generate laporan PDF** lengkap untuk manajemen.
    5. dibuat oleh Team RnD Byru.
    """
)

# --- Upload File ---
uploaded_file = st.file_uploader(
    "Upload File Excel (misalnya: Analitics.xlsx)",
    type=["xlsx"]
)

def render_segment(seg_name: str, seg_df: pd.DataFrame, main_segment: bool = False):
    """
    Render analitik untuk satu segmen:
    - seg_name: nama segmen (Gabungan, EWA, PPOB)
    - seg_df: dataframe terfilter
    - main_segment: kalau True, tampilkan KPI cards besar
    Return dict berisi hasil penting untuk PDF.
    """
    results = {
        "name": seg_name,
        "has_data": False,
        "total_kasbon": 0.0,
        "total_trx": 0,
        "total_user": 0,
        "avg_ticket": 0.0,
        "max_ticket": 0.0,
        "monthly_stats": None,
        "top_users_amount": None,
        "top_users_qty": None,
        "path_chart1": None,
        "path_chart1b": None,   # tren user & company unik
        "path_chart3": None,
        "path_chart4": None,
        "weekend_amount": 0.0,
        "weekend_trx": 0,
        "weekend_amount_pct": 0.0,
        "weekend_trx_pct": 0.0,
    }

    if seg_df is None or seg_df.empty:
        st.info(f"Segmen **{seg_name}**: tidak ada data.")
        return results

    # pastikan copy terurut
    df_seg = seg_df.sort_values("Tanggal Approved").copy()

    # ---- METRIK UTAMA ----
    total_kasbon = float(df_seg["Total Kasbon"].sum())
    total_trx = int(len(df_seg))
    total_user = int(df_seg["Username/ ID User"].nunique())
    avg_ticket = float(df_seg["Total Kasbon"].mean())
    max_ticket = float(df_seg["Total Kasbon"].max())

    results.update(
        total_kasbon=total_kasbon,
        total_trx=total_trx,
        total_user=total_user,
        avg_ticket=avg_ticket,
        max_ticket=max_ticket,
        has_data=True,
    )

    if main_segment:
        st.markdown("### 💰 Ringkasan Performa (Gabungan EWA + PPOB)")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Pencairan", format_singkat(total_kasbon))
        k2.metric("Total Transaksi", format_int(total_trx))
        k3.metric("Rata-rata Pinjaman", format_singkat(avg_ticket))
        k4.metric("User Unik", format_int(total_user))
        st.markdown("---")
    else:
        st.markdown(f"### 📂 Ringkasan Segmen: {seg_name}")
        st.write(
            f"- Total kasbon: **{format_rupiah(total_kasbon)}** dari **{format_int(total_trx)}** transaksi "
            f"oleh **{format_int(total_user)}** user unik.  \n"
            f"- Rata-rata tiket: **{format_rupiah(avg_ticket)}**, tiket terbesar: **{format_rupiah(max_ticket)}**."
        )

    # ==============================================================
    # 1. Tren Keuangan Bulanan
    # ==============================================================
    df_seg["Bulan_Str"] = df_seg["Tanggal Approved"].dt.strftime("%b-%y")

    st.subheader(f"1. Tren Keuangan Bulanan – {seg_name}")
    monthly_stats = (
        df_seg.groupby("Bulan_Str", sort=False)["Total Kasbon"]
        .agg(["sum", "count"])
        .reset_index()
    )
    results["monthly_stats"] = monthly_stats

    fig1, ax1 = plt.subplots(figsize=(11, 6))

    bars = ax1.bar(
        monthly_stats["Bulan_Str"],
        monthly_stats["sum"],
        color="#6baed6",
        alpha=0.8,
        label="Nominal (Rp)",
    )
    ax1.set_ylabel("Total Nominal (Rp)", color="#6baed6", fontweight="bold")
    ax1.tick_params(axis="y", labelcolor="#6baed6")

    for bar in bars:
        height = bar.get_height()
        ax1.text(
            bar.get_x() + bar.get_width() / 2.0,
            height,
            format_singkat(height),
            ha="center",
            va="bottom",
            fontsize=9,
            fontweight="bold",
            color="#3182bd",
        )

    ax2 = ax1.twinx()
    ax2.plot(
        monthly_stats["Bulan_Str"],
        monthly_stats["count"],
        color="#d62728",
        marker="o",
        linewidth=2,
        label="Jumlah Transaksi",
    )
    ax2.set_ylabel("Jumlah Transaksi", color="#d62728", fontweight="bold")
    ax2.tick_params(axis="y", labelcolor="#d62728")

    for i, txt in enumerate(monthly_stats["count"]):
        ax2.text(
            i,
            txt,
            str(txt),
            ha="center",
            va="bottom",
            fontsize=9,
            color="white",
            bbox=dict(
                facecolor="#d62728", edgecolor="none", boxstyle="round,pad=0.2"
            ),
        )

    plt.xticks(rotation=45)
    plt.title(f"Total Nominal Kasbon vs Jumlah Transaksi per Bulan – {seg_name}", pad=20)
    st.pyplot(fig1)
    path_chart1 = create_chart_image(fig1, f"trend_keuangan_{seg_name}.png")
    results["path_chart1"] = path_chart1

    # ==============================================================
    # 1.a Tren User & Company Unik per Bulan
    # ==============================================================
    # ==============================================================
    # 1.a Tren User & Company Unik per Bulan
    # ==============================================================
    st.markdown(f"#### 1.a Tren User & Company Unik per Bulan – {seg_name}")

    # cari kolom company (fleksibel nama kolom, termasuk 'Nama Perushaan')
    company_col = None
    for c in ["Nama Perushaan", "Nama Perusahaan", "Company", "Nama Company"]:
        if c in df_seg.columns:
            company_col = c
            break

    # Hitung user unik per bulan
    user_per_month = (
        df_seg.groupby("Bulan_Str")["Username/ ID User"]
        .nunique()
        .reset_index(name="User Unik")
    )

    # Hitung company unik per bulan (kalau kolomnya ada)
    if company_col:
        comp_per_month = (
            df_seg.groupby("Bulan_Str")[company_col]
            .nunique()
            .reset_index(name="Company Unik")
        )
        monthly_uc = pd.merge(
            user_per_month, comp_per_month, on="Bulan_Str", how="outer"
        ).fillna(0)
    else:
        monthly_uc = user_per_month.copy()

    # 🔒 Pastikan bulan di grafik user/company PERSIS sama dengan grafik keuangan
    # (ambil urutan bulan dari monthly_stats yang sudah dipakai di grafik 1)
    base_months = monthly_stats["Bulan_Str"].tolist()
    monthly_uc = (
        monthly_uc.set_index("Bulan_Str")
        .reindex(base_months, fill_value=0)
        .reset_index()
    )

    # Pakai index numerik untuk X agar mudah ditambah label
    x = list(range(len(monthly_uc)))

    fig1b, axu = plt.subplots(figsize=(11, 4))
    axu.plot(
        x,
        monthly_uc["User Unik"],
        marker="o",
        linewidth=2,
        label="User Unik",
    )

    if "Company Unik" in monthly_uc.columns:
        axu.plot(
            x,
            monthly_uc["Company Unik"],
            marker="s",
            linestyle="--",
            linewidth=2,
            label="Company Unik",
        )

    # Label sumbu X pakai nama bulan
    axu.set_xticks(x)
    axu.set_xticklabels(monthly_uc["Bulan_Str"], rotation=45)

    axu.set_ylabel("Jumlah Unik")
    axu.set_title(f"Tren User & Company Unik per Bulan – {seg_name}")
    axu.grid(axis="y", linestyle="--", alpha=0.3)
    axu.legend()

    # === Tambah value label di atas titik User Unik ===
    max_user = monthly_uc["User Unik"].max() if len(monthly_uc) > 0 else 0
    offset_user = max_user * 0.05 if max_user > 0 else 0.3
    for i, val in enumerate(monthly_uc["User Unik"]):
        axu.text(
            x[i],
            val + offset_user,
            str(int(val)),
            ha="center",
            va="bottom",
            fontsize=9,
        )

    # === Tambah value label untuk Company Unik (kalau ada) ===
    if "Company Unik" in monthly_uc.columns:
        max_comp = monthly_uc["Company Unik"].max()
        offset_comp = max_comp * 0.05 if max_comp > 0 else 0.3
        for i, val in enumerate(monthly_uc["Company Unik"]):
            axu.text(
                x[i] + 0.1,        # geser dikit supaya nggak numpuk
                val + offset_comp,
                str(int(val)),
                ha="left",
                va="bottom",
                fontsize=9,
            )

    st.pyplot(fig1b)
    path_chart1b = create_chart_image(fig1b, f"trend_user_company_{seg_name}.png")
    results["path_chart1b"] = path_chart1b

    # ==============================================================
    # 2. Top 10 Karyawan (Nominal & Frekuensi)
    # ==============================================================
    st.subheader(f"2. Top 10 Karyawan – {seg_name}")

    # Tentukan kolom nama karyawan & nama perusahaan
    nama_karyawan_col = (
        "Nama Karyawan" if "Nama Karyawan" in df_seg.columns else "Username/ ID User"
    )
    if "Nama Perusahaan" in df_seg.columns:
        nama_perusahaan_col = "Nama Perusahaan"
    elif "Nama Perushaan" in df_seg.columns:  # typo safe
        nama_perusahaan_col = "Nama Perushaan"
    else:
        nama_perusahaan_col = None

    group_cols = [nama_karyawan_col]
    if "Username/ ID User" in df_seg.columns and "Username/ ID User" not in group_cols:
        group_cols.append("Username/ ID User")
    if nama_perusahaan_col and nama_perusahaan_col not in group_cols:
        group_cols.append(nama_perusahaan_col)

    agg_users = (
        df_seg.groupby(group_cols)["Total Kasbon"]
        .agg(Qty_EWA_PPOB="count", Total_Kasbon="sum")
        .reset_index()
    )

    top_users_amount = (
        agg_users.sort_values("Total_Kasbon", ascending=False)
        .head(10)
        .reset_index(drop=True)
    )
    top_users_qty = (
        agg_users.sort_values("Qty_EWA_PPOB", ascending=False)
        .head(10)
        .reset_index(drop=True)
    )

    results["top_users_amount"] = top_users_amount
    results["top_users_qty"] = top_users_qty

    # Chart Top 10 berdasarkan nominal
    if not top_users_amount.empty:
        fig3, ax3 = plt.subplots(figsize=(12, 7))
        y_pos = range(len(top_users_amount))
        bars_h = ax3.barh(
            y_pos,
            top_users_amount["Total_Kasbon"],
            color="#0ea5e9",
            alpha=0.9,
        )

        ax3.set_yticks(y_pos)
        ax3.set_yticklabels(top_users_amount[nama_karyawan_col], fontsize=10)
        ax3.invert_yaxis()

        ax3.set_xlabel("Total Nilai Pinjaman (Rp)", fontsize=11)
        ax3.set_title(
            f"Top 10 Karyawan Paling Boros (Nominal) – {seg_name}",
            fontsize=14,
            pad=15,
        )

        ax3.xaxis.set_major_formatter(
            mtick.FuncFormatter(lambda x, pos: format_singkat(x))
        )
        ax3.grid(axis="x", linestyle="--", alpha=0.3)

        max_val_amt = float(top_users_amount["Total_Kasbon"].max())
        for bar in bars_h:
            width = bar.get_width()
            label_x = width + (max_val_amt * 0.01 if max_val_amt > 0 else 0)
            ax3.text(
                label_x,
                bar.get_y() + bar.get_height() / 2,
                format_singkat(width),
                va="center",
                fontsize=10,
                fontweight="bold",
                color="#111111",
            )

        plt.tight_layout()
        st.pyplot(fig3)
        path_chart3 = create_chart_image(fig3, f"top_users_{seg_name}.png")
        results["path_chart3"] = path_chart3
    else:
        st.info(f"Tidak ada data Top 10 karyawan untuk segmen {seg_name}.")

    # Helper tabel
    def _render_top_table(df_source, section_title: str):
        st.markdown(f"#### {section_title}")
        if df_source.empty:
            st.info("Belum ada data untuk ditampilkan.")
            return

        table_df = df_source.copy()
        table_df.insert(0, "No", table_df.index + 1)

        rename_map = {
            nama_karyawan_col: "Nama Karyawan",
            "Qty_EWA_PPOB": "Qty",
            "Total_Kasbon": "Total Amount",
        }
        if "Username/ ID User" in table_df.columns:
            rename_map["Username/ ID User"] = "Username/ ID User"
        if nama_perusahaan_col and nama_perusahaan_col in table_df.columns:
            rename_map[nama_perusahaan_col] = "Nama Perusahaan"

        table_df = table_df.rename(columns=rename_map)
        table_df["Total Amount"] = table_df["Total Amount"].apply(format_rupiah)

        display_cols = ["No", "Nama Karyawan"]
        if "Username/ ID User" in table_df.columns:
            display_cols.append("Username/ ID User")
        if "Nama Perusahaan" in table_df.columns:
            display_cols.append("Nama Perusahaan")
        display_cols += ["Qty", "Total Amount"]

        st.dataframe(table_df[display_cols], use_container_width=True)

    _render_top_table(
        top_users_amount, f"Detail Top 10 Karyawan Paling Boros – {seg_name}"
    )
    _render_top_table(
        top_users_qty,
        f"Detail Top 10 Karyawan Paling Banyak Qty Transaksinya – {seg_name}",
    )

    # ==============================================================
    # 3. Analisis Hari & Weekend
    # ==============================================================
    st.subheader(f"3. Analisis Hari & Weekend – {seg_name}")

    trx_per_day = (
        df_seg["Hari"]
        .value_counts()
        .reindex(
            [
                "Monday",
                "Tuesday",
                "Wednesday",
                "Thursday",
                "Friday",
                "Saturday",
                "Sunday",
            ]
        )
        .fillna(0)
        .astype(int)
        .reset_index()
    )
    trx_per_day.columns = ["Hari", "Jumlah"]

    fig4, ax4 = plt.subplots(figsize=(10, 5))
    colors = [
        "#b3cde3" if x < 5 else "#fdb462"
        for x in range(len(trx_per_day))
    ]
    bars_d = ax4.bar(trx_per_day["Hari"], trx_per_day["Jumlah"], color=colors)

    max_trx = trx_per_day["Jumlah"].max()
    for bar in bars_d:
        height = bar.get_height()
        ax4.text(
            bar.get_x() + bar.get_width() / 2.0,
            height + (max_trx * 0.03 if max_trx > 0 else 0.1),
            format_int(height),
            ha="center",
            va="bottom",
            fontsize=10,
        )

    ax4.set_title(f"Volume Transaksi per Hari – {seg_name}")
    if max_trx > 0:
        ax4.set_ylim(top=max_trx * 1.25)

    st.pyplot(fig4)
    path_chart4 = create_chart_image(fig4, f"daily_trx_{seg_name}.png")
    results["path_chart4"] = path_chart4

    # Weekend contribution
    weekend_mask = df_seg["Hari"].isin(["Saturday", "Sunday"])
    weekend_amount = float(df_seg.loc[weekend_mask, "Total Kasbon"].sum())
    weekend_trx = int(weekend_mask.sum())
    total_amount_all = float(total_kasbon)
    total_trx_all = int(total_trx)

    weekend_amount_pct = (
        weekend_amount / total_amount_all * 100 if total_amount_all > 0 else 0.0
    )
    weekend_trx_pct = (
        weekend_trx / total_trx_all * 100 if total_trx_all > 0 else 0.0
    )

    results["weekend_amount"] = weekend_amount
    results["weekend_trx"] = weekend_trx
    results["weekend_amount_pct"] = weekend_amount_pct
    results["weekend_trx_pct"] = weekend_trx_pct

    st.markdown(
        f"📌 **Kontribusi Akhir Pekan (Sabtu & Minggu) – {seg_name}**: "
        f"{format_rupiah(weekend_amount)} dari {format_rupiah(total_amount_all)} "
        f"({weekend_amount_pct:.1f}% dari total nominal kasbon, "
        f"{format_int(weekend_trx)} dari {format_int(total_trx_all)} transaksi / "
        f"{weekend_trx_pct:.1f}% dari total transaksi)."
    )

    return results


if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        required_columns = ["Tanggal Approved", "Username/ ID User", "Total Kasbon"]
        if not all(col in df.columns for col in required_columns):
            st.error(
                "Kolom wajib tidak ditemukan! "
                f"Pastikan ada kolom: {', '.join(required_columns)}"
            )
        else:
            if df.empty:
                st.warning("File terbaca, tapi tidak ada data di dalamnya.")
            else:
                st.success("✅ Data berhasil dimuat. Melakukan analisis...")

                # Cleaning tanggal
                df["Tanggal Approved"] = pd.to_datetime(
                    df["Tanggal Approved"], errors="coerce"
                )
                df = df[df["Tanggal Approved"].notna()].copy()

                if df.empty:
                    st.warning(
                        "Semua baris memiliki 'Tanggal Approved' yang tidak valid. "
                        "Cek format tanggal di file Excel."
                    )
                else:
                    df["Hari"] = df["Tanggal Approved"].dt.day_name()

                    # Deteksi kolom jenis EWA (EWA / PPOB)
                    jenis_candidates = [
                        "Jenis EWA",
                        "JENIS EWA",
                        "Jenis",
                        "JENIS",
                        "Jenis Transaksi",
                        "Jenis_Kasbon",
                    ]
                    jenis_col = None
                    for c in jenis_candidates:
                        if c in df.columns:
                            jenis_col = c
                            break

                    if jenis_col is None:
                        st.warning(
                            "Kolom jenis EWA (EWA/PPOB) tidak ditemukan. "
                            "Analisis hanya dilakukan sebagai gabungan (EWA+PPOB)."
                        )

                    # Bangun segmen
                    segments = {}
                    segments["Gabungan (EWA+PPOB)"] = df

                    if jenis_col is not None:
                        jenis_upper = df[jenis_col].astype(str).str.upper()
                        df_ewa = df[jenis_upper == "EWA"].copy()
                        df_ppob = df[jenis_upper == "PPOB"].copy()
                        segments["EWA"] = df_ewa
                        segments["PPOB"] = df_ppob

                    # Render gabungan dulu
                    results_all = render_segment(
                        "Gabungan (EWA+PPOB)", segments["Gabungan (EWA+PPOB)"], main_segment=True
                    )

                    # Jika ada kolom jenis, render EWA & PPOB
                    results_ewa = None
                    results_ppob = None
                    if "EWA" in segments:
                        st.markdown("---")
                        results_ewa = render_segment("EWA", segments["EWA"], main_segment=False)
                    if "PPOB" in segments:
                        st.markdown("---")
                        results_ppob = render_segment("PPOB", segments["PPOB"], main_segment=False)

                    # ==============================================================
                    # PDF REPORT (berbasis gabungan + ringkasan per jenis)
                    # ==============================================================
                    st.markdown("---")
                    st.subheader("📄 Download Laporan PDF")

                    if st.button("Generate Laporan Lengkap (PDF)"):

                        # Sanitize text agar aman untuk FPDF (latin-1)
                        def pdf_safe(text: str) -> str:
                            if not isinstance(text, str):
                                text = str(text)
                            # ganti karakter "aneh" yang sering bikin error
                            text = (
                                text.replace("–", "-")
                                    .replace("—", "-")
                                    .replace("•", "-")
                            )
                            return text.encode("latin-1", "replace").decode("latin-1")

                        class PDF(FPDF):
                            def header(self):
                                self.set_font("Arial", "B", 16)
                                self.cell(
                                    0,
                                    10,
                                    pdf_safe("Laporan Analitik Kasbon"),
                                    0,
                                    1,
                                    "C",
                                )
                                self.set_font("Arial", "I", 10)
                                self.cell(
                                    0,
                                    10,
                                    pdf_safe("Generated by Dashboard Analitik Kasbon"),
                                    0,
                                    1,
                                    "C",
                                )
                                self.line(10, 30, 200, 30)
                                self.ln(10)

                            def chapter_title(self, title: str):
                                self.set_font("Arial", "B", 14)
                                self.set_fill_color(230, 230, 230)
                                self.cell(0, 10, pdf_safe(title), 0, 1, "L", 1)
                                self.ln(4)

                            def chapter_body(self, body: str):
                                self.set_font("Arial", "", 11)
                                self.multi_cell(0, 6, pdf_safe(body))
                                self.ln()

                        pdf = PDF()
                        pdf.set_auto_page_break(auto=True, margin=15)
                        pdf.add_page()

                        # -----------------------------
                        # Ambil data utama (Gabungan)
                        # -----------------------------
                        total_kasbon = results_all["total_kasbon"]
                        total_trx = results_all["total_trx"]
                        total_user = results_all["total_user"]
                        avg_ticket = results_all["avg_ticket"]
                        max_ticket = results_all["max_ticket"]
                        monthly_stats = results_all["monthly_stats"]
                        path_chart1_all = results_all["path_chart1"]
                        path_chart1b_all = results_all.get("path_chart1b")
                        path_chart3_all = results_all["path_chart3"]
                        path_chart4_all = results_all["path_chart4"]
                        weekend_amount_all = results_all["weekend_amount"]
                        weekend_trx_all = results_all["weekend_trx"]
                        weekend_amount_pct_all = results_all["weekend_amount_pct"]
                        weekend_trx_pct_all = results_all["weekend_trx_pct"]
                        top_users_amount_all = results_all["top_users_amount"]
                        top_users_qty_all = results_all["top_users_qty"]

                        # Periode data
                        periode_start = df["Tanggal Approved"].min()
                        periode_end = df["Tanggal Approved"].max()
                        if pd.notna(periode_start) and pd.notna(periode_end):
                            periode_str = f"{periode_start:%d %b %Y} - {periode_end:%d %b %Y}"
                        else:
                            periode_str = "Tidak diketahui"

                        # -----------------------------
                        # 1. RINGKASAN EKSEKUTIF
                        # -----------------------------
                        if monthly_stats is not None and not monthly_stats.empty:
                            bulan_max = monthly_stats.loc[monthly_stats["sum"].idxmax()]
                        else:
                            bulan_max = None

                        # MoM growth (kalau minimal ada 2 bulan)
                        mom_text = ""
                        if monthly_stats is not None and len(monthly_stats) >= 2:
                            last = monthly_stats.iloc[-2]
                            current = monthly_stats.iloc[-1]
                            if last["sum"] > 0:
                                mom_pct = (current["sum"] - last["sum"]) / last["sum"] * 100
                                arah = "naik" if mom_pct >= 0 else "turun"
                                mom_text = (
                                    f"Dibanding bulan sebelumnya, total kasbon {arah} "
                                    f"{abs(mom_pct):.1f}%."
                                )

                        ringkasan_lines = [
                            f"Periode data: {periode_str}.",
                            f"Total kasbon (gabungan EWA+PPOB): {format_rupiah(total_kasbon)} "
                            f"dari {total_trx} transaksi oleh {total_user} user unik.",
                            f"Rata-rata ticket size: {format_rupiah(avg_ticket)} | "
                            f"Ticket terbesar: {format_rupiah(max_ticket)}.",
                        ]
                        if bulan_max is not None:
                            ringkasan_lines.append(
                                f"Bulan dengan pencairan tertinggi: {bulan_max['Bulan_Str']} "
                                f"sebesar {format_rupiah(bulan_max['sum'])} "
                                f"dari {bulan_max['count']} transaksi."
                            )
                        if mom_text:
                            ringkasan_lines.append(mom_text)
                        ringkasan_lines.append(
                            "Kontribusi akhir pekan (Sabtu-Minggu, gabungan): "
                            f"{format_rupiah(weekend_amount_all)} "
                            f"({weekend_amount_pct_all:.1f}% dari nominal, "
                            f"{weekend_trx_pct_all:.1f}% dari jumlah transaksi)."
                        )

                        # Ringkasan per jenis (Gabungan, EWA, PPOB)
                        jenis_lines = []
                        for res in [results_all, results_ewa, results_ppob]:
                            if not res:
                                continue
                            if not res.get("has_data", False):
                                continue
                            name = res["name"]
                            tot = format_rupiah(res["total_kasbon"])
                            trx = res["total_trx"]
                            wu = res["weekend_amount_pct"]
                            wt = res["weekend_trx_pct"]
                            jenis_lines.append(
                                f"- {name}: {tot} ({trx} trx, weekend {wu:.1f}% nominal / {wt:.1f}% trx)"
                            )
                        if jenis_lines:
                            ringkasan_lines.append(
                                "Ringkasan per jenis (Gabungan, EWA, PPOB):\n"
                                + "\n".join(jenis_lines)
                            )

                        pdf.chapter_title("1. Ringkasan Eksekutif & Perbandingan Jenis")
                        pdf.chapter_body("\n".join(ringkasan_lines))

                        # -----------------------------
                        # 2. TREN BULANAN (GABUNGAN)
                        # -----------------------------
                        pdf.chapter_title("2. Tren Keuangan Bulanan - Gabungan (EWA+PPOB)")
                        if path_chart1_all and os.path.exists(path_chart1_all):
                            pdf.image(path_chart1_all, w=180)
                            pdf.ln(5)
                        pdf.chapter_body(
                            "Grafik di atas menunjukkan perkembangan total nominal kasbon "
                            "dan jumlah transaksi per bulan untuk gabungan EWA+PPOB. "
                            "Pimpinan dapat memonitor pertumbuhan penggunaan kasbon dan "
                            "mengidentifikasi bulan dengan lonjakan signifikan."
                        )
                        # --- 2.a Tren User & Company Unik per Bulan (Gabungan) ---
                        pdf.chapter_title("2.a Tren User & Company Unik per Bulan - Gabungan")
                        if path_chart1b_all and os.path.exists(path_chart1b_all):
                            pdf.image(path_chart1b_all, w=180)
                            pdf.ln(5)

                        pdf.chapter_body(
                            "Grafik ini menunjukkan perkembangan jumlah user unik dan company unik "
                            "yang aktif menggunakan kasbon per bulan. Tren kenaikan mengindikasikan "
                            "adopsi yang semakin luas, baik dari sisi karyawan maupun perusahaan."
                        )

                        # -----------------------------
                        # 3. TOP 10 PALING BOROS
                        # -----------------------------
                        pdf.chapter_title("3. Top 10 Karyawan Paling Boros - Gabungan")
                        if path_chart3_all and os.path.exists(path_chart3_all):
                            pdf.image(path_chart3_all, w=180)
                            pdf.ln(5)
                        pdf.chapter_body(
                            "Grafik di atas menunjukkan 10 karyawan dengan total "
                            "pencairan kasbon tertinggi. Informasi ini membantu manajemen "
                            "mengidentifikasi pengguna kasbon terbesar dan potensi risiko."
                        )

                        pdf_output_path = "Laporan_Analitik_Lengkap.pdf"
                        pdf.output(pdf_output_path)

                        with open(pdf_output_path, "rb") as f:
                            pdf_bytes = f.read()

                        st.success("PDF berhasil dibuat!")
                        st.download_button(
                            label="📥 Download PDF",
                            data=pdf_bytes,
                            file_name="Laporan_Analitik_Lengkap.pdf",
                            mime="application/pdf",
                        )

    except Exception as e:
        st.error(f"Terjadi error: {e}")

else:
    st.info("Silakan upload file Excel terlebih dahulu untuk memulai analisis.")
