"""
LBP Sales Monitor - Mobile-friendly Streamlit App
Upload XLSX sales file → get full monitoring dashboard
"""
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="LBP Sales Monitor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS for mobile ──
st.markdown("""
<style>
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
    div[data-testid="stMetric"] {
        background: #f0f7ff;
        border: 1px solid #d0e3f5;
        border-radius: 10px;
        padding: 10px 14px;
    }
    div[data-testid="stMetric"] label { font-size: 0.8rem !important; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] { font-size: 1.2rem !important; }
    .stTabs [data-baseweb="tab"] { font-size: 0.85rem; padding: 8px 12px; }
    h1 { font-size: 1.5rem !important; }
    h2 { font-size: 1.2rem !important; }
    h3 { font-size: 1.05rem !important; }
    .stDataFrame { font-size: 0.8rem; }
</style>
""", unsafe_allow_html=True)


# ── Helpers ──
def fmt_rp(v):
    if pd.isna(v):
        return "-"
    return f"Rp {v:,.0f}"

def fmt_num(v):
    if pd.isna(v):
        return "-"
    return f"{v:,.0f}"

def load_lbp(file) -> pd.DataFrame:
    """Load and clean LBP-style Excel."""
    # Try header=1 first (LBP format has title row)
    df = pd.read_excel(file, header=1)
    df.columns = df.columns.str.strip()

    # If columns look wrong, try header=0
    if "No Outlet" not in df.columns and "Nama Outlet" not in df.columns:
        file.seek(0)
        df = pd.read_excel(file, header=0)
        df.columns = df.columns.str.strip()

    # Standardize expected columns
    col_map = {}
    for c in df.columns:
        cl = c.lower().replace(" ", "")
        if "nooutlet" in cl or c == "No Outlet":
            col_map[c] = "No Outlet"
        elif "namaoutlet" in cl or c == "Nama Outlet":
            col_map[c] = "Nama Outlet"
        elif "tanggalfaktur" in cl or "tanggal" in cl:
            col_map[c] = "Tanggal Faktur"
        elif c in ("QTYPCS", "Qty", "qty"):
            col_map[c] = "QTYPCS"
        elif "hargabruto" in cl or c == "Harga Bruto":
            col_map[c] = "Harga Bruto"
        elif c in ("Total", "total", "NET", "Net"):
            col_map[c] = "Total"
        elif c in ("DISC", "Disc", "Discount"):
            col_map[c] = "DISC"
        elif "proamount" in cl or c == "PROAMOUNT":
            col_map[c] = "PROAMOUNT"
        elif "namaproduk" in cl or c == "Nama Produk":
            col_map[c] = "Nama Produk"
        elif "subbrandname" in cl or c == "SUBBRANDNAME":
            col_map[c] = "SUBBRANDNAME"
        elif c in ("Pcode", "PCODE", "pcode"):
            col_map[c] = "Pcode"
        elif "salesman" in cl:
            col_map[c] = "Salesman"
        elif "kabupaten" in cl:
            col_map[c] = "Kabupaten"
        elif "kecamatan" in cl:
            col_map[c] = "Kecamatan"
        elif "channel" in cl:
            col_map[c] = "Channel"
        elif c.strip() in ("WEEK", "Week", "WEEK "):
            col_map[c] = "WEEK"
        elif "faktur" in cl and "tanggal" not in cl:
            col_map[c] = "Faktur"
        elif "transtype" in cl:
            col_map[c] = "TRANSTYPE"

    df = df.rename(columns=col_map)

    # Numeric
    for col in ["QTYPCS", "Harga Bruto", "Total", "DISC", "PROAMOUNT", "AMOUNT"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Date
    if "Tanggal Faktur" in df.columns:
        df["Tanggal Faktur"] = pd.to_datetime(
            df["Tanggal Faktur"], format="%d/%m/%Y", errors="coerce"
        )
        if df["Tanggal Faktur"].isna().all():
            df["Tanggal Faktur"] = pd.to_datetime(df["Tanggal Faktur"], errors="coerce")
        df["Tanggal"] = df["Tanggal Faktur"].dt.date
        df["Hari"] = df["Tanggal Faktur"].dt.day_name()

    # Fill missing text cols
    for col in ["Nama Outlet", "Nama Produk", "SUBBRANDNAME", "Salesman",
                "Kabupaten", "Kecamatan", "Channel", "Faktur"]:
        if col not in df.columns:
            df[col] = "-"
        else:
            df[col] = df[col].fillna("-").astype(str)

    if "WEEK" not in df.columns and "Tanggal Faktur" in df.columns:
        df["WEEK"] = df["Tanggal Faktur"].dt.isocalendar().week.astype(int)

    if "Pcode" not in df.columns:
        df["Pcode"] = df.get("Nama Produk", "-")

    if "No Outlet" not in df.columns:
        df["No Outlet"] = df.get("Nama Outlet", "-")

    if "Harga Bruto" not in df.columns:
        df["Harga Bruto"] = df.get("Total", 0)

    if "QTYPCS" not in df.columns:
        df["QTYPCS"] = 0

    if "Total" not in df.columns:
        df["Total"] = df.get("Harga Bruto", 0)

    return df


def kpi_row(df):
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Net Sales", fmt_rp(df["Total"].sum()))
    c2.metric("Omset Bruto", fmt_rp(df["Harga Bruto"].sum()))
    c3.metric("Quantity", fmt_num(df["QTYPCS"].sum()) + " pcs")
    c4.metric("Outlet", fmt_num(df["No Outlet"].nunique()))
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Faktur", fmt_num(df["Faktur"].nunique()) if "Faktur" in df.columns else "-")
    c6.metric("Produk SKU", fmt_num(df["Pcode"].nunique()))
    c7.metric("Discount", fmt_rp(df["DISC"].sum()) if "DISC" in df.columns else "-")
    ret = df[df["Total"] < 0]["Total"].sum()
    c8.metric("Return", fmt_rp(ret))


# ── Sidebar ──
st.sidebar.title("📊 LBP Monitor")
st.sidebar.caption("Upload file XLSX penjualan")

uploaded = st.sidebar.file_uploader("Pilih file Excel", type=["xlsx", "xls", "csv"])

if uploaded is None:
    st.title("📊 LBP Sales Monitor")
    st.info(
        "👈 Upload file Excel (format LBP) di sidebar kiri untuk mulai monitoring.\n\n"
        "**Fitur:**\n"
        "- KPI ringkasan (Qty, Omset Bruto, Net Sales)\n"
        "- Monitoring per Produk, Outlet, Salesman, Wilayah\n"
        "- Tren harian & mingguan\n"
        "- Filter interaktif\n"
        "- Optimasi tampilan HP"
    )
    st.markdown("---")
    st.markdown(
        "**Format yang didukung:** File XLSX dengan kolom seperti "
        "`No Outlet`, `Nama Outlet`, `Nama Produk`, `QTYPCS`, `Harga Bruto`, `Total`, "
        "`Salesman`, `Kabupaten`, `Channel`, `Tanggal Faktur`, dll."
    )
    st.stop()

# ── Load ──
with st.spinner("Memproses file..."):
    try:
        df = load_lbp(uploaded)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.stop()

st.sidebar.success(f"✅ {len(df):,} baris | {df['No Outlet'].nunique():,} outlet")

# ── Global filters ──
st.sidebar.markdown("### 🔎 Filter")
all_kab = sorted(df["Kabupaten"].unique().tolist())
sel_kab = st.sidebar.multiselect("Kabupaten", all_kab, default=all_kab)

all_sm = sorted(df["Salesman"].unique().tolist())
sel_sm = st.sidebar.multiselect("Salesman", all_sm, default=[])

all_sb = sorted(df["SUBBRANDNAME"].unique().tolist())
sel_sb = st.sidebar.multiselect("Subbrand", all_sb, default=[])

all_ch = sorted(df["Channel"].unique().tolist())
sel_ch = st.sidebar.multiselect("Channel", all_ch, default=[])

# Apply filters
fdf = df.copy()
if sel_kab:
    fdf = fdf[fdf["Kabupaten"].isin(sel_kab)]
if sel_sm:
    fdf = fdf[fdf["Salesman"].isin(sel_sm)]
if sel_sb:
    fdf = fdf[fdf["SUBBRANDNAME"].isin(sel_sb)]
if sel_ch:
    fdf = fdf[fdf["Channel"].isin(sel_ch)]

if "Tanggal Faktur" in fdf.columns and fdf["Tanggal Faktur"].notna().any():
    min_d = fdf["Tanggal Faktur"].min().date()
    max_d = fdf["Tanggal Faktur"].max().date()
    date_range = st.sidebar.date_input("Rentang Tanggal", [min_d, max_d], min_value=min_d, max_value=max_d)
    if len(date_range) == 2:
        fdf = fdf[
            (fdf["Tanggal Faktur"].dt.date >= date_range[0])
            & (fdf["Tanggal Faktur"].dt.date <= date_range[1])
        ]

st.sidebar.markdown(f"**Data aktif:** {len(fdf):,} baris")

# ── Main ──
st.title("📊 LBP Sales Monitor")
if "Tanggal Faktur" in fdf.columns and fdf["Tanggal Faktur"].notna().any():
    st.caption(
        f"Periode: {fdf['Tanggal Faktur'].min().strftime('%d/%m/%Y')} – "
        f"{fdf['Tanggal Faktur'].max().strftime('%d/%m/%Y')} | "
        f"Update: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )

kpi_row(fdf)
st.markdown("---")

tabs = st.tabs([
    "📦 Produk", "🏪 Outlet", "👤 Salesman",
    "📍 Wilayah", "📅 Tren", "🔄 Return", "⬇️ Export"
])

# ══════ TAB PRODUK ══════
with tabs[0]:
    st.subheader("Monitoring per Produk")
    prod = (
        fdf.groupby(["Pcode", "Nama Produk", "SUBBRANDNAME"])
        .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"),
             Net=("Total", "sum"), Disc=("DISC", "sum"),
             Outlet=("No Outlet", "nunique"), Faktur=("Faktur", "nunique"))
        .reset_index()
        .sort_values("Bruto", ascending=False)
    )
    total_bruto = prod["Bruto"].sum() or 1
    total_qty = prod["Qty"].sum() or 1
    prod["% Bruto"] = (prod["Bruto"] / total_bruto * 100).round(1)
    prod["% Qty"] = (prod["Qty"] / total_qty * 100).round(1)

    # Search product
    q = st.text_input("🔍 Cari produk", "", key="prod_search")
    show = prod.copy()
    if q:
        show = show[show["Nama Produk"].str.contains(q, case=False, na=False)]

    st.dataframe(
        show.rename(columns={
            "Nama Produk": "Produk", "SUBBRANDNAME": "Subbrand",
            "Qty": "Qty (pcs)", "Bruto": "Omset Bruto", "Net": "Net Sales",
            "Disc": "Discount", "Outlet": "Jml Outlet", "Faktur": "Jml Faktur"
        }),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
            "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
            "Discount": st.column_config.NumberColumn(format="Rp %d"),
            "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
            "% Bruto": st.column_config.NumberColumn(format="%.1f%%"),
            "% Qty": st.column_config.NumberColumn(format="%.1f%%"),
        },
    )

    st.markdown("#### Top 10 Produk — Omset Bruto")
    top10 = prod.head(10)
    st.bar_chart(top10.set_index("Nama Produk")["Bruto"])

    st.markdown("#### Top 10 Produk — Quantity")
    st.bar_chart(top10.set_index("Nama Produk")["Qty"])

    # Subbrand summary
    st.markdown("#### Ringkasan Subbrand")
    sb = (
        fdf.groupby("SUBBRANDNAME")
        .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"))
        .reset_index().sort_values("Bruto", ascending=False)
    )
    st.dataframe(
        sb.rename(columns={"SUBBRANDNAME": "Subbrand", "Qty": "Qty (pcs)",
                           "Bruto": "Omset Bruto", "Net": "Net Sales"}),
        use_container_width=True, hide_index=True,
        column_config={
            "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
            "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
            "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
        },
    )

# ══════ TAB OUTLET ══════
with tabs[1]:
    st.subheader("Monitoring per Outlet")
    out = (
        fdf.groupby(["No Outlet", "Nama Outlet", "Kabupaten", "Channel"])
        .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"),
             Net=("Total", "sum"), Faktur=("Faktur", "nunique"),
             SKU=("Pcode", "nunique"))
        .reset_index()
        .sort_values("Bruto", ascending=False)
    )

    q2 = st.text_input("🔍 Cari outlet", "", key="out_search")
    show2 = out.copy()
    if q2:
        show2 = show2[show2["Nama Outlet"].str.contains(q2, case=False, na=False)]

    top_n = st.slider("Tampilkan Top N outlet", 10, 100, 30, key="out_topn")
    st.dataframe(
        show2.head(top_n).rename(columns={
            "Nama Outlet": "Outlet", "Qty": "Qty (pcs)",
            "Bruto": "Omset Bruto", "Net": "Net Sales",
            "Faktur": "Jml Faktur", "SKU": "Jml SKU"
        }),
        use_container_width=True, hide_index=True,
        column_config={
            "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
            "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
            "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
        },
    )

    st.markdown("#### Top 15 Outlet — Omset Bruto")
    st.bar_chart(out.head(15).set_index("Nama Outlet")["Bruto"])

    # Product x Outlet detail
    st.markdown("#### Detail Produk di Outlet")
    sel_out = st.selectbox(
        "Pilih outlet",
        options=out["Nama Outlet"].head(50).tolist(),
        key="sel_outlet",
    )
    if sel_out:
        od = fdf[fdf["Nama Outlet"] == sel_out]
        od_prod = (
            od.groupby(["Nama Produk", "SUBBRANDNAME"])
            .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"))
            .reset_index().sort_values("Bruto", ascending=False)
        )
        st.dataframe(
            od_prod.rename(columns={"Nama Produk": "Produk", "SUBBRANDNAME": "Subbrand",
                                    "Qty": "Qty (pcs)", "Bruto": "Omset Bruto", "Net": "Net Sales"}),
            use_container_width=True, hide_index=True,
            column_config={
                "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
                "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
                "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
            },
        )

# ══════ TAB SALESMAN ══════
with tabs[2]:
    st.subheader("Monitoring per Salesman")
    sm = (
        fdf.groupby("Salesman")
        .agg(Outlet=("No Outlet", "nunique"), Qty=("QTYPCS", "sum"),
             Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"),
             Faktur=("Faktur", "nunique"), SKU=("Pcode", "nunique"))
        .reset_index().sort_values("Bruto", ascending=False)
    )
    st.dataframe(
        sm.rename(columns={
            "Outlet": "Jml Outlet", "Qty": "Qty (pcs)",
            "Bruto": "Omset Bruto", "Net": "Net Sales",
            "Faktur": "Jml Faktur", "SKU": "Jml SKU"
        }),
        use_container_width=True, hide_index=True,
        column_config={
            "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
            "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
            "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
        },
    )

    st.markdown("#### Ranking Salesman — Omset Bruto")
    st.bar_chart(sm.head(15).set_index("Salesman")["Bruto"])

    st.markdown("#### Detail Produk per Salesman")
    sel_s = st.selectbox("Pilih salesman", sm["Salesman"].tolist(), key="sel_sm")
    if sel_s:
        sd = fdf[fdf["Salesman"] == sel_s]
        sd_prod = (
            sd.groupby(["Nama Produk", "SUBBRANDNAME"])
            .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"),
                 Net=("Total", "sum"), Outlet=("No Outlet", "nunique"))
            .reset_index().sort_values("Bruto", ascending=False)
        )
        st.dataframe(
            sd_prod.rename(columns={
                "Nama Produk": "Produk", "SUBBRANDNAME": "Subbrand",
                "Qty": "Qty (pcs)", "Bruto": "Omset Bruto", "Net": "Net Sales",
                "Outlet": "Jml Outlet"
            }),
            use_container_width=True, hide_index=True,
            column_config={
                "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
                "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
                "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
            },
        )

# ══════ TAB WILAYAH ══════
with tabs[3]:
    st.subheader("Monitoring per Wilayah")
    kab = (
        fdf.groupby("Kabupaten")
        .agg(Outlet=("No Outlet", "nunique"), Qty=("QTYPCS", "sum"),
             Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"),
             Faktur=("Faktur", "nunique"))
        .reset_index().sort_values("Bruto", ascending=False)
    )
    st.dataframe(
        kab.rename(columns={
            "Outlet": "Jml Outlet", "Qty": "Qty (pcs)",
            "Bruto": "Omset Bruto", "Net": "Net Sales", "Faktur": "Jml Faktur"
        }),
        use_container_width=True, hide_index=True,
        column_config={
            "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
            "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
            "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
        },
    )
    st.bar_chart(kab.set_index("Kabupaten")["Bruto"])

    st.markdown("#### Per Kecamatan")
    kec = (
        fdf.groupby(["Kabupaten", "Kecamatan"])
        .agg(Outlet=("No Outlet", "nunique"), Qty=("QTYPCS", "sum"),
             Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"))
        .reset_index().sort_values("Bruto", ascending=False)
    )
    st.dataframe(
        kec.rename(columns={
            "Outlet": "Jml Outlet", "Qty": "Qty (pcs)",
            "Bruto": "Omset Bruto", "Net": "Net Sales"
        }),
        use_container_width=True, hide_index=True,
        column_config={
            "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
            "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
            "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
        },
    )

    st.markdown("#### Per Channel")
    ch = (
        fdf.groupby("Channel")
        .agg(Outlet=("No Outlet", "nunique"), Qty=("QTYPCS", "sum"),
             Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"))
        .reset_index().sort_values("Bruto", ascending=False)
    )
    st.dataframe(
        ch.rename(columns={
            "Outlet": "Jml Outlet", "Qty": "Qty (pcs)",
            "Bruto": "Omset Bruto", "Net": "Net Sales"
        }),
        use_container_width=True, hide_index=True,
        column_config={
            "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
            "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
            "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
        },
    )

# ══════ TAB TREN ══════
with tabs[4]:
    st.subheader("Tren Penjualan")
    if "Tanggal Faktur" in fdf.columns and fdf["Tanggal Faktur"].notna().any():
        daily = (
            fdf.groupby(fdf["Tanggal Faktur"].dt.date)
            .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"),
                 Net=("Total", "sum"), Faktur=("Faktur", "nunique"))
            .reset_index()
            .rename(columns={"Tanggal Faktur": "Tanggal"})
            .sort_values("Tanggal")
        )
        st.markdown("#### Tren Harian — Net Sales")
        st.line_chart(daily.set_index("Tanggal")["Net"])
        st.markdown("#### Tren Harian — Quantity")
        st.line_chart(daily.set_index("Tanggal")["Qty"])
        st.markdown("#### Tren Harian — Omset Bruto")
        st.line_chart(daily.set_index("Tanggal")["Bruto"])

        st.dataframe(
            daily.rename(columns={
                "Qty": "Qty (pcs)", "Bruto": "Omset Bruto",
                "Net": "Net Sales", "Faktur": "Jml Faktur"
            }),
            use_container_width=True, hide_index=True,
            column_config={
                "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
                "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
                "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
            },
        )

        if "WEEK" in fdf.columns:
            st.markdown("#### Tren Mingguan")
            weekly = (
                fdf.groupby("WEEK")
                .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"),
                     Net=("Total", "sum"), Faktur=("Faktur", "nunique"))
                .reset_index().sort_values("WEEK")
            )
            weekly["Week"] = weekly["WEEK"].apply(lambda x: f"W{int(x)}")
            st.bar_chart(weekly.set_index("Week")["Net"])
            st.dataframe(
                weekly[["Week", "Qty", "Bruto", "Net", "Faktur"]].rename(columns={
                    "Qty": "Qty (pcs)", "Bruto": "Omset Bruto",
                    "Net": "Net Sales", "Faktur": "Jml Faktur"
                }),
                use_container_width=True, hide_index=True,
                column_config={
                    "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
                    "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
                    "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
                },
            )
    else:
        st.warning("Kolom tanggal tidak tersedia di file ini.")

# ══════ TAB RETURN ══════
with tabs[5]:
    st.subheader("Transaksi Return / Credit Note")
    rets = fdf[fdf["Total"] < 0].copy()
    if len(rets) == 0:
        st.success("Tidak ada transaksi return pada filter saat ini.")
    else:
        st.warning(f"Total return: **{len(rets)}** transaksi | Nilai: **{fmt_rp(rets['Total'].sum())}**")
        ret_show = (
            rets.groupby(["Nama Outlet", "Nama Produk", "Salesman"])
            .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"))
            .reset_index().sort_values("Net")
        )
        st.dataframe(
            ret_show.rename(columns={
                "Nama Outlet": "Outlet", "Nama Produk": "Produk",
                "Qty": "Qty", "Bruto": "Omset Bruto", "Net": "Net Sales"
            }),
            use_container_width=True, hide_index=True,
            column_config={
                "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
                "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
            },
        )

# ══════ TAB EXPORT ══════
with tabs[6]:
    st.subheader("Download Ringkasan")
    st.markdown("Unduh data hasil filter dalam format Excel.")

    def to_excel_bytes(frames: dict) -> bytes:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for name, frame in frames.items():
                frame.to_excel(writer, sheet_name=name[:31], index=False)
        return buf.getvalue()

    prod_exp = (
        fdf.groupby(["Pcode", "Nama Produk", "SUBBRANDNAME"])
        .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"))
        .reset_index().sort_values("Bruto", ascending=False)
    )
    out_exp = (
        fdf.groupby(["No Outlet", "Nama Outlet", "Kabupaten"])
        .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"))
        .reset_index().sort_values("Bruto", ascending=False)
    )
    sm_exp = (
        fdf.groupby("Salesman")
        .agg(Outlet=("No Outlet", "nunique"), Qty=("QTYPCS", "sum"),
             Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"))
        .reset_index().sort_values("Bruto", ascending=False)
    )

    excel_data = to_excel_bytes({
        "Produk": prod_exp,
        "Outlet": out_exp,
        "Salesman": sm_exp,
        "Data Filtered": fdf,
    })
    st.download_button(
        "⬇️ Download Excel Ringkasan",
        data=excel_data,
        file_name=f"LBP_Monitor_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown("---")
st.caption("LBP Sales Monitor • Upload XLSX → Monitoring lengkap • Mobile-friendly")
