"""
LBP Sales Monitor - Mobile-friendly Streamlit App
Upload XLSX once → data tersimpan di session + bisa disimpan sebagai cache
"""
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import hashlib

st.set_page_config(
    page_title="LBP Sales Monitor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
    div[data-testid="stMetric"] {
        background: #f0f7ff; border: 1px solid #d0e3f5;
        border-radius: 10px; padding: 10px 14px;
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



def col_uniques(df, col):
    """Safe unique values even if duplicate column names exist."""
    if col not in df.columns:
        return []
    s = df[col]
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:, 0]
    return sorted(s.dropna().astype(str).unique().tolist())



def _dedupe_columns(df: pd.DataFrame) -> pd.DataFrame:
    """If rename created duplicate column names, keep first occurrence."""
    if not df.columns.duplicated().any():
        return df
    return df.loc[:, ~df.columns.duplicated()].copy()


def _ensure_series(df, col):
    s = df[col]
    if isinstance(s, pd.DataFrame):
        return s.iloc[:, 0]
    return s

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
    df = pd.read_excel(file, header=1)
    df.columns = df.columns.str.strip()

    if "No Outlet" not in df.columns and "Nama Outlet" not in df.columns:
        try:
            file.seek(0)
        except Exception:
            pass
        df = pd.read_excel(file, header=0)
        df.columns = df.columns.str.strip()

    col_map = {}
    for c in df.columns:
        cl = c.lower().replace(" ", "")
        if "nooutlet" in cl or c == "No Outlet":
            col_map[c] = "No Outlet"
        elif "namaoutlet" in cl or c == "Nama Outlet":
            col_map[c] = "Nama Outlet"
        elif "tanggalfaktur" in cl or ("tanggal" in cl and "faktur" in cl):
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
        elif "amount" == cl:
            col_map[c] = "AMOUNT"

    df = df.rename(columns=col_map)
    df = _dedupe_columns(df)

    for col in ["QTYPCS", "Harga Bruto", "Total", "DISC", "PROAMOUNT", "AMOUNT"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    if "Tanggal Faktur" in df.columns:
        df["Tanggal Faktur"] = pd.to_datetime(
            df["Tanggal Faktur"], format="%d/%m/%Y", errors="coerce"
        )
        if df["Tanggal Faktur"].isna().mean() > 0.5:
            df["Tanggal Faktur"] = pd.to_datetime(df["Tanggal Faktur"], errors="coerce")
        df["Tanggal"] = df["Tanggal Faktur"].dt.date
        df["Hari"] = df["Tanggal Faktur"].dt.day_name()

    for col in ["Nama Outlet", "Nama Produk", "SUBBRANDNAME", "Salesman",
                "Kabupaten", "Kecamatan", "Channel", "Faktur"]:
        if col not in df.columns:
            df[col] = "-"
        else:
            df[col] = df[col].fillna("-").astype(str)

    if "WEEK" not in df.columns and "Tanggal Faktur" in df.columns:
        df["WEEK"] = df["Tanggal Faktur"].dt.isocalendar().week.astype("Int64")

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
    if "DISC" not in df.columns:
        df["DISC"] = 0

    return df



def load_csv_pipe(file) -> pd.DataFrame:
    """Load pipe-separated CSV and normalize to LBP schema."""
    # try utf-8 then latin-1
    try:
        df = pd.read_csv(file, sep="|", dtype=str, keep_default_na=False)
    except UnicodeDecodeError:
        try:
            file.seek(0)
        except Exception:
            pass
        df = pd.read_csv(file, sep="|", dtype=str, keep_default_na=False, encoding="latin-1")

    df.columns = df.columns.str.strip()
    # drop fully empty cols
    df = df.dropna(axis=1, how="all")

    col_map = {}
    for c in df.columns:
        cl = c.lower().replace(" ", "").replace("_", "")
        if cl in ("nooutlet", "outletid", "kodeoutlet") or c == "No Outlet":
            col_map[c] = "No Outlet"
        elif cl in ("namaoutlet", "outlet") or c == "Nama Outlet":
            col_map[c] = "Nama Outlet"
        elif "tanggalfaktur" in cl or (cl.startswith("tanggal") and "faktur" in cl):
            col_map[c] = "Tanggal Faktur"
        elif cl in ("qtypcs", "qty", "quantity", "jumlah"):
            col_map[c] = "QTYPCS"
        elif "hargabruto" in cl or cl == "bruto":
            col_map[c] = "Harga Bruto"
        elif cl in ("total", "net", "netsales", "grandtotal"):
            col_map[c] = "Total"
        elif cl in ("disc", "discount", "diskon"):
            col_map[c] = "DISC"
        elif "proamount" in cl or cl == "promo":
            col_map[c] = "PROAMOUNT"
        elif "namaproduk" in cl or cl in ("produk", "product"):
            col_map[c] = "Nama Produk"
        elif "subbrandname" in cl or cl == "subbrand":
            col_map[c] = "SUBBRANDNAME"
        elif cl in ("pcode", "kodeproduk", "sku"):
            col_map[c] = "Pcode"
        elif "salesman" in cl:
            col_map[c] = "Salesman"
        elif "kabupaten" in cl:
            col_map[c] = "Kabupaten"
        elif "kecamatan" in cl:
            col_map[c] = "Kecamatan"
        elif "channel" in cl:
            col_map[c] = "Channel"
        elif cl in ("week", "minggu"):
            col_map[c] = "WEEK"
        elif "faktur" in cl and "tanggal" not in cl:
            col_map[c] = "Faktur"
        elif "transtype" in cl:
            col_map[c] = "TRANSTYPE"
        elif cl == "amount":
            col_map[c] = "AMOUNT"
        elif "grupoutlet" in cl:
            col_map[c] = "Grup Outlet"
        elif "tipeoutlet" in cl:
            col_map[c] = "Tipe Outlet"
        elif cl == "kemasan":
            col_map[c] = "Kemasan"
        elif cl == "periode":
            col_map[c] = "Periode"
        elif "salesforce" in cl:
            col_map[c] = "Salesforce"
        elif "salesteam" in cl:
            col_map[c] = "Sales Team"
        elif cl == "subbrand":
            col_map[c] = "SUBBRAND"
        elif "kelurahan" in cl:
            col_map[c] = "Kelurahan"

    df = df.rename(columns=col_map)
    df = _dedupe_columns(df)

    for col in ["QTYPCS", "Harga Bruto", "Total", "DISC", "PROAMOUNT", "AMOUNT"]:
        if col in df.columns:
            s = df[col].astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False)
            df[col] = pd.to_numeric(s, errors="coerce").fillna(0)

    if "Tanggal Faktur" in df.columns:
        df["Tanggal Faktur"] = pd.to_datetime(
            df["Tanggal Faktur"], format="%d/%m/%Y", errors="coerce"
        )
        if df["Tanggal Faktur"].isna().mean() > 0.5:
            df["Tanggal Faktur"] = pd.to_datetime(df["Tanggal Faktur"], errors="coerce")
        df["Tanggal"] = df["Tanggal Faktur"].dt.date
        df["Hari"] = df["Tanggal Faktur"].dt.day_name()

    for col in ["Nama Outlet", "Nama Produk", "SUBBRANDNAME", "Salesman",
                "Kabupaten", "Kecamatan", "Channel", "Faktur"]:
        if col not in df.columns:
            df[col] = "-"
        else:
            df[col] = df[col].fillna("-").astype(str)

    if "WEEK" not in df.columns and "Tanggal Faktur" in df.columns:
        df["WEEK"] = df["Tanggal Faktur"].dt.isocalendar().week.astype("Int64")
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
    if "DISC" not in df.columns:
        df["DISC"] = 0

    return df


def load_cache_file(file) -> pd.DataFrame:
    """Load previously saved cache (parquet / csv / csv.gz)."""
    import gzip
    name = getattr(file, "name", "").lower()
    if name.endswith(".parquet"):
        return pd.read_parquet(file)
    if name.endswith(".csv.gz") or name.endswith(".gz"):
        with gzip.open(file, "rt") as f:
            return pd.read_csv(f)
    if name.endswith(".csv"):
        return pd.read_csv(file)
    try:
        return pd.read_parquet(file)
    except Exception:
        try:
            file.seek(0)
        except Exception:
            pass
        try:
            return pd.read_csv(file)
        except Exception:
            try:
                file.seek(0)
            except Exception:
                pass
            with gzip.open(file, "rt") as f:
                return pd.read_csv(f)


def prepare_cache_df(df: pd.DataFrame) -> pd.DataFrame:
    """Slim + type-safe dataframe for cache export."""
    # Keep only useful columns (much smaller than full Excel)
    keep = [
        "No Outlet", "Nama Outlet", "Grup Outlet", "Tipe Outlet",
        "Tanggal Faktur", "Faktur", "TRANSTYPE", "Kode Sales",
        "Pcode", "Nama Produk", "Kemasan", "QTYPCS",
        "AMOUNT", "Harga Bruto", "DISC", "DISC1KH", "PROAMOUNT", "Total",
        "Channel", "Kabupaten", "Kecamatan", "Kelurahan",
        "WEEK", "Periode", "Salesman", "Salesforce", "Sales Team",
        "SUBBRAND", "SUBBRANDNAME",
    ]
    cols = [c for c in keep if c in df.columns]
    # always keep core cols even if name differs
    for c in ["No Outlet", "Nama Outlet", "Nama Produk", "QTYPCS", "Harga Bruto", "Total",
              "Salesman", "Kabupaten", "Channel", "SUBBRANDNAME", "Pcode", "DISC"]:
        if c in df.columns and c not in cols:
            cols.append(c)
    out = df[cols].copy()

    # Force object/mixed columns to string so parquet never fails
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].astype(str)
        # categorical-like mixed int/str columns
        elif str(out[c].dtype).startswith("Int") or out[c].dtype == "int64":
            pass
        else:
            try:
                # if mostly numeric but has str leftovers, stringify
                if out[c].map(lambda x: isinstance(x, str)).any():
                    out[c] = out[c].astype(str)
            except Exception:
                out[c] = out[c].astype(str)
    return out


def df_to_parquet_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    clean = prepare_cache_df(df)
    clean.to_parquet(buf, index=False, compression="zstd")
    return buf.getvalue()


def df_to_csv_gz_bytes(df: pd.DataFrame) -> bytes:
    import gzip
    clean = prepare_cache_df(df)
    raw = clean.to_csv(index=False).encode("utf-8")
    return gzip.compress(raw)


def kpi_row(df):
    def _sum(col):
        if col not in df.columns:
            return 0
        s = _ensure_series(df, col)
        return pd.to_numeric(s, errors="coerce").fillna(0).sum()

    def _nunique(col):
        if col not in df.columns:
            return 0
        s = _ensure_series(df, col)
        return s.nunique()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Net Sales", fmt_rp(_sum("Total")))
    c2.metric("Omset Bruto", fmt_rp(_sum("Harga Bruto")))
    c3.metric("Quantity", fmt_num(_sum("QTYPCS")) + " pcs")
    c4.metric("Outlet", fmt_num(_nunique("No Outlet")))
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Faktur", fmt_num(_nunique("Faktur")))
    c6.metric("Produk SKU", fmt_num(_nunique("Pcode")))
    c7.metric("Discount", fmt_rp(_sum("DISC")))
    total_s = _ensure_series(df, "Total") if "Total" in df.columns else pd.Series([0])
    total_s = pd.to_numeric(total_s, errors="coerce").fillna(0)
    ret = total_s[total_s < 0].sum()
    c8.metric("Return", fmt_rp(ret))


# ── Init session state ──
if "df" not in st.session_state:
    st.session_state.df = None
if "source_name" not in st.session_state:
    st.session_state.source_name = None
if "loaded_at" not in st.session_state:
    st.session_state.loaded_at = None


# ── Sidebar: Data Source ──
st.sidebar.title("📊 LBP Monitor")
st.sidebar.caption("Data tersimpan selama tab browser terbuka")

st.sidebar.markdown("### 📂 Sumber Data")
source_mode = st.sidebar.radio(
    "Pilih cara load data",
    ["Upload Excel (.xlsx)", "Upload CSV (| )", "Load Cache (.parquet)", "Load dari URL"],
    label_visibility="collapsed",
)

new_df = None
new_name = None

if source_mode == "Upload Excel (.xlsx)":
    uploaded = st.sidebar.file_uploader(
        "Pilih file Excel", type=["xlsx", "xls"], key="xlsx_up"
    )
    if uploaded is not None:
        file_id = f"{uploaded.name}_{uploaded.size}"
        if st.session_state.get("last_file_id") != file_id:
            with st.spinner("Memproses Excel..."):
                try:
                    new_df = load_lbp(uploaded)
                    new_name = uploaded.name
                    st.session_state.last_file_id = file_id
                except Exception as e:
                    st.sidebar.error(f"Gagal baca file: {e}")

elif source_mode == "Upload CSV (| )":
    st.sidebar.caption("CSV dengan pemisah **|** (pipe). Header di baris pertama.")
    uploaded = st.sidebar.file_uploader(
        "Pilih file CSV", type=["csv", "txt"], key="csv_up"
    )
    if uploaded is not None:
        file_id = f"csv_{uploaded.name}_{uploaded.size}"
        if st.session_state.get("last_file_id") != file_id:
            with st.spinner("Memproses CSV..."):
                try:
                    new_df = load_csv_pipe(uploaded)
                    new_name = uploaded.name
                    st.session_state.last_file_id = file_id
                except Exception as e:
                    st.sidebar.error(f"Gagal baca CSV: {e}")

elif source_mode == "Load Cache (.parquet)":
    st.sidebar.info(
        "Gunakan file cache yang sebelumnya di-download dari tab **Export**. "
        "Lebih cepat & kecil dibanding Excel."
    )
    cache_file = st.sidebar.file_uploader(
        "Pilih file cache", type=["parquet", "csv", "gz"], key="cache_up"
    )
    if cache_file is not None:
        file_id = f"cache_{cache_file.name}_{cache_file.size}"
        if st.session_state.get("last_file_id") != file_id:
            with st.spinner("Memuat cache..."):
                try:
                    new_df = load_cache_file(cache_file)
                    new_name = cache_file.name
                    st.session_state.last_file_id = file_id
                except Exception as e:
                    st.sidebar.error(f"Gagal load cache: {e}")

else:  # URL
    st.sidebar.caption(
        "Tempel link langsung ke file .xlsx / .parquet "
        "(Google Drive: gunakan link download langsung)"
    )
    url = st.sidebar.text_input("URL file", placeholder="https://...")
    if st.sidebar.button("Load dari URL", use_container_width=True) and url:
        with st.spinner("Mengunduh & memproses..."):
            try:
                if "drive.google.com" in url and "/file/d/" in url:
                    # convert sharing link to direct download
                    fid = url.split("/file/d/")[1].split("/")[0]
                    url = f"https://drive.google.com/uc?export=download&id={fid}"
                if url.lower().endswith(".parquet"):
                    new_df = pd.read_parquet(url)
                else:
                    new_df = load_lbp(url)
                new_name = url.split("/")[-1][:40]
                st.session_state.last_file_id = f"url_{url}"
            except Exception as e:
                st.sidebar.error(f"Gagal load URL: {e}")

# Apply new data to session
if new_df is not None:
    st.session_state.df = new_df
    st.session_state.source_name = new_name
    st.session_state.loaded_at = datetime.now().strftime("%d/%m/%Y %H:%M")

# Clear data
if st.session_state.df is not None:
    st.sidebar.success(
        f"✅ {len(st.session_state.df):,} baris\n\n"
        f"📄 {st.session_state.source_name or '-'}\n\n"
        f"🕒 {st.session_state.loaded_at or '-'}"
    )
    if st.sidebar.button("🗑️ Hapus data (load ulang)", use_container_width=True):
        st.session_state.df = None
        st.session_state.source_name = None
        st.session_state.loaded_at = None
        st.session_state.last_file_id = None
        st.rerun()

df = st.session_state.df

if df is None:
    st.title("📊 LBP Sales Monitor")
    st.info(
        "👈 Pilih sumber data di sidebar untuk mulai.\n\n"
        "**3 cara load data:**\n"
        "1. **Upload Excel** — file XLSX mentah (pertama kali)\n"
        "2. **Load Cache** — file `.parquet` yang pernah di-download (jauh lebih cepat)\n"
        "3. **Load dari URL** — link langsung ke file online\n\n"
        "💡 **Tips agar tidak upload ulang terus:**\n"
        "- Selama tab browser **tidak ditutup**, data tetap ada (pindah menu pun aman)\n"
        "- Di tab **Export**, download **Cache Parquet** → lain kali load file itu (lebih kecil & cepat)\n"
        "- Atau simpan file Excel di Google Drive → load via URL"
    )
    st.markdown("---")
    st.markdown(
        "**Format didukung:** XLSX dengan kolom seperti "
        "`No Outlet`, `Nama Outlet`, `Nama Produk`, `QTYPCS`, `Harga Bruto`, `Total`, "
        "`Salesman`, `Kabupaten`, `Channel`, `Tanggal Faktur`"
    )
    st.stop()


# ── Filters ──
st.sidebar.markdown("### 🔎 Filter")
all_kab = col_uniques(df, "Kabupaten")
sel_kab = st.sidebar.multiselect("Kabupaten", all_kab, default=all_kab)

all_sm = col_uniques(df, "Salesman")
sel_sm = st.sidebar.multiselect("Salesman", all_sm, default=[])

all_sb = col_uniques(df, "SUBBRANDNAME")
sel_sb = st.sidebar.multiselect("Subbrand", all_sb, default=[])

all_ch = col_uniques(df, "Channel")
sel_ch = st.sidebar.multiselect("Channel", all_ch, default=[])

fdf = df.copy()
if sel_kab and "Kabupaten" in fdf.columns:
    fdf = fdf[_ensure_series(fdf, "Kabupaten").isin(sel_kab)]
if sel_sm and "Salesman" in fdf.columns:
    fdf = fdf[_ensure_series(fdf, "Salesman").isin(sel_sm)]
if sel_sb and "SUBBRANDNAME" in fdf.columns:
    fdf = fdf[_ensure_series(fdf, "SUBBRANDNAME").isin(sel_sb)]
if sel_ch and "Channel" in fdf.columns:
    fdf = fdf[_ensure_series(fdf, "Channel").isin(sel_ch)]

if "Tanggal Faktur" in fdf.columns and fdf["Tanggal Faktur"].notna().any():
    min_d = fdf["Tanggal Faktur"].min().date()
    max_d = fdf["Tanggal Faktur"].max().date()
    date_range = st.sidebar.date_input(
        "Rentang Tanggal", [min_d, max_d], min_value=min_d, max_value=max_d
    )
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        fdf = fdf[
            (fdf["Tanggal Faktur"].dt.date >= date_range[0])
            & (fdf["Tanggal Faktur"].dt.date <= date_range[1])
        ]

st.sidebar.markdown(f"**Data aktif:** {len(fdf):,} baris")


# ── Main ──
st.title("📊 LBP Sales Monitor")
if "Tanggal Faktur" in fdf.columns and fdf["Tanggal Faktur"].notna().any():
    st.caption(
        f"📄 {st.session_state.source_name} | "
        f"Periode: {fdf['Tanggal Faktur'].min().strftime('%d/%m/%Y')} – "
        f"{fdf['Tanggal Faktur'].max().strftime('%d/%m/%Y')} | "
        f"Loaded: {st.session_state.loaded_at}"
    )
else:
    st.caption(f"📄 {st.session_state.source_name} | Loaded: {st.session_state.loaded_at}")

kpi_row(fdf)
st.markdown("---")

tabs = st.tabs([
    "📦 Produk", "🏪 Outlet", "👤 Salesman",
    "📍 Wilayah", "📅 Tren", "🔄 Return", "⬇️ Export"
])

# ══════ PRODUK ══════
with tabs[0]:
    st.subheader("Monitoring per Produk")
    prod = (
        fdf.groupby(["Pcode", "Nama Produk", "SUBBRANDNAME"], dropna=False)
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

    q = st.text_input("🔍 Cari produk", "", key="prod_search")
    show = prod if not q else prod[prod["Nama Produk"].str.contains(q, case=False, na=False)]

    st.dataframe(
        show.rename(columns={
            "Nama Produk": "Produk", "SUBBRANDNAME": "Subbrand",
            "Qty": "Qty (pcs)", "Bruto": "Omset Bruto", "Net": "Net Sales",
            "Disc": "Discount", "Outlet": "Jml Outlet", "Faktur": "Jml Faktur"
        }),
        use_container_width=True, hide_index=True,
        column_config={
            "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
            "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
            "Discount": st.column_config.NumberColumn(format="Rp %d"),
            "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
            "% Bruto": st.column_config.NumberColumn(format="%.1f%%"),
            "% Qty": st.column_config.NumberColumn(format="%.1f%%"),
        },
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Top 10 — Omset Bruto")
        st.bar_chart(prod.head(10).set_index("Nama Produk")["Bruto"])
    with c2:
        st.markdown("#### Top 10 — Quantity")
        st.bar_chart(prod.head(10).set_index("Nama Produk")["Qty"])

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

# ══════ OUTLET ══════
with tabs[1]:
    st.subheader("Monitoring per Outlet")
    out = (
        fdf.groupby(["No Outlet", "Nama Outlet", "Kabupaten", "Channel"], dropna=False)
        .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"),
             Net=("Total", "sum"), Faktur=("Faktur", "nunique"),
             SKU=("Pcode", "nunique"))
        .reset_index()
        .sort_values("Bruto", ascending=False)
    )

    q2 = st.text_input("🔍 Cari outlet", "", key="out_search")
    show2 = out if not q2 else out[out["Nama Outlet"].str.contains(q2, case=False, na=False)]
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

    st.markdown("#### Detail Produk di Outlet")
    sel_out = st.selectbox("Pilih outlet", out["Nama Outlet"].head(50).tolist(), key="sel_outlet")
    if sel_out:
        od = fdf[fdf["Nama Outlet"] == sel_out]
        od_prod = (
            od.groupby(["Nama Produk", "SUBBRANDNAME"])
            .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"), Net=("Total", "sum"))
            .reset_index().sort_values("Bruto", ascending=False)
        )
        st.dataframe(
            od_prod.rename(columns={
                "Nama Produk": "Produk", "SUBBRANDNAME": "Subbrand",
                "Qty": "Qty (pcs)", "Bruto": "Omset Bruto", "Net": "Net Sales"
            }),
            use_container_width=True, hide_index=True,
            column_config={
                "Omset Bruto": st.column_config.NumberColumn(format="Rp %d"),
                "Net Sales": st.column_config.NumberColumn(format="Rp %d"),
                "Qty (pcs)": st.column_config.NumberColumn(format="%d"),
            },
        )

# ══════ SALESMAN ══════
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
    st.bar_chart(sm.head(15).set_index("Salesman")["Bruto"])

    sel_s = st.selectbox("Detail produk salesman", sm["Salesman"].tolist(), key="sel_sm")
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

# ══════ WILAYAH ══════
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

# ══════ TREN ══════
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
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Quantity")
            st.line_chart(daily.set_index("Tanggal")["Qty"])
        with c2:
            st.markdown("#### Omset Bruto")
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

        if "WEEK" in fdf.columns and fdf["WEEK"].notna().any():
            st.markdown("#### Tren Mingguan")
            weekly = (
                fdf.groupby("WEEK")
                .agg(Qty=("QTYPCS", "sum"), Bruto=("Harga Bruto", "sum"),
                     Net=("Total", "sum"), Faktur=("Faktur", "nunique"))
                .reset_index().sort_values("WEEK")
            )
            weekly["Week"] = weekly["WEEK"].apply(lambda x: f"W{int(x)}")
            st.bar_chart(weekly.set_index("Week")["Net"])
    else:
        st.warning("Kolom tanggal tidak tersedia.")

# ══════ RETURN ══════
with tabs[5]:
    st.subheader("Transaksi Return / Credit Note")
    rets = fdf[fdf["Total"] < 0].copy()
    if len(rets) == 0:
        st.success("Tidak ada transaksi return pada filter saat ini.")
    else:
        st.warning(
            f"Total return: **{len(rets)}** transaksi | Nilai: **{fmt_rp(rets['Total'].sum())}**"
        )
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

# ══════ EXPORT ══════
with tabs[6]:
    st.subheader("Download & Cache")

    st.markdown("### ⚡ Cache Cepat (rekomendasi)")
    st.markdown(
        "Download file **Parquet** ini sekali. Lain kali pilih **Load Cache** di sidebar "
        "— jauh lebih cepat & file lebih kecil dibanding Excel mentah."
    )
    try:
        parquet_bytes = df_to_parquet_bytes(df)
        st.download_button(
            "⬇️ Download Cache Parquet (rekomendasi)",
            data=parquet_bytes,
            file_name=f"LBP_cache_{datetime.now().strftime('%Y%m%d')}.parquet",
            mime="application/octet-stream",
            use_container_width=True,
        )
        st.caption(
            f"Ukuran cache: **{len(parquet_bytes)/1024:.0f} KB** | {len(df):,} baris "
            f"(hanya kolom penting, terkompresi)"
        )
    except Exception as e:
        st.warning(f"Parquet gagal ({e}). Memakai CSV.GZ sebagai alternatif.")
        try:
            gz = df_to_csv_gz_bytes(df)
            st.download_button(
                "⬇️ Download Cache CSV.GZ",
                data=gz,
                file_name=f"LBP_cache_{datetime.now().strftime('%Y%m%d')}.csv.gz",
                mime="application/gzip",
                use_container_width=True,
            )
            st.caption(f"Ukuran: {len(gz)/1024:.0f} KB")
        except Exception as e2:
            csv_buf = BytesIO()
            prepare_cache_df(df).to_csv(csv_buf, index=False)
            st.download_button(
                "⬇️ Download Cache CSV",
                data=csv_buf.getvalue(),
                file_name=f"LBP_cache_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
            )

    st.markdown("---")
    st.markdown("### 📊 Export Ringkasan Excel")

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
        "Produk": prod_exp, "Outlet": out_exp,
        "Salesman": sm_exp, "Data Filtered": fdf,
    })
    st.download_button(
        "⬇️ Download Excel Ringkasan",
        data=excel_data,
        file_name=f"LBP_Monitor_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("---")
st.caption(
    "LBP Sales Monitor • Data tersimpan di session browser • "
    "Download Cache Parquet agar tidak perlu upload Excel berulang"
)
