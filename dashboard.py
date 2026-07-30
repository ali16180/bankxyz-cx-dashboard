# ==================================================================
# BANK XYZ — CUSTOMER EXPERIENCE DASHBOARD (v3)
# ------------------------------------------------------------------
# Jalankan dari root project:   streamlit run dashboard.py
#
# File yang dibutuhkan (relatif terhadap file ini):
#   1. data/Deka_project_dataset_BankXYZ.csv
#   2. metadata/metadata_dashboard.xlsx  (atau .csv)  -> hasil pipeline.ipynb
#
# Fitur:
#   - Navigasi halaman di SIDEBAR (Ringkasan, Brand Image, Branch
#     Facilities, Service Experience, ATM Experience)
#   - Filter tersimpan saat pindah halaman + tombol Reset Filter
#   - Setiap halaman: KPI kepuasan/kepentingan/gap/basis responden,
#     lalu IPA, prioritas berbasis gap, dan detail per sub-kategori
#
# Catatan analitik yang menentukan bentuk dashboard ini:
#   * Skala KEPUASAN (role=Atribut) dan KEPENTINGAN (role=Importance)
#     TIDAK PERNAH dicampur — keduanya pertanyaan berbeda.
#   * Rata-rata section/sub-kategori DIBOBOT jumlah responden (n),
#     karena n antar atribut sangat berbeda (70 s/d 1.730).
#   * Skor berkumpul di 5,4–5,9 dari 6 (efek plafon). Grafik memakai
#     dot plot dengan sumbu yang dizoom ke rentang data — bar dengan
#     baseline terpotong akan melebih-lebihkan perbedaan.
# ==================================================================

import math
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ==================================================================
# KONFIGURASI
# ==================================================================

BASE_DIR = Path(__file__).parent

DATA_PATH = BASE_DIR / "data" / "Deka_project_dataset_BankXYZ.csv"
META_XLSX = BASE_DIR / "metadata" / "metadata_dashboard.xlsx"
META_CSV = BASE_DIR / "metadata" / "metadata_dashboard.csv"

# ---- Palet -------------------------------------------------------
# Pasangan kategorikal XYZ/Kompetitor lulus seluruh gate CVD & kontras
# (validate_palette.js, mode light, surface #FFFFFF).
C_BRAND = "#4CA0E0"          # judul & aksen brand (dicerahkan utk latar gelap)
SERIES_XYZ = "#4A93D1"       # slot kategorikal 1
SERIES_COMP = "#EB6834"      # slot kategorikal 2
# Sekuensial satu hue (heatmap): terang -> gelap
BLUE_SEQ = ["#CDE2FB", "#9EC5F4", "#6DA7EC", "#3987E5", "#256ABF", "#104281"]
# Diverging biru <-> merah dengan titik netral abu (untuk gap)
GAP_UNDER = "#E2544C"        # kepentingan > kepuasan  -> perlu diperbaiki
GAP_OVER = "#4A93D1"         # kepuasan > kepentingan  -> sudah melampaui
# Status (tetap, tidak pernah dipakai sebagai warna seri biasa)
ST_GOOD = "#22C55E"
ST_WARN = "#FBBF24"
ST_SERIOUS = "#F0855A"
ST_CRIT = "#E2544C"
# Ink — mode gelap: teks terang di atas permukaan gelap
INK = "#EAF1FB"
INK_2 = "#AAB8CC"
MUTED = "#7C8AA0"
GRID = "#25344A"
AXIS = "#3A4E6B"
SURFACE = "#141F30"
C_PALE = "#1E2E44"
C_BG = "#0B1420"

MIN_BASE = 100               # di bawah ini basis responden dianggap tipis

PAGE_RINGKASAN = "Ringkasan"
PAGE_BI = "Brand Image"
PAGE_BF = "Branch Facilities"
PAGE_SE = "Service Experience"
PAGE_ATM = "ATM Experience"
PAGES = {
    PAGE_RINGKASAN: "📊",
    PAGE_BI: "⭐",
    PAGE_BF: "🏢",
    PAGE_SE: "🤝",
    PAGE_ATM: "🏧",
}

TP_SERVICE = ["Customer Service", "Teller", "Security",
              "Customer Advisor", "Service Electronics"]
TP_LABEL = {"Security": "Sekuriti",
            "Service Electronics": "Sarana Elektronik"}

ORDER_USIA = [
    "17 -19 tahun", "20 - 25 tahun", "26 - 30 tahun", "31 - 35 tahun",
    "36 - 40 tahun", "41 - 45 tahun", "46 - 50 tahun",
    "50 tahun dan ke atas",
]
ORDER_LAMA = [
    "1 bulan s/d 3 bulan", "3 bulan s/d 11 bulan",
    "1 tahun s/d 2 tahun 11 bulan", "3 tahun s/d 4 tahun 11 bulan",
    "5 tahun atau lebih",
]
ORDER_BY_VAR = {"S2_2": ORDER_USIA, "S4": ORDER_LAMA}

st.set_page_config(
    page_title="Bank XYZ Customer Experience Dashboard",
    page_icon="🏦",
    layout="wide",
)

# ==================================================================
# TEMA / CSS
# ==================================================================

CSS = """
<style>
.stApp { background-color: __BG__; }
.block-container { padding-top: 1.4rem; }

h1, h2, h3 { color: __DARK__ !important; font-weight: 800; }

/* ---------- SIDEBAR ---------- */
section[data-testid="stSidebar"] {
    background-color: __SURF__;
    border-right: 1px solid __BORDER__;
}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] small {
    color: __TEXT__ !important;
}
section[data-testid="stSidebar"] label { font-weight: 700 !important; }

/* kotak selectbox / multiselect: PUTIH dengan teks gelap */
section[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background-color: __SURF__ !important;
    border: 1px solid __BORDER__ !important;
    border-radius: 8px !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] span,
section[data-testid="stSidebar"] [data-baseweb="select"] div {
    color: __TEXT__ !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] svg {
    fill: __DARK__ !important;
}
/* tag pada multiselect */
section[data-testid="stSidebar"] [data-baseweb="tag"] {
    background-color: __DARK__ !important;
}
section[data-testid="stSidebar"] [data-baseweb="tag"] span,
section[data-testid="stSidebar"] [data-baseweb="tag"] svg {
    color: #FFFFFF !important;
    fill: #FFFFFF !important;
}

/* menu dropdown (muncul di luar sidebar) */
div[data-baseweb="popover"] ul[role="listbox"] {
    background-color: __SURF__ !important;
    border: 1px solid __BORDER__ !important;
}
div[data-baseweb="popover"] li[role="option"],
div[data-baseweb="popover"] li[role="option"] div,
div[data-baseweb="popover"] li[role="option"] span {
    color: __TEXT__ !important;
    background-color: __SURF__;
}
div[data-baseweb="popover"] li[role="option"]:hover,
div[data-baseweb="popover"] li[aria-selected="true"] {
    background-color: __PALE__ !important;
}

/* navigasi halaman (radio) tampil seperti menu */
section[data-testid="stSidebar"] div[role="radiogroup"] > label {
    display: block;
    background: __SURF__;
    border: 1px solid __BORDER__;
    border-radius: 10px;
    padding: 9px 12px;
    margin-bottom: 6px;
    cursor: pointer;
}
section[data-testid="stSidebar"] div[role="radiogroup"]
    > label:has(input:checked) {
    background: __DARK__;
    border-color: __DARK__;
}
section[data-testid="stSidebar"] div[role="radiogroup"]
    > label:has(input:checked) * {
    color: #FFFFFF !important;
}

/* tombol reset */
section[data-testid="stSidebar"] button {
    background-color: __DARK__ !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 8px !important;
    width: 100%;
}
section[data-testid="stSidebar"] button p { color:#FFFFFF !important; }

/* expander filter tambahan */
section[data-testid="stSidebar"] details {
    background: __SURF__;
    border: 1px solid __BORDER__;
    border-radius: 10px;
}

/* ---------- KARTU METRIK ---------- */
div[data-testid="stMetric"] {
    background-color: __SURF__;
    border: 1px solid __BORDER__;
    border-radius: 14px;
    padding: 14px 16px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.35);
}
div[data-testid="stMetric"] label { color: __DARK__ !important;
    font-weight: 700; }
div[data-testid="stMetricValue"] { color: __TEXT__ !important;
    font-weight: 800; }
div[data-testid="stMetricDelta"] { color: __TEXT2__ !important; }
/* allow delta subtext and label to wrap instead of truncating */
div[data-testid="stMetricDelta"],
div[data-testid="stMetricDelta"] *,
div[data-testid="stMetricLabel"],
div[data-testid="stMetricLabel"] *,
[data-testid="stMetric"] [data-testid="stMetricDelta"],
[data-testid="stMetric"] [data-testid="stMetricDelta"] * {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    overflow-wrap: break-word !important;
    word-wrap: break-word !important;
    max-width: 100% !important;
}

/* ---------- TAB SUB-KATEGORI ---------- */
button[data-baseweb="tab"] { color: __TEXT2__ !important; font-weight: 600; }
button[data-baseweb="tab"][aria-selected="true"] {
    color: __DARK__ !important;
    border-bottom: 3px solid __DARK__;
}

div[data-testid="stExpander"] { background:__SURF__; border-radius:12px; }
hr { border-color: __PALE__; }

/* ---------- EXPANDER — header and body text always dark on white ---------- */
details summary,
details > summary,
details summary p,
details summary span,
details summary svg,
div[data-testid="stExpander"] summary,
div[data-testid="stExpander"] summary p,
div[data-testid="stExpander"] summary span,
div[data-testid="stExpander"] summary svg {
    color: __TEXT__ !important;
    fill: __TEXT__ !important;
}
div[data-testid="stExpander"] p,
div[data-testid="stExpander"] strong,
div[data-testid="stExpander"] em,
div[data-testid="stExpander"] span,
div[data-testid="stExpander"] li,
div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] p,
div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] strong,
div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] em,
div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] span {
    color: __TEXT__ !important;
}

/* ---------- TOGGLE / CHECKBOX WIDGET — always visible ---------- */
div[data-testid="stToggle"] label,
div[data-testid="stToggle"] label p,
div[data-testid="stToggle"] label span,
div[data-testid="stToggle"] > div > label > div:last-child,
div[data-testid="stToggle"] > div > label > div:last-child *,
label[data-baseweb="checkbox"] > div:last-child,
label[data-baseweb="checkbox"] > div:last-child p,
label[data-baseweb="checkbox"] > div:last-child span,
[data-testid="stWidgetLabel"],
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] span,
[data-testid="stWidgetLabel"] * {
    color: __TEXT__ !important;
    opacity: 1 !important;
    visibility: visible !important;
}
</style>
"""
st.markdown(
    CSS.replace("__BG__", C_BG).replace("__DARK__", C_BRAND)
       .replace("__PALE__", C_PALE).replace("__TEXT__", INK)
       .replace("__SURF__", SURFACE).replace("__BORDER__", AXIS)
       .replace("__TEXT2__", INK_2),
    unsafe_allow_html=True,
)

# ==================================================================
# LOAD DATA & METADATA
# ==================================================================

META_COLS = ("variable", "question", "label", "bank", "section", "touchpoint",
             "subgroup", "role", "scale_type", "pair_key")


def to_score(series: pd.Series, max_valid: int = 6) -> pd.Series:
    """'6  SANGAT PUAS' -> 6 ; '99 TIDAK RELEVAN' / kosong -> NaN."""
    s = series.astype(str).str.extract(r"^\s*(\d+)")[0]
    s = pd.to_numeric(s, errors="coerce")
    s = s.where(s <= max_valid)
    return s.astype("Float64")


@st.cache_data(show_spinner="Memuat data…")
def load_all():
    df = pd.read_csv(DATA_PATH, sep=";", header=1,
                     low_memory=False, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]

    if META_XLSX.exists():
        meta = pd.read_excel(META_XLSX, sheet_name="Metadata")
    else:
        meta = pd.read_csv(META_CSV, encoding="utf-8-sig")
    meta.columns = [str(c).strip() for c in meta.columns]
    for c in META_COLS:
        if c not in meta.columns:
            meta[c] = ""
        meta[c] = meta[c].astype(str).str.strip()
    meta["include"] = pd.to_numeric(meta["include"],
                                    errors="coerce").fillna(0).astype(int)
    meta["scale_max"] = pd.to_numeric(meta["scale_max"],
                                      errors="coerce").fillna(6).astype(int)
    # hanya variabel yang benar-benar ada di data
    meta = meta[meta["variable"].isin(df.columns)].reset_index(drop=True)

    # Matriks skor numerik sekali jalan: semua perhitungan halaman hanya
    # menyeleksi baris dari sini, jadi tidak ada parsing string berulang.
    rating = meta[meta["scale_max"] > 0]
    num = pd.DataFrame(
        {r.variable: to_score(df[r.variable], r.scale_max)
         for r in rating.itertuples()},
        index=df.index,
    )
    return df, meta, num


if not DATA_PATH.exists():
    st.error("File **data/Deka_project_dataset_BankXYZ.csv** tidak ditemukan. "
             "Pastikan file data ada di folder `data/` pada root project.")
    st.stop()
if not META_XLSX.exists() and not META_CSV.exists():
    st.error("Metadata dashboard belum ada. Jalankan dulu **pipeline.ipynb** "
             "untuk membuat **metadata/metadata_dashboard.xlsx / .csv**.")
    st.stop()

df, META, NUM = load_all()


def _rows(role, bank=None, include_only=True):
    sub = META[META["role"] == role]
    if include_only:
        sub = sub[sub["include"] == 1]
    if bank:
        sub = sub[sub["bank"] == bank]
    return sub


ATTR = _rows("Atribut", bank="XYZ")            # KEPUASAN saja
IMPORTANCE = _rows("Importance")               # KEPENTINGAN saja
OVERALL = _rows("Overall", bank="XYZ")
LOYALTY_DRIVERS = _rows("Loyalty Driver")
EMOTION = _rows("Emotion", include_only=True)
DIGITAL = _rows("Digitalization")
EXTRA_FILTERS = META[(META["role"].str.startswith("Filter"))
                     & (META["subgroup"] == "Tambahan")]

# pair_key -> baris kepentingan pasangannya
IMP_BY_PAIR = IMPORTANCE.set_index("pair_key") if not IMPORTANCE.empty else None


def _subset(frame, section=None, touchpoint=None) -> pd.DataFrame:
    if section:
        frame = frame[frame["section"] == section]
    if touchpoint:
        frame = frame[frame["touchpoint"] == touchpoint]
    return frame


def attrs_of(section=None, touchpoint=None) -> pd.DataFrame:
    return _subset(ATTR, section, touchpoint)


def overall_of(section=None, touchpoint=None) -> pd.DataFrame:
    return _subset(OVERALL, section, touchpoint)


# ==================================================================
# SIDEBAR — NAVIGASI + FILTER (persisten antar halaman)
# ==================================================================

MAIN_FILTERS = [
    ("f_prov", "PROV", "Provinsi"),
    ("f_kab", "KABKOTA", "Kabupaten/Kota"),
    ("f_cab", "CABANG", "Cabang"),
    ("f_usia", "S2_2", "Kelompok Usia"),
    ("f_lama", "S4", "Lama Menjadi Nasabah"),
]


def ordered_options(values, var):
    vals = pd.Series(values).dropna().unique().tolist()
    order = ORDER_BY_VAR.get(var)
    if order:
        return [v for v in order if v in vals] + sorted(
            v for v in vals if v not in order)
    return sorted(vals)


def reset_filters():
    for key, _, _ in MAIN_FILTERS:
        st.session_state[key] = "Semua"
    for var in EXTRA_FILTERS["variable"]:
        st.session_state[f"fx_{var}"] = []


def select_persist(key, label, options):
    """Selectbox ber-key: nilai tersimpan di session_state sehingga
    tidak ke-reset saat pindah halaman; nilai tak valid dikembalikan
    ke 'Semua' (mis. saat pilihan kab/kota berubah karena provinsi)."""
    opts = ["Semua"] + options
    if key not in st.session_state or st.session_state[key] not in opts:
        st.session_state[key] = "Semua"
    return st.sidebar.selectbox(label, opts, key=key)


with st.sidebar:
    st.title("🏦 Bank XYZ")
    page = st.radio("Halaman", list(PAGES),
                    format_func=lambda p: f"{PAGES[p]}  {p}",
                    key="nav_page", label_visibility="collapsed")
    st.markdown("---")
    st.subheader("📌 Filter Data")

# Filter utama (berjenjang: Provinsi -> Kab/Kota -> Cabang)
prov = select_persist("f_prov", "Provinsi",
                      ordered_options(df["PROV"], "PROV"))
_t = df if prov == "Semua" else df[df["PROV"] == prov]
kab = select_persist("f_kab", "Kabupaten/Kota",
                     ordered_options(_t["KABKOTA"], "KABKOTA"))
if kab != "Semua":
    _t = _t[_t["KABKOTA"] == kab]
cab = select_persist("f_cab", "Cabang",
                     ordered_options(_t["CABANG"], "CABANG"))
usia = select_persist("f_usia", "Kelompok Usia",
                      ordered_options(df["S2_2"], "S2_2"))
lama = select_persist("f_lama", "Lama Menjadi Nasabah",
                      ordered_options(df["S4"], "S4"))

# Filter tambahan (multiselect, kosong = semua)
with st.sidebar.expander("⚙️ Filter Tambahan"):
    for _, r in EXTRA_FILTERS.iterrows():
        var = r["variable"]
        st.multiselect(r["label"], ordered_options(df[var], var),
                       key=f"fx_{var}", placeholder="Semua")

st.sidebar.button("🔄 Reset Semua Filter", on_click=reset_filters)

# Terapkan filter
fdf = df
for key, col, _ in MAIN_FILTERS:
    val = st.session_state.get(key, "Semua")
    if val != "Semua":
        fdf = fdf[fdf[col] == val]
for _, r in EXTRA_FILTERS.iterrows():
    sel = st.session_state.get(f"fx_{r['variable']}", [])
    if sel:
        fdf = fdf[fdf[r["variable"]].isin(sel)]

st.sidebar.markdown("---")
st.sidebar.metric("Jumlah Responden Terfilter", f"{len(fdf):,}")
if len(fdf) == 0:
    st.warning("Tidak ada responden yang sesuai dengan kombinasi filter ini. "
               "Silakan longgarkan filter atau tekan Reset.")
    st.stop()
if len(fdf) < 30:
    st.sidebar.caption("⚠️ Responden < 30 — hasil agregat kurang "
                       "representatif.")

# Matriks skor untuk irisan terfilter — dipakai semua perhitungan halaman.
FNUM = NUM.loc[fdf.index]

# ==================================================================
# PERHITUNGAN
# ==================================================================


def col_stats(var, scale=6):
    """(mean, n, %top-box) untuk satu variabel pada irisan terfilter."""
    if var not in FNUM.columns:
        return None, 0, None
    s = FNUM[var].dropna()
    if len(s) == 0:
        return None, 0, None
    return float(s.mean()), len(s), float((s == scale).mean() * 100)


def wmean(values, weights):
    """Rata-rata dibobot n. Bobot menghindari atribut n=70 menyetarai n=1.730."""
    pairs = [(v, w) for v, w in zip(values, weights)
             if v is not None and pd.notna(v) and w]
    if not pairs:
        return None
    total = sum(w for _, w in pairs)
    return sum(v * w for v, w in pairs) / total if total else None


def nps(data_idx, var="G1A"):
    if var not in NUM.columns:
        return None, None, None, None
    s = NUM.loc[data_idx, var].dropna()
    if len(s) == 0:
        return None, None, None, None
    p = float((s >= 9).mean() * 100)
    d = float((s <= 6).mean() * 100)
    return p - d, p, 100 - p - d, d


def attribute_frame(sat_rows: pd.DataFrame) -> pd.DataFrame:
    """Satu baris per atribut KEPUASAN, plus KEPENTINGAN pasangannya dan gap.

    Kolom: variable, Atribut, Subkategori, Skor, n, TopBox, Kepentingan,
           n_imp, Gap  (Gap = Kepentingan - Skor; positif = belum terpenuhi)
    """
    rows = []
    for r in sat_rows.itertuples():
        skor, n, tb = col_stats(r.variable)
        if skor is None:
            continue
        imp = imp_n = gap = None
        if IMP_BY_PAIR is not None and r.pair_key in IMP_BY_PAIR.index:
            hit = IMP_BY_PAIR.loc[r.pair_key]
            if isinstance(hit, pd.DataFrame):      # tidak seharusnya terjadi
                hit = hit.iloc[0]
            imp, imp_n, _ = col_stats(hit["variable"])
            if imp is not None:
                gap = imp - skor
        rows.append({"variable": r.variable, "Atribut": r.label,
                     "Subkategori": r.subgroup, "Skor": skor, "n": n,
                     "TopBox": tb, "Kepentingan": imp, "n_imp": imp_n,
                     "Gap": gap})
    return pd.DataFrame(rows)


def frame_summary(am: pd.DataFrame) -> dict:
    """Agregat berbobot n untuk satu kumpulan atribut."""
    if am.empty:
        return {}
    return {
        "skor": wmean(am["Skor"], am["n"]),
        "topbox": wmean(am["TopBox"], am["n"]),
        "kepentingan": wmean(am["Kepentingan"], am["n_imp"].fillna(0)),
        "n_median": int(am["n"].median()),
        "n_min": int(am["n"].min()),
        "n_max": int(am["n"].max()),
        "n_attrs": len(am),
        "n_thin": int((am["n"] < MIN_BASE).sum()),
    }


def subgroup_summary(am: pd.DataFrame) -> pd.DataFrame:
    """Agregat per sub-kategori, dibobot n. Loop eksplisit (bukan
    groupby.apply) supaya tidak bergantung pada perilaku pandas versi tertentu."""
    out = []
    for sub, s in am.groupby("Subkategori", sort=False):
        out.append({"Atribut": sub,
                    "Skor": wmean(s["Skor"], s["n"]),
                    "Kepentingan": wmean(s["Kepentingan"], s["n_imp"].fillna(0)),
                    "TopBox": wmean(s["TopBox"], s["n"]),
                    "n": int(s["n"].median()),
                    "Atribut_n": len(s)})
    g = pd.DataFrame(out)
    g["Gap"] = g["Kepentingan"] - g["Skor"]
    return g


# ==================================================================
# GRAFIK
# ==================================================================

PLOTLY_CFG = {"displaylogo": False,
              "modeBarButtonsToRemove": ["lasso2d", "select2d"]}


def wrap_label(text, width=48, max_lines=4):
    """Pecah label panjang ke beberapa baris — tidak pernah dipotong."""
    words, lines, cur = str(text).split(), [], ""
    for w in words:
        if cur and len(cur) + 1 + len(w) > width:
            lines.append(cur)
            cur = w
        else:
            cur = f"{cur} {w}".strip()
    if cur:
        lines.append(cur)
    if len(lines) > max_lines:                     # gabungkan sisa ke baris akhir
        lines = lines[:max_lines - 1] + [" ".join(lines[max_lines - 1:])]
    return "<br>".join(lines)


def _cat_height(labels, base=140, row=30, line_px=17):
    """Tinggi kanvas dihitung dari jumlah BARIS label, bukan jumlah kategori —
    label 3 baris butuh pita kategori yang lebih tinggi agar tidak bertumpuk."""
    total = sum(max(row, (str(lab).count("<br>") + 1) * line_px + 10)
                for lab in labels)
    return int(max(240, base + total))


def nice_dtick(span, target=6):
    """Jarak tick 'bulat' (1/2/5 × 10^n) untuk rentang tertentu, supaya sumbu
    tidak pernah menampilkan angka seperti -2.22e-17."""
    if not span or span <= 0:
        return None
    raw = span / max(target, 1)
    mag = 10.0 ** math.floor(math.log10(raw))
    for mult in (1, 2, 5):
        if raw <= mult * mag:
            return mult * mag
    return 10 * mag


def _hover_cell(value, kind):
    """Pra-format nilai hover jadi string, supaya template tidak perlu
    menebak jumlah desimal per kolom."""
    if value is None or pd.isna(value):
        return "–"
    if kind == "pct0":
        return f"{float(value):.0f}%"
    if kind == "int":
        return f"{int(value):,}"
    if kind == "signed":
        return f"{float(value):+.2f}"
    if kind == "num2":
        return f"{float(value):.2f}"
    return str(value)


def style_fig(fig, height=420, ygrid=False):
    fig.update_layout(
        height=height, plot_bgcolor=SURFACE, paper_bgcolor=SURFACE,
        font=dict(family="system-ui, -apple-system, 'Segoe UI', sans-serif",
                  color=INK, size=13),
        legend=dict(font=dict(size=12, color=INK_2)),
        title_font=dict(color=C_BRAND, size=16),
        # automargin di kedua sumbu yang mengatur ruang sebenarnya; margin di
        # sini hanya lantai minimum, supaya label tick / judul sumbu tidak
        # pernah menimpa area plot.
        margin=dict(l=16, r=28, t=58, b=52),
        hoverlabel=dict(bgcolor="white", font_size=12,
                        bordercolor=AXIS, align="left"),
    )
    fig.update_xaxes(tickfont=dict(color=INK_2, size=12),
                     title_font=dict(color=INK_2, size=12),
                     showgrid=True, gridcolor=GRID, gridwidth=1,
                     zeroline=False, linecolor=AXIS, automargin=True)
    fig.update_yaxes(tickfont=dict(color=INK_2, size=12),
                     title_font=dict(color=INK_2, size=12),
                     showgrid=ygrid, gridcolor=GRID, gridwidth=1,
                     zeroline=False, linecolor=AXIS, automargin=True)
    return fig


def zoom_range(values, pad_frac=0.12, min_span=0.25):
    """Sumbu yang mengikuti rentang data, dengan lebar minimum agar
    perbedaan 0,01 tidak tampak dramatis."""
    vals = [v for v in values if v is not None and pd.notna(v)]
    if not vals:
        return None
    lo, hi = min(vals), max(vals)
    span = max(hi - lo, min_span)
    pad = span * pad_frac
    return [lo - pad, hi + pad]


def dot_chart(am, value_col, title, x_title, hover_rows, height_extra=0,
              color=SERIES_XYZ, ascending_urgent=True):
    """Dot plot horizontal pada sumbu yang dizoom.

    Dot (bukan bar) karena sumbu tidak mulai dari nol: panjang bar dengan
    baseline terpotong melebih-lebihkan perbedaan; posisi titik tidak.
    Urutan: paling perlu perhatian di ATAS.
    """
    d = am.dropna(subset=[value_col]).copy()
    if d.empty:
        return
    d = d.sort_values(value_col, ascending=not ascending_urgent)
    labels = [wrap_label(x) for x in d["Atribut"]]

    custom = [[_hover_cell(v, kind) for (_, _, kind), v
               in zip(hover_rows, row)]
              for row in zip(*[d[c] for c, _, _ in hover_rows])]
    htmpl = ("<b>%{customdata[0]}</b><br>" + "<br>".join(
        f"{lbl}: %{{customdata[{i}]}}"
        for i, (_, lbl, _) in enumerate(hover_rows) if i > 0)
        + "<extra></extra>")

    fig = go.Figure(go.Scatter(
        x=d[value_col], y=labels, mode="markers",
        marker=dict(size=11, color=color,
                    line=dict(width=2, color=SURFACE)),
        customdata=custom, hovertemplate=htmpl, showlegend=False,
    ))
    # Label langsung hanya pada dua ekstrem — bukan angka di setiap titik.
    extremes = [(d[value_col].idxmin(), "middle right")]
    if len(d) > 1:
        extremes.append((d[value_col].idxmax(), "middle left"))
    for idx, pos in extremes:
        fig.add_trace(go.Scatter(
            x=[d.loc[idx, value_col]], y=[wrap_label(d.loc[idx, "Atribut"])],
            mode="markers+text", marker=dict(size=11, color=color,
                                             line=dict(width=2, color=SURFACE)),
            text=[f"{d.loc[idx, value_col]:.2f}"], textposition=pos,
            textfont=dict(color=INK_2, size=12), hoverinfo="skip",
            showlegend=False))

    avg = wmean(d[value_col], d["n"]) if "n" in d else d[value_col].mean()
    if avg is not None:
        fig.add_vline(x=avg, line_color=MUTED, line_width=1,
                      annotation_text=f"rata-rata {avg:.2f}",
                      annotation_position="top",
                      annotation_font=dict(color=MUTED, size=11))
    rng = zoom_range(list(d[value_col]) + ([avg] if avg else []))
    fig.update_layout(title=title,
                      xaxis=dict(title=x_title, range=rng),
                      yaxis=dict(title=None, automargin=True))
    st.plotly_chart(style_fig(fig, _cat_height(labels) + height_extra,
                              ygrid=True),
                    width="stretch", config=PLOTLY_CFG)


ATTR_HOVER = [("Atribut", "Atribut", "text"), ("Skor", "Kepuasan", "num2"),
              ("Kepentingan", "Kepentingan", "num2"), ("Gap", "Gap", "signed"),
              ("TopBox", "% sangat puas", "pct0"), ("n", "n responden", "int")]


def attribute_dots(am, title):
    dot_chart(am, "Skor", title, "Skor kepuasan (sumbu dizoom ke rentang data)",
              hover_rows=ATTR_HOVER)


def subgroup_dots(am, title):
    """Rata-rata per sub-kategori, DIBOBOT n (bukan rata-rata sederhana)."""
    dot_chart(subgroup_summary(am), "Skor", title,
              "Rata-rata kepuasan sub-kategori (dibobot n)",
              hover_rows=[("Atribut", "Sub-kategori", "text"),
                          ("Skor", "Kepuasan", "num2"),
                          ("Kepentingan", "Kepentingan", "num2"),
                          ("Gap", "Gap", "signed"),
                          ("Atribut_n", "jumlah atribut", "int"),
                          ("n", "n median", "int")])


def gap_chart(am, title, top=12):
    """Diverging bar: kepentingan − kepuasan. Nol adalah baseline sejati,
    jadi di sini bar memang bentuk yang benar."""
    d = am.dropna(subset=["Gap"]).copy()
    if d.empty:
        return
    d = d.reindex(d["Gap"].abs().sort_values(ascending=False).index).head(top)
    d = d.sort_values("Gap")                       # gap terbesar di atas
    labels = [wrap_label(x) for x in d["Atribut"]]
    colors = [GAP_UNDER if v > 0 else GAP_OVER for v in d["Gap"]]
    fig = go.Figure(go.Bar(
        x=d["Gap"], y=labels, orientation="h",
        marker=dict(color=colors, line=dict(width=0)),
        width=0.62,
        customdata=list(zip(d["Atribut"], d["Kepentingan"], d["Skor"], d["n"])),
        hovertemplate=("<b>%{customdata[0]}</b><br>Kepentingan: %{customdata[1]:.2f}"
                       "<br>Kepuasan: %{customdata[2]:.2f}"
                       "<br>Gap: %{x:+.2f}<br>n responden: %{customdata[3]:,}"
                       "<extra></extra>"),
        showlegend=False))
    fig.add_vline(x=0, line_color=MUTED, line_width=1)
    span = float(d["Gap"].max() - d["Gap"].min())
    fig.update_layout(
        title=title,
        xaxis=dict(title="Kepentingan − Kepuasan  (kanan = belum terpenuhi)",
                   tickformat=".2f", dtick=nice_dtick(span)),
        yaxis=dict(title=None, automargin=True), bargap=0.38)
    st.plotly_chart(style_fig(fig, _cat_height(labels), ygrid=False),
                    width="stretch", config=PLOTLY_CFG)
    st.caption(
        f"🔴 Kanan = nasabah menilai atribut lebih **penting** daripada tingkat "
        f"**kepuasan** yang dirasakan → prioritas perbaikan. "
        f"🔵 Kiri = kepuasan sudah melampaui tingkat kepentingan. "
        f"Menampilkan {len(d)} gap terbesar (absolut)."
    )


QUAD_ORDER = ["⚠️ Prioritas Perbaikan", "✅ Pertahankan",
              "Monitor", "Efisiensi Lebih"]
QUAD_COLOR = {"⚠️ Prioritas Perbaikan": ST_CRIT, "✅ Pertahankan": ST_GOOD,
              "Monitor": MUTED, "Efisiensi Lebih": ST_SERIOUS}


def ipa_chart(am, title):
    """Importance-Performance Analysis untuk kumpulan atribut apa pun.

    Warna mengikuti KUADRAN (status), bukan sub-kategori — sub-kategori bisa
    berjumlah 9 dan tidak ada palet kategorikal yang aman sebanyak itu.
    """
    d = am.dropna(subset=["Skor", "Kepentingan"]).copy()
    if len(d) < 2:
        return
    avg_p = wmean(d["Skor"], d["n"])
    avg_i = wmean(d["Kepentingan"], d["n_imp"].fillna(0))
    if avg_p is None or avg_i is None:
        return

    def quad(r):
        hi_imp, hi_perf = r["Kepentingan"] >= avg_i, r["Skor"] >= avg_p
        if hi_imp and not hi_perf: return QUAD_ORDER[0]
        if hi_imp and hi_perf: return QUAD_ORDER[1]
        return QUAD_ORDER[2] if not hi_perf else QUAD_ORDER[3]

    d["Kuadran"] = d.apply(quad, axis=1)
    fig = go.Figure()
    for q in QUAD_ORDER:
        s = d[d["Kuadran"] == q]
        if s.empty:
            continue
        fig.add_trace(go.Scatter(
            x=s["Skor"], y=s["Kepentingan"], mode="markers", name=q,
            marker=dict(size=11, color=QUAD_COLOR[q],
                        line=dict(width=2, color=SURFACE)),
            customdata=list(zip(s["Atribut"], s["Subkategori"], s["Gap"], s["n"])),
            hovertemplate=("<b>%{customdata[0]}</b><br>%{customdata[1]}"
                           "<br>Kepuasan: %{x:.2f}<br>Kepentingan: %{y:.2f}"
                           "<br>Gap: %{customdata[2]:+.2f}"
                           "<br>n responden: %{customdata[3]:,}<extra></extra>")))
    fig.add_vline(x=avg_p, line_color=MUTED, line_width=1)
    fig.add_hline(y=avg_i, line_color=MUTED, line_width=1)
    fig.update_layout(
        title=title,
        xaxis=dict(title="Kepuasan / Performance (skor 1–6)",
                   range=zoom_range(list(d["Skor"]) + [avg_p])),
        yaxis=dict(title="Kepentingan / Importance (skor 1–6)",
                   range=zoom_range(list(d["Kepentingan"]) + [avg_i])),
        legend=dict(orientation="h", y=-0.18, title=None))
    st.plotly_chart(style_fig(fig, 540, ygrid=True),
                    width="stretch", config=PLOTLY_CFG)
    st.caption(
        f"Setiap titik = satu atribut. Garis pemisah = rata-rata berbobot "
        f"kepuasan ({avg_p:.2f}) dan kepentingan ({avg_i:.2f}). "
        "**Kiri atas** = penting tapi kurang memuaskan → prioritas perbaikan. "
        "**Kanan atas** = penting dan sudah memuaskan → pertahankan. "
        "**Kanan bawah** = memuaskan tapi kurang penting → efisiensi lebih. "
        "**Kiri bawah** = kurang penting dan kurang memuaskan → monitor."
    )


def dumbbell(items, title, x_title, note=None, height=None,
             name_a="Bank XYZ", name_b="Kompetitor"):
    """Perbandingan dua entitas per item. items = list of dict:
    label / a / b / n_a / n_b."""
    d = pd.DataFrame(items).dropna(subset=["a", "b"])
    if d.empty:
        return
    d = d.sort_values("a", ascending=False)
    labels = [wrap_label(x, 44) for x in d["label"]]
    fig = go.Figure()
    for xa, xb, y in zip(d["a"], d["b"], labels):       # penghubung dulu
        fig.add_trace(go.Scatter(x=[xa, xb], y=[y, y], mode="lines",
                                 line=dict(color=GRID, width=2),
                                 hoverinfo="skip", showlegend=False))
    for col, name, color, ncol in (("a", name_a, SERIES_XYZ, "n_a"),
                                   ("b", name_b, SERIES_COMP, "n_b")):
        fig.add_trace(go.Scatter(
            x=d[col], y=labels, mode="markers", name=name,
            marker=dict(size=11, color=color, line=dict(width=2, color=SURFACE)),
            customdata=list(zip(d["label"], d[ncol])),
            hovertemplate=("<b>%{customdata[0]}</b><br>" + name +
                           ": %{x:.2f}<br>n responden: %{customdata[1]:,}"
                           "<extra></extra>")))
    fig.update_layout(title=title, xaxis=dict(title=x_title,
                      range=zoom_range(list(d["a"]) + list(d["b"]), 0.10)),
                      yaxis=dict(title=None, automargin=True),
                      legend=dict(orientation="h", y=-0.16, title=None))
    st.plotly_chart(style_fig(fig, height or _cat_height(labels), ygrid=True),
                    width="stretch", config=PLOTLY_CFG)
    if note:
        st.caption(note)


def heatmap_usia(am, title, max_attrs=8):
    """Atribut prioritas per kelompok usia, dengan skala warna TETAP
    sehingga perbedaan 0,1 poin tidak tampak seperti perbedaan besar."""
    if am.empty:
        return
    pick = am.nsmallest(max_attrs, "Skor")
    groups = [u for u in ORDER_USIA if u in set(fdf["S2_2"])]
    if not groups:
        return
    idx_by_group = {u: fdf.index[fdf["S2_2"] == u] for u in groups}
    matrix, labels, cells = [], [], []
    for r in pick.itertuples():
        row = []
        for u in groups:
            s = NUM.loc[idx_by_group[u], r.variable].dropna()
            row.append(round(float(s.mean()), 2) if len(s) else None)
        matrix.append(row)
        cells.append([len(NUM.loc[idx_by_group[u], r.variable].dropna())
                      for u in groups])
        labels.append(wrap_label(r.Atribut, 42, 3))
    flat = [v for row in matrix for v in row if v is not None]
    if not flat:
        return
    fig = go.Figure(go.Heatmap(
        z=matrix, x=[f"{u.replace(' tahun', '')}<br>n={len(idx_by_group[u]):,}"
                     for u in groups],
        y=labels, colorscale=[[i / (len(BLUE_SEQ) - 1), c]
                              for i, c in enumerate(BLUE_SEQ)],
        zmin=min(flat), zmax=max(flat), xgap=2, ygap=2,
        colorbar=dict(title="Skor", thickness=12, len=0.7,
                      tickfont=dict(color=INK_2, size=11)),
        text=[[f"{v:.2f}" if v is not None else "–" for v in row]
              for row in matrix],
        texttemplate="%{text}", textfont=dict(size=11),
        customdata=cells,
        hovertemplate=("%{y}<br>Usia %{x}<br>Skor: %{z:.2f}"
                       "<br>n: %{customdata:,}<extra></extra>")))
    fig.update_layout(title=title, xaxis=dict(title=None),
                      yaxis=dict(title=None, automargin=True))
    st.plotly_chart(style_fig(fig, _cat_height(labels, base=170, row=34)),
                    width="stretch", config=PLOTLY_CFG)
    st.caption(
        f"Atribut dengan kepuasan terendah (prioritas perbaikan). Semakin gelap "
        f"= semakin puas. Skala warna dipatok pada rentang nyata "
        f"{min(flat):.2f}–{max(flat):.2f} — perbedaan antar sel memang kecil."
    )


def fmt(v, suffix="", nd=2):
    return "–" if v is None or pd.isna(v) else f"{v:.{nd}f}{suffix}"


def data_table(am):
    """Tabel lengkap: pasangan setiap grafik, memuat angka persisnya."""
    with st.expander("📋 Tabel data (angka persis + basis responden)"):
        t = am.copy()
        t["Gap"] = t["Gap"].round(2)
        show = t[["Atribut", "Subkategori", "Skor", "Kepentingan", "Gap",
                  "TopBox", "n"]].rename(columns={
                      "Skor": "Kepuasan", "TopBox": "% Sangat Puas",
                      "n": "n responden", "Subkategori": "Sub-Kategori"})
        st.dataframe(
            show.sort_values("Kepuasan"), hide_index=True, width="stretch",
            column_config={
                "Kepuasan": st.column_config.NumberColumn(format="%.2f"),
                "Kepentingan": st.column_config.NumberColumn(format="%.2f"),
                "Gap": st.column_config.NumberColumn(format="%+.2f"),
                "% Sangat Puas": st.column_config.NumberColumn(format="%.0f%%"),
                "n responden": st.column_config.NumberColumn(format="%d"),
            })


def thin_base_notice(am, what="atribut"):
    thin = am[am["n"] < MIN_BASE]
    if thin.empty:
        return
    st.warning(
        f"⚠️ **Basis responden tipis.** {len(thin)} dari {len(am)} {what} "
        f"pada tampilan ini dijawab kurang dari {MIN_BASE} responden "
        f"(n {int(thin['n'].min())}–{int(thin['n'].max())}). Skor mereka "
        f"jauh lebih tidak stabil daripada atribut dengan n ribuan; "
        f"rata-rata di kartu KPI sudah dibobot n, tetapi baca skor per "
        f"atribut dengan hati-hati.")


# ==================================================================
# KARTU KPI SECTION
# ==================================================================

def section_kpis(am, ov_rows):
    s = frame_summary(am)
    if not s:
        return
    ov = [col_stats(v)[0] for v in ov_rows["variable"]]
    ov = [v for v in ov if v is not None]
    ov_avg = sum(ov) / len(ov) if ov else None
    gap = (s["kepentingan"] - s["skor"]
           if s["kepentingan"] is not None and s["skor"] is not None else None)

    c = st.columns(5)
    c[0].metric("Skor Kepuasan", fmt(s["skor"], " / 6"),
                f"{fmt(s['topbox'], '%', 0)} menjawab 'sangat puas' (6) · "
                f"{s['n_attrs']} atribut", delta_color="off")
    c[1].metric("Penilaian Keseluruhan", fmt(ov_avg, " / 6"),
                "penilaian holistik responden" if ov_avg is not None
                else "tidak ditanyakan di bagian ini", delta_color="off")
    c[2].metric("Kepentingan", fmt(s["kepentingan"], " / 6"),
                "skala terpisah: 1–6 SANGAT PENTING", delta_color="off")
    if gap is None:
        c[3].metric("Gap Kepentingan − Kepuasan", "–", "–", delta_color="off")
    else:
        c[3].metric("Gap Kepentingan − Kepuasan", f"{gap:+.2f}",
                    "nasabah mengharapkan lebih" if gap > 0
                    else "kepuasan melampaui ekspektasi", delta_color="off")
    thin_flag = "⚠️ " if s["n_median"] < MIN_BASE else ""
    c[4].metric("Basis Responden", f"{thin_flag}{s['n_median']:,}",
                f"median per atribut (n {s['n_min']:,}–{s['n_max']:,})",
                delta_color="off")


# ==================================================================
# HALAMAN SECTION (generik)
# ==================================================================

def render_section_page(title, caption, sat_rows, ov_rows, heat_key):
    st.header(f"{PAGES.get(title, '')} {title}")
    st.caption(caption)

    am = attribute_frame(sat_rows)
    if am.empty:
        st.info("Tidak ada data atribut untuk bagian ini "
                "(periksa kolom include pada metadata).")
        return

    section_kpis(am, ov_rows)
    thin_base_notice(am)
    st.markdown("---")

    if am["Subkategori"].nunique() > 1:
        subgroup_dots(am, f"Skor per Sub-Kategori — {title}")

    if am["Kepentingan"].notna().any():
        st.markdown("---")
        st.subheader("📌 Importance-Performance Analysis (IPA)")
        ipa_chart(am, f"IPA — {title}")
        st.markdown("---")
        st.subheader("🎯 Prioritas berdasarkan gap kepentingan − kepuasan")
        gap_chart(am, f"Gap terbesar — {title}")

    st.markdown("---")
    st.subheader("🔍 Detail per sub-kategori")
    order = (subgroup_summary(am).sort_values("Skor")["Atribut"].tolist())
    tabs = st.tabs(list(order)) if len(order) > 1 else [st.container()]
    for tab, sub in zip(tabs, order):
        with tab:
            sam = am[am["Subkategori"] == sub]
            ss = frame_summary(sam)
            c1, c2, c3 = st.columns(3)
            c1.metric(f"Kepuasan — {sub}", fmt(ss["skor"], " / 6"),
                      "dibobot n", delta_color="off")
            c2.metric("Kepentingan", fmt(ss["kepentingan"], " / 6"),
                      f"gap {ss['kepentingan'] - ss['skor']:+.2f}"
                      if ss["kepentingan"] is not None else "–",
                      delta_color="off")
            c3.metric("Basis Responden", f"{ss['n_median']:,}",
                      f"{ss['n_attrs']} atribut", delta_color="off")
            attribute_dots(sam, f"Kepuasan per Atribut — {sub}")

    data_table(am)

    if st.toggle("Tampilkan perbandingan antar kelompok usia (heatmap)",
                 key=heat_key):
        heatmap_usia(am, f"Atribut Prioritas {title} per Kelompok Usia")


# ==================================================================
# HALAMAN 1 — RINGKASAN
# ==================================================================

def page_ringkasan():
    st.title("🏦 Bank XYZ Customer Experience Dashboard")
    st.caption("Monitoring **CSAT**, **Loyalty**, **NPS**, dan **Customer "
               "Experience** Bank XYZ berdasarkan survei nasabah. Gunakan "
               "filter di sidebar — pilihan filter tetap tersimpan saat "
               "berpindah halaman.")

    csat_m, _, csat_tb = col_stats("E1A")
    loy_m, _, loy_tb = col_stats("F1A")
    nps_val, prom, pas, det = nps(fdf.index)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("CSAT (Kepuasan)", fmt(csat_m, " / 6"),
              f"{fmt(csat_tb, '%', 0)} menjawab 'sangat puas' (6)",
              delta_color="off")
    c2.metric("Loyalty", fmt(loy_m, " / 6"),
              f"{fmt(loy_tb, '%', 0)} menjawab 'sangat setuju' (6)",
              delta_color="off")
    c3.metric("NPS", "–" if nps_val is None else f"{nps_val:.0f}",
              None if prom is None else f"{prom:.0f}% promoter",
              delta_color="off")
    c4.metric("Responden", f"{len(fdf):,}", "menjawab CSAT/Loyalty/NPS",
              delta_color="off")

    st.markdown("---")
    col_l, col_r = st.columns(2)

    # Skor kepuasan per touchpoint (radar, rentang mengikuti data)
    with col_l:
        tp_scores = []
        for tp in ATTR["touchpoint"].unique():
            am = attribute_frame(attrs_of(touchpoint=tp))
            if not am.empty:
                tp_scores.append((TP_LABEL.get(tp, tp),
                                  wmean(am["Skor"], am["n"]),
                                  int(am["n"].median())))
        tp_scores = [t for t in tp_scores if t[1] is not None]
        if tp_scores:
            cats = [t[0] for t in tp_scores]
            vals = [round(t[1], 2) for t in tp_scores]
            ns = [t[2] for t in tp_scores]
            lo, hi = min(vals), max(vals)
            pad = max(hi - lo, 0.20) * 0.35
            fig = go.Figure(go.Scatterpolar(
                r=vals + vals[:1], theta=cats + cats[:1],
                fill="toself", line=dict(color=SERIES_XYZ, width=2),
                fillcolor="rgba(46,119,174,0.10)",
                customdata=[[n] for n in ns + ns[:1]],
                hovertemplate=("%{theta}<br>Kepuasan: %{r:.2f}"
                               "<br>n median: %{customdata[0]:,}<extra></extra>")))
            fig.update_layout(
                title="Skor Kepuasan per Touchpoint",
                polar=dict(bgcolor=SURFACE,
                           radialaxis=dict(range=[lo - pad, hi + pad],
                                           visible=True, gridcolor=GRID,
                                           nticks=4, angle=67.5, tickangle=0,
                                           tickfont=dict(color=MUTED, size=10)),
                           angularaxis=dict(gridcolor=GRID,
                                            tickfont=dict(color=INK_2, size=11))),
                showlegend=False)
            fig = style_fig(fig, 470)
            # Nama touchpoint terpanjang duduk di luar lingkaran — beri ruang
            # supaya label spoke bawah/samping tidak terpotong.
            fig.update_layout(margin=dict(l=90, r=90, t=64, b=70))
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CFG)
            st.caption(
                f"⚠️ Sumbu dizoom ke {lo - pad:.2f}–{hi + pad:.2f}: seluruh "
                f"touchpoint sebenarnya hanya terpisah {hi - lo:.2f} poin "
                f"(dari 6). Bentuk radar melebih-lebihkan jarak itu — baca "
                f"angkanya, bukan luas areanya.")

    with col_r:
        # Komposisi NPS — part-to-whole 3 kelas berurutan
        if nps_val is not None:
            fig = go.Figure()
            for name, val, color in (("Promoter (9–10)", prom, ST_GOOD),
                                     ("Passive (7–8)", pas, ST_WARN),
                                     ("Detractor (0–6)", det, ST_CRIT)):
                fig.add_trace(go.Bar(
                    x=[val], y=["NPS"], orientation="h", name=name,
                    marker=dict(color=color, line=dict(width=2, color=SURFACE)),
                    text=[f"{val:.0f}%"], textposition="inside",
                    insidetextanchor="middle",
                    textfont=dict(color="white", size=12),
                    hovertemplate=f"{name}: %{{x:.1f}}%<extra></extra>"))
            fig.update_layout(barmode="stack", title="Komposisi NPS Bank XYZ",
                              xaxis=dict(title="% responden", range=[0, 100]),
                              yaxis=dict(visible=False),
                              legend=dict(orientation="h", y=-0.55, title=None))
            st.plotly_chart(style_fig(fig, 210), width="stretch",
                            config=PLOTLY_CFG)

        # XYZ vs Kompetitor — dua skala, dua grafik terpisah
        st.markdown("**Bank XYZ vs Kompetitor**")
        sc_items = []
        for label, vx, vk in (("CSAT — kepuasan", "E1A", "E1B"),
                              ("Loyalty — tetap menggunakan", "F1A", "F1B")):
            mx, nx, _ = col_stats(vx)
            mk, nk, _ = col_stats(vk)
            sc_items.append({"label": label, "a": mx, "b": mk,
                             "n_a": nx, "n_b": nk})
        n_xyz = sc_items[0]["n_a"] if sc_items else 0
        n_comp = sc_items[0]["n_b"] if sc_items else 0
        dumbbell(sc_items, "CSAT & Loyalty (skala 1–6)",
                 "Skor (sumbu dizoom)")
        gx, nx_nps, _ = col_stats("G1A", 10)
        gk, nk_nps, _ = col_stats("G1C", 10)
        dumbbell([{"label": "NPS — rata-rata rekomendasi", "a": gx, "b": gk,
                   "n_a": nx_nps, "n_b": nk_nps}],
                 "Rekomendasi (skala 0–10)", "Skor rata-rata (sumbu dizoom)")
        st.caption(
            f"⚠️ **Basis berbeda.** Skor XYZ dihitung dari **{n_xyz:,}** "
            f"responden; skor kompetitor hanya dari **{n_comp:,}** responden "
            f"yang juga memakai bank lain — bukan sampel yang sama. "
            f"NPS (0–10) sengaja dipisah dari CSAT/Loyalty (1–6): satu sumbu "
            f"untuk dua skala berbeda akan membuat NPS menenggelamkan keduanya.")

    # Prioritas lintas touchpoint — berbasis gap, bukan skor terendah
    st.markdown("---")
    st.subheader("🎯 Prioritas Perbaikan Lintas Touchpoint")
    all_am = attribute_frame(ATTR)
    if not all_am.empty:
        tp_of = dict(zip(ATTR["variable"], ATTR["touchpoint"]))
        labelled = all_am.copy()
        labelled["Atribut"] = [
            f"[{TP_LABEL.get(tp_of.get(v, ''), tp_of.get(v, ''))}] {a}"
            for v, a in zip(labelled["variable"], labelled["Atribut"])]
        gap_chart(labelled, "Gap kepentingan − kepuasan terbesar "
                            "(seluruh touchpoint)", top=10)
        with st.expander("📋 Semua atribut kepuasan (seluruh touchpoint)"):
            data = labelled[["Atribut", "Subkategori", "Skor", "Kepentingan",
                             "Gap", "TopBox", "n"]].rename(columns={
                                 "Skor": "Kepuasan", "TopBox": "% Sangat Puas",
                                 "n": "n responden",
                                 "Subkategori": "Sub-Kategori"})
            st.dataframe(data.sort_values("Kepuasan"), hide_index=True,
                         width="stretch")

    # Driver loyalitas
    if not LOYALTY_DRIVERS.empty:
        st.markdown("---")
        st.subheader("🔗 Driver Loyalitas — Mengapa Nasabah Tetap Setia?")
        st.caption("15 dimensi yang menjelaskan *mengapa* nasabah berniat tetap "
                   "menggunakan Bank XYZ (bukan hanya seberapa tinggi niatnya).")
        ld = []
        for r in LOYALTY_DRIVERS.itertuples():
            m, n, tb = col_stats(r.variable)
            if m is not None:
                ld.append({"Atribut": r.label, "Skor": m, "n": n, "TopBox": tb})
        if ld:
            dot_chart(pd.DataFrame(ld), "Skor",
                      "Skor per Driver Loyalitas (skala 1–6 SANGAT SETUJU)",
                      "Skor (sumbu dizoom ke rentang data)",
                      hover_rows=[("Atribut", "Driver", "text"),
                                  ("Skor", "Skor", "num2"),
                                  ("TopBox", "% sangat setuju", "pct0"),
                                  ("n", "n responden", "int")])

    # Emosi — sinyal dengan variasi terbesar di seluruh studi
    if not EMOTION.empty:
        st.markdown("---")
        st.subheader("💬 Emosi Saat Menggunakan Layanan Cabang")
        st.caption(
            "Skala 1–6 (SANGAT TIDAK SESUAI → SANGAT SESUAI). Item positif dan "
            "negatif **tidak boleh dirata-ratakan bersama**: pada item negatif "
            "skor rendah justru hasil yang baik. Inilah ukuran dengan selisih "
            "XYZ–kompetitor paling lebar di seluruh survei.")
        emo = []
        for r in EMOTION.itertuples():
            m, n, _ = col_stats(r.variable)
            if m is not None:
                emo.append({"label": r.label, "polarity": r.subgroup,
                            "bank": r.bank, "mean": m, "n": n})
        E = pd.DataFrame(emo)
        for pol, note in (("Emosi Positif", "Semakin **tinggi** semakin baik."),
                          ("Emosi Negatif",
                           ("Semakin **rendah** semakin baik — "
                            "item ini reverse-coded."))):
            part = E[E["polarity"] == pol]
            if part.empty:
                continue
            xyz = part[part["bank"] == "XYZ"].set_index("label")
            comp = part[part["bank"] == "Kompetitor"].set_index("label")
            items = [{"label": lab,
                      "a": xyz.loc[lab, "mean"],
                      "b": comp.loc[lab, "mean"] if lab in comp.index else None,
                      "n_a": int(xyz.loc[lab, "n"]),
                      "n_b": int(comp.loc[lab, "n"]) if lab in comp.index else 0}
                     for lab in xyz.index]
            dumbbell(items, f"{pol} — Bank XYZ vs Kompetitor",
                     "Skor rata-rata (sumbu dizoom)", note=note)

    # Digitalisasi cabang
    if not DIGITAL.empty:
        st.markdown("---")
        st.subheader("💻 Digitalisasi Layanan Cabang")
        st.caption("Persepsi nasabah atas digitalisasi cabang "
                   "(skala 1–6 SANGAT SETUJU).")
        dg = []
        for r in DIGITAL.itertuples():
            m, n, tb = col_stats(r.variable)
            if m is not None:
                dg.append({"Atribut": r.label, "Skor": m, "n": n, "TopBox": tb})
        if dg:
            dot_chart(pd.DataFrame(dg), "Skor", "Skor per Pernyataan",
                      "Skor (sumbu dizoom ke rentang data)",
                      hover_rows=[("Atribut", "Pernyataan", "text"),
                                  ("Skor", "Skor", "num2"),
                                  ("TopBox", "% sangat setuju", "pct0"),
                                  ("n", "n responden", "int")])

    # Profil responden
    st.markdown("---")
    st.subheader("👤 Profil Responden")
    p1, p2 = st.columns([1, 2])
    if "S1" in fdf.columns:
        with p1:
            g = fdf["S1"].value_counts()
            total = int(g.sum())
            fig = go.Figure()
            for i, (name, val) in enumerate(g.items()):
                fig.add_trace(go.Bar(
                    x=[val / total * 100], y=["Gender"], orientation="h",
                    name=str(name),
                    marker=dict(color=[SERIES_XYZ, SERIES_COMP][i % 2],
                                line=dict(width=2, color=SURFACE)),
                    text=[f"{name} · {val / total * 100:.0f}%"],
                    textposition="inside", insidetextanchor="middle",
                    textfont=dict(color="white", size=12),
                    hovertemplate=f"{name}: %{{x:.1f}}% ({val:,})<extra></extra>"))
            # Setiap segmen sudah diberi label langsung, jadi sumbu % hanya
            # menambah keramaian.
            fig.update_layout(barmode="stack", title="Gender",
                              xaxis=dict(visible=False, range=[0, 100]),
                              yaxis=dict(visible=False), showlegend=False)
            st.plotly_chart(style_fig(fig, 200), width="stretch",
                            config=PLOTLY_CFG)
    if "S2_2" in fdf.columns:
        with p2:
            u = (fdf["S2_2"].value_counts().reindex(ORDER_USIA).dropna())
            fig = go.Figure(go.Bar(
                x=[x.replace(" tahun", "") for x in u.index], y=u.values,
                marker=dict(color=SERIES_XYZ), width=0.62,
                hovertemplate="Usia %{x}<br>%{y:,} responden<extra></extra>"))
            fig.update_layout(title="Kelompok Usia",
                              xaxis=dict(title=None),
                              yaxis=dict(title="Responden"), bargap=0.3)
            st.plotly_chart(style_fig(fig, 200), width="stretch",
                            config=PLOTLY_CFG)
    if "S4" in fdf.columns:
        lm = (fdf["S4"].value_counts().reindex(ORDER_LAMA).dropna())
        fig = go.Figure(go.Bar(
            x=lm.values, y=[wrap_label(x, 30) for x in lm.index],
            orientation="h", marker=dict(color=SERIES_XYZ), width=0.62,
            hovertemplate="%{y}<br>%{x:,} responden<extra></extra>"))
        fig.update_layout(title="Lama Menjadi Nasabah",
                          xaxis=dict(title="Responden"),
                          yaxis=dict(title=None, automargin=True), bargap=0.3)
        st.plotly_chart(style_fig(fig, 260), width="stretch", config=PLOTLY_CFG)


# ==================================================================
# HALAMAN SERVICE EXPERIENCE (5 touchpoint petugas)
# ==================================================================

def page_service():
    st.header(f"{PAGES[PAGE_SE]} Service Experience")
    st.caption("Pengalaman layanan dari petugas cabang: Customer Service, "
               "Teller, Sekuriti, Customer Advisor, dan Sarana Elektronik "
               "Layanan.")

    avail = [tp for tp in TP_SERVICE if not attrs_of(touchpoint=tp).empty]
    if not avail:
        st.info("Tidak ada atribut Service Experience pada metadata.")
        return

    cols = st.columns(len(avail))
    for i, tp in enumerate(avail):
        am = attribute_frame(attrs_of(touchpoint=tp))
        s = frame_summary(am)
        if not s:
            cols[i].metric(TP_LABEL.get(tp, tp), "–")
            continue
        thin = "⚠️ " if s["n_median"] < MIN_BASE else ""
        cols[i].metric(TP_LABEL.get(tp, tp), fmt(s["skor"]),
                       f"{s['n_attrs']} atribut · {thin}n={s['n_median']:,}",
                       delta_color="off")
    st.caption("Basis responden (n) adalah median jumlah penjawab per atribut — "
               "bukan jumlah responden terfilter. ⚠️ menandai basis di bawah "
               f"{MIN_BASE}.")

    st.markdown("---")
    # pilihan touchpoint juga dibuat persisten antar halaman
    cur = st.session_state.get("se_tp_keep", avail[0])
    if cur not in avail:
        cur = avail[0]
    sel = st.selectbox("Pilih area layanan untuk detail:", avail,
                       index=avail.index(cur),
                       format_func=lambda t: TP_LABEL.get(t, t))
    st.session_state["se_tp_keep"] = sel

    render_section_page(
        TP_LABEL.get(sel, sel),
        f"Detail kepuasan dan kepentingan layanan {TP_LABEL.get(sel, sel)} "
        f"per sub-kategori dan atribut.",
        attrs_of(touchpoint=sel), overall_of(touchpoint=sel),
        heat_key=f"hm_se_{sel}")


# ==================================================================
# HALAMAN LAIN
# ==================================================================

def page_brand_image():
    render_section_page(
        PAGE_BI,
        "Persepsi nasabah terhadap citra Bank XYZ (skala 1–6; jawaban "
        "'tidak relevan' dikeluarkan dari perhitungan).",
        attrs_of(section=PAGE_BI), overall_of(section=PAGE_BI),
        heat_key="hm_bi")


def page_branch_facilities():
    render_section_page(
        PAGE_BF,
        "Penilaian nasabah terhadap fasilitas fisik cabang Bank XYZ "
        "(skala 1–6).",
        attrs_of(section=PAGE_BF), overall_of(section=PAGE_BF),
        heat_key="hm_bf")


def page_atm():
    render_section_page(
        PAGE_ATM,
        "Pengalaman nasabah menggunakan ATM Bank XYZ (skala 1–6).",
        attrs_of(section=PAGE_ATM), overall_of(section=PAGE_ATM),
        heat_key="hm_atm")


# ==================================================================
# ROUTING UTAMA
# ==================================================================

if page == PAGE_RINGKASAN:
    page_ringkasan()
elif page == PAGE_BI:
    page_brand_image()
elif page == PAGE_BF:
    page_branch_facilities()
elif page == PAGE_SE:
    page_service()
elif page == PAGE_ATM:
    page_atm()
