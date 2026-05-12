"""
Marketplace Listing Checker
===========================
Upload: Zecom Tracker · Lazada · Shopee · Zalora · Inventory
Output: Summary view + downloadable Excel audit report

Run:  streamlit run marketplace_checker_app.py
"""

import io
import warnings
from datetime import date

import pandas as pd
import streamlit as st
import xlsxwriter

warnings.filterwarnings("ignore")

TODAY = pd.Timestamp(date.today())

# ─────────────────────────────────────────────────────────────────────────────
# PAGE SETUP
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Marketplace Listing Checker",
    page_icon="🛍️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  .main .block-container { padding-top: 1.4rem; padding-bottom: 2rem; }
  h1  { color: #1F4E79; }
  h2  { color: #2E4057; font-size: 1.1rem; margin-bottom: 4px; }
  .stMetric label { font-size: 0.8rem; }
  div[data-testid="stMetricValue"] { font-size: 1.5rem; font-weight: 700; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────────────────────
def norm_ean(series: pd.Series) -> pd.Series:
    """Convert EAN column to clean integer-string (handles float, leading-zero issues)."""
    def _n(x):
        if pd.isna(x):
            return ""
        try:
            return str(int(float(str(x).strip())))
        except Exception:
            return str(x).strip()
    return series.apply(_n)


def find_col(df: pd.DataFrame, *keywords) -> str | None:
    """Return first column whose lower-case name contains any keyword."""
    for kw in keywords:
        for c in df.columns:
            if kw.lower() in c.lower():
                return c
    return None


# ─────────────────────────────────────────────────────────────────────────────
# FILE LOADERS  (return slim DataFrames with standard column names)
# ─────────────────────────────────────────────────────────────────────────────
def load_zecom(file_bytes: bytes):
    """
    Returns (brand_status_df, content_df).
    Brand Status expected columns: Article No, Launch Dates, LAZADA, SHOPEE, ZALORA, Stock
    Content file expected columns: Article No, EAN
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {s.lower(): s for s in xl.sheet_names}

    # Find sheets by keyword
    bs_sheet = next((v for k, v in sheets.items() if "brand" in k or "status" in k), xl.sheet_names[0])
    ct_sheet = next((v for k, v in sheets.items() if "content" in k),
                    xl.sheet_names[1] if len(xl.sheet_names) > 1 else xl.sheet_names[0])

    df_bs = pd.read_excel(xl, sheet_name=bs_sheet)
    df_ct = pd.read_excel(xl, sheet_name=ct_sheet)
    df_bs.columns = df_bs.columns.str.strip()
    df_ct.columns = df_ct.columns.str.strip()

    # Normalise Brand Status column names
    rename_bs = {}
    for c in df_bs.columns:
        cl = c.lower()
        if "article" in cl:          rename_bs[c] = "Article No"
        elif "launch" in cl:         rename_bs[c] = "Launch Date"
        elif "lazada" in cl:         rename_bs[c] = "Int_Lazada"
        elif "shopee" in cl:         rename_bs[c] = "Int_Shopee"
        elif "zalora" in cl:         rename_bs[c] = "Int_Zalora"
        elif "stock" in cl:          rename_bs[c] = "BS_Stock"
    df_bs = df_bs.rename(columns=rename_bs)
    df_bs["Article No"] = df_bs["Article No"].astype(str).str.strip()

    # Normalise Content file column names
    rename_ct = {}
    for c in df_ct.columns:
        cl = c.lower()
        if "article" in cl:          rename_ct[c] = "Article No"
        elif "ean" in cl:            rename_ct[c] = "EAN_raw"
    df_ct = df_ct.rename(columns=rename_ct)
    df_ct["Article No"] = df_ct["Article No"].astype(str).str.strip()
    df_ct["EAN"]        = norm_ean(df_ct["EAN_raw"].astype(str))

    return df_bs, df_ct[["Article No", "EAN"]]


def load_lazada(file_bytes: bytes) -> pd.DataFrame:
    """
    Lazada export has 2 description rows after the header; skiprows=[1,2].
    Key columns: SellerSKU (EAN), status, Product Name
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = next((s for s in xl.sheet_names if "template" in s.lower()), xl.sheet_names[0])
    df = pd.read_excel(xl, sheet_name=sheet, header=0, skiprows=[1, 2])
    df.columns = df.columns.str.strip()

    # Drop description / metadata rows (Product ID must be numeric)
    pid_col = find_col(df, "product id")
    if pid_col:
        df = df[df[pid_col].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x))].copy()

    ean_col  = find_col(df, "sellersku", "seller_sku", "seller sku")
    stat_col = find_col(df, "status")
    name_col = find_col(df, "product name", "name")

    df["EAN"]    = norm_ean(df[ean_col].astype(str)) if ean_col else ""
    df["Status"] = df[stat_col].str.strip().str.lower() if stat_col else "unknown"
    df["Name"]   = df[name_col].fillna("") if name_col else ""

    return df[["EAN", "Status", "Name"]].drop_duplicates("EAN", keep="last")


def load_shopee(file_bytes: bytes) -> pd.DataFrame:
    """
    Key columns: SKU (EAN), Status, Product Name
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    df = pd.read_excel(xl, sheet_name=xl.sheet_names[0])
    df.columns = df.columns.str.strip()

    ean_col  = find_col(df, "^sku$", "sku")  # SKU column (not Parent SKU)
    # Prefer exact "SKU" over "Parent SKU"
    ean_col  = "SKU" if "SKU" in df.columns else find_col(df, "sku")
    stat_col = find_col(df, "status")
    name_col = find_col(df, "product name", "name")

    df["EAN"]    = norm_ean(df[ean_col].astype(str)) if ean_col else ""
    df["Status"] = df[stat_col].str.strip().str.lower() if stat_col else "unknown"
    df["Name"]   = df[name_col].fillna("") if name_col else ""

    return df[["EAN", "Status", "Name"]].drop_duplicates("EAN", keep="last")


def load_zalora(file_bytes: bytes) -> pd.DataFrame:
    """
    Key columns: SellerSku (EAN), Status, Name
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    df = pd.read_excel(xl, sheet_name=xl.sheet_names[0])
    df.columns = df.columns.str.strip()

    ean_col  = find_col(df, "sellersku", "seller_sku", "seller sku")
    stat_col = find_col(df, "status")
    name_col = find_col(df, "name")

    df["EAN"]    = norm_ean(df[ean_col].astype(str)) if ean_col else ""
    df["Status"] = df[stat_col].str.strip().str.lower() if stat_col else "unknown"
    df["Name"]   = df[name_col].fillna("") if name_col else ""

    return df[["EAN", "Status", "Name"]].drop_duplicates("EAN", keep="last")


def load_inventory(file_bytes: bytes, filename: str) -> tuple[pd.Series, str]:
    """
    Returns (ean_to_qty Series, match_key) where:
      - ean_to_qty : EAN (str) → Avail_Qty  for xlsx files (full EAN precision)
      - match_key  : "EAN" always — raises if CSV has precision-lost EANs

    CSV exports from Excel truncate 13-digit EANs to scientific notation
    (e.g. 4063699714586 → 4.0637E+12), permanently destroying the last digits.
    The app therefore REQUIRES the inventory file in .xlsx format for EAN-level
    stock matching.  If a CSV with truncated EANs is detected, a clear error is
    raised so the user knows to re-export as xlsx.
    """
    is_csv = filename.lower().endswith(".csv")

    if is_csv:
        # Read a small sample to check for scientific-notation EANs
        sample_raw = io.BytesIO(file_bytes).read(2048).decode("utf-8", errors="ignore")
        import re
        if re.search(r"\d+\.\d+E\+\d+", sample_raw, re.IGNORECASE):
            raise ValueError(
                "PRECISION_LOSS: The inventory CSV has EANs stored in scientific notation "
                "(e.g. 4.0637E+12). The full 13-digit EAN cannot be recovered from this file.\n\n"
                "Please re-export your inventory from Excel as .xlsx (not CSV) and upload that instead. "
                "Excel .xlsx preserves full EAN integer precision."
            )
        df = pd.read_csv(io.BytesIO(file_bytes), dtype=str)
    else:
        # xlsx — read EAN as integer via openpyxl; norm to string afterwards
        df = pd.read_excel(io.BytesIO(file_bytes))

    df.columns = df.columns.str.strip()

    ean_col = find_col(df, "ean")
    qty_col = find_col(df, "avail_qty", "avail qty", "available")
    if not qty_col:
        qty_col = find_col(df, "sumstock", "sum_stock")
    if not qty_col:
        qty_col = find_col(df, "qty", "stock", "quantity")

    if not ean_col or not qty_col:
        raise ValueError("Could not find EAN or quantity column in inventory file.")

    df["_EAN"] = norm_ean(df[ean_col].astype(str))
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    # One row per EAN — keep last occurrence if any duplicates
    return df.drop_duplicates("_EAN", keep="last").set_index("_EAN")[qty_col], "EAN"


# ─────────────────────────────────────────────────────────────────────────────
# STATUS + COMMENT ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def status_and_comment(mp_raw: str | None, intended: str, launch_ts,
                        inv_qty: float, bs_stock: float) -> tuple[str, str]:
    """
    Derive (listing_status, comment) for a single EAN on one marketplace.

    mp_raw    : 'active' | 'inactive' | None (not in marketplace file)
    intended  : 'Active' | 'Inactive' (from Zecom Brand Status)
    launch_ts : pd.Timestamp or NaT
    inv_qty   : available qty from inventory file
    bs_stock  : stock value from Brand Status tab
    """
    # Effective stock: inventory is authoritative; fall back to brand status stock
    stock = inv_qty if (not pd.isna(inv_qty) and inv_qty >= 0) else \
            (bs_stock if not pd.isna(bs_stock) else 0)
    zero_stock   = (stock == 0)
    future_launch = pd.notna(launch_ts) and pd.Timestamp(launch_ts) > TODAY
    inactive_intent = str(intended).strip().lower() == "inactive"

    ld_str = pd.Timestamp(launch_ts).strftime("%Y-%m-%d") if pd.notna(launch_ts) else "N/A"

    # ── Listing Status ────────────────────────────────────────
    if mp_raw == "active":
        listing = "Active"
    elif mp_raw == "inactive":
        listing = "Listed but Inactive"
    else:
        listing = "Not Listed"

    # ── Comment ──────────────────────────────────────────────
    # Collect flags for reason building
    flags = []
    if future_launch:    flags.append(f"Future Launch Date ({ld_str})")
    if zero_stock:       flags.append("0 Stock")
    if inactive_intent:  flags.append("Marked Inactive in Zecom")

    flag_str = " | ".join(flags)

    if listing == "Active":
        if zero_stock and future_launch:
            comment = f"⚠️ Active — 0 Stock + Future Launch Date ({ld_str}). Review listing."
        elif zero_stock:
            comment = "⚠️ Active — 0 Stock. Product live but no inventory."
        elif future_launch:
            comment = f"⚠️ Active — Future Launch Date ({ld_str}). Listed before launch."
        elif inactive_intent:
            comment = "⚠️ Active — Zecom marks as Inactive. Check intended status."
        else:
            comment = "✅ Active — Stock available, past launch, intended Active."

    elif listing == "Listed but Inactive":
        if zero_stock and future_launch:
            comment = f"ℹ️ Listed Inactive — 0 Stock + Future Launch ({ld_str}). Expected inactive."
        elif zero_stock:
            comment = "ℹ️ Listed Inactive — 0 Stock. Expected inactive until restocked."
        elif future_launch:
            comment = f"ℹ️ Listed Inactive — Future Launch Date ({ld_str}). Will activate on launch."
        elif inactive_intent:
            comment = "ℹ️ Listed Inactive — Intended Inactive in Zecom. Expected."
        else:
            comment = "🔴 Listed Inactive — No clear reason. Stock available, past launch. Review."

    else:  # Not Listed
        if inactive_intent and zero_stock:
            comment = "ℹ️ Not Listed — Intended Inactive + 0 Stock. Expected not to be listed."
        elif inactive_intent:
            comment = "ℹ️ Not Listed — Marked Inactive in Zecom. Expected not to be listed."
        elif zero_stock and future_launch:
            comment = f"ℹ️ Not Listed — 0 Stock + Future Launch ({ld_str}). OK to be unlisted."
        elif zero_stock:
            comment = "ℹ️ Not Listed — 0 Stock. Acceptable; list when stock arrives."
        elif future_launch:
            comment = f"ℹ️ Not Listed — Future Launch Date ({ld_str}). OK; list on launch."
        else:
            comment = "🔴 Not Listed — ACTIVE INTENT + Stock Available + Past Launch. ACTION NEEDED."

    return listing, comment


# ─────────────────────────────────────────────────────────────────────────────
# CORE PROCESSING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def run_analysis(zecom_b, laz_b, sho_b, zal_b, inv_b, inv_name):
    """Build the master EAN-level DataFrame with status + comments for all 3 MPs."""

    # Load files
    df_bs, df_ct = load_zecom(zecom_b)
    laz_df = load_lazada(laz_b)
    sho_df = load_shopee(sho_b)
    zal_df = load_zalora(zal_b)
    inv_s, _ = load_inventory(inv_b, inv_name)   # EAN (str) → Avail_Qty

    # Filter content to Brand Status articles only
    bs_set = set(df_bs["Article No"])
    df = df_ct[df_ct["Article No"].isin(bs_set)].copy()

    # Merge Brand Status fields onto every EAN row
    df = df.merge(df_bs, on="Article No", how="left")

    # EAN-level inventory stock — join on EAN string
    df["Inv_Stock"] = pd.to_numeric(df["EAN"].map(inv_s), errors="coerce").fillna(0)
    df["BS_Stock"]  = pd.to_numeric(df.get("BS_Stock", 0), errors="coerce").fillna(0)

    # Effective stock: inventory Avail_Qty is authoritative; fall back to Brand Status stock
    df["Eff_Stock"] = df.apply(
        lambda r: r["Inv_Stock"] if r["Inv_Stock"] > 0 else r["BS_Stock"], axis=1
    )

    # EAN → marketplace status lookups
    laz_stat = laz_df.set_index("EAN")["Status"].to_dict()
    sho_stat = sho_df.set_index("EAN")["Status"].to_dict()
    zal_stat = zal_df.set_index("EAN")["Status"].to_dict()
    laz_name = laz_df.set_index("EAN")["Name"].to_dict()
    sho_name = sho_df.set_index("EAN")["Name"].to_dict()
    zal_name = zal_df.set_index("EAN")["Name"].to_dict()

    # Derive per-MP status and comment (vectorised where possible)
    results = []
    for row in df.itertuples(index=False):
        d     = row._asdict()
        ean   = d.get("EAN", "")
        art   = d.get("Article No", "")
        ld    = d.get("Launch Date", None)
        bs_s  = d.get("BS_Stock", 0)
        inv_s_val = d.get("Inv_Stock", 0)
        i_laz = d.get("Int_Lazada", "")
        i_sho = d.get("Int_Shopee", "")
        i_zal = d.get("Int_Zalora", "")

        ld_ts     = pd.Timestamp(ld) if pd.notna(ld) else pd.NaT
        ld_str    = ld_ts.strftime("%Y-%m-%d") if pd.notna(ld_ts) else ""
        is_future = pd.notna(ld_ts) and ld_ts > TODAY

        laz_s, laz_c = status_and_comment(laz_stat.get(ean), i_laz, ld_ts, inv_s_val, bs_s)
        sho_s, sho_c = status_and_comment(sho_stat.get(ean), i_sho, ld_ts, inv_s_val, bs_s)
        zal_s, zal_c = status_and_comment(zal_stat.get(ean), i_zal, ld_ts, inv_s_val, bs_s)

        results.append({
            "Article No"           : art,
            "EAN"                  : ean,
            "Launch Date"          : ld_str,
            "Future Launch"        : "Yes" if is_future else "No",
            "Inventory Stock"      : int(inv_s_val),
            "Brand Status Stock"   : int(bs_s),
            "Effective Stock"      : int(d.get("Eff_Stock", 0)),
            # Lazada
            "Intended - Lazada"    : str(i_laz) if pd.notna(i_laz) else "",
            "Lazada Status"        : laz_s,
            "Lazada Comment"       : laz_c,
            "Lazada Product Name"  : laz_name.get(ean, ""),
            # Shopee
            "Intended - Shopee"    : str(i_sho) if pd.notna(i_sho) else "",
            "Shopee Status"        : sho_s,
            "Shopee Comment"       : sho_c,
            "Shopee Product Name"  : sho_name.get(ean, ""),
            # Zalora
            "Intended - Zalora"    : str(i_zal) if pd.notna(i_zal) else "",
            "Zalora Status"        : zal_s,
            "Zalora Comment"       : zal_c,
            "Zalora Product Name"  : zal_name.get(ean, ""),
        })

    return pd.DataFrame(results)


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────────────────────────────────────
def build_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"constant_memory": True, "in_memory": True})

    # ── Format helpers ────────────────────────────────────────
    def hf(bg, fc="white", sz=10, bold=True):
        return wb.add_format({"bold": bold, "bg_color": bg, "font_color": fc,
                               "border": 1, "text_wrap": True, "valign": "vcenter",
                               "align": "center", "font_size": sz})

    H_BLUE   = hf("#1F4E79"); H_ORANGE = hf("#C55A11")
    H_GREEN  = hf("#375623"); H_RED    = hf("#C00000"); H_PURPLE = hf("#4472C4")
    H_GREY   = hf("#404040")

    CELL     = wb.add_format({"border": 1, "valign": "vcenter", "font_size": 9})
    C_GRN    = wb.add_format({"border": 1, "valign": "vcenter", "font_color": "#375623", "font_size": 9})
    C_RED    = wb.add_format({"border": 1, "valign": "vcenter", "font_color": "#C00000", "font_size": 9})
    C_ORG    = wb.add_format({"border": 1, "valign": "vcenter", "font_color": "#BF8F00", "font_size": 9})
    C_BOLD   = wb.add_format({"border": 1, "valign": "vcenter", "bold": True, "font_size": 9})
    C_NUM    = wb.add_format({"border": 1, "valign": "vcenter", "num_format": "#,##0", "font_size": 9})
    C_WRAP   = wb.add_format({"border": 1, "valign": "vcenter", "text_wrap": True, "font_size": 9})
    TITLE_F  = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#1F4E79",
                               "font_color": "white", "align": "center", "valign": "vcenter"})
    NOTE_F   = wb.add_format({"italic": True, "font_color": "#595959", "font_size": 9})
    SUB_F    = wb.add_format({"italic": True, "font_color": "#404040", "font_size": 9, "align": "center"})

    today_str = pd.Timestamp("today").strftime("%Y-%m-%d")

    STATUS_FMT = {
        "Active"              : C_GRN,
        "Not Listed"          : C_RED,
        "Listed but Inactive" : C_ORG,
    }

    def write_block(ws, data_df, hdr_fmt, col_widths, start_row=0):
        cols = list(data_df.columns)
        for c, (col, w) in enumerate(zip(cols, col_widths)):
            ws.write(start_row, c, col, hdr_fmt)
            ws.set_column(c, c, w)
        r = start_row + 1
        for row_t in data_df.itertuples(index=False):
            for c, val in enumerate(row_t):
                isna = False
                if not isinstance(val, str):
                    try: isna = pd.isna(val)
                    except: pass
                if isna:
                    ws.write(r, c, "", CELL)
                elif isinstance(val, (int, float)):
                    ws.write_number(r, c, val, C_NUM)
                else:
                    v   = str(val)
                    fmt = STATUS_FMT.get(v, C_WRAP if len(v) > 35 else CELL)
                    ws.write(r, c, v, fmt)
            r += 1
        ws.freeze_panes(start_row + 1, 0)
        if r > start_row + 1:
            ws.autofilter(start_row, 0, r - 1, len(cols) - 1)

    # ─── SHEET 1: SUMMARY ─────────────────────────────────────
    ws0 = wb.add_worksheet("Summary")
    ws0.set_tab_color("#1F4E79")
    ws0.merge_range(0, 0, 0, 8, f"MARKETPLACE LISTING AUDIT  |  {today_str}", TITLE_F)
    ws0.set_row(0, 22)
    ws0.merge_range(1, 0, 1, 8, "Stock source: Inventory File (Avail_Qty)  |  Zecom Tracker: Brand Status + Content File", SUB_F)

    SH = ["Marketplace", "Total EANs", "Active",
          "Not Listed", "Listed but Inactive",
          "Not Listed — 0 Stock (OK)",
          "Not Listed — Future Launch (OK)",
          "Not Listed — Intended Inactive (OK)",
          "🔴 Not Listed — ACTION NEEDED"]
    SHF = [H_BLUE, H_BLUE, H_GREEN, H_RED, H_ORANGE, H_ORANGE, H_ORANGE, H_ORANGE, H_RED]
    SHW = [14, 12, 10, 12, 20, 24, 26, 28, 30]

    for c, (h, f, w) in enumerate(zip(SH, SHF, SHW)):
        ws0.write(2, c, h, f)
        ws0.set_column(c, c, w)

    for r, (mp, sc) in enumerate(
        [("Lazada", "Lazada Status"), ("Shopee", "Shopee Status"), ("Zalora", "Zalora Status")], start=3
    ):
        total   = len(df)
        active  = (df[sc] == "Active").sum()
        nl      = (df[sc] == "Not Listed").sum()
        li      = (df[sc] == "Listed but Inactive").sum()
        nl_0s   = ((df[sc] == "Not Listed") & (df["Effective Stock"] == 0)
                   & (df["Future Launch"] == "No")).sum()
        nl_fut  = ((df[sc] == "Not Listed") & (df["Future Launch"] == "Yes")).sum()
        nl_ina  = ((df[sc] == "Not Listed") &
                   (df[f"Intended - {mp}"].str.lower() == "inactive")).sum()
        nl_act  = ((df[sc] == "Not Listed") &
                   (df["Effective Stock"] > 0) &
                   (df["Future Launch"] == "No") &
                   (df[f"Intended - {mp}"].str.lower() != "inactive")).sum()
        vals = [mp, total, active, nl, li, nl_0s, nl_fut, nl_ina, nl_act]
        for c2, v in enumerate(vals):
            if isinstance(v, str): ws0.write(r, c2, v, C_BOLD)
            else: ws0.write_number(r, c2, int(v), C_NUM)

    # Legend
    row_l = 7
    ws0.write(row_l, 0, "LEGEND", wb.add_format({"bold": True, "font_size": 10})); row_l += 1
    legends = [
        ("✅ Active", "EAN is live and active on the marketplace."),
        ("⚠️ Active — 0 Stock", "Live on MP but no inventory — review."),
        ("⚠️ Active — Future Launch", "Listed before launch date — review."),
        ("ℹ️ Not Listed — 0 Stock", "No inventory; expected to be inactive until restocked."),
        ("ℹ️ Not Listed — Future Launch", "Launch date is in the future; listing expected after launch."),
        ("ℹ️ Not Listed — Intended Inactive", "Zecom Tracker marks article as Inactive; expected."),
        ("⚠️ Listed Inactive — No reason", "On MP but inactive with no clear justification; review."),
        ("🔴 Not Listed — ACTION NEEDED", "Intended Active + stock > 0 + past launch date but missing from marketplace."),
    ]
    for label, desc in legends:
        ws0.write(row_l, 0, label, NOTE_F)
        ws0.write(row_l, 1, desc,  NOTE_F)
        row_l += 1
    ws0.freeze_panes(3, 0)

    # ─── PER-MARKETPLACE SHEETS ────────────────────────────────
    BASE_COLS = ["Article No", "EAN", "Launch Date", "Future Launch",
                 "Inventory Stock", "Brand Status Stock", "Effective Stock"]
    BASE_W    = [14, 18, 12, 12, 14, 16, 14]

    MP_CFG = {
        "Lazada": (H_BLUE,   "#1F4E79", "#BDD7EE",
                   "Intended - Lazada", "Lazada Status", "Lazada Comment", "Lazada Product Name"),
        "Shopee": (H_ORANGE, "#C55A11", "#FCE4D6",
                   "Intended - Shopee", "Shopee Status", "Shopee Comment", "Shopee Product Name"),
        "Zalora": (H_GREEN,  "#375623", "#E2EFDA",
                   "Intended - Zalora", "Zalora Status", "Zalora Comment", "Zalora Product Name"),
    }
    MP_EXTRA_W = [16, 20, 40, 35]

    for mp, (hf_mp, dark, light, ic, sc, cc, nc) in MP_CFG.items():
        all_cols = BASE_COLS + [ic, sc, nc, cc]
        all_w    = BASE_W   + MP_EXTRA_W
        sub      = df[all_cols].copy()

        # ① All EANs
        ws_a = wb.add_worksheet(f"{mp} - All EANs")
        ws_a.set_tab_color(dark)
        write_block(ws_a, sub, hf_mp, all_w)

        # ② Action Required  (not listed, active intent, stock > 0, past launch)
        action_df = sub[
            (sub[sc] == "Not Listed") &
            (sub[ic].str.lower() != "inactive") &
            (sub["Effective Stock"] > 0) &
            (sub["Future Launch"] == "No")
        ].copy()
        ws_ar = wb.add_worksheet(f"{mp} - Action Required")
        ws_ar.set_tab_color(dark)
        ws_ar.merge_range(0, 0, 0, len(all_cols) - 1,
            f"{mp}  ·  ACTION REQUIRED — Not Listed | Active Intent | Stock Available  ({len(action_df):,} EANs)", TITLE_F)
        write_block(ws_ar, action_df, hf_mp, all_w, start_row=1)

        # ③ Not Listed — all (with reason column)
        nl_df = sub[sub[sc] == "Not Listed"].copy()
        ws_nl = wb.add_worksheet(f"{mp} - Not Listed")
        ws_nl.set_tab_color(light)
        ws_nl.merge_range(0, 0, 0, len(all_cols) - 1,
            f"{mp}  ·  Not Listed — All Reasons  ({len(nl_df):,} EANs)", TITLE_F)
        write_block(ws_nl, nl_df, hf_mp, all_w, start_row=1)

        # ④ Listed but Inactive
        li_df = sub[sub[sc] == "Listed but Inactive"].copy()
        ws_li = wb.add_worksheet(f"{mp} - Listed Inactive")
        ws_li.set_tab_color(light)
        ws_li.merge_range(0, 0, 0, len(all_cols) - 1,
            f"{mp}  ·  Listed but Inactive  ({len(li_df):,} EANs)", TITLE_F)
        write_block(ws_li, li_df, hf_mp, all_w, start_row=1)

    # ─── ALL MARKETPLACES COMBINED ─────────────────────────────
    combo_cols = BASE_COLS + [
        "Intended - Lazada", "Lazada Status", "Lazada Comment",
        "Intended - Shopee", "Shopee Status", "Shopee Comment",
        "Intended - Zalora", "Zalora Status", "Zalora Comment",
    ]
    combo_w = BASE_W + [16, 20, 38, 16, 20, 38, 16, 20, 38]
    ws_combo = wb.add_worksheet("All Marketplaces")
    ws_combo.set_tab_color("#7030A0")
    write_block(ws_combo, df[combo_cols], H_PURPLE, combo_w)

    wb.close()
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────
st.title("🛍️ Marketplace Listing Checker")
st.caption("Upload your files → run analysis → view summary → download full Excel report.")

# ── Sidebar ─────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 File Upload")
    zecom_f = st.file_uploader("Zecom Tracker (.xlsx)",      type=["xlsx"])
    laz_f   = st.file_uploader("Lazada MP File (.xlsx)",     type=["xlsx"])
    sho_f   = st.file_uploader("Shopee MP File (.xlsx)",     type=["xlsx"])
    zal_f   = st.file_uploader("Zalora MP File (.xlsx)",     type=["xlsx"])
    inv_f   = st.file_uploader("Inventory (.xlsx preferred — CSV loses EAN precision)",
                               type=["csv", "xlsx"])

    st.divider()
    all_ready = all([zecom_f, laz_f, sho_f, zal_f, inv_f])
    run_btn   = st.button("▶ Run Analysis", type="primary",
                          use_container_width=True, disabled=not all_ready)
    if not all_ready:
        missing = [n for n, f in [("Zecom",zecom_f),("Lazada",laz_f),
                                   ("Shopee",sho_f),("Zalora",zal_f),
                                   ("Inventory",inv_f)] if not f]
        st.caption(f"Still needed: {', '.join(missing)}")

    st.divider()
    st.markdown("""
**Stock logic**
- Stock source: Inventory `Avail_Qty` per EAN
- Upload inventory as **.xlsx** (CSV loses EAN digit precision)
- Stock = 0 → expected inactive
- Not listed + stock 0 → ✅ OK
- Not listed + future launch → ✅ OK
- Not listed + active intent + stock > 0 + past launch → 🔴 Action
""")

# ── Main area: landing / results ──────────────────────────────
if not all_ready:
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("#### Files required")
        for label, uploaded in [
            ("Zecom Tracker (.xlsx)", zecom_f), ("Lazada MP (.xlsx)", laz_f),
            ("Shopee MP (.xlsx)", sho_f), ("Zalora MP (.xlsx)", zal_f),
            ("Inventory (.csv/.xlsx)", inv_f)
        ]:
            icon = "✅" if uploaded else "⬜"
            st.markdown(f"{icon} {label}")
    with c2:
        st.markdown("#### Output sheets")
        for s in ["Summary", "Lazada – All EANs", "Lazada – Action Required",
                  "Lazada – Not Listed", "Lazada – Listed Inactive",
                  "(same 4 for Shopee & Zalora)", "All Marketplaces"]:
            st.markdown(f"- {s}")
    with c3:
        st.markdown("#### Comment logic")
        st.markdown("""
- ✅ **Active – OK**: Live, stocked, past launch
- ⚠️ **Active – 0 Stock**: Live but no inventory
- ⚠️ **Active – Future Launch**: Live before launch date
- ℹ️ **Not Listed – 0 Stock**: Expected, restock to list
- ℹ️ **Not Listed – Future Launch**: Will launch on date
- ℹ️ **Not Listed – Intended Inactive**: Zecom says inactive
- 🔴 **Not Listed – Action Required**: Gap — needs listing
- ⚠️ **Listed Inactive – No reason**: Review needed
        """)
    st.stop()

# ── Run & display results ──────────────────────────────────────
if run_btn:
    st.session_state.pop("df", None)   # clear previous cache
    with st.spinner("Processing… (may take ~30 s for large files)"):
        try:
            df = run_analysis(
                zecom_f.read(), laz_f.read(), sho_f.read(),
                zal_f.read(), inv_f.read(), inv_f.name,
            )
            st.session_state["df"] = df
        except ValueError as exc:
            if "PRECISION_LOSS" in str(exc):
                st.error("⚠️ Inventory file: EAN precision lost in CSV export")
                st.warning(
                    "Your inventory CSV stores EANs in scientific notation "
                    "(e.g. `4.0637E+12` instead of `4063699714586`). "
                    "The last digits are permanently lost in the CSV — EAN-level "
                    "stock matching is not possible from this file.\n\n"
                    "**Fix:** Open your inventory in Excel → **File → Save As → "
                    "Excel Workbook (.xlsx)** → upload the `.xlsx` file instead."
                )
            else:
                st.error(f"Processing error: {exc}")
            st.stop()
        except Exception as exc:
            st.error(f"Processing error: {exc}")
            st.exception(exc)
            st.stop()

df = st.session_state.get("df")
if df is None:
    st.info("Upload all files and click ▶ Run Analysis.")
    st.stop()

# ── SUMMARY METRICS ────────────────────────────────────────────
st.subheader("📊 Summary")
MP_LIST = [("Lazada", "Lazada Status", "#1F4E79"),
           ("Shopee", "Shopee Status", "#C55A11"),
           ("Zalora", "Zalora Status", "#375623")]

col1, col2, col3 = st.columns(3)
for col_ui, (mp, sc, _) in zip([col1, col2, col3], MP_LIST):
    total  = len(df)
    active = int((df[sc] == "Active").sum())
    nl     = int((df[sc] == "Not Listed").sum())
    li     = int((df[sc] == "Listed but Inactive").sum())
    action = int(((df[sc] == "Not Listed") &
                  (df["Effective Stock"] > 0) &
                  (df["Future Launch"] == "No") &
                  (df[f"Intended - {mp}"].str.lower() != "inactive")).sum())
    pct    = f"{active/total*100:.1f}%" if total else "0%"
    with col_ui:
        st.markdown(f"**{mp}**")
        m1, m2 = st.columns(2)
        m1.metric("Active EANs",        f"{active:,}",  pct)
        m2.metric("Not Listed",         f"{nl:,}")
        m3, m4 = st.columns(2)
        m3.metric("Listed Inactive",    f"{li:,}")
        m4.metric("🔴 Action Needed",   f"{action:,}",  delta_color="inverse")
        st.divider()

# ── DETAILED TABLE ────────────────────────────────────────────
st.subheader("📋 Breakdown Table")
rows = []
for mp, sc, _ in MP_LIST:
    total  = len(df)
    active = int((df[sc] == "Active").sum())
    nl     = int((df[sc] == "Not Listed").sum())
    li     = int((df[sc] == "Listed but Inactive").sum())
    nl_0s  = int(((df[sc]=="Not Listed")&(df["Effective Stock"]==0)&(df["Future Launch"]=="No")).sum())
    nl_ft  = int(((df[sc]=="Not Listed")&(df["Future Launch"]=="Yes")).sum())
    nl_ina = int(((df[sc]=="Not Listed")&(df[f"Intended - {mp}"].str.lower()=="inactive")).sum())
    nl_act = int(((df[sc]=="Not Listed")&(df["Effective Stock"]>0)&(df["Future Launch"]=="No")
                  &(df[f"Intended - {mp}"].str.lower()!="inactive")).sum())
    rows.append({
        "Marketplace"                     : mp,
        "Total EANs"                      : total,
        "Active"                          : active,
        "Not Listed (Total)"              : nl,
        "  → 0 Stock (OK)"               : nl_0s,
        "  → Future Launch (OK)"         : nl_ft,
        "  → Intended Inactive (OK)"     : nl_ina,
        "  → 🔴 Action Required"         : nl_act,
        "Listed but Inactive"             : li,
    })
tbl = pd.DataFrame(rows).set_index("Marketplace")
st.dataframe(tbl, use_container_width=True)

# ── TABS PER MARKETPLACE ──────────────────────────────────────
st.subheader("🔍 EAN-level Preview")
tab_laz, tab_sho, tab_zal = st.tabs(["Lazada", "Shopee", "Zalora"])

PREVIEW_BASE = ["Article No", "EAN", "Launch Date", "Future Launch", "Effective Stock"]

for tab, (mp, sc, _) in zip([tab_laz, tab_sho, tab_zal], MP_LIST):
    with tab:
        sc_col = f"{mp} Status"
        cc_col = f"{mp} Comment"
        ic_col = f"Intended - {mp}"
        view_cols = PREVIEW_BASE + [ic_col, sc_col, cc_col]

        # Filter selector
        flt = st.radio(f"Show:", ["All", "🔴 Action Required", "Not Listed", "Active", "Listed but Inactive"],
                       horizontal=True, key=f"flt_{mp}")
        sub = df[view_cols].copy()
        if flt == "🔴 Action Required":
            sub = sub[(sub[sc_col]=="Not Listed") &
                      (sub["Effective Stock"]>0) &
                      (sub["Future Launch"]=="No") &
                      (sub[ic_col].str.lower()!="inactive")]
        elif flt != "All":
            status_map = {"Not Listed":"Not Listed","Active":"Active","Listed but Inactive":"Listed but Inactive"}
            sub = sub[sub[sc_col] == status_map.get(flt, flt)]

        st.caption(f"{len(sub):,} rows")
        st.dataframe(sub.head(500), use_container_width=True, height=340)

# ── DOWNLOAD ──────────────────────────────────────────────────
st.subheader("⬇️ Download Full Report")
with st.spinner("Building Excel…"):
    xlsx_bytes = build_excel(df)

fname = f"Marketplace_Audit_{pd.Timestamp('today').strftime('%Y%m%d')}.xlsx"
st.download_button(
    label            = "📥 Download Excel Report",
    data             = xlsx_bytes,
    file_name        = fname,
    mime             = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type             = "primary",
    use_container_width = True,
)
st.caption(
    f"Report contains 14 sheets: Summary · Lazada/Shopee/Zalora (All EANs, Action Required, "
    f"Not Listed, Listed Inactive) · All Marketplaces  —  {len(df):,} EANs total"
)
