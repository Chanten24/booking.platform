import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
import base64
from io import BytesIO
from datetime import date
from typing import Optional, Tuple, List, Dict

# DOCX (python-docx)
try:
    from docx import Document
    from docx.shared import Pt
    from docx.oxml.ns import qn
except Exception:
    Document = None

# ------------------------------------------------------------
# Config
# ------------------------------------------------------------
st.set_page_config(page_title="Booking Platform", layout="wide")

BASE_DIR = Path(__file__).parent
ASSETS_DIR = BASE_DIR / "assets"
DATA_DIR = BASE_DIR / "data"
DOOH_MASTER_PATH = DATA_DIR / "dooh_master.csv"

# ------------------------------------------------------------
# Helpers: assets + styling
# ------------------------------------------------------------
def find_asset(stem: str) -> Optional[Path]:
    """Find an asset by stem in /assets (with or without extension)."""
    if not ASSETS_DIR.exists():
        return None

    exact = ASSETS_DIR / stem
    if exact.exists() and exact.is_file():
        return exact

    matches = list(ASSETS_DIR.glob(f"{stem}.*"))
    return matches[0] if matches else None


def sniff_mime_type(path: Path) -> str:
    header = path.read_bytes()[:16]
    if header.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if header.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if header.startswith(b"RIFF") and b"WEBP" in header:
        return "image/webp"
    return "image/png"


def file_to_base64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("utf-8")


def inject_background(background_path: Path, white_overlay_opacity: float = 0.45):
    """Full-page background image + white overlay."""
    b64 = file_to_base64(background_path)
    mime = sniff_mime_type(background_path)

    st.markdown(
        f"""
        <style>
          .stApp {{
            background-image: url("data:{mime};base64,{b64}");
            background-size: cover;
            background-position: center top;
            background-repeat: no-repeat;
            background-attachment: fixed;
          }}

          /* white overlay */
          .stApp::before {{
            content: "";
            position: fixed;
            inset: 0;
            background: rgba(255,255,255,{white_overlay_opacity});
            pointer-events: none;
            z-index: 0;
          }}

          /* keep content above overlay */
          section[data-testid="stSidebar"], .main, header {{
            position: relative;
            z-index: 2;
          }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def inject_app_css():
    st.markdown(
        """
        <style>
          /* Brand accent colour for radio/checkbox */
          input[type="radio"], input[type="checkbox"] {
            accent-color: #0B2A4A !important;
          }

          @keyframes fadeInUp {
            from { opacity: 0; transform: translateY(10px); }
            to   { opacity: 1; transform: translateY(0); }
          }
          .fade-in { animation: fadeInUp 520ms ease-out both; }

          /* Sticky header (logo must NOT be clipped) */
          .sticky-wrap {
            position: sticky;
            top: 0;
            z-index: 99999;
            background: rgba(255,255,255,0.97);
            backdrop-filter: blur(6px);
            padding-top: 10px;
          }

          .logo-bar {
            width: 100%;
            padding: 12px 0 4px 0;
            margin: 0;
          }
          .logo-inner {
            display: flex;
            justify-content: center;
            align-items: center;
            overflow: visible;
          }
          .logo-inner img {
            height: 84px;        /* ensure full visibility */
            width: auto;
            display: block;
          }
          .logo-divider {
            height: 1px;
            width: 100%;
            background: rgba(15, 23, 42, 0.08);
            margin-top: 10px;
          }

          .title-wrap {
            text-align: center;
            padding: 12px 0 16px 0;
          }
          .app-title {
            color: #0B2A4A;
            font-size: 24px;
            font-weight: 800;
            margin: 0;
            line-height: 1.2;
          }
          .app-caption {
            color: #475569;
            font-size: 13px;
            margin-top: 6px;
          }

          /* Do NOT push content down (sticky header sits in flow) */
          .block-container {
            padding-top: 1rem;
          }

          div[data-testid="stDataFrame"] { border-radius: 10px; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header(logo_path: Optional[Path]):
    logo_html = ""
    if logo_path and logo_path.exists():
        b64 = file_to_base64(logo_path)
        mime = sniff_mime_type(logo_path)
        logo_html = f'<img src="data:{mime};base64,{b64}" alt="Vicinity Logo" />'

    st.markdown(
        f"""
        <div class="sticky-wrap fade-in">
          <div class="logo-bar">
            <div class="logo-inner">{logo_html}</div>
            <div class="logo-divider"></div>
          </div>

          <div class="title-wrap">
            <div class="app-title">Booking Platform</div>
            <div class="app-caption">Proximity, budgeting, and automated Insertion Order generation</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ------------------------------------------------------------
# Helpers: column detection
# ------------------------------------------------------------
def pick_lat_lon(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    cols = [c.strip() for c in df.columns]
    lower = {c: c.lower() for c in cols}

    lat_candidates = [c for c in cols if "lat" in lower[c]]
    lon_candidates = [c for c in cols if ("lon" in lower[c]) or ("lng" in lower[c]) or ("long" in lower[c])]

    if not lat_candidates or not lon_candidates:
        return None, None
    return lat_candidates[0], lon_candidates[0]


def pick_name_col(df: pd.DataFrame) -> Optional[str]:
    priority = [
        "branch_name", "branch", "store_name", "store", "name", "location",
        "site_name", "site", "title", "outlet", "outlet_name", "id"
    ]
    low_map = {c.lower(): c for c in df.columns}
    for p in priority:
        if p in low_map:
            return low_map[p]

    for c in df.columns:
        cl = c.lower()
        if ("lat" in cl) or ("lon" in cl) or ("lng" in cl) or ("long" in cl):
            continue
        return c
    return None


def pick_first_existing(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    low_map = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        key = cand.lower().strip()
        if key in low_map:
            return low_map[key]
    return None

# ------------------------------------------------------------
# Distance
# ------------------------------------------------------------
def haversine_vec(lat1, lon1, lat2_arr, lon2_arr):
    R = 6371.0
    lat1, lon1 = np.radians(lat1), np.radians(lon1)
    lat2 = np.radians(lat2_arr)
    lon2 = np.radians(lon2_arr)

    dlat = lat2 - lat1
    dlon = lon2 - lon1

    a = np.sin(dlat / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2) ** 2
    return 2 * R * np.arcsin(np.sqrt(a))

# ------------------------------------------------------------
# Formatting helpers
# ------------------------------------------------------------
def fmt_currency_rands(x: float) -> str:
    try:
        return f"R {x:,.2f}"
    except Exception:
        return "R 0.00"


def fmt_int(x: float) -> str:
    try:
        return f"{int(round(x)):,}"
    except Exception:
        return "0"

# ------------------------------------------------------------
# Selection / counting logic
# ------------------------------------------------------------
def load_selected_sites_from_upload(file) -> pd.DataFrame:
    """
    DOOH Selection format (expected):
      - Site ID/Number (or similar)
      - Selected (1/0 or TRUE/FALSE)
    """
    df = pd.read_csv(file)
    site_col = pick_first_existing(df, ["Site ID/Number", "Site ID", "SiteID", "SiteID/Number", "Site_Number", "Site Number"])
    sel_col = pick_first_existing(df, ["Selected", "selected", "SELECTED"])

    if not site_col or not sel_col:
        raise ValueError("Selection file must contain 'Site ID/Number' and 'Selected' columns.")

    df[sel_col] = df[sel_col].astype(str).str.strip().str.lower()
    df["_selected_flag"] = df[sel_col].isin(["1", "true", "yes", "y"])
    selected = df[df["_selected_flag"]].copy()
    selected = selected[[site_col]].rename(columns={site_col: "Site ID/Number"})
    selected["Site ID/Number"] = selected["Site ID/Number"].astype(str).str.strip()
    selected = selected.dropna().drop_duplicates()
    return selected


def count_mobile_locations_any_format(file) -> int:
    """
    Supports:
    - DOOH selection-like format (Selected column), OR
    - ABSA-style location list: count rows with valid lat/lon
    """
    df = pd.read_csv(file)

    sel_col = pick_first_existing(df, ["Selected", "selected", "SELECTED"])
    site_col = pick_first_existing(df, ["Site ID/Number", "Site ID", "SiteID", "Site Number"])

    if sel_col and site_col:
        # treat like selection file
        tmp = df.copy()
        tmp[sel_col] = tmp[sel_col].astype(str).str.strip().str.lower()
        tmp["_selected_flag"] = tmp[sel_col].isin(["1", "true", "yes", "y"])
        return int(tmp["_selected_flag"].sum())

    # otherwise count rows with valid lat/lon
    lat, lon = pick_lat_lon(df)
    if not lat or not lon:
        # fallback: count all rows
        return int(len(df))

    df[lat] = pd.to_numeric(df[lat], errors="coerce")
    df[lon] = pd.to_numeric(df[lon], errors="coerce")
    df = df.dropna(subset=[lat, lon])
    return int(len(df))


def dooh_cpm_from_count(n_locations: int) -> int:
    return 240 if n_locations >= 50 else 260


def mobile_cpm_from_count(n_locations: int) -> int:
    return 140 if n_locations >= 50 else 160


def impressions_from_budget_and_cpm(budget: float, cpm: float) -> float:
    if cpm <= 0:
        return 0
    return (budget / cpm) * 1000.0

# ------------------------------------------------------------
# DOCX generation helpers
# ------------------------------------------------------------
def find_default_io_template() -> Optional[Path]:
    """
    Uses any file in /data matching *Template_Insertion Order*.docx
    Else first .docx in /data.
    """
    if not DATA_DIR.exists():
        return None

    candidates = sorted(DATA_DIR.glob("*Template_Insertion Order*.docx"))
    if candidates:
        return candidates[0]

    any_docx = sorted(DATA_DIR.glob("*.docx"))
    return any_docx[0] if any_docx else None


def set_run_font(run, pt_size: int = 9, font_name: Optional[str] = None):
    try:
        run.font.size = Pt(pt_size)
        if font_name:
            run.font.name = font_name
            # ensure font applies in Word
            rFonts = run._element.rPr.rFonts
            rFonts.set(qn("w:ascii"), font_name)
            rFonts.set(qn("w:hAnsi"), font_name)
    except Exception:
        pass


def write_value_in_paragraph_if_label(paragraph, label: str, value: str, pt_size: int = 9) -> bool:
    txt = (paragraph.text or "").strip()
    if not txt.lower().startswith(label.lower()):
        return False

    paragraph.text = label + " "
    run = paragraph.add_run(value)
    set_run_font(run, pt_size)
    return True


def fill_media_buy_total_cell(doc: "Document", media_buy_total: float) -> bool:
    target_label = "media buy total"
    amount_text = fmt_currency_rands(media_buy_total)

    for table in doc.tables:
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                cell_txt = (cell.text or "").strip().lower()
                if target_label in cell_txt:
                    if i + 1 < len(row.cells):
                        row.cells[i + 1].text = amount_text
                        for p in row.cells[i + 1].paragraphs:
                            for r in p.runs:
                                set_run_font(r, 9)
                        return True
    return False


def fill_top_line_numbers(doc: "Document", total_budget: float, total_impressions: int, note: str = "CPM: Mixed"):
    for p in doc.paragraphs:
        t = (p.text or "").strip().lower()
        if "media buy total" in t and "cpm" in t and "impressions" in t:
            p.text = f"MEDIA BUY TOTAL: {fmt_currency_rands(total_budget)} | {note} | IMPRESSIONS: {fmt_int(total_impressions)}"
            for r in p.runs:
                set_run_font(r, 10)
            break


def fill_media_buy_rows(doc: "Document", line_items: List[Dict]) -> bool:
    def normalize(s: str) -> str:
        return (s or "").strip().lower()

    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            headers = [normalize(c.text) for c in row.cells]
            if ("product" in headers) and ("rate" in headers) and ("quantity" in headers):
                def col_idx(name_exact: str, contains: Optional[str] = None):
                    for i, h in enumerate(headers):
                        if h == name_exact:
                            return i
                    if contains:
                        for i, h in enumerate(headers):
                            if contains in h:
                                return i
                    return None

                i_product = col_idx("product")
                i_desc = col_idx("description")
                i_start = col_idx("start date")
                i_end = col_idx("end date")
                i_publisher = col_idx("publisher")
                i_targeting = col_idx("targeting")
                i_rate = col_idx("rate")
                i_metric = col_idx("metric")
                i_qty = col_idx("quantity")
                i_gross = col_idx("gross rate", contains="gross")
                i_net = col_idx("net rate(excl. agency comm)", contains="net rate")

                first_data_idx = r_idx + 1
                if first_data_idx >= len(table.rows):
                    return False

                needed = len(line_items)
                existing = len(table.rows) - first_data_idx
                while existing < needed:
                    table.add_row()
                    existing += 1

                for k, item in enumerate(line_items):
                    drow = table.rows[first_data_idx + k]

                    def set_cell(ci, val):
                        if ci is None:
                            return
                        drow.cells[ci].text = val
                        for p in drow.cells[ci].paragraphs:
                            for rr in p.runs:
                                set_run_font(rr, 9)

                    set_cell(i_product, item.get("product", ""))
                    set_cell(i_desc, item.get("description", ""))
                    set_cell(i_start, item.get("start_date", ""))
                    set_cell(i_end, item.get("end_date", ""))
                    set_cell(i_publisher, item.get("publisher", ""))
                    set_cell(i_targeting, item.get("targeting", ""))
                    set_cell(i_rate, item.get("rate", ""))
                    set_cell(i_metric, item.get("metric", "Impressions"))
                    set_cell(i_qty, item.get("quantity", ""))
                    set_cell(i_gross, item.get("gross_rate", ""))
                    if i_net is not None and item.get("net_rate"):
                        set_cell(i_net, item.get("net_rate", ""))

                return True

    return False


def fill_sales_contact_block(doc: "Document", sales_name: str, sales_email: str) -> None:
    """
    Find the Sales Contact box/table and update Name/Email inside that area.
    This avoids clashing with Billing 'Contact Name:'.
    """
    if not (sales_name or sales_email):
        return

    for table in doc.tables:
        table_text = " ".join([(c.text or "") for row in table.rows for c in row.cells]).lower()
        if "sales contact" in table_text:
            # update within this table only
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        t = (p.text or "").strip()
                        if sales_name and t.lower().startswith("name:"):
                            p.text = "Name: "
                            r = p.add_run(sales_name)
                            set_run_font(r, 9)
                        if sales_email and t.lower().startswith("email:"):
                            p.text = "Email: "
                            r = p.add_run(sales_email)
                            set_run_font(r, 9)
            return


def insert_signature_above_signature_line(doc: "Document", signer_name: str, signer_title: str) -> None:
    """
    Inserts typed signature ABOVE the "Customer Authorized Signature:" line,
    using a script-like font where possible.
    """
    if not signer_name:
        return

    script_font = "Brush Script MT"  # fallback to whatever Word has if missing

    # Search paragraphs in all table cells too
    def process_paragraph_list(paragraphs):
        for p in paragraphs:
            text = (p.text or "").lower()
            if "customer authorized signature" in text:
                # insert BEFORE this paragraph
                parent = p._p.getparent()
                idx = parent.index(p._p)

                # Name line
                new_p = p._p.__class__()  # empty paragraph element
                parent.insert(idx, new_p)
                np = p._parent.add_paragraph()  # creates at end; we'll move content instead
                np._p.getparent().remove(np._p)  # remove from end
                np._p = new_p  # re-bind
                run = np.add_run(signer_name)
                set_run_font(run, 18, font_name=script_font)

                # Title line (optional)
                if signer_title:
                    new_p2 = p._p.__class__()
                    parent.insert(idx + 1, new_p2)
                    np2 = p._parent.add_paragraph()
                    np2._p.getparent().remove(np2._p)
                    np2._p = new_p2
                    run2 = np2.add_run(signer_title)
                    set_run_font(run2, 9)
                return True
        return False

    # Try doc paragraphs first
    if process_paragraph_list(doc.paragraphs):
        return

    # Then tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if process_paragraph_list(cell.paragraphs):
                    return


def generate_io_docx_bytes(
    template_path: Path,
    # customer/campaign blocks
    advertiser_name: str,
    advertiser_contact: str,
    agency_name: str,
    agency_contact: str,
    campaign_name: str,
    customer_ref: str,
    campaign_date: str,
    # billing block
    billing_customer_name: str,
    billing_contact_name: str,
    billing_address: str,
    billing_phone: str,
    billing_email: str,
    # sales contact
    sales_contact_name: str,
    sales_contact_email: str,
    # signature (typed)
    signer_name: str,
    signer_title: str,
    # line items + totals
    line_items: List[Dict],
    total_budget: float,
    total_impressions: int,
) -> bytes:
    doc = Document(str(template_path))

    # 1) Fill labels (Customer / Campaign / Billing)
    label_map = [
        ("Advertiser Name:", advertiser_name),
        ("Advertiser Contact:", advertiser_contact),
        ("Agency Name:", agency_name),
        ("Agency Contact:", agency_contact),

        ("Campaign Name:", campaign_name),
        ("Customer Reference Number:", customer_ref),

        # Customer Job Number should be BLANK by design (do not fill)
        ("Customer Job Number:", ""),  # ensures template doesn't keep old value if present

        ("Date:", campaign_date),

        # Billing
        ("Customer Name:", billing_customer_name),
        ("Contact Name:", billing_contact_name),
        ("Address:", billing_address),
        ("Phone:", billing_phone),
        ("Email:", billing_email),
    ]

    def apply_to_paragraphs(paragraphs):
        for p in paragraphs:
            for lbl, val in label_map:
                # allow blanking "Customer Job Number:"
                if lbl.lower().startswith("customer job number"):
                    if write_value_in_paragraph_if_label(p, lbl, "", pt_size=9):
                        break
                    continue

                if val:
                    if write_value_in_paragraph_if_label(p, lbl, val, pt_size=9):
                        break

    apply_to_paragraphs(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                apply_to_paragraphs(cell.paragraphs)

    # 2) Fill Sales contact block precisely
    fill_sales_contact_block(doc, sales_contact_name, sales_contact_email)

    # 3) Insert signature ABOVE signature line
    insert_signature_above_signature_line(doc, signer_name, signer_title)

    # 4) Fill Media Buy Total summary cell
    fill_media_buy_total_cell(doc, total_budget)

    # 5) Fill Media Buy Specifications rows
    fill_media_buy_rows(doc, line_items)

    # 6) Top line (Mixed CPM for combo)
    fill_top_line_numbers(doc, total_budget, total_impressions, note="CPM: Mixed")

    out = BytesIO()
    doc.save(out)
    return out.getvalue()

# ------------------------------------------------------------
# Background + header
# ------------------------------------------------------------
bg_path = find_asset("background")
if bg_path:
    inject_background(bg_path, white_overlay_opacity=0.45)

inject_app_css()

# IMPORTANT: logo file should be assets/logo.png
logo_path = find_asset("logo")
render_header(logo_path)

# ------------------------------------------------------------
# Tabs
# ------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["Proximity", "Budget & Selection", "Insertion Order (DOCX)"])

# ============================================================
# TAB 1: Proximity
# ============================================================
with tab1:
    st.subheader("Analysis type")
    mode = st.radio(
        "Select what you want to calculate",
        ["Store to DOOH", "Store to Store"],
        horizontal=True,
        key="mode_radio"
    )

    st.subheader("Upload files")

    if mode == "Store to DOOH":
        stores_file = st.file_uploader("Upload Stores (CSV)", type="csv", key="stores_dooh")

        dooh_df = None
        if DOOH_MASTER_PATH.exists():
            try:
                dooh_df = pd.read_csv(DOOH_MASTER_PATH)
                st.caption(f"Using DOOH master: {DOOH_MASTER_PATH}")
            except Exception as e:
                st.warning(f"Found {DOOH_MASTER_PATH} but couldn't read it: {e}")
                dooh_df = None

        if dooh_df is None:
            st.info("DOOH master not found. Temporary option: upload a DOOH CSV for this run.")
            dooh_file = st.file_uploader("Upload DOOH (CSV) [temporary]", type="csv", key="dooh_temp")
            if dooh_file:
                dooh_df = pd.read_csv(dooh_file)

    else:
        colA, colB = st.columns(2)
        with colA:
            stores_a_file = st.file_uploader("Upload Store List A (CSV)", type="csv", key="stores_a")
        with colB:
            stores_b_file = st.file_uploader("Upload Store List B (CSV)", type="csv", key="stores_b")

    st.subheader("Settings")
    radius = st.slider("Radius (km)", 1, 50, 10, key="radius_slider")

    show_summary = st.checkbox("Show summary table", value=True, key="show_summary")
    include_pairwise = st.checkbox("Generate pairwise rows (within radius)", value=True, key="include_pairwise")

    run = st.button("Run proximity", key="run_prox")

    if run:
        if mode == "Store to DOOH":
            if stores_file is None:
                st.error("Please upload the Stores CSV first.")
                st.stop()
            if dooh_df is None:
                st.error("DOOH data not available. Add data/dooh_master.csv or upload a temporary DOOH CSV above.")
                st.stop()

            stores = pd.read_csv(stores_file)
            targets = dooh_df.copy()
            target_label = "DOOH"

        else:
            if (stores_a_file is None) or (stores_b_file is None):
                st.error("Please upload BOTH Store List A and Store List B.")
                st.stop()

            stores = pd.read_csv(stores_a_file)
            targets = pd.read_csv(stores_b_file)
            target_label = "Store"

        store_lat, store_lon = pick_lat_lon(stores)
        store_name_col = pick_name_col(stores)
        if store_lat is None or store_lon is None:
            st.error("Could not detect latitude/longitude in Stores. Ensure columns include 'lat' and 'lon' (or 'lng').")
            st.stop()

        tgt_lat, tgt_lon = pick_lat_lon(targets)
        tgt_name_col = pick_name_col(targets)
        if tgt_lat is None or tgt_lon is None:
            st.error(f"Could not detect latitude/longitude in {target_label} dataset.")
            st.stop()

        # DOOH identifiers (optional)
        dooh_site_id_col = None
        dooh_network_col = None
        if mode == "Store to DOOH":
            dooh_site_id_col = pick_first_existing(
                targets,
                ["Site ID/Number", "Site ID", "SiteID", "Site Number", "Site_Number", "SiteID/Number", "SiteId", "Site"]
            )
            dooh_network_col = pick_first_existing(
                targets,
                ["Network:", "Network", "Network Type", "Network_Type", "NetworkType"]
            )

        # Clean coords
        stores[store_lat] = pd.to_numeric(stores[store_lat], errors="coerce")
        stores[store_lon] = pd.to_numeric(stores[store_lon], errors="coerce")
        targets[tgt_lat] = pd.to_numeric(targets[tgt_lat], errors="coerce")
        targets[tgt_lon] = pd.to_numeric(targets[tgt_lon], errors="coerce")

        stores = stores.dropna(subset=[store_lat, store_lon]).reset_index(drop=True)
        targets = targets.dropna(subset=[tgt_lat, tgt_lon]).reset_index(drop=True)

        if stores.empty or targets.empty:
            st.error("After cleaning coordinates, one of the datasets has no valid lat/lon rows.")
            st.stop()

        tgt_lats = targets[tgt_lat].to_numpy()
        tgt_lons = targets[tgt_lon].to_numpy()

        summary_rows = []
        pairwise_rows = []

        for i, s in stores.iterrows():
            s_name = str(s[store_name_col]) if (store_name_col and store_name_col in stores.columns and pd.notna(s[store_name_col])) else f"Location {i}"

            dists = haversine_vec(s[store_lat], s[store_lon], tgt_lats, tgt_lons)
            within_mask = dists <= radius
            idxs = np.where(within_mask)[0]

            nearest_idx = int(np.nanargmin(dists)) if len(dists) else None
            nearest_dist = float(dists[nearest_idx]) if nearest_idx is not None else np.nan

            if show_summary:
                row = {
                    "Location": s_name,
                    f"{target_label} sites within {radius}km": int(len(idxs)),
                    f"Nearest {target_label} distance (km)": round(nearest_dist, 3) if np.isfinite(nearest_dist) else None,
                }

                if mode == "Store to DOOH":
                    row["Nearest DOOH Site ID/Number"] = (
                        str(targets.iloc[nearest_idx][dooh_site_id_col])
                        if (nearest_idx is not None and dooh_site_id_col and pd.notna(targets.iloc[nearest_idx][dooh_site_id_col]))
                        else ""
                    )
                    row["Nearest DOOH Network"] = (
                        str(targets.iloc[nearest_idx][dooh_network_col])
                        if (nearest_idx is not None and dooh_network_col and pd.notna(targets.iloc[nearest_idx][dooh_network_col]))
                        else ""
                    )
                else:
                    row["Nearest Store"] = (
                        str(targets.iloc[nearest_idx][tgt_name_col])
                        if (nearest_idx is not None and tgt_name_col and tgt_name_col in targets.columns and pd.notna(targets.iloc[nearest_idx][tgt_name_col]))
                        else ""
                    )
                summary_rows.append(row)

            if include_pairwise:
                for j in idxs:
                    pr = {
                        "Location": s_name,
                        "Distance (km)": round(float(dists[j]), 3),
                    }

                    if mode == "Store to DOOH":
                        pr["Site ID/Number"] = (
                            str(targets.iloc[j][dooh_site_id_col])
                            if (dooh_site_id_col and pd.notna(targets.iloc[j][dooh_site_id_col]))
                            else ""
                        )
                        pr["Network"] = (
                            str(targets.iloc[j][dooh_network_col])
                            if (dooh_network_col and pd.notna(targets.iloc[j][dooh_network_col]))
                            else ""
                        )
                    else:
                        pr["Store"] = (
                            str(targets.iloc[j][tgt_name_col])
                            if (tgt_name_col and tgt_name_col in targets.columns and pd.notna(targets.iloc[j][tgt_name_col]))
                            else f"Store {j}"
                        )

                    pairwise_rows.append(pr)

        if show_summary:
            summary_df = pd.DataFrame(summary_rows)
            st.subheader("Results (summary)")
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

            st.download_button(
                "Download summary (CSV)",
                summary_df.to_csv(index=False).encode("utf-8"),
                file_name="proximity_summary.csv",
                mime="text/csv",
            )

        if include_pairwise:
            pairwise_df = pd.DataFrame(pairwise_rows)
            st.subheader("Results (pairwise within radius)")
            st.dataframe(pairwise_df, use_container_width=True, hide_index=True)

            st.download_button(
                "Download pairwise (CSV)",
                pairwise_df.to_csv(index=False).encode("utf-8"),
                file_name="location_proximity_results.csv",
                mime="text/csv",
            )

        if (not show_summary) and (not include_pairwise):
            st.info("Select at least one output option (summary and/or pairwise).")

# ============================================================
# TAB 2: Budget & Selection
# ============================================================
with tab2:
    st.subheader("Budget, selection, CPM & impressions")

    campaign_type = st.radio(
        "Campaign type",
        ["DOOH", "Mobile", "DOOH + Mobile"],
        horizontal=True,
        key="campaign_type"
    )

    # Defaults
    dooh_selected_count = int(st.session_state.get("selected_sites_count_dooh", 0))
    mobile_selected_count = int(st.session_state.get("selected_sites_count_mobile", 0))

    dooh_budget = float(st.session_state.get("budget_dooh", 200000.0))
    mobile_budget = float(st.session_state.get("budget_mobile", 150000.0))

    if campaign_type == "DOOH":
        st.caption("DOOH selection upload must include: 'Site ID/Number' and 'Selected' (1/0).")
        selection_file_dooh = st.file_uploader("Upload DOOH selected sites (CSV)", type="csv", key="selection_upload_dooh")
        dooh_budget = st.number_input("DOOH budget (R)", min_value=0.0, value=float(dooh_budget), step=1000.0, key="budget_dooh_input")

        if selection_file_dooh:
            try:
                _df = load_selected_sites_from_upload(selection_file_dooh)
                dooh_selected_count = int(len(_df))
                st.success(f"Loaded {dooh_selected_count} selected DOOH sites.")
            except Exception as e:
                st.error(str(e))
                dooh_selected_count = 0

        dooh_cpm = dooh_cpm_from_count(dooh_selected_count)
        dooh_imps = impressions_from_budget_and_cpm(dooh_budget, dooh_cpm)

        st.session_state["selected_sites_count_dooh"] = dooh_selected_count
        st.session_state["budget_dooh"] = float(dooh_budget)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Selected DOOH sites", dooh_selected_count)
        with c2:
            st.metric("DOOH CPM (R)", dooh_cpm)
        with c3:
            st.metric("Estimated DOOH impressions", fmt_int(dooh_imps))

    elif campaign_type == "Mobile":
        st.caption("Mobile upload supports ABSA-style location list OR selection format.")
        selection_file_mobile = st.file_uploader("Upload Mobile locations (CSV)", type="csv", key="selection_upload_mobile")
        mobile_budget = st.number_input("Mobile budget (R)", min_value=0.0, value=float(mobile_budget), step=1000.0, key="budget_mobile_input")

        if selection_file_mobile:
            try:
                mobile_selected_count = count_mobile_locations_any_format(selection_file_mobile)
                st.success(f"Loaded {mobile_selected_count} Mobile locations.")
            except Exception as e:
                st.error(str(e))
                mobile_selected_count = 0

        mob_cpm = mobile_cpm_from_count(mobile_selected_count)
        mob_imps = impressions_from_budget_and_cpm(mobile_budget, mob_cpm)

        st.session_state["selected_sites_count_mobile"] = mobile_selected_count
        st.session_state["budget_mobile"] = float(mobile_budget)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Mobile locations", mobile_selected_count)
        with c2:
            st.metric("Mobile CPM (R)", mob_cpm)
        with c3:
            st.metric("Estimated Mobile impressions", fmt_int(mob_imps))

    else:
        colA, colB = st.columns(2)
        with colA:
            st.caption("DOOH selection upload: 'Site ID/Number' + 'Selected'")
            selection_file_dooh = st.file_uploader("Upload DOOH selected sites (CSV)", type="csv", key="selection_upload_dooh_combo")
            dooh_budget = st.number_input("DOOH budget (R)", min_value=0.0, value=float(dooh_budget), step=1000.0, key="budget_dooh_input_combo")
        with colB:
            st.caption("Mobile upload supports ABSA-style list OR selection format.")
            selection_file_mobile = st.file_uploader("Upload Mobile locations (CSV)", type="csv", key="selection_upload_mobile_combo")
            mobile_budget = st.number_input("Mobile budget (R)", min_value=0.0, value=float(mobile_budget), step=1000.0, key="budget_mobile_input_combo")

        if selection_file_dooh:
            try:
                _df = load_selected_sites_from_upload(selection_file_dooh)
                dooh_selected_count = int(len(_df))
                st.success(f"Loaded {dooh_selected_count} selected DOOH sites.")
            except Exception as e:
                st.error(f"DOOH: {e}")
                dooh_selected_count = 0

        if selection_file_mobile:
            try:
                mobile_selected_count = count_mobile_locations_any_format(selection_file_mobile)
                st.success(f"Loaded {mobile_selected_count} Mobile locations.")
            except Exception as e:
                st.error(f"Mobile: {e}")
                mobile_selected_count = 0

        dooh_cpm = dooh_cpm_from_count(dooh_selected_count)
        mob_cpm = mobile_cpm_from_count(mobile_selected_count)

        dooh_imps = impressions_from_budget_and_cpm(dooh_budget, dooh_cpm)
        mob_imps = impressions_from_budget_and_cpm(mobile_budget, mob_cpm)

        total_budget = float(dooh_budget) + float(mobile_budget)
        total_imps = int(round(dooh_imps + mob_imps, 0))

        st.session_state["selected_sites_count_dooh"] = dooh_selected_count
        st.session_state["selected_sites_count_mobile"] = mobile_selected_count
        st.session_state["budget_dooh"] = float(dooh_budget)
        st.session_state["budget_mobile"] = float(mobile_budget)

        st.subheader("Summary")
        r1, r2, r3 = st.columns(3)
        with r1:
            st.metric("Total budget", fmt_currency_rands(total_budget))
        with r2:
            st.metric("Total impressions", fmt_int(total_imps))
        with r3:
            st.metric("CPM", "Mixed")

        st.divider()

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.metric("DOOH sites", dooh_selected_count)
        with c2:
            st.metric("DOOH CPM", f"R {dooh_cpm}")
        with c3:
            st.metric("DOOH imps", fmt_int(dooh_imps))
        with c4:
            st.metric("DOOH budget", fmt_currency_rands(dooh_budget))
        with c5:
            st.metric("Mobile CPM", f"R {mob_cpm}")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Mobile locations", mobile_selected_count)
        with c2:
            st.metric("Mobile imps", fmt_int(mob_imps))
        with c3:
            st.metric("Mobile budget", fmt_currency_rands(mobile_budget))
        with c4:
            st.metric("Mobile CPM (repeat)", f"R {mob_cpm}")

# ============================================================
# TAB 3: Insertion Order (DOCX)
# ============================================================
with tab3:
    st.subheader("Insertion Order (DOCX)")

    default_template = find_default_io_template()

    if Document is None:
        st.warning("DOCX generation needs python-docx. Install it with: pip install python-docx")
        st.stop()

    if not default_template:
        st.error("No default IO template found in /data. Please place your template .docx in the data folder.")
        st.stop()

    st.caption(f"Using template from /data: {default_template.name}")

    st.divider()
    st.subheader("Client + campaign details")

    col1, col2 = st.columns(2)
    with col1:
        advertiser_name = st.text_input("Advertiser Name", value="")
        advertiser_contact = st.text_input("Advertiser Contact", value="")
        agency_name = st.text_input("Agency Name", value="")
        agency_contact = st.text_input("Agency Contact", value="")
    with col2:
        campaign_name = st.text_input("Campaign Name", value="")
        customer_ref = st.text_input("Customer Reference Number", value="")
        # Customer Job Number must be blank by design -> do not show input
        campaign_date = st.text_input("Date (DD/MM/YYYY)", value=date.today().strftime("%d/%m/%Y"))

    st.subheader("Billing details")
    b1, b2 = st.columns(2)
    with b1:
        billing_customer_name = st.text_input("Billing - Customer Name", value="")
        billing_contact_name = st.text_input("Billing - Contact Name", value="")
        billing_phone = st.text_input("Billing - Phone", value="")
    with b2:
        billing_email = st.text_input("Billing - Email", value="")
        billing_address = st.text_area("Billing - Address", value="", height=80)

    st.subheader("Sales contact")
    s1, s2 = st.columns(2)
    with s1:
        sales_contact_name = st.text_input("Sales Contact - Name", value="")
    with s2:
        sales_contact_email = st.text_input("Sales Contact - Email", value="")

    st.subheader("Campaign start and end date")
    d1, d2 = st.columns(2)
    with d1:
        start_date = st.text_input("Start Date (DD/MM/YYYY)", value="")
    with d2:
        end_date = st.text_input("End Date (DD/MM/YYYY)", value="")

    st.subheader("Customer signature (typed)")
    sig1, sig2 = st.columns(2)
    with sig1:
        signer_name = st.text_input("Signed by (name)", value="")
    with sig2:
        signer_title = st.text_input("Title (optional)", value="")

    st.caption("Signature will be inserted above the 'Customer Authorized Signature' line in the template.")

    # Pull campaign type & budgets/counts from Tab 2
    campaign_type = st.session_state.get("campaign_type", "DOOH")
    dooh_budget = float(st.session_state.get("budget_dooh", 0.0))
    mobile_budget = float(st.session_state.get("budget_mobile", 0.0))
    dooh_count = int(st.session_state.get("selected_sites_count_dooh", 0))
    mobile_count = int(st.session_state.get("selected_sites_count_mobile", 0))

    # Build line items
    line_items = []
    total_budget = 0.0
    total_impressions = 0
    dooh_cpm = 0
    mob_cpm = 0

    if campaign_type == "DOOH":
        dooh_cpm = dooh_cpm_from_count(dooh_count)
        imps = int(round(impressions_from_budget_and_cpm(dooh_budget, dooh_cpm), 0))
        total_budget = dooh_budget
        total_impressions = imps

        line_items.append({
            "product": "DOOH",
            "description": "",
            "start_date": start_date.strip(),
            "end_date": end_date.strip(),
            "publisher": "",
            "targeting": "",
            "rate": f"R {dooh_cpm}",
            "metric": "Impressions",
            "quantity": fmt_int(imps),
            "gross_rate": fmt_currency_rands(dooh_budget),
        })

    elif campaign_type == "Mobile":
        mob_cpm = mobile_cpm_from_count(mobile_count)
        imps = int(round(impressions_from_budget_and_cpm(mobile_budget, mob_cpm), 0))
        total_budget = mobile_budget
        total_impressions = imps

        line_items.append({
            "product": "Mobile",
            "description": "",
            "start_date": start_date.strip(),
            "end_date": end_date.strip(),
            "publisher": "",
            "targeting": "",
            "rate": f"R {mob_cpm}",
            "metric": "Impressions",
            "quantity": fmt_int(imps),
            "gross_rate": fmt_currency_rands(mobile_budget),
        })

    else:
        dooh_cpm = dooh_cpm_from_count(dooh_count)
        mob_cpm = mobile_cpm_from_count(mobile_count)

        dooh_imps = int(round(impressions_from_budget_and_cpm(dooh_budget, dooh_cpm), 0))
        mob_imps = int(round(impressions_from_budget_and_cpm(mobile_budget, mob_cpm), 0))

        total_budget = dooh_budget + mobile_budget
        total_impressions = dooh_imps + mob_imps

        line_items.append({
            "product": "DOOH",
            "description": "",
            "start_date": start_date.strip(),
            "end_date": end_date.strip(),
            "publisher": "",
            "targeting": "",
            "rate": f"R {dooh_cpm}",
            "metric": "Impressions",
            "quantity": fmt_int(dooh_imps),
            "gross_rate": fmt_currency_rands(dooh_budget),
        })
        line_items.append({
            "product": "Mobile",
            "description": "",
            "start_date": start_date.strip(),
            "end_date": end_date.strip(),
            "publisher": "",
            "targeting": "",
            "rate": f"R {mob_cpm}",
            "metric": "Impressions",
            "quantity": fmt_int(mob_imps),
            "gross_rate": fmt_currency_rands(mobile_budget),
        })

    st.subheader("Auto-calculated (from Budget & Selection)")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("Media Buy Total", fmt_currency_rands(total_budget))
    with c2:
        st.metric("Campaign type", campaign_type)
    with c3:
        st.metric("Total impressions", fmt_int(total_impressions))
    with c4:
        st.metric("DOOH CPM", f"R {dooh_cpm}" if dooh_cpm else "—")
    with c5:
        st.metric("Mobile CPM", f"R {mob_cpm}" if mob_cpm else "—")

    gen = st.button("Generate IO (pre-populated DOCX)")

    if gen:
        try:
            doc_bytes = generate_io_docx_bytes(
                template_path=default_template,
                advertiser_name=advertiser_name.strip(),
                advertiser_contact=advertiser_contact.strip(),
                agency_name=agency_name.strip(),
                agency_contact=agency_contact.strip(),
                campaign_name=campaign_name.strip(),
                customer_ref=customer_ref.strip(),
                campaign_date=campaign_date.strip(),

                billing_customer_name=billing_customer_name.strip(),
                billing_contact_name=billing_contact_name.strip(),
                billing_address=billing_address.strip(),
                billing_phone=billing_phone.strip(),
                billing_email=billing_email.strip(),

                sales_contact_name=sales_contact_name.strip(),
                sales_contact_email=sales_contact_email.strip(),

                signer_name=signer_name.strip(),
                signer_title=signer_title.strip(),

                line_items=line_items,
                total_budget=total_budget,
                total_impressions=int(total_impressions),
            )

            st.success("IO generated.")
            st.download_button(
                "Download IO (DOCX)",
                data=doc_bytes,
                file_name="Insertion_Order_Prepopulated.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

            # Open Gmail compose (can’t auto-attach from Streamlit)
            subject = "Insertion Order – Pre-populated"
            body = (
                "Hi Sales Team,\n\n"
                "Please find the pre-populated Insertion Order attached.\n\n"
                f"Campaign type: {campaign_type}\n"
                f"Media Buy Total: {fmt_currency_rands(total_budget)}\n"
                f"Total Impressions: {fmt_int(total_impressions)}\n"
                f"DOOH CPM: {('R ' + str(dooh_cpm)) if dooh_cpm else '—'}\n"
                f"Mobile CPM: {('R ' + str(mob_cpm)) if mob_cpm else '—'}\n\n"
                "Regards,\n"
                f"{signer_name or ''}"
            )
            mailto = (
                "mailto:?"
                f"subject={subject.replace(' ', '%20')}"
                f"&body={body.replace(' ', '%20').replace('\\n', '%0A')}"
            )
            st.markdown(f"**Send to Sales (Gmail draft):** [Open email draft]({mailto})")
            st.caption("Tip: Download the IO first, then attach it to the Gmail draft.")

        except Exception as e:
            st.error(f"Could not generate IO: {e}")
