from __future__ import annotations

import json
import re
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

# ==========================================
# KONFIGURASI HALAMAN & KONSTANTA
# ==========================================
APP_TITLE = "Monitoring Kontrak Outlet"
DATA_FILE = Path("data") / "DRP & Outlet MCD - DSO Surabaya Barat.xlsx"
SHEET_NAME = "Monitoring Kontrak Outlet"
DRP_SHEET_NAME = "Pemesanan Layar DRP"

MATERIAL_COLUMNS = ["Panel Toko", "Showcase", "Layar Toko", "Shop Sign / Neon Box"]
SYNC_PANEL_SCREEN_COLUMNS = [
    "Ukuran Sticker Panel Toko",
    "Brand Panel",
    "Versi",
    "Ukuran Layar",
    "Brand Layar",
]
DATE_COLUMNS = [
    "Periode Kontrak Awal",
    "Periode Kontrak Akhir",
    "Tanggal Penggantian Sticker Panel",
]
STICKER_ORDER_COLUMNS = [
    "ID Outlet",
    "Nama Outlet",
    "Material",
    "Brand",
    "Version",
    "Jumlah Sticker / Vinyl",
    "Tanggal OrderPenggantian Sticker / Vinyl",
    "Nama Vendor",
]
COLUMNS = [
    "ID Outlet", "Nama Customer", "Alamat", "Zona", "Sektor", "Pos Budget",
    "Kategori Outlet", "Kategori Pekerjaan", "Brand", "Nominal Kontrak",
    "Periode Kontrak Awal", "Periode Kontrak Akhir", "No. Dokumen", *MATERIAL_COLUMNS,
    "Ukuran Panel Toko", "Ukuran Sticker Panel Toko", "Brand Panel", "Versi",
    "Tanggal Penggantian Sticker Panel", "Nama Vendor",
    "Ukuran Layar", "Brand Layar",
]
CONTENT_COLUMNS = [
    "ID Outlet", "Nama Customer", "Alamat", "Zona", "Sektor", "Pos Budget",
    "Kategori Outlet", "Kategori Pekerjaan", "Brand", "Nominal Kontrak",
    "Periode Kontrak Awal", "Periode Kontrak Akhir", "No. Dokumen",
    "Ukuran Panel Toko", "Ukuran Sticker Panel Toko", "Brand Panel", "Versi",
    "Tanggal Penggantian Sticker Panel", "Nama Vendor",
    "Ukuran Layar", "Brand Layar",
]
ALIASES = {
    "ID Outlet": ("ID Outlet", "ID Customer", "Customer ID", "Kode Outlet", "Outlet ID"),
    "Nama Customer": ("Nama Customer", "Customer Name", "Customer", "Nama Outlet", "Nama Toko", "Outlet"),
    "Alamat": ("Alamat", "Alamat Outlet", "Alamat Toko", "Address", "Outlet Address", "Customer Address"),
    "Zona": ("Zona", "Zone"),
    "Sektor": ("Sektor", "Sector", "Kecamatan"),
    "Pos Budget": ("Pos Budget", "Budget Position"),
    "Kategori Outlet": ("Kategori Outlet", "Customer Group Category", "Kategori"),
    "Kategori Pekerjaan": ("Kategori Pekerjaan", "Kategori Pekerjaan Outlet", "Work Category", "Job Category", "Kategori Kerja"),
    "Brand": ("Brand", "Finance Brand", "Merek"),
    "Nominal Kontrak": ("Nominal Kontrak", "Nominal", "Nominal.1", "Nominal SPKS", "Contract Amount", "Nilai Kontrak"),
    "Periode Kontrak Awal": ("Periode Kontrak Awal", "Periode Awal Kontrak", "Start Kontrak", "Mulai", "Tanggal Mulai", "Start Date"),
    "Periode Kontrak Akhir": ("Periode Kontrak Akhir", "Periode Akhir Kontrak", "End Kontrak", "Akhir", "Tanggal Akhir", "End Date", "Expiry Date"),
    "No. Dokumen": ("No. Dokumen", "No Dokumen", "Nomor Dokumen", "Document No", "No SPKS"),
    "Panel Toko": ("Panel Toko", "Materi Panel Toko"),
    "Showcase": ("Showcase", "Show Case", "Materi Showcase"),
    "Layar Toko": ("Layar Toko", "Layar", "Materi Layar Toko"),
    "Shop Sign / Neon Box": ("Shop Sign / Neon Box", "Shop Sign", "Neon Box", "Materi Shop Sign / Neon Box"),
    "Ukuran Panel Toko": ("Ukuran Panel Toko", "Size Panel Toko", "Panel"),
    "Ukuran Sticker Panel Toko": (
        "Ukuran Sticker Panel Toko",
        "Ukuran Sticker Panel",
        "Ukuran Sticker",
        "Ukuran Sticker / Vinyl",
        "Ukuran Sticker Panel (diluar Nama Toko)",
        "Ukuran Stiker Panel Toko",
        "Size Sticker Panel Toko",
        "Material",
    ),
    "Brand Panel": (
        "Brand Panel",
        "Brand Sticker",
        "Brand Planning - Sticker Panel",
        "Merek Panel",
    ),
    "Versi": (
        "Versi",
        "Version",
        "Versi Sticker",
        "Versi Panel",
    ),
    "Tanggal Penggantian Sticker Panel": (
        "Tanggal Penggantian Sticker Panel",
        "Tanggal OrderPenggantian Sticker / Vinyl",
        "Tanggal Order Penggantian Sticker / Vinyl",
        "Tanggal Penggantian Terakhir Sticker Panel",
        "Tgl Penggantian Sticker Panel",
        "Tgl Penggantian Terakhir",
        "Tanggal Ganti Sticker",
    ),
    "Nama Vendor": (
        "Nama Vendor",
        "Vendor",
        "Nama Vendor Sticker",
    ),
    "Ukuran Layar": ("Ukuran Layar", "Layar", "Materi Layar Toko", "Size Layar"),
    "Brand Layar": ("Brand Layar", "Brand Planning - Layar DRP", "Merek Layar"),
}

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==========================================
# CUSTOM CSS UNTUK TAMPILAN MODERN
# ==========================================
def inject_custom_css():
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
            max-width: 96%;
        }
        .hero-header {
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
            padding: 1.5rem 2rem;
            border-radius: 16px;
            color: white;
            margin-bottom: 1.5rem;
            box-shadow: 0 4px 20px rgba(15, 23, 42, 0.12);
        }
        .hero-header h1 {
            margin: 0;
            font-size: 1.75rem;
            font-weight: 700;
            color: #ffffff;
            letter-spacing: -0.02em;
        }
        .hero-header p {
            margin: 0.35rem 0 0 0;
            font-size: 0.95rem;
            color: #cbd5e1;
        }
        .kpi-card {
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 14px;
            padding: 1.15rem 1.35rem;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
            transition: transform 0.15s ease, box-shadow 0.15s ease;
            height: 100%;
        }
        .kpi-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(0, 0, 0, 0.08);
        }
        .kpi-label {
            font-size: 0.82rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            color: #64748b;
            margin-bottom: 0.35rem;
        }
        .kpi-value {
            font-size: 1.65rem;
            font-weight: 700;
            color: #0f172a;
            line-height: 1.2;
        }
        .kpi-sub {
            font-size: 0.78rem;
            color: #64748b;
            margin-top: 0.35rem;
        }
        .border-blue { border-left: 5px solid #2563eb; }
        .border-emerald { border-left: 5px solid #10b981; }
        .border-amber { border-left: 5px solid #f59e0b; }
        .border-rose { border-left: 5px solid #ef4444; }

        section[data-testid="stSidebar"] {
            background-color: #f8fafc;
            border-right: 1px solid #e2e8f0;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
            background-color: #f1f5f9;
            padding: 6px;
            border-radius: 12px;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 8px;
            padding: 8px 16px;
            font-weight: 600;
        }
        .stTabs [aria-selected="true"] {
            background-color: #ffffff !important;
            box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        }

        @media (prefers-color-scheme: dark) {
            .kpi-card {
                background-color: #1e293b;
                border-color: #334155;
            }
            .kpi-value { color: #f8fafc; }
            .kpi-label, .kpi-sub { color: #94a3b8; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==========================================
# FUNGSI RESET FILTER & STATE EDIT
# ==========================================
def reset_dashboard_filters() -> None:
    st.session_state["dash_search_query"] = ""
    st.session_state["dash_status"] = "Semua Status"
    st.session_state["dash_zone"] = "Semua"
    st.session_state["dash_sector"] = "Semua"
    st.session_state["dash_budget"] = "Semua"
    st.session_state["dash_job_cat"] = "Semua"
    st.session_state["dash_sticker_size"] = "Semua"
    st.session_state["dash_screen_size"] = "Semua"


def reset_multi_edit_state() -> None:
    st.session_state["multi_edit_quick_search"] = ""
    st.session_state["multi_edit_zone"] = "Semua Zona"
    st.session_state["multi_edit_status"] = "Semua Status"
    st.session_state["multi_select_all_state"] = False
    st.session_state["confirm_delete_multi_cb"] = False
    st.session_state.pop("multi_customer_data_editor", None)


# ==========================================
# FUNGSI PEMROSESAN DATA (BACKEND LOGIC)
# ==========================================
def clean_headers(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame


def parse_amount(value):
    if pd.isna(value) or str(value).strip() == "":
        return pd.NA
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace("Rp", "").replace("rp", "").replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text:
        tail = text.rsplit(",", 1)[-1]
        text = text.replace(",", ".") if len(tail) <= 2 else text.replace(",", "")
    elif text.count(".") > 1:
        text = text.replace(".", "")
    elif "." in text:
        tail = text.rsplit(".", 1)[-1]
        if len(tail) == 3:
            text = text.replace(".", "")
    try:
        return float(text)
    except ValueError:
        return pd.NA


def parse_tickmark(value) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    return str(value).strip().casefold() in {"✓", "✔", "☑", "x", "yes", "ya", "true", "1", "centang", "check"}


def parse_material(value, column: str) -> bool:
    if parse_tickmark(value):
        return True
    if pd.isna(value):
        return False
    text = str(value).strip().casefold()
    patterns = {
        "Panel Toko": r"\bpanel\s+toko\b",
        "Showcase": r"\b(?:display|show\s*case|showcase)\b",
        "Layar Toko": r"\blayar\b",
        "Shop Sign / Neon Box": r"\b(?:shop\s*sign|neon\s*box)\b",
    }
    return bool(re.search(patterns[column], text))


def extract_sizes(value):
    """
    Mengambil hanya ukuran dimensi (contoh: 'Vinyl Panel Toko 1 x 2 m (BL)' -> '1 x 2 m').
    """
    if pd.isna(value) or str(value).strip() == "":
        return pd.NA
    text = str(value).strip()
    if text.casefold() in {"<na>", "nan", "none"}:
        return pd.NA
    matches = re.findall(
        r"(?<!\d)(\d+(?:[.,]\d+)?)\s*[x×]\s*(\d+(?:[.,]\d+)?)\s*(?:m\b)?",
        text,
        flags=re.IGNORECASE,
    )
    if matches:
        sizes = dict.fromkeys(f"{width.strip()} x {height.strip()} m" for width, height in matches)
        return ", ".join(sizes)
    cleaned = re.sub(
        r"\b(?:sticker|stiker|vinyl|panel\s*toko|panel|toko|materi|material|ukuran|size|\(bl\))\b",
        "",
        text,
        flags=re.IGNORECASE,
    )
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" -:/,;()")
    return cleaned if cleaned else pd.NA


def customer_id_key(value) -> str:
    if pd.isna(value):
        return ""
    key = str(value).strip()
    if re.fullmatch(r"\d+\.0", key):
        key = key[:-2]
    return key.casefold()


def normalize_name_key(value) -> str:
    if pd.isna(value):
        return ""
    text = re.sub(r"\s+", " ", str(value).strip()).casefold()
    if text in {"", "<na>", "nan", "none"}:
        return ""
    return text


def normalize_data(frame: pd.DataFrame) -> pd.DataFrame:
    frame = clean_headers(frame)
    normalized = pd.DataFrame(index=frame.index)
    lowered = {str(column).strip().casefold(): column for column in frame.columns}
    for target, names in ALIASES.items():
        sources = [lowered[name.casefold()] for name in names if name.casefold() in lowered]
        if target == "Nama Customer" and sources:
            candidates = frame[sources].replace(r"^\s*$", pd.NA, regex=True)
            normalized[target] = candidates.bfill(axis=1).iloc[:, 0]
        else:
            normalized[target] = frame[sources[0]] if sources else pd.NA
    text_columns = [
        column
        for column in COLUMNS
        if column not in MATERIAL_COLUMNS
        and column not in ("Nominal Kontrak", *DATE_COLUMNS)
    ]
    for column in text_columns:
        normalized[column] = normalized[column].astype("string").str.strip()
        normalized[column] = normalized[column].replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA, "None": pd.NA})
    for column in ("Ukuran Panel Toko", "Ukuran Sticker Panel Toko", "Ukuran Layar"):
        normalized[column] = normalized[column].map(extract_sizes)
    normalized["Nominal Kontrak"] = normalized["Nominal Kontrak"].map(parse_amount)
    contract_material_source = lowered.get("materi kontrak")
    contract_material = frame[contract_material_source] if contract_material_source else pd.Series(pd.NA, index=frame.index)
    has_contract_material = contract_material.notna() & contract_material.astype("string").str.strip().ne("")
    for column in MATERIAL_COLUMNS:
        source_values = normalized[column].map(lambda value: parse_material(value, column))
        if column != "Shop Sign / Neon Box":
            patterns = {
                "Panel Toko": r"\bpanel\s+toko\b",
                "Showcase": r"\b(?:display|show\s*case|showcase)\b",
                "Layar Toko": r"\blayar\b",
            }
            from_contract = contract_material.astype("string").str.contains(
                patterns[column], case=False, regex=True, na=False
            )
            source_values = source_values.where(~has_contract_material, from_contract)
        elif not any(name.casefold() in lowered for name in ALIASES[column]):
            source_values = contract_material.astype("string").str.contains(
                r"\b(?:shop\s*sign|neon\s*box)\b", case=False, regex=True, na=False
            )
        normalized[column] = source_values.astype(bool)
    for column in DATE_COLUMNS:
        normalized[column] = pd.to_datetime(normalized[column], errors="coerce", dayfirst=True).dt.tz_localize(None)
    normalized = normalized[COLUMNS]
    existing_content_cols = [col for col in CONTENT_COLUMNS if col in normalized.columns]
    has_data = normalized[existing_content_cols].notna().any(axis=1)
    return normalized.loc[has_data.to_numpy(dtype=bool)].reset_index(drop=True)


def apply_sticker_replacement_orders(
    contracts: pd.DataFrame,
    orders_df: pd.DataFrame,
) -> tuple[pd.DataFrame, int, int, list[str]]:
    target = normalize_data(contracts).copy()
    id_keys = target["ID Outlet"].map(customer_id_key)
    name_keys = target["Nama Customer"].map(normalize_name_key)

    id_to_rows: dict[str, list[int]] = {}
    for idx, k in id_keys.items():
        if k:
            id_to_rows.setdefault(k, []).append(idx)

    name_to_rows: dict[str, list[int]] = {}
    for idx, k in name_keys.items():
        if k:
            name_to_rows.setdefault(k, []).append(idx)

    updated_rows_set: set[int] = set()
    created_count = 0
    unmatched_labels: list[str] = []

    for _, order in orders_df.iterrows():
        raw_id = order.get("ID Outlet", pd.NA)
        raw_name = order.get("Nama Outlet", pd.NA)
        raw_material = order.get("Material", order.get("Ukuran Sticker", pd.NA))
        raw_brand = order.get("Brand", order.get("Brand Sticker", pd.NA))
        raw_version = order.get("Version", order.get("Versi", pd.NA))
        raw_vendor = order.get("Nama Vendor", pd.NA)
        raw_date = order.get(
            "Tanggal OrderPenggantian Sticker / Vinyl",
            order.get("Tanggal Penggantian Sticker Panel", pd.NaT),
        )

        id_str = "" if pd.isna(raw_id) else str(raw_id).strip()
        name_str = "" if pd.isna(raw_name) else str(raw_name).strip()
        if not id_str and not name_str:
            continue

        size_val = extract_sizes(raw_material)
        brand_val = (
            str(raw_brand).strip()
            if pd.notna(raw_brand) and str(raw_brand).strip() not in {"", "<NA>", "nan", "None"}
            else pd.NA
        )
        version_val = (
            str(raw_version).strip()
            if pd.notna(raw_version) and str(raw_version).strip() not in {"", "<NA>", "nan", "None"}
            else pd.NA
        )
        vendor_val = (
            str(raw_vendor).strip()
            if pd.notna(raw_vendor) and str(raw_vendor).strip() not in {"", "<NA>", "nan", "None"}
            else pd.NA
        )
        parsed_dt = pd.to_datetime(raw_date, errors="coerce", dayfirst=True)
        date_val = parsed_dt.tz_localize(None) if not pd.isna(parsed_dt) else pd.Timestamp(date.today()).normalize()

        matched_indices: list[int] = []
        k_id = customer_id_key(id_str)
        k_name = normalize_name_key(name_str)

        if k_id and k_id in id_to_rows:
            matched_indices = id_to_rows[k_id]
        elif k_name and k_name in name_to_rows:
            matched_indices = name_to_rows[k_name]

        if matched_indices:
            for r_idx in matched_indices:
                if pd.notna(size_val) and str(size_val).strip() != "":
                    target.at[r_idx, "Ukuran Sticker Panel Toko"] = size_val
                if pd.notna(brand_val) and str(brand_val).strip() != "":
                    target.at[r_idx, "Brand Panel"] = brand_val
                if pd.notna(version_val) and str(version_val).strip() != "":
                    target.at[r_idx, "Versi"] = version_val
                if pd.notna(vendor_val) and str(vendor_val).strip() != "":
                    target.at[r_idx, "Nama Vendor"] = vendor_val
                if not pd.isna(date_val):
                    target.at[r_idx, "Tanggal Penggantian Sticker Panel"] = date_val
                if (pd.notna(size_val) and str(size_val).strip() != "") or (pd.notna(brand_val) and str(brand_val).strip() != ""):
                    target.at[r_idx, "Panel Toko"] = True
                updated_rows_set.add(r_idx)
        else:
            label = f"{id_str} - {name_str}".strip(" -")
            unmatched_labels.append(label)

    return normalize_data(target), len(updated_rows_set), created_count, unmatched_labels


def read_sticker_order_excel(source) -> pd.DataFrame:
    if hasattr(source, "seek"):
        source.seek(0)
    raw = pd.read_excel(source, sheet_name=0, header=0)
    raw = clean_headers(raw)
    if raw.empty:
        raise ValueError("File Excel kosong.")

    lowered = {c.casefold(): c for c in raw.columns}

    def pick(*aliases):
        for a in aliases:
            if a.casefold() in lowered:
                return raw[lowered[a.casefold()]]
        return pd.Series(pd.NA, index=raw.index)

    id_col = pick("ID Outlet", "ID Customer", "Kode Outlet", "Outlet ID")
    name_col = pick("Nama Outlet", "Nama Customer", "Customer Name", "Nama Toko", "Outlet")
    material_col = pick(
        "Material",
        "Ukuran Sticker",
        "Ukuran Sticker / Vinyl",
        "Ukuran Sticker Panel Toko",
        "Ukuran Sticker Panel",
        "Ukuran Stiker",
    )
    brand_col = pick("Brand", "Brand Sticker", "Brand Panel", "Merek Sticker", "Merek Panel")
    version_col = pick("Version", "Versi", "Versi Sticker", "Versi Panel")
    qty_col = pick("Jumlah Sticker / Vinyl", "Jumlah", "Qty", "Jumlah Sticker")
    date_col = pick(
        "Tanggal OrderPenggantian Sticker / Vinyl",
        "Tanggal Order Penggantian Sticker / Vinyl",
        "Tanggal Penggantian Sticker Panel",
        "Tanggal Penggantian",
        "Tanggal Penggantian Terakhir Sticker Panel",
        "Tgl Penggantian",
    )
    vendor_col = pick("Nama Vendor", "Vendor", "Nama Vendor Sticker")

    df = pd.DataFrame({
        "ID Outlet": id_col,
        "Nama Outlet": name_col,
        "Material": material_col,
        "Brand": brand_col,
        "Version": version_col,
        "Jumlah Sticker / Vinyl": qty_col,
        "Tanggal OrderPenggantian Sticker / Vinyl": date_col,
        "Nama Vendor": vendor_col,
    })

    for col in ("ID Outlet", "Nama Outlet", "Material", "Brand", "Version", "Nama Vendor"):
        df[col] = df[col].astype("string").str.strip()
        df[col] = df[col].replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA, "None": pd.NA})

    df["Jumlah Sticker / Vinyl"] = pd.to_numeric(df["Jumlah Sticker / Vinyl"], errors="coerce")
    df["Tanggal OrderPenggantian Sticker / Vinyl"] = pd.to_datetime(
        df["Tanggal OrderPenggantian Sticker / Vinyl"], errors="coerce", dayfirst=True
    ).dt.tz_localize(None)

    has_any = df[["ID Outlet", "Nama Outlet", "Material", "Brand", "Version", "Nama Vendor"]].notna().any(axis=1)
    df = df.loc[has_any.to_numpy(dtype=bool)].reset_index(drop=True)

    missing_name_mask = df["Nama Outlet"].isna()
    if missing_name_mask.any():
        bad_rows = [str(i + 2) for i in df.index[missing_name_mask]]
        raise ValueError(f"Kolom 'Nama Outlet' wajib diisi pada baris Excel: {', '.join(bad_rows)}.")

    return df


def merge_panel_screen_by_id(
    contracts: pd.DataFrame,
    source_df: pd.DataFrame,
    overwrite_existing: bool = False,
) -> tuple[pd.DataFrame, int, list[str]]:
    target = normalize_data(contracts).copy()
    src = normalize_data(source_df).copy()

    target_keys = target["ID Outlet"].map(customer_id_key)
    key_to_rows = {
        key: list(idx_list)
        for key, idx_list in target_keys.groupby(target_keys).groups.items()
        if key
    }

    updated_rows = 0
    unmatched_ids: list[str] = []

    for _, row in src.iterrows():
        raw_id = row.get("ID Outlet", pd.NA)
        key = customer_id_key(raw_id)
        if not key:
            continue
        if key not in key_to_rows:
            if pd.notna(raw_id) and str(raw_id).strip():
                unmatched_ids.append(str(raw_id).strip())
            continue

        row_changed = False
        for r_idx in key_to_rows[key]:
            for col in SYNC_PANEL_SCREEN_COLUMNS:
                val = row.get(col, pd.NA)
                if pd.isna(val) or str(val).strip() == "":
                    continue
                current_val = target.at[r_idx, col]
                current_is_empty = pd.isna(current_val) or str(current_val).strip() == ""
                if overwrite_existing or current_is_empty:
                    if str(current_val) != str(val):
                        target.at[r_idx, col] = val
                        row_changed = True
            screen_val = target.at[r_idx, "Ukuran Layar"]
            if pd.notna(screen_val) and str(screen_val).strip() != "":
                target.at[r_idx, "Layar Toko"] = True

        if row_changed:
            updated_rows += 1

    return normalize_data(target), updated_rows, sorted(set(unmatched_ids))


def read_drp_sheet(source) -> pd.DataFrame | None:
    try:
        if hasattr(source, "seek"):
            source.seek(0)
        excel_file = pd.ExcelFile(source)
        if DRP_SHEET_NAME not in excel_file.sheet_names:
            return None
        if hasattr(source, "seek"):
            source.seek(0)
        raw_drp = pd.read_excel(source, sheet_name=DRP_SHEET_NAME, header=0)
        return normalize_data(raw_drp)
    except Exception:
        return None


def read_workbook(source) -> pd.DataFrame:
    last_error = None
    base_df: pd.DataFrame | None = None
    for sheet, header in (("COSPKR06", 3), (SHEET_NAME, 0), (0, 0)):
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            result = normalize_data(pd.read_excel(source, sheet_name=sheet, header=header))
            has_identity = result["ID Outlet"].notna() | result["Nama Customer"].notna()
            if not result.empty and has_identity.any():
                base_df = result.loc[has_identity.to_numpy(dtype=bool)].reset_index(drop=True)
                break
        except Exception as error:
            last_error = error
            continue

    if base_df is None:
        detail = f" ({last_error})" if last_error else ""
        raise ValueError(f"Tidak dapat membaca sheet/kolom data kontrak{detail}. Periksa file Excel dan nama sheet-nya.")

    drp_df = read_drp_sheet(source)
    if drp_df is not None and not drp_df.empty:
        base_df, _, _ = merge_panel_screen_by_id(base_df, drp_df, overwrite_existing=False)

    return base_df


def read_branding_workbook(source) -> pd.DataFrame:
    if hasattr(source, "seek"):
        source.seek(0)
    frame = pd.read_excel(source, sheet_name="COSPKR06", header=11)
    frame = frame.iloc[1:].copy()
    columns = list(frame.columns)
    required = [
        "ID Customer", "Customer", "Customer Group Category", "Alamat",
        "Zona", "Finance Brand", "Nominal",
    ]
    missing = [column for column in required if column not in columns]
    if missing:
        raise ValueError(f"Kolom wajib tidak ditemukan: {', '.join(missing)}.")
    start_index = columns.index("Tanggal Pekerjaan")
    if start_index + 1 >= len(columns):
        raise ValueError("Kolom Tanggal Pekerjaan Sampai tidak ditemukan.")
    nominal_index = columns.index("Nominal")
    if nominal_index + 1 >= len(columns):
        raise ValueError("Kolom nilai Nominal tidak ditemukan.")
    columns[start_index] = "Tanggal Pekerjaan Dari"
    columns[start_index + 1] = "Tanggal Pekerjaan Sampai"
    columns[nominal_index] = "Mata Uang"
    columns[nominal_index + 1] = "Nominal"
    frame.columns = columns
    result_columns = [
        "ID Customer", "Customer", "Customer Group Category", "Alamat",
        "Zona", "Finance Brand", "Nominal", "Tanggal Pekerjaan Dari",
        "Tanggal Pekerjaan Sampai", "No. Dokumen",
    ]
    if "No. Dokumen" not in frame:
        frame["No. Dokumen"] = pd.NA
    result = frame[result_columns].dropna(how="all").reset_index(drop=True)
    result["ID Customer"] = result["ID Customer"].astype("string").str.strip()
    result["ID Customer"] = result["ID Customer"].replace("", pd.NA)
    for column in ("Customer", "Customer Group Category", "Alamat", "Zona", "Finance Brand", "No. Dokumen"):
        result[column] = result[column].astype("string").str.strip()
        result[column] = result[column].replace("", pd.NA)
    result["Nominal"] = result["Nominal"].map(parse_amount)
    for column in ("Tanggal Pekerjaan Dari", "Tanggal Pekerjaan Sampai"):
        result[column] = pd.to_datetime(result[column], errors="coerce", dayfirst=True).dt.tz_localize(None)
    result["_source_order"] = range(len(result))
    return result


def read_drp_crosscheck_workbook(source) -> pd.DataFrame:
    if hasattr(source, "seek"):
        source.seek(0)
    excel_file = pd.ExcelFile(source)
    candidate_sheets = ["CODPR03"] + [s for s in excel_file.sheet_names if s != "CODPR03"]
    parsed_df: pd.DataFrame | None = None

    for sheet in candidate_sheets:
        for hdr in (8, 0, 3, 9, 10):
            try:
                if hasattr(source, "seek"):
                    source.seek(0)
                raw = pd.read_excel(source, sheet_name=sheet, header=hdr)
                raw = clean_headers(raw)
                unnamed_cols = [c for c in raw.columns if c.startswith("Unnamed:")]
                raw = raw.drop(columns=unnamed_cols, errors="ignore")
                lowered = {c.casefold(): c for c in raw.columns}
                id_col = next((lowered[k] for k in ("id customer", "id outlet", "customer id") if k in lowered), None)
                if id_col is not None:
                    parsed_df = raw.copy()
                    break
            except Exception:
                continue
        if parsed_df is not None:
            break

    if parsed_df is None:
        raise ValueError("Kolom 'ID Customer' atau 'ID Outlet' tidak ditemukan pada file DRP yang diunggah.")

    lowered = {c.casefold(): c for c in parsed_df.columns}

    def pick_col(*names):
        for n in names:
            if n.casefold() in lowered:
                return parsed_df[lowered[n.casefold()]]
        return pd.Series(pd.NA, index=parsed_df.index)

    result = pd.DataFrame({
        "ID Customer": pick_col("ID Customer", "ID Outlet", "Customer ID"),
        "Customer": pick_col("Customer", "Nama Customer", "Nama Outlet"),
        "Zona": pick_col("Zona", "Zone"),
        "Kategori Outlet": pick_col("Kategori Outlet", "Customer Group Category"),
        "Kategori DRP": pick_col("Kategori DRP", "Kategori Pekerjaan"),
        "Brand": pick_col("Finance Brand", "Brand", "Merek"),
        "No. Dokumen": pick_col("No. Dokumen", "No Dokumen", "Nomor Dokumen"),
        "Periode Kontrak Awal": pick_col("Start Kontrak", "Periode Kontrak Awal", "Tanggal Pekerjaan Dari"),
        "Periode Kontrak Akhir": pick_col("End Kontrak", "Periode Kontrak Akhir", "Tanggal Pekerjaan Sampai"),
        "Nominal Kontrak": pick_col("Nominal SPKS", "Nominal", "Nominal Kontrak"),
        "Status Outlet": pick_col("Status Outlet"),
        "Status DRP": pick_col("Status DRP"),
        "Kontrak SPKS": pick_col("Kontrak SPKS"),
    })

    for col in ("ID Customer", "Customer", "Zona", "Kategori Outlet", "Kategori DRP", "Brand", "No. Dokumen", "Status Outlet", "Status DRP", "Kontrak SPKS"):
        result[col] = result[col].astype("string").str.strip()
        result[col] = result[col].replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA, "None": pd.NA})

    result["Nominal Kontrak"] = result["Nominal Kontrak"].map(parse_amount)
    for col in ("Periode Kontrak Awal", "Periode Kontrak Akhir"):
        result[col] = pd.to_datetime(result[col], errors="coerce", dayfirst=True).dt.tz_localize(None)

    result = result[result["ID Customer"].notna()].reset_index(drop=True)
    return result


def _fmt_val_cmp(col: str, val) -> str:
    if pd.isna(val) or str(val).strip() == "":
        return "-"
    if col == "Nominal Kontrak":
        num = parse_amount(val)
        return "-" if pd.isna(num) else f"Rp {float(num):,.0f}".replace(",", ".")
    if col in DATE_COLUMNS:
        ts = pd.to_datetime(val, errors="coerce")
        return "-" if pd.isna(ts) else ts.strftime("%d/%m/%Y")
    return str(val).strip()


def _is_field_different(col: str, dash_val, drp_val) -> bool:
    if pd.isna(drp_val) or str(drp_val).strip() == "":
        return False
    if col == "Nominal Kontrak":
        d_num = parse_amount(dash_val)
        r_num = parse_amount(drp_val)
        if pd.isna(d_num) and pd.isna(r_num):
            return False
        if pd.isna(d_num) != pd.isna(r_num):
            return True
        return abs(float(d_num) - float(r_num)) > 0.5
    if col in DATE_COLUMNS:
        d_ts = pd.to_datetime(dash_val, errors="coerce")
        r_ts = pd.to_datetime(drp_val, errors="coerce")
        if pd.isna(d_ts) and pd.isna(r_ts):
            return False
        if pd.isna(d_ts) != pd.isna(r_ts):
            return True
        return d_ts.normalize() != r_ts.normalize()
    d_str = "" if pd.isna(dash_val) else str(dash_val).strip().casefold()
    r_str = "" if pd.isna(drp_val) else str(drp_val).strip().casefold()
    return d_str != r_str


def build_drp_crosscheck_report(
    contracts: pd.DataFrame,
    drp_df: pd.DataFrame,
    compare_fields: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    dash = normalize_data(contracts)
    dash_keys = dash["ID Outlet"].map(customer_id_key)
    key_to_dash_idx: dict[str, int] = {}
    for idx, k in dash_keys.items():
        if k and k not in key_to_dash_idx:
            key_to_dash_idx[k] = idx

    mismatch_rows = []
    matched_ok_rows = []
    missing_rows = []

    for _, drp_row in drp_df.iterrows():
        raw_id = str(drp_row["ID Customer"]).strip()
        key = customer_id_key(raw_id)
        if not key:
            continue
        if key not in key_to_dash_idx:
            missing_rows.append(drp_row.to_dict())
            continue

        d_idx = key_to_dash_idx[key]
        d_row = dash.loc[d_idx]
        diff_cols = [
            col for col in compare_fields
            if _is_field_different(
                col,
                d_row.get(col, pd.NA),
                drp_row.get("Customer", pd.NA) if col == "Nama Customer" else drp_row.get(col, pd.NA),
            )
        ]

        record = {
            "_dash_index": d_idx,
            "ID Customer": raw_id,
            "Nama Customer (Dashboard)": safe_text(d_row.get("Nama Customer")) or "-",
            "Nama Customer (DRP)": safe_text(drp_row.get("Customer")) or "-",
            "Kolom Tidak Sesuai": ", ".join(diff_cols) if diff_cols else "Sesuai",
            "Brand (Dash)": _fmt_val_cmp("Brand", d_row.get("Brand")),
            "Brand (DRP)": _fmt_val_cmp("Brand", drp_row.get("Brand")),
            "Nominal (Dash)": _fmt_val_cmp("Nominal Kontrak", d_row.get("Nominal Kontrak")),
            "Nominal (DRP)": _fmt_val_cmp("Nominal Kontrak", drp_row.get("Nominal Kontrak")),
            "Awal Kontrak (Dash)": _fmt_val_cmp("Periode Kontrak Awal", d_row.get("Periode Kontrak Awal")),
            "Awal Kontrak (DRP)": _fmt_val_cmp("Periode Kontrak Awal", drp_row.get("Periode Kontrak Awal")),
            "Akhir Kontrak (Dash)": _fmt_val_cmp("Periode Kontrak Akhir", d_row.get("Periode Kontrak Akhir")),
            "Akhir Kontrak (DRP)": _fmt_val_cmp("Periode Kontrak Akhir", drp_row.get("Periode Kontrak Akhir")),
            "No. Dokumen (Dash)": _fmt_val_cmp("No. Dokumen", d_row.get("No. Dokumen")),
            "No. Dokumen (DRP)": _fmt_val_cmp("No. Dokumen", drp_row.get("No. Dokumen")),
            "Zona (Dash)": _fmt_val_cmp("Zona", d_row.get("Zona")),
            "Zona (DRP)": _fmt_val_cmp("Zona", drp_row.get("Zona")),
            "Kategori Outlet (Dash)": _fmt_val_cmp("Kategori Outlet", d_row.get("Kategori Outlet")),
            "Kategori Outlet (DRP)": _fmt_val_cmp("Kategori Outlet", drp_row.get("Kategori Outlet")),
            "_raw_Nama Customer": drp_row.get("Customer", pd.NA),
            "_raw_Brand": drp_row.get("Brand", pd.NA),
            "_raw_Nominal Kontrak": drp_row.get("Nominal Kontrak", pd.NA),
            "_raw_Periode Kontrak Awal": drp_row.get("Periode Kontrak Awal", pd.NA),
            "_raw_Periode Kontrak Akhir": drp_row.get("Periode Kontrak Akhir", pd.NA),
            "_raw_No. Dokumen": drp_row.get("No. Dokumen", pd.NA),
            "_raw_Zona": drp_row.get("Zona", pd.NA),
            "_raw_Kategori Outlet": drp_row.get("Kategori Outlet", pd.NA),
        }

        if diff_cols:
            mismatch_rows.append(record)
        else:
            matched_ok_rows.append(record)

    return (
        pd.DataFrame(mismatch_rows),
        pd.DataFrame(matched_ok_rows),
        pd.DataFrame(missing_rows),
    )


def deduplicate_branding_data(frame: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    records = frame.copy()
    records["_id_key"] = records["ID Customer"].map(customer_id_key)
    with_id = records[records["_id_key"].ne("")]
    without_id = records[records["_id_key"].eq("")]
    ordered = with_id.sort_values(
        ["Tanggal Pekerjaan Sampai", "_source_order"],
        kind="stable",
        na_position="first",
    )
    latest = ordered.drop_duplicates("_id_key", keep="last")
    duplicated = with_id[with_id.duplicated("_id_key", keep=False)].copy()
    latest_orders = set(latest["_source_order"])
    duplicated["Dipakai"] = duplicated["_source_order"].isin(latest_orders)
    selected = pd.concat([latest, without_id], ignore_index=True).sort_values(
        "_source_order", kind="stable"
    )
    return selected, duplicated


def apply_branding_updates(contracts: pd.DataFrame, branding: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, int]:
    selected, duplicates = deduplicate_branding_data(branding)
    updated = normalize_data(contracts)
    existing_keys = updated["ID Outlet"].map(customer_id_key)
    existing_by_key = {
        key: indexes for key, indexes in existing_keys.groupby(existing_keys).groups.items() if key
    }
    missing_rows = []
    matched_keys = set()
    field_map = {
        "Customer": "Nama Customer",
        "Customer Group Category": "Kategori Outlet",
        "Alamat": "Alamat",
        "Zona": "Zona",
        "Finance Brand": "Brand",
        "Nominal": "Nominal Kontrak",
        "Tanggal Pekerjaan Dari": "Periode Kontrak Awal",
        "Tanggal Pekerjaan Sampai": "Periode Kontrak Akhir",
        "No. Dokumen": "No. Dokumen",
    }
    for _, record in selected.iterrows():
        key = customer_id_key(record["ID Customer"])
        if not key or key not in existing_by_key:
            missing_rows.append(record)
            continue
        matched_keys.add(key)
        for index in existing_by_key[key]:
            for source_column, target_column in field_map.items():
                value = record[source_column]
                if pd.isna(value) or (isinstance(value, str) and not value.strip()):
                    continue
                updated.at[index, target_column] = value
    unmatched = pd.DataFrame(missing_rows, columns=selected.columns).drop(
        columns=["_id_key", "_source_order"], errors="ignore"
    )
    duplicate_details = duplicates.drop(columns=["_id_key", "_source_order"], errors="ignore")
    return normalize_data(updated), unmatched, duplicate_details, len(matched_keys)


def load_initial_data() -> tuple[pd.DataFrame, str | None]:
    if not DATA_FILE.exists():
        return pd.DataFrame(columns=COLUMNS), None
    try:
        return read_workbook(DATA_FILE), None
    except Exception as error:
        return pd.DataFrame(columns=COLUMNS), f"File data awal tidak dapat dibaca ({error}). Unggah workbook Excel yang valid untuk memulai."


def validate_data(frame: pd.DataFrame) -> list[str]:
    errors = []
    normalized = normalize_data(frame)
    if normalized.empty:
        return errors
    if (normalized["Nama Customer"].isna() & normalized["ID Outlet"].isna()).any():
        errors.append("Setiap baris harus berisi Nama Customer atau ID Outlet.")
    if (pd.to_numeric(normalized["Nominal Kontrak"], errors="coerce") < 0).any():
        errors.append("Nominal kontrak tidak boleh negatif.")
    invalid_period = (
        normalized["Periode Kontrak Awal"].notna()
        & normalized["Periode Kontrak Akhir"].notna()
        & (normalized["Periode Kontrak Akhir"] < normalized["Periode Kontrak Awal"])
    )
    if invalid_period.any():
        errors.append("Tanggal akhir kontrak tidak boleh sebelum tanggal awal.")
    return errors


def add_contract_status(frame: pd.DataFrame, today: date | None = None) -> pd.DataFrame:
    result = normalize_data(frame)
    remaining = (result["Periode Kontrak Akhir"] - pd.Timestamp(today or date.today()).normalize()).dt.days
    result["Sisa Hari"] = remaining

    def status(days):
        if pd.isna(days):
            return "⚪ Data belum lengkap"
        if days < 0:
            return "🔴 Expired"
        if days <= 30:
            return "🟠 Kritis (≤ 30 hari)"
        if days <= 90:
            return "🟡 Warning (31–90 hari)"
        return "🟢 Aktif (> 90 hari)"

    result["Status Kontrak"] = remaining.map(status)
    return result


def excel_bytes(frame: pd.DataFrame, sheet_name: str = SHEET_NAME) -> bytes:
    output = BytesIO()
    export = frame.copy()
    for column in MATERIAL_COLUMNS:
        if column in export:
            export[column] = export[column].map(lambda checked: "✓" if parse_tickmark(checked) else "")
    for column in (*DATE_COLUMNS, "Tanggal OrderPenggantian Sticker / Vinyl"):
        if column in export:
            export[column] = pd.to_datetime(export[column], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()


def backup_to_google_sheet(frame: pd.DataFrame, url: str) -> int:
    if not url.startswith("https://script.google.com/"):
        raise ValueError("URL backup harus menggunakan alamat HTTPS Google Apps Script.")
    normalized = normalize_data(frame)
    records = json.loads(normalized.to_json(orient="records", date_format="iso", force_ascii=False))
    payload = {
        "action": "backup",
        "sheet_name": SHEET_NAME,
        "columns": COLUMNS,
        "records": records,
        "backup_at": pd.Timestamp.now(tz="UTC").isoformat(),
    }
    response = requests.post(url, json=payload, timeout=60)
    response.raise_for_status()
    try:
        result = response.json()
    except ValueError as error:
        raise ValueError("Respons Google Apps Script bukan JSON yang valid.") from error
    if not isinstance(result, dict) or result.get("ok") is not True:
        message = result.get("error", "Backup ditolak oleh Google Apps Script.") if isinstance(result, dict) else "Respons Google Apps Script tidak valid."
        raise ValueError(message)
    return int(result.get("row_count", len(records)))


def _supabase_headers(api_key: str) -> dict[str, str]:
    key = str(api_key or "").strip()
    return {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }


def _supabase_config() -> tuple[str, str, str]:
    try:
        url = str(st.secrets.get("SUPABASE_URL", "") or "").strip().rstrip("/")
    except FileNotFoundError:
        url = ""
    try:
        key = (
            str(
                st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")
                or st.secrets.get("SUPABASE_ANON_KEY", "")
                or st.secrets.get("SUPABASE_KEY", "")
                or ""
            )
            .strip()
        )
    except FileNotFoundError:
        key = ""
    try:
        table = str(st.secrets.get("SUPABASE_TABLE", "monitoring_kontrak_outlet") or "monitoring_kontrak_outlet").strip()
    except FileNotFoundError:
        table = "monitoring_kontrak_outlet"
    return url, key, table


def backup_to_supabase(frame: pd.DataFrame, url: str, api_key: str, table: str) -> int:
    clean_url = str(url or "").strip().rstrip("/")
    clean_table = str(table or "").strip()
    if not clean_url.startswith("https://") or not clean_table:
        raise ValueError("SUPABASE_URL dan nama tabel Supabase harus valid.")
    normalized = normalize_data(frame)
    records = json.loads(normalized.to_json(orient="records", date_format="iso", force_ascii=False))
    headers = _supabase_headers(api_key)
    endpoint = f"{clean_url}/rest/v1/{clean_table}"

    del_resp = requests.delete(f"{endpoint}?payload_json=not.is.null", headers=headers, timeout=60)
    if del_resp.status_code >= 400:
        raise ValueError(f"Gagal membersihkan data lama di Supabase ({del_resp.status_code}): {del_resp.text}")

    if records:
        payload = [{"payload_json": records}]
        ins_resp = requests.post(
            endpoint,
            headers={**headers, "Prefer": "return=representation"},
            json=payload,
            timeout=60,
        )
        if ins_resp.status_code == 400 and "payload_json" in ins_resp.text:
            ins_resp = requests.post(
                endpoint,
                headers={**headers, "Prefer": "return=representation"},
                json=records,
                timeout=60,
            )
        ins_resp.raise_for_status()
    return len(records)


def restore_from_supabase(url: str, api_key: str, table: str) -> pd.DataFrame:
    clean_url = str(url or "").strip().rstrip("/")
    clean_table = str(table or "").strip()
    if not clean_url.startswith("https://") or not clean_table:
        raise ValueError("SUPABASE_URL dan nama tabel Supabase harus valid.")
    headers = _supabase_headers(api_key)
    endpoint = f"{clean_url}/rest/v1/{clean_table}?select=payload_json,created_at&order=created_at.desc&limit=1"
    resp = requests.get(endpoint, headers=headers, timeout=60)
    resp.raise_for_status()
    rows = resp.json()
    if not isinstance(rows, list) or not rows:
        return pd.DataFrame(columns=COLUMNS)
    raw_records = rows[0].get("payload_json")
    if isinstance(raw_records, str):
        raw_records = json.loads(raw_records)
    if not isinstance(raw_records, list) or not raw_records:
        return pd.DataFrame(columns=COLUMNS)
    return normalize_data(pd.DataFrame(raw_records))


def format_outlet_table(frame: pd.DataFrame) -> pd.DataFrame:
    display_frame = frame.copy()
    if "Nominal Kontrak" in display_frame:
        display_frame["Nominal Kontrak"] = display_frame["Nominal Kontrak"].map(
            lambda value: "" if pd.isna(value) else f"Rp {float(value):,.0f}".replace(",", ".")
        )
    for column in DATE_COLUMNS:
        if column in display_frame:
            display_frame[column] = pd.to_datetime(display_frame[column], errors="coerce").map(
                lambda value: "" if pd.isna(value) else value.strftime("%d/%m/%Y")
            )
    for column in MATERIAL_COLUMNS:
        if column in display_frame:
            display_frame[column] = display_frame[column].map(lambda checked: "✅" if parse_tickmark(checked) else "—")
    priority_cols = [
        "Status Kontrak", "Sisa Hari", "ID Outlet", "Nama Customer",
        "Brand", "Nominal Kontrak", "Periode Kontrak Awal", "Periode Kontrak Akhir", "No. Dokumen",
        "Ukuran Sticker Panel Toko", "Brand Panel", "Versi", "Tanggal Penggantian Sticker Panel", "Nama Vendor",
        "Zona", "Sektor", "Pos Budget", "Kategori Outlet", "Kategori Pekerjaan",
    ]
    existing_priority = [c for c in priority_cols if c in display_frame.columns]
    remaining_cols = [c for c in display_frame.columns if c not in existing_priority]
    return display_frame[existing_priority + remaining_cols]


def budget_group(value) -> str:
    if pd.isna(value):
        return ""
    text = re.sub(r"\s+", " ", str(value).strip().casefold())
    if re.match(r"others?\s+permanen\b", text):
        return "Others Permanen"
    if text == "venue":
        return "Venue"
    return ""


def render_dashboard_table(title: str, frame: pd.DataFrame, download_filename: str, dl_key: str) -> None:
    col_head, col_dl = st.columns([3, 1])
    with col_head:
        st.markdown(f"#### {title} ({len(frame):,} outlet)".replace(",", "."))
    with col_dl:
        st.download_button(
            "📥 Export Excel",
            data=excel_bytes(frame.sort_values("Sisa Hari", na_position="last")),
            file_name=download_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=dl_key,
            use_container_width=True,
        )

    if frame.empty:
        st.info(f"Tidak ada data customer pada kategori **{title}** yang sesuai dengan filter saat ini.")
        return

    sorted_frame = frame.sort_values("Sisa Hari", na_position="last")
    st.dataframe(
        format_outlet_table(sorted_frame),
        use_container_width=True,
        hide_index=True,
        height=440,
        column_config={
            "Status Kontrak": st.column_config.TextColumn("Status", width="medium"),
            "Sisa Hari": st.column_config.NumberColumn("Sisa Hari", format="%d hr", width="small"),
            "ID Outlet": st.column_config.TextColumn("ID Outlet", width="small"),
            "Nama Customer": st.column_config.TextColumn("Nama Customer", width="large"),
            "Nominal Kontrak": st.column_config.TextColumn("Nominal Kontrak", width="medium"),
            "Ukuran Sticker Panel Toko": st.column_config.TextColumn("Ukuran Sticker Panel", width="medium"),
            "Brand Panel": st.column_config.TextColumn("Brand Sticker", width="medium"),
            "Versi": st.column_config.TextColumn("Versi", width="small"),
            "Tanggal Penggantian Sticker Panel": st.column_config.TextColumn("Tgl Ganti Sticker Terakhir", width="medium"),
            "Nama Vendor": st.column_config.TextColumn("Nama Vendor", width="medium"),
        },
    )


def save_data(frame: pd.DataFrame) -> None:
    errors = validate_data(frame)
    if errors:
        raise ValueError("\n".join(errors))
    normalized = normalize_data(frame)
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    normalized.to_excel(DATA_FILE, sheet_name=SHEET_NAME, index=False, engine="openpyxl")


def safe_text(value) -> str:
    return "" if pd.isna(value) else str(value)


def safe_date(value):
    return None if pd.isna(value) else pd.Timestamp(value).date()


def render_outlet_form(current: pd.Series, row_index: int | None) -> tuple[bool, dict]:
    suffix = "new" if row_index is None else f"row_{row_index}"
    with st.form(f"outlet_form_{suffix}", border=True):
        st.markdown("##### 1. Identitas & Lokasi Outlet")
        c1, c2, c3 = st.columns([1, 2, 1])
        outlet_id = c1.text_input("ID Outlet *", value=safe_text(current["ID Outlet"]), placeholder="Contoh: 501X1234", key=f"id_{suffix}")
        customer_name = c2.text_input("Nama Customer *", value=safe_text(current["Nama Customer"]), placeholder="Nama Toko / Outlet", key=f"name_{suffix}")
        outlet_category = c3.text_input("Kategori Outlet", value=safe_text(current["Kategori Outlet"]), placeholder="Warung / Minimarket", key=f"outlet_category_{suffix}")

        address = st.text_input("Alamat Lengkap", value=safe_text(current["Alamat"]), placeholder="Jl. Raya...", key=f"address_{suffix}")

        z1, z2, z3, z4 = st.columns(4)
        zone = z1.text_input("Zona", value=safe_text(current["Zona"]), key=f"zone_{suffix}")
        sector = z2.text_input("Sektor", value=safe_text(current["Sektor"]), key=f"sector_{suffix}")
        budget_pos = z3.text_input("Pos Budget", value=safe_text(current["Pos Budget"]), key=f"budget_{suffix}")
        job_category = z4.text_input("Kategori Pekerjaan", value=safe_text(current["Kategori Pekerjaan"]), key=f"job_category_{suffix}")

        st.divider()
        st.markdown("##### 2. Detail Kontrak & Nilai")
        k1, k2, k3, k4, k5 = st.columns(5)
        brand = k1.text_input("Brand Utama", value=safe_text(current["Brand"]), key=f"brand_{suffix}")
        parsed_amount = parse_amount(current["Nominal Kontrak"])
        amount = k2.number_input(
            "Nominal Kontrak (Rp)",
            min_value=0.0,
            step=100000.0,
            value=0.0 if pd.isna(parsed_amount) else float(parsed_amount),
            key=f"amount_{suffix}",
        )
        contract_start = k3.date_input("Tanggal Mulai", value=safe_date(current["Periode Kontrak Awal"]), key=f"start_{suffix}")
        contract_end = k4.date_input("Tanggal Berakhir", value=safe_date(current["Periode Kontrak Akhir"]), key=f"end_{suffix}")
        doc_no = k5.text_input("No. Dokumen", value=safe_text(current.get("No. Dokumen", pd.NA)), key=f"doc_no_{suffix}")

        st.divider()
        st.markdown("##### 3. Materi Promosi & Spesifikasi Ukuran")
        material_cols = st.columns(4)
        material_values = {
            column: material_cols[index].checkbox(column, value=parse_tickmark(current[column]), key=f"material_{index}_{suffix}")
            for index, column in enumerate(MATERIAL_COLUMNS)
        }

        p1, p2, p3, p4, p5, p6 = st.columns(6)
        panel_size = p1.text_input("Ukuran Panel Toko", value=safe_text(current["Ukuran Panel Toko"]), placeholder="1 x 3 m", key=f"panel_size_{suffix}")
        sticker_size = p2.text_input("Ukuran Sticker Panel", value=safe_text(current["Ukuran Sticker Panel Toko"]), placeholder="1 x 2 m", key=f"sticker_size_{suffix}")
        panel_brand = p3.text_input("Brand Panel", value=safe_text(current["Brand Panel"]), key=f"panel_brand_{suffix}")
        sticker_version = p4.text_input("Versi (Opsional)", value=safe_text(current.get("Versi", pd.NA)), placeholder="V1 / V2", key=f"sticker_version_{suffix}")
        sticker_replace_date = p5.date_input(
            "Tgl Ganti Sticker Terakhir",
            value=safe_date(current.get("Tanggal Penggantian Sticker Panel", pd.NaT)),
            key=f"sticker_replace_date_{suffix}",
        )
        vendor_name = p6.text_input("Nama Vendor", value=safe_text(current.get("Nama Vendor", pd.NA)), placeholder="ARTISTIC", key=f"vendor_name_{suffix}")

        s1, s2 = st.columns(2)
        screen_size = s1.text_input("Ukuran Layar", value=safe_text(current["Ukuran Layar"]), placeholder="2 x 3 m", key=f"screen_size_{suffix}")
        screen_brand = s2.text_input("Brand Layar", value=safe_text(current["Brand Layar"]), key=f"screen_brand_{suffix}")

        st.markdown("")
        submitted = st.form_submit_button("💾 Simpan Data Customer", type="primary", use_container_width=True)

    values = {
        "ID Outlet": outlet_id.strip() or pd.NA,
        "Nama Customer": customer_name.strip() or pd.NA,
        "Alamat": address.strip() or pd.NA,
        "Zona": zone.strip() or pd.NA,
        "Sektor": sector.strip() or pd.NA,
        "Pos Budget": budget_pos.strip() or pd.NA,
        "Kategori Outlet": outlet_category.strip() or pd.NA,
        "Kategori Pekerjaan": job_category.strip() or pd.NA,
        "Brand": brand.strip() or pd.NA,
        "Nominal Kontrak": amount,
        "Periode Kontrak Awal": pd.Timestamp(contract_start) if contract_start else pd.NaT,
        "Periode Kontrak Akhir": pd.Timestamp(contract_end) if contract_end else pd.NaT,
        "No. Dokumen": doc_no.strip() or pd.NA,
        **material_values,
        "Ukuran Panel Toko": panel_size.strip() or pd.NA,
        "Ukuran Sticker Panel Toko": extract_sizes(sticker_size) if sticker_size.strip() else pd.NA,
        "Brand Panel": panel_brand.strip() or pd.NA,
        "Versi": sticker_version.strip() or pd.NA,
        "Tanggal Penggantian Sticker Panel": pd.Timestamp(sticker_replace_date) if sticker_replace_date else pd.NaT,
        "Nama Vendor": vendor_name.strip() or pd.NA,
        "Ukuran Layar": screen_size.strip() or pd.NA,
        "Brand Layar": screen_brand.strip() or pd.NA,
    }
    return submitted, values


# ==========================================
# INISIALISASI STATE & LAYOUT UTAMA
# ==========================================
inject_custom_css()

if "contracts" not in st.session_state:
    initial, warning = load_initial_data()
    st.session_state.contracts = initial
    st.session_state.load_warning = warning

if "multi_select_all_state" not in st.session_state:
    st.session_state.multi_select_all_state = False

contracts = add_contract_status(st.session_state.contracts)
expired_count = int(contracts["Status Kontrak"].eq("🔴 Expired").sum())
critical_count = int(contracts["Status Kontrak"].eq("🟠 Kritis (≤ 30 hari)").sum())
warning_count = int(contracts["Status Kontrak"].eq("🟡 Warning (31–90 hari)").sum())
soon_count = critical_count + warning_count
total_value = pd.to_numeric(contracts["Nominal Kontrak"], errors="coerce").sum()

# SIDEBAR NAVIGASI
with st.sidebar:
    st.markdown("### 📊 DSO Surabaya Barat")
    st.caption("Sistem Monitoring Kontrak Outlet")
    st.divider()

    active_menu = st.radio(
        "Navigasi Utama",
        [
            "📊 Dashboard & Monitoring",
            "👥 Kelola Data Customer",
            "🏷️ Pemesanan Sticker Panel",
            "🔄 Import & Sinkronisasi",
            "☁️ Backup & Restore",
        ],
        label_visibility="collapsed",
        key="main_menu",
    )

    st.divider()
    st.markdown("##### Ringkasan Cepat")
    st.write(f"• **Total Outlet:** `{len(contracts):,}`".replace(",", "."))
    st.write(f"• **Kritis (≤30 hr):** `{critical_count}`")
    st.write(f"• **Expired:** `{expired_count}`")

    st.divider()
    st.download_button(
        "⬇️ Download Semua Data (.xlsx)",
        data=excel_bytes(st.session_state.contracts),
        file_name="Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# HEADER BANNER
st.markdown(
    """
    <div class="hero-header">
        <h1>📊 Monitoring Kontrak Outlet</h1>
        <p>Pantau masa berlaku kontrak, kelola database outlet, dan sinkronisasi data branding secara terpusat.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if st.session_state.load_warning:
    st.warning(st.session_state.load_warning)


# ==========================================
# HALAMAN 1: DASHBOARD & MONITORING
# ==========================================
if active_menu == "📊 Dashboard & Monitoring":
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f"""
            <div class="kpi-card border-blue">
                <div class="kpi-label">Total Customer</div>
                <div class="kpi-value">{len(contracts):,}</div>
                <div class="kpi-sub">Outlet terdaftar aktif</div>
            </div>
            """.replace(",", "."),
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"""
            <div class="kpi-card border-emerald">
                <div class="kpi-label">Total Nilai Kontrak</div>
                <div class="kpi-value">Rp {total_value/1_000_000:,.1f} Jt</div>
                <div class="kpi-sub">Rp {total_value:,.0f}</div>
            </div>
            """.replace(",", "."),
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f"""
            <div class="kpi-card border-amber">
                <div class="kpi-label">Akan Berakhir (≤ 90 Hari)</div>
                <div class="kpi-value">{soon_count}</div>
                <div class="kpi-sub">{critical_count} kritis (≤30 hr) • {warning_count} warning</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f"""
            <div class="kpi-card border-rose">
                <div class="kpi-label">Sudah Expired</div>
                <div class="kpi-value">{expired_count}</div>
                <div class="kpi-sub">Perlu tindak lanjut perpanjangan</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("")

    if expired_count or critical_count:
        alert_col1, alert_col2 = st.columns(2)
        if expired_count:
            alert_col1.error(f"🚨 **{expired_count} kontrak** sudah melewati masa berlaku (Expired).")
        if critical_count:
            alert_col2.warning(f"⚠️ **{critical_count} kontrak** memasuki masa kritis (berakhir dalam ≤ 30 hari).")

    # PENCARIAN & FILTER + TOMBOL RESET
    head_f, head_rst = st.columns([4, 1])
    with head_f:
        st.markdown("### 🔍 Pencarian & Filter Data")
    with head_rst:
        st.button(
            "🔄 Reset Filter",
            use_container_width=True,
            key="btn_reset_dash_filter",
            on_click=reset_dashboard_filters,
        )

    search_col, status_col = st.columns([2, 1])
    search_query = search_col.text_input(
        "Cari cepat (ID Outlet, Nama Customer, Alamat, No. Dokumen, atau Brand)",
        placeholder="Ketik kata kunci pencarian...",
        key="dash_search_query",
    )
    statuses = [
        "Semua Status",
        "🔴 Expired",
        "🟠 Kritis (≤ 30 hari)",
        "🟡 Warning (31–90 hari)",
        "🟢 Aktif (> 90 hari)",
        "⚪ Data belum lengkap",
    ]
    selected_status = status_col.selectbox("Status Masa Berlaku", statuses, key="dash_status")

    with st.expander("🎛️ Filter Lanjutan (Zona, Sektor, Pos Budget, Ukuran Panel/Layar)", expanded=False):
        f1, f2, f3, f4 = st.columns(4)
        zones = ["Semua"] + sorted(contracts["Zona"].dropna().astype(str).unique().tolist())
        sectors = ["Semua"] + sorted(contracts["Sektor"].dropna().astype(str).unique().tolist())
        budget_positions = ["Semua"] + sorted(contracts["Pos Budget"].dropna().astype(str).unique().tolist())
        job_categories = ["Semua"] + sorted(contracts["Kategori Pekerjaan"].dropna().astype(str).unique().tolist())

        selected_zone = f1.selectbox("Zona", zones, key="dash_zone")
        selected_sector = f2.selectbox("Sektor", sectors, key="dash_sector")
        selected_budget = f3.selectbox("Pos Budget", budget_positions, key="dash_budget")
        selected_job_category = f4.selectbox("Kategori Pekerjaan", job_categories, key="dash_job_cat")

        f5, f6 = st.columns(2)
        sticker_sizes = ["Semua"] + sorted(contracts["Ukuran Sticker Panel Toko"].dropna().astype(str).unique().tolist())
        screen_sizes = ["Semua"] + sorted(contracts["Ukuran Layar"].dropna().astype(str).unique().tolist())
        selected_sticker_size = f5.selectbox("Ukuran Sticker Panel", sticker_sizes, key="dash_sticker_size")
        selected_screen_size = f6.selectbox("Ukuran Layar", screen_sizes, key="dash_screen_size")

    view = contracts.copy()
    if search_query.strip():
        q = search_query.strip().casefold()
        mask = (
            view["ID Outlet"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
            | view["Nama Customer"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
            | view["Alamat"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
            | view["No. Dokumen"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
            | view["Brand"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
            | view["Brand Panel"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
            | view["Versi"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
            | view["Nama Vendor"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
        )
        view = view[mask]

    if selected_status != "Semua Status":
        view = view[view["Status Kontrak"] == selected_status]
    if selected_zone != "Semua":
        view = view[view["Zona"] == selected_zone]
    if selected_sector != "Semua":
        view = view[view["Sektor"] == selected_sector]
    if selected_budget != "Semua":
        view = view[view["Pos Budget"] == selected_budget]
    if selected_job_category != "Semua":
        view = view[view["Kategori Pekerjaan"] == selected_job_category]
    if selected_sticker_size != "Semua":
        view = view[view["Ukuran Sticker Panel Toko"] == selected_sticker_size]
    if selected_screen_size != "Semua":
        view = view[view["Ukuran Layar"] == selected_screen_size]

    st.caption(f"Menampilkan **{len(view):,}** dari **{len(contracts):,}** total customer.".replace(",", "."))

    if view.empty:
        st.info("Tidak ada data yang cocok dengan kriteria pencarian/filter. Coba ubah filter atau impor data baru di menu **Import & Sinkronisasi**.")
    else:
        groups = view["Pos Budget"].map(budget_group)
        regular_view = view.loc[groups.eq("")]
        venue_view = view.loc[groups.eq("Venue")]
        other_permanent_view = view.loc[groups.eq("Others Permanen")]

        tab_all, tab_retail, tab_venue, tab_other = st.tabs([
            f"📋 Semua ({len(view)})",
            f"🏪 Retail Supply ({len(regular_view)})",
            f"🏟️ Venue ({len(venue_view)})",
            f"🏢 Others Permanen ({len(other_permanent_view)})",
        ])

        with tab_all:
            render_dashboard_table("Semua Data Customer Terfilter", view, "Daftar_Semua_Customer_Terfilter.xlsx", "dl_all_filtered")
        with tab_retail:
            render_dashboard_table("Daftar Customer - Retail Supply", regular_view, "Daftar_Customer_Retail_Supply.xlsx", "dl_retail_filtered")
        with tab_venue:
            render_dashboard_table("Daftar Customer - Venue", venue_view, "Daftar_Customer_Venue.xlsx", "dl_venue_filtered")
        with tab_other:
            render_dashboard_table("Daftar Customer - Others Permanen", other_permanent_view, "Daftar_Customer_Others_Permanen.xlsx", "dl_other_filtered")


# ==========================================
# HALAMAN 2: KELOLA DATA CUSTOMER (SINGLE & MULTIPLE)
# ==========================================
elif active_menu == "👥 Kelola Data Customer":
    tab_multi_edit, tab_single_edit, tab_add_manual, tab_add_batch = st.tabs([
        "⚡ Edit Multiple (Quick Search)",
        "✏️ Edit Satuan (Form)",
        "➕ Tambah Customer (Form)",
        "📑 Tambah Massal (Excel)",
    ])

    with tab_multi_edit:
        if contracts.empty:
            st.info("Belum ada data customer untuk diedit. Silakan tambah customer baru atau import Excel.")
        else:
            me_title_col, me_reset_col = st.columns([3.5, 1.5])
            with me_title_col:
                st.markdown("#### 🔍 Pencarian Instan & Edit Massal")
                st.caption("Cari beberapa outlet sekaligus, lalu edit langsung di tabel atau gunakan **Update Massal** untuk mengubah banyak baris dalam 1 klik.")
            with me_reset_col:
                st.button(
                    "🔄 Reset & Batalkan Edit",
                    use_container_width=True,
                    key="btn_reset_multi_edit",
                    on_click=reset_multi_edit_state,
                )

            qs_col1, qs_col2, qs_col3 = st.columns([2.2, 1, 1])
            multi_search = qs_col1.text_input(
                "Quick Search (ID Outlet, Nama Customer, Alamat, Brand, atau Sektor)",
                placeholder="Ketik kata kunci, misal: 'Warung', 'Surabaya', 'Super', atau kode ID...",
                key="multi_edit_quick_search",
            )
            multi_zones = ["Semua Zona"] + sorted(contracts["Zona"].dropna().astype(str).unique().tolist())
            multi_sel_zone = qs_col2.selectbox("Filter Zona", multi_zones, key="multi_edit_zone")

            multi_statuses = [
                "Semua Status",
                "🔴 Expired",
                "🟠 Kritis (≤ 30 hari)",
                "🟡 Warning (31–90 hari)",
                "🟢 Aktif (> 90 hari)",
                "⚪ Data belum lengkap",
            ]
            multi_sel_status = qs_col3.selectbox("Filter Status", multi_statuses, key="multi_edit_status")

            filtered_multi = contracts.copy()
            if multi_search.strip():
                q_multi = multi_search.strip().casefold()
                mask_multi = (
                    filtered_multi["ID Outlet"].fillna("").astype(str).str.casefold().str.contains(q_multi, regex=False)
                    | filtered_multi["Nama Customer"].fillna("").astype(str).str.casefold().str.contains(q_multi, regex=False)
                    | filtered_multi["Alamat"].fillna("").astype(str).str.casefold().str.contains(q_multi, regex=False)
                    | filtered_multi["Brand"].fillna("").astype(str).str.casefold().str.contains(q_multi, regex=False)
                    | filtered_multi["Sektor"].fillna("").astype(str).str.casefold().str.contains(q_multi, regex=False)
                )
                filtered_multi = filtered_multi[mask_multi]

            if multi_sel_zone != "Semua Zona":
                filtered_multi = filtered_multi[filtered_multi["Zona"] == multi_sel_zone]
            if multi_sel_status != "Semua Status":
                filtered_multi = filtered_multi[filtered_multi["Status Kontrak"] == multi_sel_status]

            st.caption(f"Ditemukan **{len(filtered_multi):,}** baris customer sesuai pencarian.".replace(",", "."))

            if filtered_multi.empty:
                st.warning("Tidak ada customer yang cocok dengan kata kunci pencarian Anda.")
            else:
                sel_info_col, btn_all_col, btn_none_col = st.columns([3.5, 1, 1])
                sel_info_col.markdown("**Daftar Hasil Pencarian (Editable)**")
                if btn_all_col.button("☑️ Centang Semua", use_container_width=True, key="btn_check_all_multi"):
                    st.session_state.multi_select_all_state = True
                    st.session_state.pop("multi_customer_data_editor", None)
                    st.rerun()
                if btn_none_col.button("⬜ Batal Centang", use_container_width=True, key="btn_uncheck_all_multi"):
                    st.session_state.multi_select_all_state = False
                    st.session_state.pop("multi_customer_data_editor", None)
                    st.rerun()

                editor_df = filtered_multi[COLUMNS].copy()
                editor_df.insert(0, "Pilih", bool(st.session_state.multi_select_all_state))

                for col in COLUMNS:
                    if col not in MATERIAL_COLUMNS and col not in ("Nominal Kontrak", *DATE_COLUMNS):
                        editor_df[col] = editor_df[col].fillna("").astype(str)
                    elif col == "Nominal Kontrak":
                        editor_df[col] = pd.to_numeric(editor_df[col], errors="coerce").fillna(0.0)

                with st.expander("🪄 Aksi Cepat Massal (Ubah Kolom Secara Serentak untuk Banyak Outlet)", expanded=False):
                    st.caption("Klik **Centang Semua** di atas atau centang kolom **Pilih** pada baris tertentu, lalu terapkan perubahan.")
                    ba1, ba2, ba3, ba4 = st.columns([1.3, 1.5, 1.2, 1])

                    bulk_target_col = ba1.selectbox(
                        "Kolom yang Diubah",
                        [
                            "Nama Customer",
                            "Brand",
                            "Nominal Kontrak",
                            "Periode Kontrak Awal",
                            "Periode Kontrak Akhir",
                            "Tanggal Penggantian Sticker Panel",
                            "Nama Vendor",
                            "Versi",
                            "No. Dokumen",
                            "Zona",
                            "Sektor",
                            "Pos Budget",
                            "Kategori Outlet",
                            "Kategori Pekerjaan",
                            "Brand Panel",
                            "Ukuran Panel Toko",
                            "Ukuran Sticker Panel Toko",
                            "Brand Layar",
                            "Ukuran Layar",
                            "Panel Toko",
                            "Showcase",
                            "Layar Toko",
                            "Shop Sign / Neon Box",
                        ],
                        key="bulk_target_col",
                    )

                    if bulk_target_col in DATE_COLUMNS:
                        bulk_new_val = ba2.date_input("Nilai Baru", value=date.today(), key="bulk_val_date")
                    elif bulk_target_col == "Nominal Kontrak":
                        bulk_new_val = ba2.number_input("Nilai Nominal Baru (Rp)", min_value=0.0, step=100000.0, key="bulk_val_num")
                    elif bulk_target_col in MATERIAL_COLUMNS:
                        bulk_new_val = ba2.selectbox(
                            "Status Centang Baru",
                            [True, False],
                            format_func=lambda x: "✅ Aktif / Terpasang" if x else "❌ Tidak Ada",
                            key="bulk_val_bool",
                        )
                    else:
                        bulk_new_val = ba2.text_input("Nilai Teks Baru", placeholder="Masukkan nilai baru...", key="bulk_val_text")

                    bulk_scope = ba3.selectbox(
                        "Target Baris",
                        ["Hanya yang dicentang (Pilih)", f"Semua hasil pencarian ({len(filtered_multi)} baris)"],
                        key="bulk_scope_choice",
                    )
                    ba4.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                    apply_bulk_clicked = ba4.button("Terapkan Massal", type="secondary", use_container_width=True, key="btn_apply_bulk")

                edited_df = st.data_editor(
                    editor_df,
                    use_container_width=True,
                    height=460,
                    hide_index=True,
                    key="multi_customer_data_editor",
                    column_config={
                        "Pilih": st.column_config.CheckboxColumn("Pilih", default=False, width="small"),
                        "ID Outlet": st.column_config.TextColumn("ID Outlet", width="small"),
                        "Nama Customer": st.column_config.TextColumn("Nama Customer", width="medium"),
                        "Nominal Kontrak": st.column_config.NumberColumn("Nominal (Rp)", min_value=0, step=50000, format="Rp %d"),
                        "Periode Kontrak Awal": st.column_config.DateColumn("Mulai Kontrak", format="DD/MM/YYYY"),
                        "Periode Kontrak Akhir": st.column_config.DateColumn("Akhir Kontrak", format="DD/MM/YYYY"),
                        "Tanggal Penggantian Sticker Panel": st.column_config.DateColumn("Tgl Ganti Sticker", format="DD/MM/YYYY"),
                        "Nama Vendor": st.column_config.TextColumn("Nama Vendor", width="medium"),
                        "Panel Toko": st.column_config.CheckboxColumn("Panel", width="small"),
                        "Showcase": st.column_config.CheckboxColumn("Showcase", width="small"),
                        "Layar Toko": st.column_config.CheckboxColumn("Layar", width="small"),
                        "Shop Sign / Neon Box": st.column_config.CheckboxColumn("Neon Box", width="small"),
                    },
                )

                if apply_bulk_clicked:
                    if bulk_scope.startswith("Hanya"):
                        target_indices = edited_df.index[edited_df["Pilih"] == True].tolist()
                    else:
                        target_indices = edited_df.index.tolist()

                    if not target_indices:
                        st.warning("⚠️ Klik **Centang Semua** atau centang minimal 1 baris pada kolom **Pilih** terlebih dahulu.")
                    else:
                        full_updated = st.session_state.contracts[COLUMNS].copy()
                        for orig_idx in edited_df.index:
                            for col in COLUMNS:
                                full_updated.at[orig_idx, col] = edited_df.at[orig_idx, col]

                        val_to_set = (
                            pd.Timestamp(bulk_new_val)
                            if bulk_target_col in DATE_COLUMNS
                            else bulk_new_val
                        )
                        for orig_idx in target_indices:
                            full_updated.at[orig_idx, bulk_target_col] = val_to_set

                        problems = validate_data(full_updated)
                        if problems:
                            st.error("\n".join(problems))
                        else:
                            try:
                                save_data(full_updated)
                                st.session_state.contracts = normalize_data(full_updated)
                                st.session_state.multi_select_all_state = False
                                st.session_state.pop("multi_customer_data_editor", None)
                                st.success(f"✅ Berhasil memperbarui kolom **{bulk_target_col}** untuk **{len(target_indices)}** customer sekaligus!")
                                st.rerun()
                            except Exception as error:
                                st.error(f"Gagal menyimpan pembaruan massal: {error}")

                selected_rows_count = int(edited_df["Pilih"].sum())
                act_col1, act_col2, act_col3 = st.columns([2, 1.3, 1.3])
                act_col1.info(f"💡 Terpilih: **{selected_rows_count}** dari **{len(edited_df)}** baris pada hasil pencarian ini.")

                if act_col2.button("💾 Simpan Perubahan Tabel", type="primary", use_container_width=True, key="btn_save_multi_table"):
                    full_updated = st.session_state.contracts[COLUMNS].copy()
                    for orig_idx in edited_df.index:
                        for col in COLUMNS:
                            full_updated.at[orig_idx, col] = edited_df.at[orig_idx, col]

                    problems = validate_data(full_updated)
                    if problems:
                        st.error("\n".join(problems))
                    else:
                        try:
                            save_data(full_updated)
                            st.session_state.contracts = normalize_data(full_updated)
                            st.session_state.multi_select_all_state = False
                            st.session_state.pop("multi_customer_data_editor", None)
                            st.success(f"✅ Perubahan pada {len(edited_df)} baris hasil pencarian berhasil disimpan!")
                            st.rerun()
                        except Exception as error:
                            st.error(f"Gagal menyimpan perubahan: {error}")

                confirm_delete_all = act_col3.checkbox(
                    "Konfirmasi hapus",
                    value=False,
                    disabled=(selected_rows_count == 0),
                    key="confirm_delete_multi_cb",
                )
                if st.button(
                    f"🗑️ Hapus Terpilih ({selected_rows_count} baris)",
                    type="secondary",
                    use_container_width=True,
                    disabled=(selected_rows_count == 0 or not confirm_delete_all),
                    key="btn_delete_multi_selected",
                ):
                    try:
                        indices_to_drop = edited_df.index[edited_df["Pilih"] == True].tolist()
                        remaining_df = st.session_state.contracts.drop(index=indices_to_drop).reset_index(drop=True)
                        save_data(remaining_df)
                        st.session_state.contracts = normalize_data(remaining_df)
                        st.session_state.multi_select_all_state = False
                        st.session_state["confirm_delete_multi_cb"] = False
                        st.session_state.pop("multi_customer_data_editor", None)
                        st.success(f"✅ {len(indices_to_drop)} customer terpilih berhasil dihapus.")
                        st.rerun()
                    except Exception as error:
                        st.error(f"Gagal menghapus customer terpilih: {error}")

    with tab_single_edit:
        if contracts.empty:
            st.info("Belum ada data customer. Silakan tambah customer baru atau import Excel.")
        else:
            record_options = [
                f"{i} | {safe_text(row['ID Outlet']) or 'Tanpa ID'} — {safe_text(row['Nama Customer']) or 'Tanpa Nama'} ({safe_text(row['Zona']) or '-'})"
                for i, row in contracts.iterrows()
            ]
            selected_record = st.selectbox(
                "Pilih Customer yang Ingin Diedit (ketik nama atau ID untuk mencari)",
                record_options,
                key="selected_customer_record",
            )
            selected_index = int(selected_record.split(" | ", 1)[0])
            submitted, updated_record = render_outlet_form(contracts.loc[selected_index], selected_index)

            if submitted:
                updated_data = contracts[COLUMNS].copy()
                for column, value in updated_record.items():
                    updated_data.at[selected_index, column] = value
                problems = validate_data(updated_data)
                if problems:
                    st.error("\n".join(problems))
                else:
                    try:
                        save_data(updated_data)
                        st.session_state.contracts = normalize_data(updated_data)
                        st.success("✅ Data customer berhasil diperbarui!")
                        st.rerun()
                    except Exception as error:
                        st.error(f"Gagal menyimpan data: {error}")

            with st.expander("🗑️ Zona Berbahaya: Hapus Customer Ini"):
                st.warning("Tindakan ini akan menghapus data customer secara permanen dari daftar aktif.")
                if st.button("Hapus Customer Terpilih", type="secondary", key="delete_customer"):
                    try:
                        updated_data = st.session_state.contracts.drop(index=selected_index).reset_index(drop=True)
                        save_data(updated_data)
                        st.session_state.contracts = normalize_data(updated_data)
                        st.success("Customer berhasil dihapus.")
                        st.rerun()
                    except Exception as error:
                        st.error(f"Gagal menghapus customer: {error}")

    with tab_add_manual:
        blank_record = pd.Series({column: pd.NA for column in COLUMNS})
        submitted, new_record = render_outlet_form(blank_record, None)
        if submitted:
            updated_data = pd.concat([contracts[COLUMNS], pd.DataFrame([new_record])], ignore_index=True)
            problems = validate_data(updated_data)
            if problems:
                st.error("\n".join(problems))
            else:
                try:
                    save_data(updated_data)
                    st.session_state.contracts = normalize_data(updated_data)
                    st.success("✅ Customer baru berhasil ditambahkan!")
                    st.rerun()
                except Exception as error:
                    st.error(f"Gagal menyimpan customer baru: {error}")

    with tab_add_batch:
        st.markdown("##### Langkah 1: Unduh Template Excel")
        sample_row = pd.DataFrame(
            [{
                "ID Outlet": "501X9999",
                "Nama Customer": "CONTOH TOKO BARU",
                "Alamat": "Jl. Contoh No. 123, Surabaya",
                "Zona": "Bubutan",
                "Sektor": "Retail Supply",
                "Pos Budget": "Small Retailer",
                "Kategori Outlet": "Warung",
                "Kategori Pekerjaan": "DRP",
                "Brand": "DJARUM SUPER",
                "Nominal Kontrak": 1500000,
                "Periode Kontrak Awal": pd.Timestamp("2026-01-01"),
                "Periode Kontrak Akhir": pd.Timestamp("2026-12-31"),
                "No. Dokumen": "BRN-26/1/5SBB/999",
                "Panel Toko": True,
                "Showcase": False,
                "Layar Toko": False,
                "Shop Sign / Neon Box": False,
                "Ukuran Panel Toko": "1 x 3 m",
                "Ukuran Sticker Panel Toko": "1 x 2 m",
                "Brand Panel": "Djarum Super",
                "Versi": "V1",
                "Tanggal Penggantian Sticker Panel": pd.Timestamp(date.today()),
                "Nama Vendor": "ARTISTIC",
                "Ukuran Layar": pd.NA,
                "Brand Layar": pd.NA,
            }]
        )
        dl1, dl2 = st.columns(2)
        dl1.download_button(
            "📄 Template Kosong (.xlsx)",
            data=excel_bytes(pd.DataFrame(columns=COLUMNS)),
            file_name="Template_Tambah_Customer_Baru.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        dl2.download_button(
            "📝 Template + Contoh Pengisian (.xlsx)",
            data=excel_bytes(sample_row[COLUMNS]),
            file_name="Template_Tambah_Customer_Baru_Contoh.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.divider()
        st.markdown("##### Langkah 2: Upload File Excel Customer Baru")
        batch_upload = st.file_uploader("Pilih file Excel (.xlsx)", type=["xlsx"], key="batch_new_customer_upload")
        skip_duplicate_id = st.checkbox(
            "Lewati baris jika ID Outlet sudah terdaftar (cegah duplikat)",
            value=True,
            key="skip_dup_id_new_customer",
        )
        if st.button("🚀 Tambahkan ke Database", type="primary", key="btn_save_batch_new_customer"):
            if batch_upload is None:
                st.warning("Pilih file Excel terlebih dahulu.")
            else:
                try:
                    new_rows = read_workbook(batch_upload)
                    if new_rows.empty:
                        st.warning("Tidak ditemukan baris data valid pada file Excel.")
                    else:
                        existing_keys = {
                            customer_id_key(val)
                            for val in contracts["ID Outlet"].dropna()
                            if customer_id_key(val)
                        }
                        skipped_ids: list[str] = []
                        rows_to_add: list[pd.Series] = []
                        seen_batch_keys: set[str] = set()

                        for _, r in new_rows.iterrows():
                            k = customer_id_key(r.get("ID Outlet"))
                            if skip_duplicate_id and k:
                                if k in existing_keys or k in seen_batch_keys:
                                    skipped_ids.append(str(r.get("ID Outlet")).strip())
                                    continue
                                seen_batch_keys.add(k)
                            rows_to_add.append(r[COLUMNS])

                        if not rows_to_add:
                            st.info(f"Tidak ada data baru ditambahkan ({len(skipped_ids)} ID sudah terdaftar).")
                        else:
                            combined = pd.concat([contracts[COLUMNS], pd.DataFrame(rows_to_add)], ignore_index=True)
                            problems = validate_data(combined)
                            if problems:
                                st.error("\n".join(problems))
                            else:
                                save_data(combined)
                                st.session_state.contracts = normalize_data(combined)
                                st.session_state.load_warning = None
                                st.success(f"✅ Berhasil menambahkan {len(rows_to_add)} customer baru!")
                                if skipped_ids:
                                    st.info(f"Dilewati ({len(skipped_ids)} ID duplikat): {', '.join(skipped_ids[:15])}")
                                st.rerun()
                except Exception as error:
                    st.error(f"Gagal memproses file: {error}")


# ==========================================
# HALAMAN 3: PEMESANAN PENGGANTIAN STICKER PANEL
# ==========================================
elif active_menu == "🏷️ Pemesanan Sticker Panel":
    st.markdown("### 🏷️ Pemesanan Penggantian Sticker Panel")
    st.caption(
        "Catat pemesanan penggantian sticker panel secara satuan maupun massal (Excel). "
        "Kolom **Material** hanya akan dibaca dimensinya saja (contoh: `Vinyl Panel Toko 1 x 2 m (BL)` → `1 x 2 m`), "
        "lalu ukuran sticker, **Brand**, **Version**, **Tanggal OrderPenggantian Sticker / Vinyl**, dan **Nama Vendor** "
        "langsung terupdate ke **Semua Data Customer Terfilter**."
    )

    tab_sticker_single, tab_sticker_batch = st.tabs([
        "✏️ Form Pemesanan Satuan",
        "📤 Upload Penggantian Massal (Excel)",
    ])

    with tab_sticker_single:
        with st.form("form_pemesanan_sticker_panel", border=True):
            st.markdown("##### Detail Pemesanan Penggantian Sticker")
            col_s1, col_s2 = st.columns([1.2, 2.3])
            in_id_outlet = col_s1.text_input(
                "ID Outlet",
                placeholder="Opsional — contoh: 501X0001",
                key="sticker_order_id",
            )
            in_nama_outlet = col_s2.text_input(
                "Nama Outlet *",
                placeholder="Wajib diisi — contoh: TOKO CONTOH 1",
                key="sticker_order_name",
            )

            col_s3, col_s4, col_s5 = st.columns([1.6, 1.3, 1.4])
            in_material = col_s3.text_input(
                "Material",
                placeholder="Contoh: Vinyl Panel Toko 1 x 2 m (BL) atau 1 x 2 m",
                key="sticker_order_material",
            )
            in_brand = col_s4.text_input(
                "Brand",
                placeholder="Contoh: DJARUM BLACK MATCHA KRETEK",
                key="sticker_order_brand",
            )
            in_version = col_s5.text_input(
                "Version",
                placeholder="Opsional — contoh: Matcha Powder - Rp 17.000/12 btg",
                key="sticker_order_version",
            )

            col_s6, col_s7, col_s8 = st.columns([1.0, 1.2, 1.3])
            in_jumlah = col_s6.number_input(
                "Jumlah Sticker / Vinyl",
                min_value=0,
                step=1,
                value=1,
                key="sticker_order_qty",
            )
            in_tgl_order = col_s7.date_input(
                "Tanggal OrderPenggantian Sticker / Vinyl",
                value=date.today(),
                key="sticker_order_date",
            )
            in_nama_vendor = col_s8.text_input(
                "Nama Vendor",
                placeholder="Contoh: ARTISTIC atau 77 ADVERTISING",
                key="sticker_order_vendor",
            )

            st.markdown("")
            submitted_sticker = st.form_submit_button(
                "💾 Simpan & Update Data Customer",
                type="primary",
                use_container_width=True,
            )

        if submitted_sticker:
            if not in_nama_outlet.strip():
                st.error("⚠️ **Nama Outlet** wajib diisi.")
            else:
                order_row = pd.DataFrame([{
                    "ID Outlet": in_id_outlet.strip() or pd.NA,
                    "Nama Outlet": in_nama_outlet.strip(),
                    "Material": in_material.strip() or pd.NA,
                    "Brand": in_brand.strip() or pd.NA,
                    "Version": in_version.strip() or pd.NA,
                    "Jumlah Sticker / Vinyl": int(in_jumlah),
                    "Tanggal OrderPenggantian Sticker / Vinyl": pd.Timestamp(in_tgl_order) if in_tgl_order else pd.Timestamp(date.today()),
                    "Nama Vendor": in_nama_vendor.strip() or pd.NA,
                }])
                try:
                    updated_df, upd_cnt, _, _ = apply_sticker_replacement_orders(
                        st.session_state.contracts, order_row
                    )
                    if upd_cnt == 0:
                        st.warning(
                            f"⚠️ Outlet **{in_nama_outlet.strip()}** (ID: `{in_id_outlet.strip() or '-'}`) tidak ditemukan pada data customer yang ada. "
                            "Pastikan ID Outlet atau Nama Outlet sesuai dengan data customer."
                        )
                    else:
                        save_data(updated_df)
                        st.session_state.contracts = updated_df
                        st.session_state.load_warning = None
                        parsed_sz = extract_sizes(in_material) if in_material.strip() else "-"
                        st.success(
                            f"✅ Berhasil mengupdate **{upd_cnt}** data customer untuk **{in_nama_outlet.strip()}** "
                            f"(Dimensi terbaca: `{parsed_sz}`, Qty: `{int(in_jumlah)}`). Data sudah terupdate di **Semua Data Customer Terfilter**."
                        )
                        st.rerun()
                except Exception as error:
                    st.error(f"Gagal menyimpan pemesanan penggantian sticker: {error}")

        st.markdown("##### Riwayat Penggantian Sticker Panel Terbaru")
        sticker_history = contracts[
            contracts["Tanggal Penggantian Sticker Panel"].notna()
            | contracts["Ukuran Sticker Panel Toko"].notna()
            | contracts["Brand Panel"].notna()
            | contracts["Versi"].notna()
            | contracts["Nama Vendor"].notna()
        ].copy()
        if sticker_history.empty:
            st.info("Belum ada data penggantian sticker panel yang tercatat.")
        else:
            sticker_history = sticker_history.sort_values(
                "Tanggal Penggantian Sticker Panel", ascending=False, na_position="last"
            )
            hist_cols = [
                "ID Outlet",
                "Nama Customer",
                "Zona",
                "Ukuran Sticker Panel Toko",
                "Brand Panel",
                "Versi",
                "Tanggal Penggantian Sticker Panel",
                "Nama Vendor",
            ]
            hist_disp = sticker_history[hist_cols].copy()
            hist_disp["Tanggal Penggantian Sticker Panel"] = pd.to_datetime(
                hist_disp["Tanggal Penggantian Sticker Panel"], errors="coerce"
            ).map(lambda d: "" if pd.isna(d) else d.strftime("%d/%m/%Y"))
            st.dataframe(
                hist_disp,
                use_container_width=True,
                hide_index=True,
                height=300,
                column_config={
                    "Nama Customer": st.column_config.TextColumn("Nama Outlet", width="large"),
                    "Ukuran Sticker Panel Toko": st.column_config.TextColumn("Ukuran Sticker", width="medium"),
                    "Brand Panel": st.column_config.TextColumn("Brand", width="medium"),
                    "Versi": st.column_config.TextColumn("Version", width="medium"),
                    "Tanggal Penggantian Sticker Panel": st.column_config.TextColumn("Tanggal OrderPenggantian Sticker / Vinyl", width="medium"),
                    "Nama Vendor": st.column_config.TextColumn("Nama Vendor", width="medium"),
                },
            )

    with tab_sticker_batch:
        st.markdown("##### Langkah 1: Download Template Excel Penggantian Sticker Panel")
        sticker_sample = pd.DataFrame([
            {
                "ID Outlet": "501X0001",
                "Nama Outlet": "TOKO CONTOH 1",
                "Material": "Vinyl Panel Toko 1 x 2 m (BL)",
                "Brand": "DJARUM BLACK MATCHA KRETEK",
                "Version": "Matcha Powder - Rp 17.000/12 btg",
                "Jumlah Sticker / Vinyl": 2,
                "Tanggal OrderPenggantian Sticker / Vinyl": pd.Timestamp("2026-11-05"),
                "Nama Vendor": "ARTISTIC",
            },
            {
                "ID Outlet": "S50113D250917120644415",
                "Nama Outlet": "TOKO CONTOH 2",
                "Material": "Sticker Panel Toko 1 x 2 m",
                "Brand": "LA LIGHTS",
                "Version": "Cat Eyes",
                "Jumlah Sticker / Vinyl": 1,
                "Tanggal OrderPenggantian Sticker / Vinyl": pd.Timestamp("2026-11-05"),
                "Nama Vendor": "77 ADVERTISING",
            },
        ])[STICKER_ORDER_COLUMNS]

        tpl_col1, tpl_col2 = st.columns(2)
        tpl_col1.download_button(
            "📄 Download Template Kosong (.xlsx)",
            data=excel_bytes(pd.DataFrame(columns=STICKER_ORDER_COLUMNS), sheet_name="Penggantian Sticker Panel"),
            file_name="Template_Penggantian_Sticker_Panel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        tpl_col2.download_button(
            "📝 Download Template + Contoh (.xlsx)",
            data=excel_bytes(sticker_sample, sheet_name="Penggantian Sticker Panel"),
            file_name="Template_Penggantian_Sticker_Panel_Contoh.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.divider()
        st.markdown("##### Langkah 2: Upload File Excel Penggantian Massal")
        sticker_batch_file = st.file_uploader(
            "Pilih file Excel (.xlsx) penggantian sticker panel",
            type=["xlsx"],
            key="sticker_batch_uploader",
        )

        if st.button("🚀 Proses & Update Massal Sticker Panel", type="primary", key="btn_process_sticker_batch"):
            if sticker_batch_file is None:
                st.warning("Pilih file Excel terlebih dahulu.")
            else:
                try:
                    orders_df = read_sticker_order_excel(sticker_batch_file)
                    if orders_df.empty:
                        st.warning("Tidak ditemukan baris data pemesanan pada file Excel.")
                    else:
                        updated_df, upd_cnt, _, unmatched = apply_sticker_replacement_orders(
                            st.session_state.contracts, orders_df
                        )
                        if upd_cnt > 0:
                            save_data(updated_df)
                            st.session_state.contracts = updated_df
                            st.session_state.load_warning = None
                        st.success(
                            f"✅ Proses selesai: **{upd_cnt}** baris data customer terupdate di **Semua Data Customer Terfilter** "
                            f"(dari **{len(orders_df)}** baris pemesanan)."
                        )
                        if unmatched:
                            st.warning(
                                f"⚠️ **{len(unmatched)}** baris tidak ditemukan kecocokan ID/Nama Outlet: {', '.join(unmatched[:20])}"
                            )
                        if upd_cnt > 0:
                            st.rerun()
                except Exception as error:
                    st.error(f"Gagal memproses file Excel penggantian sticker: {error}")


# ==========================================
# HALAMAN 4: IMPORT & SINKRONISASI
# ==========================================
elif active_menu == "🔄 Import & Sinkronisasi":
    tab_main_import, tab_branding, tab_drp_crosscheck, tab_panel_sync = st.tabs([
        "📥 Replace Workbook Utama",
        "🏷️ Update SPKS Branding",
        "🔎 Crosscheck Data DRP",
        "📐 Sinkronisasi Ukuran Panel & Layar",
    ])

    with tab_main_import:
        st.info("Fitur ini akan **mengganti seluruh database** dengan isi file Excel yang Anda unggah.")
        col_up, col_tpl = st.columns([2, 1])
        with col_up:
            upload = st.file_uploader("Upload Workbook Excel (.xlsx)", type=["xlsx"], key="workbook_upload")
            if st.button("Import & Ganti Semua Data", type="primary", key="import_contract_workbook"):
                if upload is None:
                    st.warning("Pilih file Excel terlebih dahulu.")
                else:
                    try:
                        imported = read_workbook(upload)
                        problems = validate_data(imported)
                        if problems:
                            st.error("\n".join(problems))
                        else:
                            save_data(imported)
                            st.session_state.contracts = imported
                            st.session_state.load_warning = None
                            st.success(f"✅ Berhasil mengimpor {len(imported)} baris data.")
                            st.rerun()
                    except Exception as error:
                        st.error(f"Gagal mengimpor Excel: {error}")
        with col_tpl:
            st.markdown("##### Butuh Format Standar?")
            st.download_button(
                "📄 Download Template Utama",
                data=excel_bytes(pd.DataFrame(columns=COLUMNS)),
                file_name="Template_Monitoring_Kontrak_Outlet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with tab_branding:
        st.caption(
            "Mengupdate data customer yang cocok berdasarkan **ID Customer** (mengambil tanggal pekerjaan paling baru jika terdapat duplikat)."
        )
        branding_upload = st.file_uploader("Upload Laporan SPKS Branding (.xlsx)", type=["xlsx"], key="branding_workbook_upload")
        if st.button("Proses & Update Data Customer", type="primary"):
            st.session_state.pop("branding_import_result", None)
            if branding_upload is None:
                st.warning("Pilih file laporan branding terlebih dahulu.")
            else:
                try:
                    branding_data = read_branding_workbook(branding_upload)
                    updated, unmatched, duplicates, updated_count = apply_branding_updates(
                        st.session_state.contracts, branding_data
                    )
                    problems = validate_data(updated)
                    if problems:
                        st.error("\n".join(problems))
                    else:
                        if updated_count:
                            save_data(updated)
                            st.session_state.contracts = updated
                        unmatched_has_id = unmatched["ID Customer"].map(customer_id_key).ne("")
                        st.session_state.branding_import_result = {
                            "duplicate_ids": int(duplicates["ID Customer"].nunique()),
                            "duplicate_rows": len(duplicates),
                            "updated_count": updated_count,
                            "unmatched_id_count": int(unmatched_has_id.sum()),
                            "missing_id_count": int((~unmatched_has_id).sum()),
                            "unmatched": unmatched,
                            "duplicates": duplicates,
                        }
                        st.success(f"✅ Pembaruan selesai: {updated_count} ID berhasil diperbarui.")
                except Exception as error:
                    st.error(f"Gagal memproses laporan branding: {error}")

        import_result = st.session_state.get("branding_import_result")
        if import_result:
            s1, s2, s3, s4 = st.columns(4)
            s1.metric("ID Duplikat", f"{import_result['duplicate_ids']} ({import_result['duplicate_rows']} brs)")
            s2.metric("ID Diperbarui", import_result["updated_count"])
            s3.metric("ID Tidak Ditemukan", import_result["unmatched_id_count"])
            s4.metric("Baris Tanpa ID", import_result["missing_id_count"])

            if import_result["duplicate_rows"]:
                with st.expander("Lihat Detail ID Duplikat"):
                    st.dataframe(
                        import_result["duplicates"][["ID Customer", "Customer", "Tanggal Pekerjaan Sampai", "No. Dokumen", "Dipakai"]],
                        use_container_width=True,
                        hide_index=True,
                    )
            missing_ids = import_result["unmatched"][import_result["unmatched"]["ID Customer"].map(customer_id_key).ne("")]
            if not missing_ids.empty:
                with st.expander(f"Lihat {len(missing_ids)} ID yang Tidak Ditemukan di Database"):
                    st.dataframe(missing_ids, use_container_width=True, hide_index=True)

    with tab_drp_crosscheck:
        st.markdown("#### 🔎 Crosscheck Data DRP dengan Dashboard & Monitoring")
        st.caption(
            "Unggah file Excel DRP (contoh: `CODPR03 (DRP).xlsx`) untuk mencocokkan data berdasarkan **ID Customer**, "
            "memperbarui data yang tidak sesuai, dan menambahkan data yang belum ada ke Dashboard."
        )

        drp_upload = st.file_uploader(
            "Upload File Excel DRP (.xlsx)",
            type=["xlsx"],
            key="drp_crosscheck_upload",
        )

        available_compare_fields = [
            "Nama Customer",
            "Brand",
            "Nominal Kontrak",
            "Periode Kontrak Awal",
            "Periode Kontrak Akhir",
            "No. Dokumen",
            "Zona",
            "Kategori Outlet",
        ]
        selected_compare_fields = st.multiselect(
            "Pilih Kolom yang Dicocokkan",
            options=available_compare_fields,
            default=["Nama Customer", "Brand", "Nominal Kontrak", "Periode Kontrak Awal", "Periode Kontrak Akhir", "No. Dokumen"],
            key="drp_compare_fields",
        )

        if st.button("🚀 Jalankan Crosscheck DRP", type="primary", key="btn_run_drp_crosscheck"):
            if drp_upload is None:
                st.warning("Pilih file Excel DRP terlebih dahulu.")
            elif not selected_compare_fields:
                st.warning("Pilih minimal satu kolom untuk dicocokkan.")
            else:
                try:
                    drp_parsed = read_drp_crosscheck_workbook(drp_upload)
                    st.session_state["drp_crosscheck_source"] = drp_parsed
                    st.session_state["drp_crosscheck_fields"] = list(selected_compare_fields)
                except Exception as error:
                    st.error(f"Gagal membaca file DRP: {error}")

        drp_source_df = st.session_state.get("drp_crosscheck_source")
        active_fields = st.session_state.get("drp_crosscheck_fields", selected_compare_fields)

        if drp_source_df is not None and not drp_source_df.empty:
            mismatch_df, matched_ok_df, missing_df = build_drp_crosscheck_report(
                st.session_state.contracts,
                drp_source_df,
                active_fields,
            )

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Data DRP", len(drp_source_df))
            m2.metric("Sesuai Semua", len(matched_ok_df))
            m3.metric("Tidak Sesuai", len(mismatch_df))
            m4.metric("ID Belum Ada di Dashboard", len(missing_df))

            sub_mismatch, sub_ok, sub_missing = st.tabs([
                f"⚠️ Tidak Sesuai ({len(mismatch_df)})",
                f"✅ Sudah Sesuai ({len(matched_ok_df)})",
                f"❓ ID Tidak di Dashboard ({len(missing_df)})",
            ])

            with sub_mismatch:
                if mismatch_df.empty:
                    st.success("✅ Seluruh data DRP yang cocok dengan ID di Dashboard sudah sesuai!")
                else:
                    st.markdown("**Data Tidak Sesuai — Pilih Baris & Kolom untuk Update ke Dashboard**")
                    fields_to_apply = st.multiselect(
                        "Kolom yang akan di-update ke Dashboard",
                        options=available_compare_fields,
                        default=active_fields,
                        key="drp_fields_to_apply",
                    )
                    select_all_mismatch = st.checkbox(
                        "Pilih semua baris tidak sesuai",
                        value=True,
                        key="drp_select_all_mismatch",
                    )

                    display_cols = [
                        "ID Customer",
                        "Nama Customer (Dashboard)",
                        "Nama Customer (DRP)",
                        "Kolom Tidak Sesuai",
                        "Brand (Dash)",
                        "Brand (DRP)",
                        "Nominal (Dash)",
                        "Nominal (DRP)",
                        "Awal Kontrak (Dash)",
                        "Awal Kontrak (DRP)",
                        "Akhir Kontrak (Dash)",
                        "Akhir Kontrak (DRP)",
                        "No. Dokumen (Dash)",
                        "No. Dokumen (DRP)",
                        "Zona (Dash)",
                        "Zona (DRP)",
                        "Kategori Outlet (Dash)",
                        "Kategori Outlet (DRP)",
                    ]

                    editor_df = mismatch_df[display_cols].copy()
                    editor_df.insert(0, "Pilih Update", select_all_mismatch)

                    edited_mismatch = st.data_editor(
                        editor_df,
                        use_container_width=True,
                        hide_index=True,
                        height=420,
                        disabled=display_cols,
                        key="drp_mismatch_editor",
                    )

                    dl_col, upd_col = st.columns(2)
                    dl_col.download_button(
                        "📥 Download Laporan Selisih DRP (.xlsx)",
                        data=excel_bytes(mismatch_df[display_cols]),
                        file_name="Crosscheck_Selisih_DRP.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                    if upd_col.button(
                        "💾 Update Data Terpilih ke Dashboard",
                        type="primary",
                        use_container_width=True,
                        key="btn_apply_drp_updates",
                    ):
                        if not fields_to_apply:
                            st.warning("Pilih minimal satu kolom yang akan di-update.")
                        else:
                            selected_mask = edited_mismatch["Pilih Update"].fillna(False).to_numpy(dtype=bool)
                            rows_to_update = mismatch_df.loc[selected_mask]
                            if rows_to_update.empty:
                                st.warning("Centang minimal satu baris data untuk di-update.")
                            else:
                                updated_contracts = normalize_data(st.session_state.contracts).copy()
                                updated_counter = 0
                                for _, r in rows_to_update.iterrows():
                                    target_idx = int(r["_dash_index"])
                                    changed = False
                                    for col in fields_to_apply:
                                        raw_val = r.get(f"_raw_{col}", pd.NA)
                                        if pd.isna(raw_val) or (isinstance(raw_val, str) and not raw_val.strip()):
                                            continue
                                        if _is_field_different(col, updated_contracts.at[target_idx, col], raw_val):
                                            updated_contracts.at[target_idx, col] = raw_val
                                            changed = True
                                    if changed:
                                        updated_counter += 1

                                problems = validate_data(updated_contracts)
                                if problems:
                                    st.error("\n".join(problems))
                                else:
                                    save_data(updated_contracts)
                                    st.session_state.contracts = normalize_data(updated_contracts)
                                    st.success(f"✅ Berhasil meng-update {updated_counter} outlet di Dashboard sesuai data DRP!")
                                    st.rerun()

            with sub_ok:
                if matched_ok_df.empty:
                    st.info("Belum ada data yang sesuai sepenuhnya pada kolom yang dipilih.")
                else:
                    ok_cols = [
                        "ID Customer",
                        "Nama Customer (Dashboard)",
                        "Brand (DRP)",
                        "Nominal (DRP)",
                        "Awal Kontrak (DRP)",
                        "Akhir Kontrak (DRP)",
                        "No. Dokumen (DRP)",
                    ]
                    st.dataframe(matched_ok_df[ok_cols], use_container_width=True, hide_index=True)

            with sub_missing:
                if missing_df.empty:
                    st.info("Seluruh ID Customer pada file DRP sudah terdaftar di Dashboard.")
                else:
                    st.markdown("**Data DRP Belum Ada di Dashboard — Pilih Baris untuk Ditambahkan Langsung**")
                    st.caption("Centang baris pada kolom **Pilih Tambah**, atur informasi bawaan (opsional), lalu klik **Tambahkan Data Terpilih ke Dashboard**.")

                    m_cfg1, m_cfg2, m_cfg3 = st.columns(3)
                    default_sector = m_cfg1.text_input("Sektor Bawaan", value="Retail Supply", key="drp_add_default_sector")
                    default_budget = m_cfg2.text_input("Pos Budget Bawaan", value="Small Retailer", key="drp_add_default_budget")
                    default_job_cat = m_cfg3.text_input("Kategori Pekerjaan Bawaan", value="DRP", key="drp_add_default_job_cat")

                    select_all_missing = st.checkbox(
                        "Pilih semua baris yang belum ada di Dashboard",
                        value=False,
                        key="drp_select_all_missing",
                    )

                    missing_view_df = pd.DataFrame({
                        "ID Customer": missing_df["ID Customer"].map(safe_text),
                        "Nama Customer": missing_df["Customer"].map(safe_text),
                        "Zona": missing_df["Zona"].map(safe_text),
                        "Kategori Outlet": missing_df["Kategori Outlet"].map(safe_text),
                        "Kategori DRP": missing_df["Kategori DRP"].map(safe_text),
                        "Brand": missing_df["Brand"].map(safe_text),
                        "Nominal Kontrak": missing_df["Nominal Kontrak"].map(lambda v: _fmt_val_cmp("Nominal Kontrak", v)),
                        "Periode Kontrak Awal": missing_df["Periode Kontrak Awal"].map(lambda v: _fmt_val_cmp("Periode Kontrak Awal", v)),
                        "Periode Kontrak Akhir": missing_df["Periode Kontrak Akhir"].map(lambda v: _fmt_val_cmp("Periode Kontrak Akhir", v)),
                        "No. Dokumen": missing_df["No. Dokumen"].map(safe_text),
                    })
                    missing_view_df.insert(0, "Pilih Tambah", select_all_missing)

                    edited_missing = st.data_editor(
                        missing_view_df,
                        use_container_width=True,
                        hide_index=True,
                        height=400,
                        disabled=list(missing_view_df.columns[1:]),
                        key="drp_missing_editor",
                    )

                    m_dl_col, m_add_col = st.columns(2)
                    m_dl_col.download_button(
                        "📥 Download Data Belum Ada di Dashboard (.xlsx)",
                        data=excel_bytes(missing_view_df.drop(columns=["Pilih Tambah"])),
                        file_name="DRP_Belum_Ada_Di_Dashboard.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                    if m_add_col.button(
                        "➕ Tambahkan Data Terpilih ke Dashboard",
                        type="primary",
                        use_container_width=True,
                        key="btn_add_missing_drp_to_dash",
                    ):
                        add_mask = edited_missing["Pilih Tambah"].fillna(False).to_numpy(dtype=bool)
                        rows_to_insert_raw = missing_df.loc[add_mask]
                        if rows_to_insert_raw.empty:
                            st.warning("Centang minimal satu baris data untuk ditambahkan ke Dashboard.")
                        else:
                            current_contracts = normalize_data(st.session_state.contracts).copy()
                            existing_keys = {
                                customer_id_key(val)
                                for val in current_contracts["ID Outlet"].dropna()
                                if customer_id_key(val)
                            }
                            new_records = []
                            skipped_existing = 0

                            for _, r in rows_to_insert_raw.iterrows():
                                k = customer_id_key(r.get("ID Customer"))
                                if k and k in existing_keys:
                                    skipped_existing += 1
                                    continue
                                if k:
                                    existing_keys.add(k)

                                new_records.append({
                                    "ID Outlet": r.get("ID Customer", pd.NA),
                                    "Nama Customer": r.get("Customer", pd.NA),
                                    "Alamat": pd.NA,
                                    "Zona": r.get("Zona", pd.NA),
                                    "Sektor": default_sector.strip() or pd.NA,
                                    "Pos Budget": default_budget.strip() or pd.NA,
                                    "Kategori Outlet": r.get("Kategori Outlet", pd.NA),
                                    "Kategori Pekerjaan": default_job_cat.strip() or r.get("Kategori DRP", pd.NA),
                                    "Brand": r.get("Brand", pd.NA),
                                    "Nominal Kontrak": r.get("Nominal Kontrak", pd.NA),
                                    "Periode Kontrak Awal": r.get("Periode Kontrak Awal", pd.NaT),
                                    "Periode Kontrak Akhir": r.get("Periode Kontrak Akhir", pd.NaT),
                                    "No. Dokumen": r.get("No. Dokumen", pd.NA),
                                    "Panel Toko": False,
                                    "Showcase": False,
                                    "Layar Toko": True,
                                    "Shop Sign / Neon Box": False,
                                    "Ukuran Panel Toko": pd.NA,
                                    "Ukuran Sticker Panel Toko": pd.NA,
                                    "Brand Panel": pd.NA,
                                    "Versi": pd.NA,
                                    "Tanggal Penggantian Sticker Panel": pd.NaT,
                                    "Nama Vendor": pd.NA,
                                    "Ukuran Layar": pd.NA,
                                    "Brand Layar": pd.NA,
                                })

                            if not new_records:
                                st.info("Tidak ada baris baru ditambahkan karena ID sudah terdaftar di Dashboard.")
                            else:
                                combined_contracts = pd.concat(
                                    [current_contracts[COLUMNS], pd.DataFrame(new_records)],
                                    ignore_index=True,
                                )
                                problems = validate_data(combined_contracts)
                                if problems:
                                    st.error("\n".join(problems))
                                else:
                                    save_data(combined_contracts)
                                    st.session_state.contracts = normalize_data(combined_contracts)
                                    msg = f"✅ Berhasil menambahkan {len(new_records)} outlet baru ke Dashboard!"
                                    if skipped_existing:
                                        msg += f" ({skipped_existing} ID dilewati karena sudah ada)."
                                    st.success(msg)
                                    st.rerun()

    with tab_panel_sync:
        st.caption("Mencocokkan ukuran sticker panel, brand panel, versi, dan ukuran layar dari sheet DRP berdasarkan ID Customer.")
        sync_upload = st.file_uploader("Upload File Sumber Panel & Layar (.xlsx)", type=["xlsx"], key="sync_panel_screen_upload")
        overwrite_sync = st.checkbox("Timpa nilai yang sudah terisi jika ID Customer cocok", value=False, key="overwrite_panel_screen_sync")
        if st.button("🔗 Hubungkan & Sinkronkan", type="primary", key="btn_sync_panel_screen"):
            if sync_upload is None:
                st.warning("Pilih file Excel sumber terlebih dahulu.")
            else:
                try:
                    drp_source = read_drp_sheet(sync_upload)
                    if drp_source is not None and not drp_source.empty:
                        source_frame = drp_source
                    else:
                        sync_upload.seek(0)
                        source_frame = read_workbook(sync_upload)
                    merged_df, updated_cnt, missing_list = merge_panel_screen_by_id(
                        st.session_state.contracts,
                        source_frame,
                        overwrite_existing=overwrite_sync,
                    )
                    save_data(merged_df)
                    st.session_state.contracts = merged_df
                    st.success(f"✅ Sinkronisasi selesai: {updated_cnt} baris customer diperbarui.")
                    if missing_list:
                        st.info(f"{len(missing_list)} ID tidak ditemukan: {', '.join(missing_list[:15])}")
                except Exception as error:
                    st.error(f"Gagal sinkronisasi: {error}")


# ==========================================
# HALAMAN 4: BACKUP & RESTORE CLOUD
# ==========================================
elif active_menu == "☁️ Backup & Restore":
    col_sb, col_gs = st.columns(2)

    with col_sb:
        with st.container(border=True):
            st.subheader("🗄️ Supabase Cloud")
            st.caption("Simpan snapshot database atau pulihkan data terakhir dari Supabase.")
            sb_url, sb_key, sb_table = _supabase_config()
            if not sb_url or not sb_key:
                st.warning("Konfigurasi `SUPABASE_URL` dan `SUPABASE_SERVICE_ROLE_KEY` belum diatur di Streamlit Secrets.")
            else:
                st.info(f"Tabel: `{sb_table}` • Data saat ini: **{len(st.session_state.contracts)}** outlet")
                b1, b2 = st.columns(2)
                if b1.button("⬆️ Backup Sekarang", type="primary", use_container_width=True, key="btn_backup_supabase"):
                    try:
                        saved_cnt = backup_to_supabase(st.session_state.contracts, sb_url, sb_key, sb_table)
                        st.success(f"✅ {saved_cnt} baris tersimpan.")
                    except Exception as error:
                        st.error(f"Gagal backup: {error}")

                if b2.button("⬇️ Restore Data", type="secondary", use_container_width=True, key="btn_restore_supabase"):
                    try:
                        restored_df = restore_from_supabase(sb_url, sb_key, sb_table)
                        problems = validate_data(restored_df)
                        if problems:
                            st.error("\n".join(problems))
                        else:
                            save_data(restored_df)
                            st.session_state.contracts = restored_df
                            st.session_state.load_warning = None
                            st.success(f"✅ {len(restored_df)} baris dipulihkan.")
                            st.rerun()
                    except Exception as error:
                        st.error(f"Gagal restore: {error}")

    with col_gs:
        with st.container(border=True):
            st.subheader("📗 Google Sheets")
            st.caption(f"Kirim seluruh data aktif ke tab `{SHEET_NAME}` melalui Google Apps Script.")
            try:
                backup_url = str(st.secrets.get("GOOGLE_APPS_SCRIPT_URL", "") or "").strip()
            except FileNotFoundError:
                backup_url = ""
            if not backup_url:
                st.warning("Konfigurasi `GOOGLE_APPS_SCRIPT_URL` belum diatur di Streamlit Secrets.")
            else:
                st.info(f"Siap mengirim **{len(st.session_state.contracts)}** data outlet (akan menimpa tab tujuan).")
                if st.button("⬆️ Backup ke Google Sheet", type="primary", use_container_width=True):
                    try:
                        row_count = backup_to_google_sheet(st.session_state.contracts, backup_url)
                        st.success(f"✅ {row_count} baris berhasil dikirim ke Google Sheet.")
                    except Exception as error:
                        st.error(f"Gagal backup: {error}")
