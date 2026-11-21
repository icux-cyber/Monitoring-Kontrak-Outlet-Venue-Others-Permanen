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
    "Ukuran Layar",
    "Brand Layar",
]
COLUMNS = [
    "ID Outlet", "Nama Customer", "Alamat", "Zona", "Sektor", "Pos Budget",
    "Kategori Outlet", "Kategori Pekerjaan", "Brand", "Nominal Kontrak",
    "Periode Kontrak Awal", "Periode Kontrak Akhir", *MATERIAL_COLUMNS,
    "Ukuran Panel Toko", "Ukuran Sticker Panel Toko", "Brand Panel",
    "Ukuran Layar", "Brand Layar",
]
CONTENT_COLUMNS = [
    "ID Outlet", "Nama Customer", "Alamat", "Zona", "Sektor", "Pos Budget",
    "Kategori Outlet", "Kategori Pekerjaan", "Brand", "Nominal Kontrak",
    "Periode Kontrak Awal", "Periode Kontrak Akhir",
    "Ukuran Panel Toko", "Ukuran Sticker Panel Toko", "Brand Panel",
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
    "Nominal Kontrak": ("Nominal Kontrak", "Nominal", "Nominal.1", "Contract Amount", "Nilai Kontrak"),
    "Periode Kontrak Awal": ("Periode Kontrak Awal", "Periode Awal Kontrak", "Mulai", "Tanggal Mulai", "Start Date"),
    "Periode Kontrak Akhir": ("Periode Kontrak Akhir", "Periode Akhir Kontrak", "Akhir", "Tanggal Akhir", "End Date", "Expiry Date"),
    "Panel Toko": ("Panel Toko", "Materi Panel Toko"),
    "Showcase": ("Showcase", "Show Case", "Materi Showcase"),
    "Layar Toko": ("Layar Toko", "Layar", "Materi Layar Toko"),
    "Shop Sign / Neon Box": ("Shop Sign / Neon Box", "Shop Sign", "Neon Box", "Materi Shop Sign / Neon Box"),
    "Ukuran Panel Toko": ("Ukuran Panel Toko", "Size Panel Toko", "Panel"),
    "Ukuran Sticker Panel Toko": (
        "Ukuran Sticker Panel Toko",
        "Ukuran Sticker Panel (diluar Nama Toko)",
        "Ukuran Stiker Panel Toko",
        "Size Sticker Panel Toko",
    ),
    "Brand Panel": ("Brand Panel", "Brand Planning - Sticker Panel", "Merek Panel"),
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
        /* Latar belakang & tipografi utama */
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
            max-width: 96%;
        }

        /* Header Banner */
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

        /* KPI Cards */
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

        /* Section Card */
        .section-box {
            background-color: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 1.2rem;
            margin-bottom: 1rem;
        }

        /* Sidebar styling */
        section[data-testid="stSidebar"] {
            background-color: #f8fafc;
            border-right: 1px solid #e2e8f0;
        }
        div[data-testid="stSidebarNav"] {
            padding-top: 1rem;
        }

        /* Tab styling */
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

        /* Dark mode compatibility */
        @media (prefers-color-scheme: dark) {
            .kpi-card {
                background-color: #1e293b;
                border-color: #334155;
            }
            .kpi-value { color: #f8fafc; }
            .kpi-label, .kpi-sub { color: #94a3b8; }
            .section-box {
                background-color: #1e293b;
                border-color: #334155;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


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
    if pd.isna(value) or str(value).strip() == "":
        return pd.NA
    text = str(value).strip()
    matches = re.findall(
        r"(?<!\d)(\d+(?:[.,]\d+)?)\s*[x×]\s*(\d+(?:[.,]\d+)?)\s*m\b",
        text,
        flags=re.IGNORECASE,
    )
    if not matches:
        return text
    sizes = dict.fromkeys(f"{width} x {height} m" for width, height in matches)
    return ", ".join(sizes)


def customer_id_key(value) -> str:
    if pd.isna(value):
        return ""
    key = str(value).strip()
    if re.fullmatch(r"\d+\.0", key):
        key = key[:-2]
    return key.casefold()


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
        and column not in ("Nominal Kontrak", "Periode Kontrak Awal", "Periode Kontrak Akhir")
    ]
    for column in text_columns:
        normalized[column] = normalized[column].astype("string").str.strip()
        normalized[column] = normalized[column].replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA})
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
    for column in ("Periode Kontrak Awal", "Periode Kontrak Akhir"):
        normalized[column] = pd.to_datetime(normalized[column], errors="coerce", dayfirst=True).dt.tz_localize(None)
    normalized = normalized[COLUMNS]
    existing_content_cols = [col for col in CONTENT_COLUMNS if col in normalized.columns]
    has_data = normalized[existing_content_cols].notna().any(axis=1)
    return normalized.loc[has_data.to_numpy(dtype=bool)].reset_index(drop=True)


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


def excel_bytes(frame: pd.DataFrame) -> bytes:
    output = BytesIO()
    export = frame.copy()
    for column in MATERIAL_COLUMNS:
        if column in export:
            export[column] = export[column].map(lambda checked: "✓" if parse_tickmark(checked) else "")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export.to_excel(writer, sheet_name=SHEET_NAME, index=False)
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
    for column in ("Periode Kontrak Awal", "Periode Kontrak Akhir"):
        if column in display_frame:
            display_frame[column] = pd.to_datetime(display_frame[column], errors="coerce").map(
                lambda value: "" if pd.isna(value) else value.strftime("%d/%m/%Y")
            )
    for column in MATERIAL_COLUMNS:
        if column in display_frame:
            display_frame[column] = display_frame[column].map(lambda checked: "✅" if parse_tickmark(checked) else "—")
    # Atur urutan kolom agar informasi paling krusial berada di depan
    priority_cols = [
        "Status Kontrak", "Sisa Hari", "ID Outlet", "Nama Customer",
        "Brand", "Nominal Kontrak", "Periode Kontrak Awal", "Periode Kontrak Akhir",
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
        k1, k2, k3, k4 = st.columns(4)
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

        st.divider()
        st.markdown("##### 3. Materi Promosi & Spesifikasi Ukuran")
        material_cols = st.columns(4)
        material_values = {
            column: material_cols[index].checkbox(column, value=parse_tickmark(current[column]), key=f"material_{index}_{suffix}")
            for index, column in enumerate(MATERIAL_COLUMNS)
        }

        p1, p2, p3 = st.columns(3)
        panel_size = p1.text_input("Ukuran Panel Toko", value=safe_text(current["Ukuran Panel Toko"]), placeholder="1 x 3 m", key=f"panel_size_{suffix}")
        sticker_size = p2.text_input("Ukuran Sticker Panel", value=safe_text(current["Ukuran Sticker Panel Toko"]), placeholder="1 x 2 m", key=f"sticker_size_{suffix}")
        panel_brand = p3.text_input("Brand Panel", value=safe_text(current["Brand Panel"]), key=f"panel_brand_{suffix}")

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
        **material_values,
        "Ukuran Panel Toko": panel_size.strip() or pd.NA,
        "Ukuran Sticker Panel Toko": sticker_size.strip() or pd.NA,
        "Brand Panel": panel_brand.strip() or pd.NA,
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
    # KPI CARDS
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

    # NOTIFIKASI RINGKAS
    if expired_count or critical_count:
        alert_col1, alert_col2 = st.columns(2)
        if expired_count:
            alert_col1.error(f"🚨 **{expired_count} kontrak** sudah melewati masa berlaku (Expired).")
        if critical_count:
            alert_col2.warning(f"⚠️ **{critical_count} kontrak** memasuki masa kritis (berakhir dalam ≤ 30 hari).")

    # PENCARIAN & FILTER
    st.markdown("### 🔍 Pencarian & Filter Data")
    search_col, status_col = st.columns([2, 1])
    search_query = search_col.text_input(
        "Cari cepat (ID Outlet, Nama Customer, Alamat, atau Brand)",
        placeholder="Ketik kata kunci pencarian...",
    )
    statuses = [
        "Semua Status",
        "🔴 Expired",
        "🟠 Kritis (≤ 30 hari)",
        "🟡 Warning (31–90 hari)",
        "🟢 Aktif (> 90 hari)",
        "⚪ Data belum lengkap",
    ]
    selected_status = status_col.selectbox("Status Masa Berlaku", statuses)

    with st.expander("🎛️ Filter Lanjutan (Zona, Sektor, Pos Budget, Ukuran Panel/Layar)", expanded=False):
        f1, f2, f3, f4 = st.columns(4)
        zones = ["Semua"] + sorted(contracts["Zona"].dropna().astype(str).unique().tolist())
        sectors = ["Semua"] + sorted(contracts["Sektor"].dropna().astype(str).unique().tolist())
        budget_positions = ["Semua"] + sorted(contracts["Pos Budget"].dropna().astype(str).unique().tolist())
        job_categories = ["Semua"] + sorted(contracts["Kategori Pekerjaan"].dropna().astype(str).unique().tolist())

        selected_zone = f1.selectbox("Zona", zones)
        selected_sector = f2.selectbox("Sektor", sectors)
        selected_budget = f3.selectbox("Pos Budget", budget_positions)
        selected_job_category = f4.selectbox("Kategori Pekerjaan", job_categories)

        f5, f6 = st.columns(2)
        sticker_sizes = ["Semua"] + sorted(contracts["Ukuran Sticker Panel Toko"].dropna().astype(str).unique().tolist())
        screen_sizes = ["Semua"] + sorted(contracts["Ukuran Layar"].dropna().astype(str).unique().tolist())
        selected_sticker_size = f5.selectbox("Ukuran Sticker Panel", sticker_sizes)
        selected_screen_size = f6.selectbox("Ukuran Layar", screen_sizes)

    # TERAPKAN FILTER
    view = contracts.copy()
    if search_query.strip():
        q = search_query.strip().casefold()
        mask = (
            view["ID Outlet"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
            | view["Nama Customer"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
            | view["Alamat"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
            | view["Brand"].fillna("").astype(str).str.casefold().str.contains(q, regex=False)
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

    # TAMPILAN TABEL BERBASIS TAB
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
            render_dashboard_table(
                "Semua Data Customer Terfilter",
                view,
                "Daftar_Semua_Customer_Terfilter.xlsx",
                "dl_all_filtered",
            )
        with tab_retail:
            render_dashboard_table(
                "Daftar Customer - Retail Supply",
                regular_view,
                "Daftar_Customer_Retail_Supply.xlsx",
                "dl_retail_filtered",
            )
        with tab_venue:
            render_dashboard_table(
                "Daftar Customer - Venue",
                venue_view,
                "Daftar_Customer_Venue.xlsx",
                "dl_venue_filtered",
            )
        with tab_other:
            render_dashboard_table(
                "Daftar Customer - Others Permanen",
                other_permanent_view,
                "Daftar_Customer_Others_Permanen.xlsx",
                "dl_other_filtered",
            )


# ==========================================
# HALAMAN 2: KELOLA DATA CUSTOMER
# ==========================================
elif active_menu == "👥 Kelola Data Customer":
    tab_edit, tab_add_manual, tab_add_batch = st.tabs([
        "✏️ Edit / Hapus Customer",
        "➕ Tambah Customer (Form)",
        "📑 Tambah Massal (Excel)",
    ])

    # TAB 1: EDIT & HAPUS
    with tab_edit:
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

    # TAB 2: TAMBAH MANUAL
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

    # TAB 3: TAMBAH MASSAL EXCEL
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
                "Panel Toko": True,
                "Showcase": False,
                "Layar Toko": False,
                "Shop Sign / Neon Box": False,
                "Ukuran Panel Toko": "1 x 3 m",
                "Ukuran Sticker Panel Toko": "1 x 2 m",
                "Brand Panel": "Djarum Super",
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
# HALAMAN 3: IMPORT & SINKRONISASI
# ==========================================
elif active_menu == "🔄 Import & Sinkronisasi":
    tab_main_import, tab_branding, tab_panel_sync = st.tabs([
        "📥 Replace Workbook Utama",
        "🏷️ Update SPKS Branding",
        "📐 Sinkronisasi Ukuran Panel & Layar",
    ])

    # TAB 1: REPLACE WORKBOOK UTAMA
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

    # TAB 2: UPDATE SPKS BRANDING
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

    # TAB 3: SINKRONISASI PANEL & LAYAR
    with tab_panel_sync:
        st.caption("Mencocokkan ukuran sticker panel, brand panel, dan ukuran layar dari sheet DRP berdasarkan ID Customer.")
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

    # KOLOM SUPABASE
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

    # KOLOM GOOGLE SHEETS
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
