from __future__ import annotations

import hashlib
import io
import re
import warnings
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

DATA_DIR = Path("data")
DATA_FILE = DATA_DIR / "DRP & Outlet MCD - DSO Surabaya Barat.xlsx"
SHEET_NAME = "Monitoring Kontrak Outlet"
DRP_SHEET_NAME = "Pemesanan Layar DRP"

MATERIAL_COLUMNS = ["Panel Toko", "Showcase", "Layar Toko", "Shop Sign / Neon Box"]
SYNC_PANEL_SCREEN_COLUMNS = [
    "Ukuran Sticker / Vinyl Panel Toko",
    "Brand Panel",
    "Versi",
    "Ukuran Layar",
    "Brand Layar",
]
DATE_COLUMNS = [
    "Periode Kontrak Awal",
    "Periode Kontrak Akhir",
    "Tanggal Penggantian Sticker / Vinyl Panel",
]
COLUMNS = [
    "ID Outlet", "Nama Customer", "Alamat", "Zona", "Sektor", "Pos Budget",
    "Kategori Outlet", "Kategori Pekerjaan", "Brand", "Nominal Kontrak",
    "Periode Kontrak Awal", "Periode Kontrak Akhir", "No. Dokumen", *MATERIAL_COLUMNS,
    "Ukuran Panel Toko", "Ukuran Sticker / Vinyl Panel Toko", "Brand Panel", "Versi",
    "Tanggal Penggantian Sticker Panel / Vinyl", "Nama Vendor",
    "Ukuran Layar", "Brand Layar",
]
CONTENT_COLUMNS = [
    "ID Outlet", "Nama Customer", "Alamat", "Zona", "Sektor", "Pos Budget",
    "Kategori Outlet", "Kategori Pekerjaan", "Brand", "Nominal Kontrak",
    "Periode Kontrak Awal", "Periode Kontrak Akhir", "No. Dokumen",
    "Ukuran Panel Toko", "Ukuran Sticker / Vinyl Panel Toko", "Brand Panel", "Versi",
    "Tanggal Penggantian Sticker Panel / Vinyl", "Nama Vendor",
    "Ukuran Layar", "Brand Layar",
]
SEARCH_SOURCE_COLUMNS = [
    "ID Outlet", "Nama Customer", "Alamat", "No. Dokumen",
    "Brand", "Brand Panel", "Versi", "Nama Vendor", "Sektor",
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
    "Ukuran Sticker / Vinyl Panel Toko": (
        "Ukuran Sticker / Vinyl Panel Toko",
        "Ukuran Sticker Panel / Vinyl",
        "Ukuran Sticker / Vinyl",
        "Ukuran Sticker / Vinyl",
        "Ukuran Sticker / Vinyl Panel (diluar Nama Toko)",
        "Ukuran Sticker / Vinyl Panel Toko",
        "Size Sticker / Vinyl Panel Toko",
        "Material",
    ),
    "Brand Panel": (
        "Brand Panel",
        "Brand Sticker / Vinyl",
        "Brand Planning - Sticker / Vinyl Panel",
        "Merek Panel",
    ),
    "Versi": (
        "Versi",
        "Version",
        "Versi Sticker / Vinyl",
        "Versi Panel",
    ),
    "Tanggal Penggantian Sticker / Vinyl Panel": (
        "Tanggal Penggantian Sticker Panel",
        "Tanggal OrderPenggantian Sticker / Vinyl",
        "Tanggal Order Penggantian Sticker / Vinyl",
        "Tanggal Penggantian Terakhir Sticker / Vinyl Panel",
        "Tgl Penggantian Sticker / Vinyl Panel",
        "Tgl Penggantian Terakhir",
        "Tanggal Ganti Sticker / Vinyl",
    ),
    "Nama Vendor": (
        "Nama Vendor",
        "Vendor",
        "Nama Vendor Sticker / Vinyl",
    ),
    "Ukuran Layar": ("Ukuran Layar", "Layar", "Materi Layar Toko", "Size Layar"),
    "Brand Layar": ("Brand Layar", "Brand Planning - Layar DRP", "Merek Layar"),
}


def clean_headers(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame


def safe_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.casefold() in {"<na>", "nan", "none"} else text


def safe_date(value):
    return None if pd.isna(value) else pd.Timestamp(value).date()


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
    if key.casefold() in {"", "<na>", "nan", "none"}:
        return ""
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


def dataframe_fingerprint(df: pd.DataFrame | None) -> str:
    """Fingerprint cepat berbasis ukuran, kolom kunci, dan sampel hash baris."""
    if df is None or df.empty:
        return "empty"
    h = hashlib.blake2b(digest_size=12)
    h.update(f"{len(df)}:{len(df.columns)}".encode("utf-8"))
    key_cols = [c for c in ("ID Outlet", "Nama Customer", "Nominal Kontrak", "Periode Kontrak Akhir") if c in df.columns]
    for col in key_cols:
        s = df[col].fillna("").astype(str)
        h.update(col.encode("utf-8"))
        h.update(s.head(25).str.cat(sep="|").encode("utf-8", errors="ignore"))
        if len(s) > 25:
            h.update(s.tail(25).str.cat(sep="|").encode("utf-8", errors="ignore"))
    return h.hexdigest()


def attach_search_blob(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    parts: list[pd.Series] = []
    for col in SEARCH_SOURCE_COLUMNS:
        if col in out.columns:
            parts.append(out[col].fillna("").astype(str).str.casefold())
    if parts:
        combined = parts[0]
        for s in parts[1:]:
            combined = combined + " " + s
        out["_search_blob"] = combined
    else:
        out["_search_blob"] = ""
    return out


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
    for column in ("Ukuran Panel Toko", "Ukuran Sticker / Vinyl Panel Toko", "Ukuran Layar"):
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


def read_drp_sheet(source) -> pd.DataFrame | None:
    try:
        if hasattr(source, "seek"):
            source.seek(0)
        excel_file = pd.ExcelFile(source, engine="openpyxl")
        if DRP_SHEET_NAME not in excel_file.sheet_names:
            return None
        if hasattr(source, "seek"):
            source.seek(0)
        raw_drp = pd.read_excel(source, sheet_name=DRP_SHEET_NAME, header=0, engine="openpyxl")
        return normalize_data(raw_drp)
    except Exception:
        return None


def merge_panel_screen_by_id(
    contracts: pd.DataFrame,
    source_df: pd.DataFrame,
    overwrite_existing: bool = False,
    already_normalized: bool = False,
) -> tuple[pd.DataFrame, int, list[str]]:
    target = contracts.copy() if already_normalized else normalize_data(contracts).copy()
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

    return attach_search_blob(target.reset_index(drop=True)), updated_rows, sorted(set(unmatched_ids))


@st.cache_data(show_spinner=False)
def _read_workbook_from_bytes(file_bytes: bytes) -> pd.DataFrame:
    bio = io.BytesIO(file_bytes)
    last_error = None
    base_df: pd.DataFrame | None = None
    for sheet, header in (("COSPKR06", 3), (SHEET_NAME, 0), (0, 0)):
        try:
            bio.seek(0)
            result = normalize_data(pd.read_excel(bio, sheet_name=sheet, header=header, engine="openpyxl"))
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

    bio.seek(0)
    drp_df = read_drp_sheet(bio)
    if drp_df is not None and not drp_df.empty:
        base_df, _, _ = merge_panel_screen_by_id(base_df, drp_df, overwrite_existing=False, already_normalized=True)

    return attach_search_blob(base_df)


def read_workbook(source) -> pd.DataFrame:
    if isinstance(source, (str, Path)):
        return _read_workbook_from_bytes(Path(source).read_bytes())
    if hasattr(source, "getvalue"):
        return _read_workbook_from_bytes(bytes(source.getvalue()))
    if hasattr(source, "read"):
        if hasattr(source, "seek"):
            source.seek(0)
        return _read_workbook_from_bytes(bytes(source.read()))
    raise ValueError("Sumber workbook tidak dikenali.")


def read_branding_workbook(source) -> pd.DataFrame:
    if hasattr(source, "seek"):
        source.seek(0)
    frame = pd.read_excel(source, sheet_name="COSPKR06", header=11, engine="openpyxl")
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


def apply_branding_updates(
    contracts: pd.DataFrame,
    branding: pd.DataFrame,
    already_normalized: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, int]:
    selected, duplicates = deduplicate_branding_data(branding)
    updated = contracts.copy() if already_normalized else normalize_data(contracts)
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
    return attach_search_blob(updated.reset_index(drop=True)), unmatched, duplicate_details, len(matched_keys)


def load_initial_data() -> tuple[pd.DataFrame, str | None]:
    if not DATA_FILE.exists():
        return attach_search_blob(pd.DataFrame(columns=COLUMNS)), None
    try:
        return attach_search_blob(read_workbook(DATA_FILE)), None
    except Exception as error:
        return (
            attach_search_blob(pd.DataFrame(columns=COLUMNS)),
            f"File data awal tidak dapat dibaca ({error}). Unggah workbook Excel yang valid untuk memulai.",
        )


def validate_data(frame: pd.DataFrame, already_normalized: bool = False) -> list[str]:
    errors = []
    normalized = frame if already_normalized else normalize_data(frame)
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


def compute_contract_status(frame: pd.DataFrame, today: date | None = None) -> pd.DataFrame:
    result = frame.copy()
    end_dates = pd.to_datetime(result["Periode Kontrak Akhir"], errors="coerce")
    ref_day = pd.Timestamp(today or date.today()).normalize()
    remaining = (end_dates - ref_day).dt.days
    result["Sisa Hari"] = remaining

    status_series = pd.Series("⚪ Data belum lengkap", index=result.index, dtype="object")
    status_series = status_series.mask(remaining < 0, "🔴 Expired")
    status_series = status_series.mask((remaining >= 0) & (remaining <= 30), "🟠 Kritis (≤ 30 hari)")
    status_series = status_series.mask((remaining > 30) & (remaining <= 90), "🟡 Warning (31–90 hari)")
    status_series = status_series.mask(remaining > 90, "🟢 Aktif (> 90 hari)")
    result["Status Kontrak"] = status_series
    return result


def add_contract_status(frame: pd.DataFrame, today: date | None = None) -> pd.DataFrame:
    return compute_contract_status(normalize_data(frame), today=today)


def save_data(frame: pd.DataFrame, already_normalized: bool = False) -> pd.DataFrame:
    normalized = frame.drop(columns=["_search_blob", "Status Kontrak", "Sisa Hari"], errors="ignore").copy() if already_normalized else normalize_data(frame)
    errors = validate_data(normalized, already_normalized=True)
    if errors:
        raise ValueError("\n".join(errors))
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    normalized.to_excel(DATA_FILE, sheet_name=SHEET_NAME, index=False, engine="openpyxl")
    return attach_search_blob(normalized)


def budget_group(value) -> str:
    if pd.isna(value):
        return ""
    text = re.sub(r"\s+", " ", str(value).strip().casefold())
    if re.match(r"others?\s+permanen\b", text):
        return "Others Permanen"
    if text == "venue":
        return "Venue"
    return ""


def format_outlet_table(frame: pd.DataFrame) -> pd.DataFrame:
    display_frame = frame.drop(columns=["_search_blob"], errors="ignore").copy()
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
        "Ukuran Sticker / Vinyl Panel Toko", "Brand Panel", "Versi", "Tanggal Penggantian Sticker / Vinyl Panel", "Nama Vendor",
        "Zona", "Sektor", "Pos Budget", "Kategori Outlet", "Kategori Pekerjaan",
    ]
    existing_priority = [c for c in priority_cols if c in display_frame.columns]
    remaining_cols = [c for c in display_frame.columns if c not in existing_priority]
    return display_frame[existing_priority + remaining_cols]
