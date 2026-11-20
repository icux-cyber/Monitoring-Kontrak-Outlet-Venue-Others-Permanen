from __future__ import annotations

import json
import re
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

APP_TITLE = "Monitoring Kontrak Outlet"
DATA_FILE = Path("data") / "DRP & Outlet MCD - DSO Surabaya Barat.xlsx"
SHEET_NAME = "Monitoring Kontrak Outlet"
DRP_SHEET_NAME = "Pemesanan Layar DRP"
DEFAULT_SUPABASE_TABLE = "monitoring_kontrak_outlet"
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

st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")


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
        missing_text = ", ".join(missing)
        raise ValueError(f"Kolom wajib tidak ditemukan: {missing_text}.")
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


def _get_supabase_config() -> tuple[str, str, str, bool]:
    try:
        url = str(st.secrets.get("SUPABASE_URL", "") or "").strip().rstrip("/")
        key = str(st.secrets.get("SUPABASE_KEY", "") or "").strip()
        table = str(st.secrets.get("SUPABASE_TABLE", DEFAULT_SUPABASE_TABLE) or DEFAULT_SUPABASE_TABLE).strip()
        raw_auto = st.secrets.get("SUPABASE_AUTO_RESTORE", True)
        auto_restore = raw_auto.strip().lower() in {"1", "true", "yes", "ya"} if isinstance(raw_auto, str) else bool(raw_auto)
    except FileNotFoundError:
        url, key, table, auto_restore = "", "", DEFAULT_SUPABASE_TABLE, True
    return url, key, table, auto_restore


def _supabase_headers(api_key: str) -> dict[str, str]:
    return {
        "apikey": api_key,
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    }


def _serialize_records_for_supabase(frame: pd.DataFrame, columns: list[str]) -> list[dict]:
    records = json.loads(frame[columns].to_json(orient="records", date_format="iso", force_ascii=False))
    cleaned: list[dict] = []
    for row in records:
        item: dict = {}
        for col in columns:
            val = row.get(col)
            if val is None or (isinstance(val, float) and pd.isna(val)):
                item[col] = None
            elif col in MATERIAL_COLUMNS:
                item[col] = bool(val)
            else:
                item[col] = val
        cleaned.append(item)
    return cleaned


def backup_to_supabase(
    frame: pd.DataFrame,
    supabase_url: str,
    supabase_key: str,
    table_name: str = DEFAULT_SUPABASE_TABLE,
    batch_size: int = 200,
) -> int:
    url = str(supabase_url).strip().rstrip("/")
    key = str(supabase_key).strip()
    table = str(table_name).strip()
    if not url.startswith("https://") or not key or not table:
        raise ValueError("Konfigurasi SUPABASE_URL, SUPABASE_KEY, dan nama tabel belum lengkap.")

    normalized = normalize_data(frame)
    base_records = _serialize_records_for_supabase(normalized, COLUMNS)
    now_iso = pd.Timestamp.now(tz="UTC").isoformat()
    payload_rows = [
        {"row_no": idx + 1, **rec, "updated_at": now_iso}
        for idx, rec in enumerate(base_records)
    ]

    endpoint = f"{url}/rest/v1/{table}"
    headers = _supabase_headers(key)

    delete_resp = requests.delete(
        f"{endpoint}?row_no=gte.0",
        headers=headers,
        timeout=60,
    )
    delete_resp.raise_for_status()

    if not payload_rows:
        return 0

    inserted_total = 0
    for start in range(0, len(payload_rows), batch_size):
        chunk = payload_rows[start : start + batch_size]
        resp = requests.post(
            f"{endpoint}?on_conflict=row_no",
            headers={**headers, "Prefer": "resolution=merge-duplicates,return=minimal"},
            json=chunk,
            timeout=60,
        )
        resp.raise_for_status()
        inserted_total += len(chunk)

    return inserted_total


def restore_from_supabase(
    supabase_url: str,
    supabase_key: str,
    table_name: str = DEFAULT_SUPABASE_TABLE,
    page_size: int = 1000,
) -> pd.DataFrame:
    url = str(supabase_url).strip().rstrip("/")
    key = str(supabase_key).strip()
    table = str(table_name).strip()
    if not url.startswith("https://") or not key or not table:
        raise ValueError("Konfigurasi SUPABASE_URL, SUPABASE_KEY, dan nama tabel belum lengkap.")

    endpoint = f"{url}/rest/v1/{table}"
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Accept": "application/json",
    }

    all_rows: list[dict] = []
    offset = 0
    while True:
        resp = requests.get(
            endpoint,
            headers={**headers, "Range": f"{offset}-{offset + page_size - 1}"},
            params={"select": "*", "order": "row_no.asc.nullslast"},
            timeout=60,
        )
        resp.raise_for_status()
        batch = resp.json()
        if not isinstance(batch, list):
            raise ValueError("Respons dari Supabase bukan list data yang valid.")
        all_rows.extend(batch)
        if len(batch) < page_size:
            break
        offset += page_size

    if not all_rows:
        return pd.DataFrame(columns=COLUMNS)

    raw_df = pd.DataFrame(all_rows)
    return normalize_data(raw_df)


def load_local_data() -> tuple[pd.DataFrame, str | None]:
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


def save_data(frame: pd.DataFrame) -> None:
    errors = validate_data(frame)
    if errors:
        raise ValueError("\n".join(errors))
    normalized = normalize_data(frame)
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    normalized.to_excel(DATA_FILE, sheet_name=SHEET_NAME, index=False, engine="openpyxl")


def init_session_data() -> None:
    if "auto_restore_enabled" not in st.session_state:
        _, _, _, default_auto = _get_supabase_config()
        st.session_state.auto_restore_enabled = default_auto

    if "contracts" in st.session_state:
        return

    sb_url, sb_key, sb_table, _ = _get_supabase_config()
    if st.session_state.auto_restore_enabled and sb_url and sb_key:
        try:
            restored = restore_from_supabase(sb_url, sb_key, sb_table)
            if not restored.empty and not validate_data(restored):
                try:
                    save_data(restored)
                except Exception:
                    pass
                st.session_state.contracts = restored
                st.session_state.load_warning = None
                st.session_state.last_restore_info = f"Auto-restore dari Supabase ({len(restored)} baris)."
                return
        except Exception as error:
            local_df, local_warn = load_local_data()
            st.session_state.contracts = local_df
            fallback_msg = f"Auto-restore Supabase gagal ({error}). Memuat data lokal."
            st.session_state.load_warning = f"{fallback_msg} {local_warn}" if local_warn else fallback_msg
            return

    local_df, local_warn = load_local_data()
    st.session_state.contracts = local_df
    st.session_state.load_warning = local_warn


def add_contract_status(frame: pd.DataFrame, today: date | None = None) -> pd.DataFrame:
    result = normalize_data(frame)
    remaining = (result["Periode Kontrak Akhir"] - pd.Timestamp(today or date.today()).normalize()).dt.days
    result["Sisa Hari"] = remaining

    def status(days):
        if pd.isna(days):
            return "Data belum lengkap"
        if days < 0:
            return "Expired"
        if days <= 30:
            return "Kritis (≤ 30 hari)"
        if days <= 90:
            return "Warning (31–90 hari)"
        return "Aktif (> 90 hari)"

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


def format_outlet_table(frame: pd.DataFrame) -> pd.DataFrame:
    display_frame = frame.copy()
    if "Nominal Kontrak" in display_frame:
        display_frame["Nominal Kontrak"] = display_frame["Nominal Kontrak"].map(
            lambda value: "" if pd.isna(value) else f"{float(value):,.0f}".replace(",", ".")
        )
    for column in ("Periode Kontrak Awal", "Periode Kontrak Akhir"):
        if column in display_frame:
            display_frame[column] = pd.to_datetime(display_frame[column], errors="coerce").map(
                lambda value: "" if pd.isna(value) else value.strftime("%d/%m/%Y")
            )
    for column in MATERIAL_COLUMNS:
        if column in display_frame:
            display_frame[column] = display_frame[column].map(lambda checked: "✓" if parse_tickmark(checked) else "")
    return display_frame


def budget_group(value) -> str:
    if pd.isna(value):
        return ""
    text = re.sub(r"\s+", " ", str(value).strip().casefold())
    if re.match(r"others?\s+permanen\b", text):
        return "Others Permanen"
    if text == "venue":
        return "Venue"
    return ""


def render_dashboard_table(title: str, frame: pd.DataFrame) -> None:
    st.subheader(title)
    if frame.empty:
        st.info(f"Tidak ada customer di tabel {title.lower()} yang sesuai dengan filter.")
        return
    st.metric("Jumlah customer pada tabel", f"{len(frame):,}".replace(",", "."))
    sorted_frame = frame.sort_values("Sisa Hari", na_position="last")
    st.dataframe(format_outlet_table(sorted_frame), use_container_width=True, hide_index=True)


def render_excel_io_section(contracts_df: pd.DataFrame) -> None:
    st.subheader("Upload / Download Excel")
    upload = st.file_uploader("Upload workbook Excel (.xlsx)", type=["xlsx"], key="workbook_upload")
    if st.button("Import dan ganti data", type="primary", key="import_contract_workbook"):
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
                    st.success(f"Berhasil mengimpor {len(imported)} baris.")
                    st.rerun()
            except Exception as error:
                st.error(f"Gagal mengimpor Excel: {error}")

    with st.expander("Sinkronkan Ukuran & Brand Panel / Layar berdasarkan ID Customer"):
        sync_upload = st.file_uploader(
            "Upload file/sheet sumber Panel & Layar (.xlsx)",
            type=["xlsx"],
            key="sync_panel_screen_upload",
        )
        overwrite_sync = st.checkbox(
            "Timpa nilai yang sudah terisi jika ID Customer cocok",
            value=False,
            key="overwrite_panel_screen_sync",
        )
        if st.button("Hubungkan ke Daftar customer", key="btn_sync_panel_screen"):
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
                        contracts_df,
                        source_frame,
                        overwrite_existing=overwrite_sync,
                    )
                    save_data(merged_df)
                    st.session_state.contracts = merged_df
                    st.success(f"Sinkronisasi selesai: {updated_cnt} baris customer diperbarui.")
                    if missing_list:
                        missing_preview = ", ".join(missing_list[:15])
                        st.info(f"{len(missing_list)} ID Customer tidak ditemukan di Daftar customer: {missing_preview}")
                    st.rerun()
                except Exception as error:
                    st.error(f"Gagal sinkronisasi data panel/layar: {error}")

    excel_col1, excel_col2 = st.columns(2)
    excel_col1.download_button(
        "Download data Excel",
        data=excel_bytes(contracts_df),
        file_name="Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="btn_download_data_excel",
    )
    excel_col2.download_button(
        "Download template Excel",
        data=excel_bytes(pd.DataFrame(columns=COLUMNS)),
        file_name="Template_Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="btn_download_template_excel",
    )
    st.caption("Import akan mengganti seluruh data customer. Penyimpanan lokal Streamlit Cloud dapat bersifat sementara.")


def safe_text(value) -> str:
    return "" if pd.isna(value) else str(value)


def safe_date(value):
    return None if pd.isna(value) else pd.Timestamp(value).date()


def render_outlet_form(current: pd.Series, row_index: int | None) -> tuple[bool, dict]:
    suffix = "new" if row_index is None else f"row_{row_index}"
    with st.form(f"outlet_form_{suffix}"):
        st.subheader("Data customer dan kontrak")
        left, right = st.columns(2)
        with left:
            outlet_id = st.text_input("ID Outlet", value=safe_text(current["ID Outlet"]), key=f"id_{suffix}")
            customer_name = st.text_input("Nama Customer", value=safe_text(current["Nama Customer"]), key=f"name_{suffix}")
            address = st.text_area("Alamat", value=safe_text(current["Alamat"]), key=f"address_{suffix}")
            zone = st.text_input("Zona", value=safe_text(current["Zona"]), key=f"zone_{suffix}")
            sector = st.text_input("Sektor", value=safe_text(current["Sektor"]), key=f"sector_{suffix}")
            budget_pos = st.text_input("Pos Budget", value=safe_text(current["Pos Budget"]), key=f"budget_{suffix}")
            outlet_category = st.text_input("Kategori Outlet", value=safe_text(current["Kategori Outlet"]), key=f"outlet_category_{suffix}")
            job_category = st.text_input("Kategori Pekerjaan", value=safe_text(current["Kategori Pekerjaan"]), key=f"job_category_{suffix}")
        with right:
            brand = st.text_input("Brand", value=safe_text(current["Brand"]), key=f"brand_{suffix}")
            parsed_amount = parse_amount(current["Nominal Kontrak"])
            amount = st.number_input(
                "Nominal Kontrak (Rp)",
                min_value=0.0,
                step=100000.0,
                value=0.0 if pd.isna(parsed_amount) else float(parsed_amount),
                key=f"amount_{suffix}",
            )
            contract_start = st.date_input("Periode Kontrak Awal", value=safe_date(current["Periode Kontrak Awal"]), key=f"start_{suffix}")
            contract_end = st.date_input("Periode Kontrak Akhir", value=safe_date(current["Periode Kontrak Akhir"]), key=f"end_{suffix}")
        st.markdown("**Materi kontrak**")
        material_cols = st.columns(4)
        material_values = {
            column: material_cols[index].checkbox(column, value=parse_tickmark(current[column]), key=f"material_{index}_{suffix}")
            for index, column in enumerate(MATERIAL_COLUMNS)
        }
        st.markdown("**Detail panel dan layar**")
        panel_col1, panel_col2, panel_col3 = st.columns(3)
        panel_size = panel_col1.text_input("Ukuran Panel Toko", value=safe_text(current["Ukuran Panel Toko"]), key=f"panel_size_{suffix}")
        sticker_size = panel_col2.text_input("Ukuran Sticker Panel Toko", value=safe_text(current["Ukuran Sticker Panel Toko"]), key=f"sticker_size_{suffix}")
        panel_brand = panel_col3.text_input("Brand Panel", value=safe_text(current["Brand Panel"]), key=f"panel_brand_{suffix}")
        screen_col1, screen_col2 = st.columns(2)
        screen_size = screen_col1.text_input("Ukuran Layar", value=safe_text(current["Ukuran Layar"]), key=f"screen_size_{suffix}")
        screen_brand = screen_col2.text_input("Brand Layar", value=safe_text(current["Brand Layar"]), key=f"screen_brand_{suffix}")
        submitted = st.form_submit_button("Simpan data customer", type="primary")
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


st.title("📊 Monitoring Kontrak Outlet")
st.caption("Pantau data customer dan kontrak, lihat peringatan masa berlaku, serta kelola data melalui form, Excel, dan Supabase.")

init_session_data()

if st.session_state.get("load_warning"):
    st.warning(st.session_state.load_warning)
elif st.session_state.get("last_restore_info"):
    st.caption(st.session_state.last_restore_info)

with st.sidebar:
    st.header("Menu")
    active_menu = st.radio(
        "Pilih halaman",
        [
            "Dashboard",
            "Edit data customer",
            "Tambah customer baru",
            "Eksport Data Branding (Excel) - Data SPKS",
            "Backup & Restore Supabase",
            "Backup ke Google Sheet",
        ],
        key="main_menu",
    )

contracts = add_contract_status(st.session_state.contracts)
if active_menu == "Dashboard":
    expired = int(contracts["Status Kontrak"].eq("Expired").sum())
    soon = int(contracts["Status Kontrak"].isin(["Kritis (≤ 30 hari)", "Warning (31–90 hari)"]).sum())
    total_value = pd.to_numeric(contracts["Nominal Kontrak"], errors="coerce").sum()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Jumlah customer", f"{len(contracts):,}".replace(",", "."))
    m2.metric("Total nominal kontrak", f"Rp {total_value:,.0f}".replace(",", "."))
    m3.metric("Akan berakhir ≤ 90 hari", soon)
    m4.metric("Sudah berakhir", expired)
    if expired:
        st.error(f"Perhatian: {expired} kontrak sudah berakhir.")
    if soon:
        st.warning(f"Pengingat: {soon} kontrak akan berakhir dalam 90 hari.")

    st.subheader("Filter data kontrak")
    f1, f2, f3, f4, f5 = st.columns(5)
    zones = ["Semua"] + sorted(contracts["Zona"].dropna().astype(str).unique().tolist())
    sectors = ["Semua"] + sorted(contracts["Sektor"].dropna().astype(str).unique().tolist())
    budget_positions = ["Semua"] + sorted(contracts["Pos Budget"].dropna().astype(str).unique().tolist())
    job_categories = ["Semua"] + sorted(contracts["Kategori Pekerjaan"].dropna().astype(str).unique().tolist())
    statuses = ["Semua", "Expired", "Kritis (≤ 30 hari)", "Warning (31–90 hari)", "Aktif (> 90 hari)", "Data belum lengkap"]
    selected_zone = f1.selectbox("Zona", zones)
    selected_sector = f2.selectbox("Sektor", sectors)
    selected_budget = f3.selectbox("Pos Budget", budget_positions)
    selected_job_category = f4.selectbox("Kategori Pekerjaan", job_categories)
    selected_status = f5.selectbox("Status kontrak", statuses)

    f6, f7 = st.columns(2)
    sticker_sizes = ["Semua"] + sorted(contracts["Ukuran Sticker Panel Toko"].dropna().astype(str).unique().tolist())
    screen_sizes = ["Semua"] + sorted(contracts["Ukuran Layar"].dropna().astype(str).unique().tolist())
    selected_sticker_size = f6.selectbox("Ukuran Sticker Panel", sticker_sizes)
    selected_screen_size = f7.selectbox("Ukuran Layar", screen_sizes)

    view = contracts.copy()
    if selected_zone != "Semua":
        view = view[view["Zona"] == selected_zone]
    if selected_sector != "Semua":
        view = view[view["Sektor"] == selected_sector]
    if selected_budget != "Semua":
        view = view[view["Pos Budget"] == selected_budget]
    if selected_job_category != "Semua":
        view = view[view["Kategori Pekerjaan"] == selected_job_category]
    if selected_status != "Semua":
        view = view[view["Status Kontrak"] == selected_status]
    if selected_sticker_size != "Semua":
        view = view[view["Ukuran Sticker Panel Toko"] == selected_sticker_size]
    if selected_screen_size != "Semua":
        view = view[view["Ukuran Layar"] == selected_screen_size]
    st.subheader("Hasil filter data kontrak")
    st.metric("Jumlah customer sesuai filter", f"{len(view):,}".replace(",", "."))
    groups = view["Pos Budget"].map(budget_group)
    regular_view = view.loc[groups.eq("")]
    venue_view = view.loc[groups.eq("Venue")]
    other_permanent_view = view.loc[groups.eq("Others Permanen")]
    download_columns = st.columns(4)
    download_specs = [
        (
            "Download Daftar customer - All Kategori Excel",
            view,
            "Daftar_Customer_All_Kategori.xlsx",
            "download_all_category_contracts",
        ),
        (
            "Download Daftar customer - Retail Suplay kontrak Excel",
            regular_view,
            "Daftar_Customer_dan_Status_Kontrak.xlsx",
            "download_regular_contracts",
        ),
        (
            "Download Daftar customer - Venue Excel",
            venue_view,
            "Daftar_Customer_Venue.xlsx",
            "download_venue_contracts",
        ),
        (
            "Download Daftar customer - Others Permanen Excel",
            other_permanent_view,
            "Daftar_Customer_Others_Permanen.xlsx",
            "download_other_permanent_contracts",
        ),
    ]
    for column, (label, data, filename, key) in zip(download_columns, download_specs):
        column.download_button(
            label,
            data=excel_bytes(data.sort_values("Sisa Hari", na_position="last")),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=key,
            use_container_width=True,
        )
    if view.empty:
        st.info("Belum ada data yang cocok. Buka menu **Eksport Data Branding (Excel) - Data SPKS** untuk mengimpor data Excel.")
    else:
        selected_group = budget_group(selected_budget)
        if selected_budget == "Semua":
            render_dashboard_table("Daftar customer - Retail Supply", regular_view)
            render_dashboard_table("Daftar customer - Venue", venue_view)
            render_dashboard_table("Daftar customer - Others Permanen", other_permanent_view)
        elif selected_group == "Venue":
            render_dashboard_table("Daftar customer - Venue", venue_view)
        elif selected_group == "Others Permanen":
            render_dashboard_table("Daftar customer - Others Permanen", other_permanent_view)
        else:
            render_dashboard_table("Daftar customer dan status kontrak", regular_view)

elif active_menu == "Edit data customer":
    st.header("Edit data customer")
    if contracts.empty:
        st.info("Data customer belum ada. Gunakan menu **Tambah customer baru** atau **Eksport Data Branding (Excel) - Data SPKS**.")
    else:
        record_options = [
            f"{i} | {safe_text(row['ID Outlet']) or 'Tanpa ID'} | {safe_text(row['Nama Customer']) or 'Tanpa nama'}"
            for i, row in contracts.iterrows()
        ]
        selected_record = st.selectbox("Pilih customer yang akan diedit", record_options, key="selected_customer_record")
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
                    st.success("Data customer berhasil diperbarui.")
                    st.rerun()
                except Exception as error:
                    st.error(f"Gagal menyimpan data: {error}")
        st.divider()
        st.subheader("Hapus customer")
        if st.button("Hapus customer yang dipilih", type="secondary", key="delete_customer"):
            try:
                updated_data = st.session_state.contracts.drop(index=selected_index).reset_index(drop=True)
                save_data(updated_data)
                st.session_state.contracts = normalize_data(updated_data)
                st.success("Customer berhasil dihapus.")
                st.rerun()
            except Exception as error:
                st.error(f"Gagal menghapus customer: {error}")

elif active_menu == "Tambah customer baru":
    st.header("Tambah customer baru")
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
                st.success("Customer baru berhasil ditambahkan.")
                st.rerun()
            except Exception as error:
                st.error(f"Gagal menyimpan customer baru: {error}")

elif active_menu == "Eksport Data Branding (Excel) - Data SPKS":
    st.header("Eksport Data Branding (Excel) - Data SPKS")
    render_excel_io_section(st.session_state.contracts)
    st.divider()
    st.subheader("Import data outlet branding (SPKS)")
    st.caption(
        "Data dengan ID Customer ganda dipilih berdasarkan Tanggal Pekerjaan Sampai terbaru. "
        "ID yang ditemukan memperbarui data customer; ID yang belum ada tidak ditambahkan ke tabel."
    )
    branding_upload = st.file_uploader(
        "Upload laporan SPKS Branding (.xlsx)",
        type=["xlsx"],
        key="branding_workbook_upload",
    )
    if st.button("Proses dan update data customer", type="primary"):
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
                    st.success(
                        f"Proses selesai: {updated_count} ID diperbarui, "
                        f"{int(unmatched_has_id.sum())} ID tidak ditemukan, "
                        f"dan {int((~unmatched_has_id).sum())} baris tidak memiliki ID."
                    )
            except Exception as error:
                st.error(f"Gagal memproses laporan branding: {error}")

    import_result = st.session_state.get("branding_import_result")
    if import_result:
        summary_cols = st.columns(4)
        summary_cols[0].metric(
            "ID duplikat (baris)",
            f"{import_result['duplicate_ids']} ({import_result['duplicate_rows']})",
        )
        summary_cols[1].metric("ID diperbarui", import_result["updated_count"])
        summary_cols[2].metric("ID tidak ditemukan", import_result["unmatched_id_count"])
        summary_cols[3].metric("Baris tanpa ID", import_result["missing_id_count"])
        if import_result["duplicate_rows"]:
            st.subheader("Data ID duplikat")
            st.caption("Baris dengan tanda Dipakai adalah data yang dipilih karena tanggal pekerjaan sampai paling baru.")
            duplicate_columns = [
                "ID Customer", "Customer", "Tanggal Pekerjaan Sampai",
                "No. Dokumen", "Dipakai",
            ]
            st.dataframe(
                import_result["duplicates"][duplicate_columns],
                use_container_width=True,
                hide_index=True,
            )
        missing_ids = import_result["unmatched"][
            import_result["unmatched"]["ID Customer"].map(customer_id_key).ne("")
        ]
        if missing_ids.empty:
            st.info("Semua ID Customer pada laporan ditemukan di Daftar customer.")
        else:
            st.subheader("ID branding yang tidak ditemukan di Daftar customer")
            missing_columns = [
                "ID Customer", "Customer", "Customer Group Category", "Alamat",
                "Zona", "Finance Brand", "Nominal", "Tanggal Pekerjaan Dari",
                "Tanggal Pekerjaan Sampai", "No. Dokumen",
            ]
            st.dataframe(
                missing_ids[missing_columns],
                use_container_width=True,
                hide_index=True,
            )
        without_ids = import_result["unmatched"][
            import_result["unmatched"]["ID Customer"].map(customer_id_key).eq("")
        ]
        if not without_ids.empty:
            st.subheader("Baris laporan tanpa ID Customer")
            st.caption("Baris ini tidak dapat dibandingkan atau diperbarui karena ID Customer kosong.")
            missing_columns = [
                "ID Customer", "Customer", "Customer Group Category", "Alamat",
                "Zona", "Finance Brand", "Nominal", "Tanggal Pekerjaan Dari",
                "Tanggal Pekerjaan Sampai", "No. Dokumen",
            ]
            st.dataframe(without_ids[missing_columns], use_container_width=True, hide_index=True)

elif active_menu == "Backup & Restore Supabase":
    st.header("Backup & Restore Supabase")
    st.caption(
        "Simpan snapshot data Daftar customer ke tabel Supabase dan pulihkan otomatis saat membuka atau me-reload halaman web."
    )
    cfg_url, cfg_key, cfg_table, cfg_auto = _get_supabase_config()

    c1, c2 = st.columns(2)
    sb_url = c1.text_input("Supabase Project URL", value=cfg_url, placeholder="https://xxxx.supabase.co", key="sb_url_input")
    sb_table = c2.text_input("Nama Tabel Supabase", value=cfg_table, key="sb_table_input")
    sb_key = st.text_input(
        "Supabase API Key (anon / service_role)",
        value=cfg_key,
        type="password",
        help="Disarankan menyimpan SUPABASE_URL, SUPABASE_KEY, dan SUPABASE_TABLE di Streamlit Secrets.",
        key="sb_key_input",
    )
    auto_restore_toggle = st.toggle(
        "Auto-restore dari Supabase saat buka / reload halaman",
        value=bool(st.session_state.get("auto_restore_enabled", cfg_auto)),
        key="toggle_auto_restore_supabase",
    )
    st.session_state.auto_restore_enabled = auto_restore_toggle

    st.divider()
    act_left, act_right = st.columns(2)

    with act_left:
        st.subheader("Backup ke Supabase")
        st.info(f"Data aktif siap dibackup: {len(st.session_state.contracts):,} outlet.".replace(",", "."))
        st.warning("Backup akan mengganti seluruh isi tabel Supabase dengan data Daftar customer saat ini.")
        if st.button("Backup data ke Supabase", type="primary", key="btn_backup_supabase"):
            try:
                saved_count = backup_to_supabase(
                    st.session_state.contracts,
                    supabase_url=sb_url,
                    supabase_key=sb_key,
                    table_name=sb_table,
                )
                st.success(f"Backup ke Supabase berhasil: {saved_count} baris tersimpan di tabel '{sb_table}'.")
            except Exception as error:
                st.error(f"Backup ke Supabase gagal: {error}")

    with act_right:
        st.subheader("Restore dari Supabase")
        st.info("Mengambil seluruh data dari tabel Supabase dan menjadikannya data aktif di aplikasi.")
        confirm_restore = st.checkbox(
            "Saya mengerti restore akan menimpa data customer yang sedang tampil di aplikasi",
            key="confirm_restore_supabase",
        )
        if st.button("Restore data dari Supabase", type="secondary", key="btn_restore_supabase"):
            if not confirm_restore:
                st.warning("Centang konfirmasi terlebih dahulu sebelum melakukan restore.")
            else:
                try:
                    restored_df = restore_from_supabase(
                        supabase_url=sb_url,
                        supabase_key=sb_key,
                        table_name=sb_table,
                    )
                    problems = validate_data(restored_df)
                    if problems:
                        st.error("\n".join(problems))
                    else:
                        save_data(restored_df)
                        st.session_state.contracts = restored_df
                        st.session_state.load_warning = None
                        st.session_state.last_restore_info = f"Restore manual dari Supabase ({len(restored_df)} baris)."
                        st.success(f"Restore dari Supabase berhasil: {len(restored_df)} baris dimuat.")
                        st.rerun()
                except Exception as error:
                    st.error(f"Restore dari Supabase gagal: {error}")

else:
    st.header("Backup ke Google Sheet")
    st.caption(
        "Backup mengirim seluruh data Daftar customer dan mengganti isi tab "
        f"'{SHEET_NAME}' pada spreadsheet yang terhubung ke Apps Script."
    )
    try:
        backup_url = str(st.secrets.get("GOOGLE_APPS_SCRIPT_URL", "") or "").strip()
    except FileNotFoundError:
        backup_url = ""
    if not backup_url:
        st.warning("Atur GOOGLE_APPS_SCRIPT_URL di Streamlit secrets untuk mengaktifkan backup.")
    else:
        st.info(f"Data yang akan dibackup: {len(st.session_state.contracts)} outlet.")
        st.warning("Backup akan mengganti isi tab tujuan dengan data outlet saat ini.")
        st.warning("Siapa pun yang mengetahui URL Web App dapat menjalankan backup dan menimpa isi tab tujuan. Jaga URL tersebut.")
        if st.button("Backup data sekarang", type="primary"):
            try:
                row_count = backup_to_google_sheet(
                    st.session_state.contracts,
                    backup_url,
                )
                st.success(f"Backup berhasil: {row_count} baris tersimpan di Google Sheet.")
            except Exception as error:
                st.error(f"Backup ke Google Sheet gagal: {error}")
