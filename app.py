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
MATERIAL_COLUMNS = ["Panel Toko", "Showcase", "Layar Toko", "Shop Sign / Neon Box"]
COLUMNS = [
    "ID Outlet", "Nama Customer", "Alamat", "Zona", "Sektor", "Pos Budget",
    "Kategori Outlet", "Kategori Pekerjaan", "Brand", "Nominal Kontrak",
    "Periode Kontrak Awal", "Periode Kontrak Akhir", *MATERIAL_COLUMNS,
    "Ukuran Panel Toko", "Ukuran Sticker Panel Toko",
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
    text_columns = [column for column in COLUMNS if column not in MATERIAL_COLUMNS and column not in ("Nominal Kontrak", "Periode Kontrak Awal", "Periode Kontrak Akhir")]
    for column in text_columns:
        normalized[column] = normalized[column].astype("string").str.strip()
        normalized[column] = normalized[column].replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA})
    for column in ("Ukuran Panel Toko", "Ukuran Sticker Panel Toko"):
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
    has_data = normalized.notna().any(axis=1)
    return normalized.loc[has_data.to_numpy(dtype=bool)].reset_index(drop=True)


def read_workbook(source) -> pd.DataFrame:
    for sheet, header in (("COSPKR06", 3), (0, 0), (SHEET_NAME, 0)):
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            result = normalize_data(pd.read_excel(source, sheet_name=sheet, header=header))
            if not result.empty:
                return result
        except (ValueError, KeyError, OSError, ImportError):
            continue
    raise ValueError("Tidak dapat membaca sheet/kolom data kontrak. Periksa file Excel dan nama sheet-nya.")


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


def customer_id_key(value) -> str:
    if pd.isna(value):
        return ""
    key = str(value).strip()
    if re.fullmatch(r"\d+\.0", key):
        key = key[:-2]
    return key.casefold()


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
    invalid_period = (normalized["Periode Kontrak Awal"].notna() & normalized["Periode Kontrak Akhir"].notna() & (normalized["Periode Kontrak Akhir"] < normalized["Periode Kontrak Awal"]))
    if invalid_period.any():
        errors.append("Tanggal akhir kontrak tidak boleh sebelum tanggal awal.")
    return errors


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
        display_frame["Nominal Kontrak"] = display_frame["Nominal Kontrak"].map(lambda value: "" if pd.isna(value) else f"{float(value):,.0f}".replace(",", "."))
    for column in ("Periode Kontrak Awal", "Periode Kontrak Akhir"):
        if column in display_frame:
            display_frame[column] = pd.to_datetime(display_frame[column], errors="coerce").map(lambda value: "" if pd.isna(value) else value.strftime("%d/%m/%Y"))
    for column in MATERIAL_COLUMNS:
        if column in display_frame:
            display_frame[column] = display_frame[column].map(lambda checked: "✓" if parse_tickmark(checked) else "")
    return display_frame


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
            amount = st.number_input("Nominal Kontrak (Rp)", min_value=0.0, step=100000.0, value=0.0 if pd.isna(parsed_amount) else float(parsed_amount), key=f"amount_{suffix}")
            contract_start = st.date_input("Periode Kontrak Awal", value=safe_date(current["Periode Kontrak Awal"]), key=f"start_{suffix}")
            contract_end = st.date_input("Periode Kontrak Akhir", value=safe_date(current["Periode Kontrak Akhir"]), key=f"end_{suffix}")
        st.markdown("**Materi kontrak**")
        material_cols = st.columns(4)
        material_values = {column: material_cols[index].checkbox(column, value=parse_tickmark(current[column]), key=f"material_{index}_{suffix}") for index, column in enumerate(MATERIAL_COLUMNS)}
        size_col1, size_col2 = st.columns(2)
        panel_size = size_col1.text_input("Ukuran Panel Toko", value=safe_text(current["Ukuran Panel Toko"]), key=f"panel_size_{suffix}")
        sticker_size = size_col2.text_input("Ukuran Sticker Panel Toko", value=safe_text(current["Ukuran Sticker Panel Toko"]), key=f"sticker_size_{suffix}")
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
    }
    return submitted, values


st.title("📊 Monitoring Kontrak Outlet")
st.caption("Pantau data customer dan kontrak, lihat peringatan masa berlaku, serta kelola data melalui form dan Excel.")
if "contracts" not in st.session_state:
    initial, warning = load_initial_data()
    st.session_state.contracts = initial
    st.session_state.load_warning = warning
if st.session_state.load_warning:
    st.warning(st.session_state.load_warning)

with st.sidebar:
    st.header("Menu")
    active_menu = st.radio(
        "Pilih halaman",
        ["Dashboard", "Edit data customer", "Tambah customer baru", "Import data branding", "Backup ke Google Sheet"],
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
    excel_col1, excel_col2 = st.columns(2)
    excel_col1.download_button(
        "Download data Excel",
        data=excel_bytes(st.session_state.contracts),
        file_name="Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    excel_col2.download_button(
        "Download template Excel",
        data=excel_bytes(pd.DataFrame(columns=COLUMNS)),
        file_name="Template_Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Import akan mengganti seluruh data customer. Penyimpanan lokal Streamlit Cloud dapat bersifat sementara.")
    st.subheader("Filter data kontrak")
    f1, f2, f3, f4 = st.columns(4)
    zones = ["Semua"] + sorted(contracts["Zona"].dropna().astype(str).unique().tolist())
    sectors = ["Semua"] + sorted(contracts["Sektor"].dropna().astype(str).unique().tolist())
    job_categories = ["Semua"] + sorted(contracts["Kategori Pekerjaan"].dropna().astype(str).unique().tolist())
    statuses = ["Semua", "Expired", "Kritis (≤ 30 hari)", "Warning (31–90 hari)", "Aktif (> 90 hari)", "Data belum lengkap"]
    selected_zone = f1.selectbox("Zona", zones)
    selected_sector = f2.selectbox("Sektor", sectors)
    selected_job_category = f3.selectbox("Kategori Pekerjaan", job_categories)
    selected_status = f4.selectbox("Status kontrak", statuses)
    view = contracts.copy()
    if selected_zone != "Semua":
        view = view[view["Zona"] == selected_zone]
    if selected_sector != "Semua":
        view = view[view["Sektor"] == selected_sector]
    if selected_job_category != "Semua":
        view = view[view["Kategori Pekerjaan"] == selected_job_category]
    if selected_status != "Semua":
        view = view[view["Status Kontrak"] == selected_status]
    st.subheader("Daftar customer dan status kontrak")
    st.metric("Jumlah customer sesuai filter", f"{len(view):,}".replace(",", "."))
    sorted_view = view.sort_values("Sisa Hari", na_position="last")
    st.download_button("Download daftar customer ke Excel", data=excel_bytes(format_outlet_table(sorted_view)), file_name="Daftar_Customer_dan_Status_Kontrak.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if view.empty:
        st.info("Belum ada data yang cocok. Gunakan bagian Upload / Download Excel pada dashboard untuk memasukkan data.")
    else:
        st.dataframe(format_outlet_table(sorted_view), use_container_width=True, hide_index=True)

elif active_menu == "Edit data customer":
    st.header("Edit data customer")
    if contracts.empty:
        st.info("Data customer belum ada. Gunakan menu **Tambah customer baru** atau bagian Upload / Download Excel pada dashboard.")
    else:
        record_options = [f"{i} | {safe_text(row['ID Outlet']) or 'Tanpa ID'} | {safe_text(row['Nama Customer']) or 'Tanpa nama'}" for i, row in contracts.iterrows()]
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
elif active_menu == "Import data branding":
    st.header("Import data outlet branding")
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
                st.error(f"Backup ke Google Sheet gagal: {error}")    "Showcase": ("Showcase", "Show Case", "Materi Showcase"),
    "Layar Toko": ("Layar Toko", "Layar", "Materi Layar Toko"),
    "Shop Sign / Neon Box": ("Shop Sign / Neon Box", "Shop Sign", "Neon Box", "Materi Shop Sign / Neon Box"),
    "Ukuran Panel Toko": ("Ukuran Panel Toko", "Size Panel Toko", "Panel"),
    "Ukuran Sticker Panel Toko": (
        "Ukuran Sticker Panel Toko",
        "Ukuran Sticker Panel (diluar Nama Toko)",
        "Ukuran Stiker Panel Toko",
        "Size Sticker Panel Toko",
    ),
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
    text_columns = [column for column in COLUMNS if column not in MATERIAL_COLUMNS and column not in ("Nominal Kontrak", "Periode Kontrak Awal", "Periode Kontrak Akhir")]
    for column in text_columns:
        normalized[column] = normalized[column].astype("string").str.strip()
        normalized[column] = normalized[column].replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA})
    for column in ("Ukuran Panel Toko", "Ukuran Sticker Panel Toko"):
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
    has_data = normalized.notna().any(axis=1)
    return normalized.loc[has_data.to_numpy(dtype=bool)].reset_index(drop=True)


def read_workbook(source) -> pd.DataFrame:
    for sheet, header in (("COSPKR06", 3), (0, 0), (SHEET_NAME, 0)):
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            result = normalize_data(pd.read_excel(source, sheet_name=sheet, header=header))
            if not result.empty:
                return result
        except (ValueError, KeyError, OSError, ImportError):
            continue
    raise ValueError("Tidak dapat membaca sheet/kolom data kontrak. Periksa file Excel dan nama sheet-nya.")


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


def customer_id_key(value) -> str:
    if pd.isna(value):
        return ""
    key = str(value).strip()
    if re.fullmatch(r"\d+\.0", key):
        key = key[:-2]
    return key.casefold()


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
    invalid_period = (normalized["Periode Kontrak Awal"].notna() & normalized["Periode Kontrak Akhir"].notna() & (normalized["Periode Kontrak Akhir"] < normalized["Periode Kontrak Awal"]))
    if invalid_period.any():
        errors.append("Tanggal akhir kontrak tidak boleh sebelum tanggal awal.")
    return errors


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
        display_frame["Nominal Kontrak"] = display_frame["Nominal Kontrak"].map(lambda value: "" if pd.isna(value) else f"{float(value):,.0f}".replace(",", "."))
    for column in ("Periode Kontrak Awal", "Periode Kontrak Akhir"):
        if column in display_frame:
            display_frame[column] = pd.to_datetime(display_frame[column], errors="coerce").map(lambda value: "" if pd.isna(value) else value.strftime("%d/%m/%Y"))
    for column in MATERIAL_COLUMNS:
        if column in display_frame:
            display_frame[column] = display_frame[column].map(lambda checked: "✓" if parse_tickmark(checked) else "")
    return display_frame


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
            amount = st.number_input("Nominal Kontrak (Rp)", min_value=0.0, step=100000.0, value=0.0 if pd.isna(parsed_amount) else float(parsed_amount), key=f"amount_{suffix}")
            contract_start = st.date_input("Periode Kontrak Awal", value=safe_date(current["Periode Kontrak Awal"]), key=f"start_{suffix}")
            contract_end = st.date_input("Periode Kontrak Akhir", value=safe_date(current["Periode Kontrak Akhir"]), key=f"end_{suffix}")
        st.markdown("**Materi kontrak**")
        material_cols = st.columns(4)
        material_values = {column: material_cols[index].checkbox(column, value=parse_tickmark(current[column]), key=f"material_{index}_{suffix}") for index, column in enumerate(MATERIAL_COLUMNS)}
        size_col1, size_col2 = st.columns(2)
        panel_size = size_col1.text_input("Ukuran Panel Toko", value=safe_text(current["Ukuran Panel Toko"]), key=f"panel_size_{suffix}")
        sticker_size = size_col2.text_input("Ukuran Sticker Panel Toko", value=safe_text(current["Ukuran Sticker Panel Toko"]), key=f"sticker_size_{suffix}")
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
    }
    return submitted, values


st.title("📊 Monitoring Kontrak Outlet")
st.caption("Pantau data customer dan kontrak, lihat peringatan masa berlaku, serta kelola data melalui form dan Excel.")
if "contracts" not in st.session_state:
    initial, warning = load_initial_data()
    st.session_state.contracts = initial
    st.session_state.load_warning = warning
if st.session_state.load_warning:
    st.warning(st.session_state.load_warning)

with st.sidebar:
    st.header("Menu")
    active_menu = st.radio(
        "Pilih halaman",
        ["Dashboard", "Edit data customer", "Tambah customer baru", "Import data branding", "Backup ke Google Sheet"],
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
    excel_col1, excel_col2 = st.columns(2)
    excel_col1.download_button(
        "Download data Excel",
        data=excel_bytes(st.session_state.contracts),
        file_name="Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    excel_col2.download_button(
        "Download template Excel",
        data=excel_bytes(pd.DataFrame(columns=COLUMNS)),
        file_name="Template_Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Import akan mengganti seluruh data customer. Penyimpanan lokal Streamlit Cloud dapat bersifat sementara.")
    st.subheader("Filter data kontrak")
    f1, f2, f3, f4 = st.columns(4)
    zones = ["Semua"] + sorted(contracts["Zona"].dropna().astype(str).unique().tolist())
    sectors = ["Semua"] + sorted(contracts["Sektor"].dropna().astype(str).unique().tolist())
    job_categories = ["Semua"] + sorted(contracts["Kategori Pekerjaan"].dropna().astype(str).unique().tolist())
    statuses = ["Semua", "Expired", "Kritis (≤ 30 hari)", "Warning (31–90 hari)", "Aktif (> 90 hari)", "Data belum lengkap"]
    selected_zone = f1.selectbox("Zona", zones)
    selected_sector = f2.selectbox("Sektor", sectors)
    selected_job_category = f3.selectbox("Kategori Pekerjaan", job_categories)
    selected_status = f4.selectbox("Status kontrak", statuses)
    view = contracts.copy()
    if selected_zone != "Semua":
        view = view[view["Zona"] == selected_zone]
    if selected_sector != "Semua":
        view = view[view["Sektor"] == selected_sector]
    if selected_job_category != "Semua":
        view = view[view["Kategori Pekerjaan"] == selected_job_category]
    if selected_status != "Semua":
        view = view[view["Status Kontrak"] == selected_status]
    st.subheader("Daftar customer dan status kontrak")
    st.metric("Jumlah customer sesuai filter", f"{len(view):,}".replace(",", "."))
    sorted_view = view.sort_values("Sisa Hari", na_position="last")
    st.download_button("Download daftar customer ke Excel", data=excel_bytes(format_outlet_table(sorted_view)), file_name="Daftar_Customer_dan_Status_Kontrak.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if view.empty:
        st.info("Belum ada data yang cocok. Gunakan bagian Upload / Download Excel pada dashboard untuk memasukkan data.")
    else:
        st.dataframe(format_outlet_table(sorted_view), use_container_width=True, hide_index=True)

elif active_menu == "Edit data customer":
    st.header("Edit data customer")
    if contracts.empty:
        st.info("Data customer belum ada. Gunakan menu **Tambah customer baru** atau bagian Upload / Download Excel pada dashboard.")
    else:
        record_options = [f"{i} | {safe_text(row['ID Outlet']) or 'Tanpa ID'} | {safe_text(row['Nama Customer']) or 'Tanpa nama'}" for i, row in contracts.iterrows()]
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
elif active_menu == "Import data branding":
    st.header("Import data outlet branding")
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
                st.error(f"Backup ke Google Sheet gagal: {error}")    "Showcase": ("Showcase", "Show Case", "Materi Showcase"),
    "Layar Toko": ("Layar Toko", "Layar", "Materi Layar Toko"),
    "Shop Sign / Neon Box": ("Shop Sign / Neon Box", "Shop Sign", "Neon Box", "Materi Shop Sign / Neon Box"),
    "Ukuran Panel Toko": ("Ukuran Panel Toko", "Size Panel Toko", "Panel"),
    "Ukuran Sticker Panel Toko": (
        "Ukuran Sticker Panel Toko",
        "Ukuran Sticker Panel (diluar Nama Toko)",
        "Ukuran Stiker Panel Toko",
        "Size Sticker Panel Toko",
    ),
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
    text_columns = [column for column in COLUMNS if column not in MATERIAL_COLUMNS and column not in ("Nominal Kontrak", "Periode Kontrak Awal", "Periode Kontrak Akhir")]
    for column in text_columns:
        normalized[column] = normalized[column].astype("string").str.strip()
        normalized[column] = normalized[column].replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA})
    for column in ("Ukuran Panel Toko", "Ukuran Sticker Panel Toko"):
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
    has_data = normalized.notna().any(axis=1)
    return normalized.loc[has_data.to_numpy(dtype=bool)].reset_index(drop=True)


def read_workbook(source) -> pd.DataFrame:
    for sheet, header in (("COSPKR06", 3), (0, 0), (SHEET_NAME, 0)):
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            result = normalize_data(pd.read_excel(source, sheet_name=sheet, header=header))
            if not result.empty:
                return result
        except (ValueError, KeyError, OSError, ImportError):
            continue
    raise ValueError("Tidak dapat membaca sheet/kolom data kontrak. Periksa file Excel dan nama sheet-nya.")


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


def customer_id_key(value) -> str:
    if pd.isna(value):
        return ""
    key = str(value).strip()
    if re.fullmatch(r"\d+\.0", key):
        key = key[:-2]
    return key.casefold()


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
    invalid_period = (normalized["Periode Kontrak Awal"].notna() & normalized["Periode Kontrak Akhir"].notna() & (normalized["Periode Kontrak Akhir"] < normalized["Periode Kontrak Awal"]))
    if invalid_period.any():
        errors.append("Tanggal akhir kontrak tidak boleh sebelum tanggal awal.")
    return errors


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
    if not url.startswith("https://script.google.com/macros/s/AKfycbzXaUkLDCO7NJlUyuAVpx7qWD4icEk5kw-9IDcXwfJALUjUAmsvNcNt6N14hQVS93Y/exec"):
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
        display_frame["Nominal Kontrak"] = display_frame["Nominal Kontrak"].map(lambda value: "" if pd.isna(value) else f"{float(value):,.0f}".replace(",", "."))
    for column in ("Periode Kontrak Awal", "Periode Kontrak Akhir"):
        if column in display_frame:
            display_frame[column] = pd.to_datetime(display_frame[column], errors="coerce").map(lambda value: "" if pd.isna(value) else value.strftime("%d/%m/%Y"))
    for column in MATERIAL_COLUMNS:
        if column in display_frame:
            display_frame[column] = display_frame[column].map(lambda checked: "✓" if parse_tickmark(checked) else "")
    return display_frame


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
            amount = st.number_input("Nominal Kontrak (Rp)", min_value=0.0, step=100000.0, value=0.0 if pd.isna(parsed_amount) else float(parsed_amount), key=f"amount_{suffix}")
            contract_start = st.date_input("Periode Kontrak Awal", value=safe_date(current["Periode Kontrak Awal"]), key=f"start_{suffix}")
            contract_end = st.date_input("Periode Kontrak Akhir", value=safe_date(current["Periode Kontrak Akhir"]), key=f"end_{suffix}")
        st.markdown("**Materi kontrak**")
        material_cols = st.columns(4)
        material_values = {column: material_cols[index].checkbox(column, value=parse_tickmark(current[column]), key=f"material_{index}_{suffix}") for index, column in enumerate(MATERIAL_COLUMNS)}
        size_col1, size_col2 = st.columns(2)
        panel_size = size_col1.text_input("Ukuran Panel Toko", value=safe_text(current["Ukuran Panel Toko"]), key=f"panel_size_{suffix}")
        sticker_size = size_col2.text_input("Ukuran Sticker Panel Toko", value=safe_text(current["Ukuran Sticker Panel Toko"]), key=f"sticker_size_{suffix}")
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
    }
    return submitted, values


st.title("📊 Monitoring Kontrak Outlet")
st.caption("Pantau data customer dan kontrak, lihat peringatan masa berlaku, serta kelola data melalui form dan Excel.")
if "contracts" not in st.session_state:
    initial, warning = load_initial_data()
    st.session_state.contracts = initial
    st.session_state.load_warning = warning
if st.session_state.load_warning:
    st.warning(st.session_state.load_warning)

with st.sidebar:
    st.header("Menu")
    active_menu = st.radio(
        "Pilih halaman",
        ["Dashboard", "Edit data customer", "Tambah customer baru", "Import data branding", "Backup ke Google Sheet"],
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
    excel_col1, excel_col2 = st.columns(2)
    excel_col1.download_button(
        "Download data Excel",
        data=excel_bytes(st.session_state.contracts),
        file_name="Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    excel_col2.download_button(
        "Download template Excel",
        data=excel_bytes(pd.DataFrame(columns=COLUMNS)),
        file_name="Template_Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Import akan mengganti seluruh data customer. Penyimpanan lokal Streamlit Cloud dapat bersifat sementara.")
    st.subheader("Filter data kontrak")
    f1, f2, f3, f4 = st.columns(4)
    zones = ["Semua"] + sorted(contracts["Zona"].dropna().astype(str).unique().tolist())
    sectors = ["Semua"] + sorted(contracts["Sektor"].dropna().astype(str).unique().tolist())
    job_categories = ["Semua"] + sorted(contracts["Kategori Pekerjaan"].dropna().astype(str).unique().tolist())
    statuses = ["Semua", "Expired", "Kritis (≤ 30 hari)", "Warning (31–90 hari)", "Aktif (> 90 hari)", "Data belum lengkap"]
    selected_zone = f1.selectbox("Zona", zones)
    selected_sector = f2.selectbox("Sektor", sectors)
    selected_job_category = f3.selectbox("Kategori Pekerjaan", job_categories)
    selected_status = f4.selectbox("Status kontrak", statuses)
    view = contracts.copy()
    if selected_zone != "Semua":
        view = view[view["Zona"] == selected_zone]
    if selected_sector != "Semua":
        view = view[view["Sektor"] == selected_sector]
    if selected_job_category != "Semua":
        view = view[view["Kategori Pekerjaan"] == selected_job_category]
    if selected_status != "Semua":
        view = view[view["Status Kontrak"] == selected_status]
    st.subheader("Daftar customer dan status kontrak")
    st.metric("Jumlah customer sesuai filter", f"{len(view):,}".replace(",", "."))
    sorted_view = view.sort_values("Sisa Hari", na_position="last")
    st.download_button("Download daftar customer ke Excel", data=excel_bytes(format_outlet_table(sorted_view)), file_name="Daftar_Customer_dan_Status_Kontrak.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if view.empty:
        st.info("Belum ada data yang cocok. Gunakan bagian Upload / Download Excel pada dashboard untuk memasukkan data.")
    else:
        st.dataframe(format_outlet_table(sorted_view), use_container_width=True, hide_index=True)

elif active_menu == "Edit data customer":
    st.header("Edit data customer")
    if contracts.empty:
        st.info("Data customer belum ada. Gunakan menu **Tambah customer baru** atau bagian Upload / Download Excel pada dashboard.")
    else:
        record_options = [f"{i} | {safe_text(row['ID Outlet']) or 'Tanpa ID'} | {safe_text(row['Nama Customer']) or 'Tanpa nama'}" for i, row in contracts.iterrows()]
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
elif active_menu == "Import data branding":
    st.header("Import data outlet branding")
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
