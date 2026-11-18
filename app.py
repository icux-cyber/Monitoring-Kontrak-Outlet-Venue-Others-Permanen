from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


APP_TITLE = "Monitoring Kontrak Outlet"
DATA_FILE = Path("data") / "Monitoring Kontrak Outlet.xlsx"
SHEET_NAME = "Monitoring Kontrak Outlet"
COLUMNS = [
    "ID Outlet",
    "Nama Outlet",
    "Zona",
    "Sektor",
    "Kategori Outlet",
    "Brand",
    "Nominal Kontrak",
    "Periode Kontrak Awal",
    "Periode Kontrak Akhir",
]
ALIASES = {
    "ID Outlet": ("ID Outlet", "ID Customer", "Customer ID", "Kode Outlet", "Outlet ID"),
    "Nama Outlet": ("Nama Outlet", "Nama Toko", "Customer", "Outlet", "Customer Name"),
    "Zona": ("Zona", "Zone"),
    "Sektor": ("Sektor", "Sector", "Kecamatan"),
    "Kategori Outlet": ("Kategori Outlet", "Customer Group Category", "Pos Budget", "Kategori"),
    "Brand": ("Brand", "Finance Brand", "Merek"),
    "Nominal Kontrak": ("Nominal Kontrak", "Nominal", "Nominal.1", "Contract Amount", "Nilai Kontrak"),
    "Periode Kontrak Awal": ("Periode Kontrak Awal", "Periode Awal Kontrak", "Mulai", "Tanggal Mulai", "Start Date"),
    "Periode Kontrak Akhir": ("Periode Kontrak Akhir", "Periode Akhir Kontrak", "Akhir", "Tanggal Akhir", "End Date", "Expiry Date"),
}

st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")


def _clean_headers(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame


def normalize_data(frame: pd.DataFrame) -> pd.DataFrame:
    """Normalize common Indonesian/English workbook headers into app columns."""
    frame = _clean_headers(frame)
    normalized = pd.DataFrame(index=frame.index)
    lowered = {str(column).strip().casefold(): column for column in frame.columns}
    for target, names in ALIASES.items():
        source = next((lowered[name.casefold()] for name in names if name.casefold() in lowered), None)
        normalized[target] = frame[source] if source is not None else pd.NA

    for column in ("ID Outlet", "Nama Outlet", "Zona", "Sektor", "Kategori Outlet", "Brand"):
        normalized[column] = normalized[column].astype("string").str.strip()
        normalized[column] = normalized[column].replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA})

    normalized["Nominal Kontrak"] = normalized["Nominal Kontrak"].map(parse_amount)
    for column in ("Periode Kontrak Awal", "Periode Kontrak Akhir"):
        normalized[column] = pd.to_datetime(normalized[column], errors="coerce", dayfirst=True).dt.tz_localize(None)

    normalized = normalized[COLUMNS]
    # Ignore completely blank rows, while preserving rows with partial data for editing.
    has_data = normalized.astype("string").fillna("").apply(lambda row: row.str.strip().ne("").any(), axis=1)
    return normalized.loc[has_data].reset_index(drop=True)


def parse_amount(value):
    if pd.isna(value) or str(value).strip() == "":
        return pd.NA
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace("Rp", "").replace("rp", "").replace(" ", "")
    # Handle both Indonesian (1.234.567,89) and English (1,234,567.89) formatting.
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        tail = text.rsplit(",", 1)[-1]
        text = text.replace(",", ".") if len(tail) <= 2 else text.replace(",", "")
    elif text.count(".") > 1:
        text = text.replace(".", "")
    try:
        return float(text)
    except ValueError:
        return pd.NA


def read_workbook(source) -> pd.DataFrame:
    """Read the supplied template or the original COSPKR06 workbook."""
    errors = []
    for sheet, header in (("COSPKR06", 3), (0, 0), ("Monitoring Kontrak Outlet", 0)):
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            raw = pd.read_excel(source, sheet_name=sheet, header=header)
            candidate = normalize_data(raw)
            if not raw.empty and any(candidate[column].notna().any() for column in COLUMNS):
                return candidate
        except (ValueError, KeyError, OSError, ImportError) as error:
            errors.append(str(error))
    raise ValueError("Tidak dapat membaca sheet/kolom data kontrak. Periksa file Excel dan nama sheet-nya.")


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
    if (normalized["Nama Outlet"].isna() & normalized["ID Outlet"].isna()).any():
        errors.append("Setiap baris harus berisi Nama Outlet atau ID Outlet.")
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
    today_ts = pd.Timestamp(today or date.today()).normalize()
    remaining = (result["Periode Kontrak Akhir"] - today_ts).dt.days
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
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    return output.getvalue()


def save_data(frame: pd.DataFrame) -> None:
    errors = validate_data(frame)
    if errors:
        raise ValueError("\n".join(errors))
    normalized = normalize_data(frame)
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    normalized.to_excel(DATA_FILE, sheet_name=SHEET_NAME, index=False, engine="openpyxl")


st.title("📊 Monitoring Kontrak Outlet")
st.caption("Pantau data outlet dan kontrak, lihat peringatan masa berlaku, serta kelola data melalui Excel.")

if "contracts" not in st.session_state:
    initial, load_warning = load_initial_data()
    st.session_state.contracts = initial
    st.session_state.load_warning = load_warning

if st.session_state.load_warning:
    st.warning(st.session_state.load_warning)

with st.sidebar:
    st.header("Kelola data")
    upload = st.file_uploader("Upload workbook Excel (.xlsx)", type=["xlsx"], key="workbook_upload")
    if st.button("Import dan ganti data", type="primary", use_container_width=True):
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

    st.download_button(
        "Download data Excel",
        data=excel_bytes(st.session_state.contracts),
        file_name="Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.download_button(
        "Download template Excel",
        data=excel_bytes(pd.DataFrame(columns=COLUMNS)),
        file_name="Template_Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Catatan: penyimpanan file lokal Streamlit Cloud dapat bersifat sementara. Unduh backup data secara berkala.")

contracts = add_contract_status(st.session_state.contracts)
expired = int(contracts["Status Kontrak"].eq("Expired").sum())
soon = int(contracts["Status Kontrak"].isin(["Kritis (≤ 30 hari)", "Warning (31–90 hari)"]).sum())
total_value = pd.to_numeric(contracts["Nominal Kontrak"], errors="coerce").sum()

m1, m2, m3, m4 = st.columns(4)
m1.metric("Jumlah outlet", f"{len(contracts):,}".replace(",", "."))
m2.metric("Total nominal kontrak", f"Rp {total_value:,.0f}".replace(",", "."))
m3.metric("Akan berakhir ≤ 90 hari", soon)
m4.metric("Sudah berakhir", expired)
if expired:
    st.error(f"Perhatian: {expired} kontrak sudah berakhir.")
if soon:
    st.warning(f"Pengingat: {soon} kontrak akan berakhir dalam 90 hari.")

st.subheader("Filter data kontrak")
f1, f2, f3 = st.columns(3)
zones = ["Semua"] + sorted(contracts["Zona"].dropna().astype(str).unique().tolist())
sectors = ["Semua"] + sorted(contracts["Sektor"].dropna().astype(str).unique().tolist())
statuses = ["Semua"] + ["Expired", "Kritis (≤ 30 hari)", "Warning (31–90 hari)", "Aktif (> 90 hari)", "Data belum lengkap"]
selected_zone = f1.selectbox("Zona", zones)
selected_sector = f2.selectbox("Sektor", sectors)
selected_status = f3.selectbox("Status kontrak", statuses)
view = contracts.copy()
if selected_zone != "Semua":
    view = view[view["Zona"] == selected_zone]
if selected_sector != "Semua":
    view = view[view["Sektor"] == selected_sector]
if selected_status != "Semua":
    view = view[view["Status Kontrak"] == selected_status]

st.subheader("Daftar outlet dan status kontrak")
if view.empty:
    st.info("Belum ada data yang cocok. Gunakan menu Upload Excel di sidebar untuk memasukkan data.")
else:
    st.dataframe(view.sort_values("Sisa Hari", na_position="last"), use_container_width=True, hide_index=True)

st.divider()
st.subheader("Edit data outlet")
st.caption("Ubah sel di tabel, tambah baris di bagian bawah, lalu simpan. Kolom status dan sisa hari dihitung otomatis.")
with st.form("edit_contracts_form"):
    edited = st.data_editor(
        st.session_state.contracts,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="contracts_editor",
        column_config={
            "Nominal Kontrak": st.column_config.NumberColumn("Nominal Kontrak (Rp)", min_value=0, step=100000),
            "Periode Kontrak Awal": st.column_config.DateColumn("Periode Kontrak Awal", format="DD/MM/YYYY"),
            "Periode Kontrak Akhir": st.column_config.DateColumn("Periode Kontrak Akhir", format="DD/MM/YYYY"),
        },
    )
    save_button = st.form_submit_button("Simpan perubahan", type="primary")

if save_button:
    problems = validate_data(edited)
    if problems:
        st.error("\n".join(problems))
    else:
        try:
            save_data(edited)
            st.session_state.contracts = normalize_data(edited)
            st.success("Perubahan berhasil disimpan.")
            st.rerun()
        except Exception as error:
            st.error(f"Gagal menyimpan data: {error}")
