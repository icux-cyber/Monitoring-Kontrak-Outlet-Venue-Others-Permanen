from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

APP_TITLE = "Monitoring Kontrak Outlet"
DATA_FILE = Path("data") / "DRP & Outlet MCD - DSO Surabaya Barat.xlsx"
SHEET_NAME = "Monitoring Kontrak Outlet"
COLUMNS = [
    "ID Outlet", "Nama Outlet", "Alamat", "Zona", "Sektor",
    "Kategori Outlet", "Kategori Pekerjaan", "Brand", "Nominal Kontrak",
    "Periode Kontrak Awal", "Periode Kontrak Akhir",
]
ALIASES = {
    "ID Outlet": ("ID Outlet", "ID Customer", "Customer ID", "Kode Outlet", "Outlet ID"),
    "Nama Outlet": ("Nama Outlet", "Nama Toko", "Customer", "Outlet", "Customer Name"),
    "Alamat": ("Alamat", "Alamat Outlet", "Alamat Toko", "Address", "Outlet Address", "Customer Address"),
    "Zona": ("Zona", "Zone"),
    "Sektor": ("Sektor", "Sector", "Kecamatan"),
    "Kategori Outlet": ("Kategori Outlet", "Customer Group Category", "Pos Budget", "Kategori"),
    "Kategori Pekerjaan": ("Kategori Pekerjaan", "Kategori Pekerjaan Outlet", "Work Category", "Job Category", "Kategori Kerja"),
    "Brand": ("Brand", "Finance Brand", "Merek"),
    "Nominal Kontrak": ("Nominal Kontrak", "Nominal", "Nominal.1", "Contract Amount", "Nilai Kontrak"),
    "Periode Kontrak Awal": ("Periode Kontrak Awal", "Periode Awal Kontrak", "Mulai", "Tanggal Mulai", "Start Date"),
    "Periode Kontrak Akhir": ("Periode Kontrak Akhir", "Periode Akhir Kontrak", "Akhir", "Tanggal Akhir", "End Date", "Expiry Date"),
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


def normalize_data(frame: pd.DataFrame) -> pd.DataFrame:
    frame = clean_headers(frame)
    normalized = pd.DataFrame(index=frame.index)
    lowered = {str(column).strip().casefold(): column for column in frame.columns}
    for target, names in ALIASES.items():
        source = next((lowered[name.casefold()] for name in names if name.casefold() in lowered), None)
        normalized[target] = frame[source] if source is not None else pd.NA
    text_columns = [
        "ID Outlet", "Nama Outlet", "Alamat", "Zona", "Sektor",
        "Kategori Outlet", "Kategori Pekerjaan", "Brand",
    ]
    for column in text_columns:
        normalized[column] = normalized[column].astype("string").str.strip()
        normalized[column] = normalized[column].replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA})
    normalized["Nominal Kontrak"] = normalized["Nominal Kontrak"].map(parse_amount)
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


def safe_text(value) -> str:
    return "" if pd.isna(value) else str(value)


def safe_date(value):
    return None if pd.isna(value) else pd.Timestamp(value).date()


def render_outlet_form(current: pd.Series, row_index: int | None) -> dict:
    suffix = "new" if row_index is None else f"row_{row_index}"
    with st.form(f"outlet_form_{suffix}"):
        st.subheader("Data outlet")
        left, right = st.columns(2)
        with left:
            outlet_id = st.text_input("ID Outlet", value=safe_text(current["ID Outlet"]), key=f"id_{suffix}")
            outlet_name = st.text_input("Nama Outlet", value=safe_text(current["Nama Outlet"]), key=f"name_{suffix}")
            address = st.text_area("Alamat", value=safe_text(current["Alamat"]), key=f"address_{suffix}")
            zone = st.text_input("Zona", value=safe_text(current["Zona"]), key=f"zone_{suffix}")
            sector = st.text_input("Sektor", value=safe_text(current["Sektor"]), key=f"sector_{suffix}")
            outlet_category = st.text_input("Kategori Outlet", value=safe_text(current["Kategori Outlet"]), key=f"outlet_category_{suffix}")
        with right:
            job_category = st.text_input("Kategori Pekerjaan", value=safe_text(current["Kategori Pekerjaan"]), key=f"job_category_{suffix}")
            brand = st.text_input("Brand", value=safe_text(current["Brand"]), key=f"brand_{suffix}")
            parsed_amount = parse_amount(current["Nominal Kontrak"])
            amount = st.number_input(
                "Nominal Kontrak (Rp)", min_value=0.0, step=100000.0,
                value=0.0 if pd.isna(parsed_amount) else float(parsed_amount), key=f"amount_{suffix}",
            )
            contract_start = st.date_input(
                "Periode Kontrak Awal", value=safe_date(current["Periode Kontrak Awal"]), key=f"start_{suffix}"
            )
            contract_end = st.date_input(
                "Periode Kontrak Akhir", value=safe_date(current["Periode Kontrak Akhir"]), key=f"end_{suffix}"
            )
        submitted = st.form_submit_button("Simpan data outlet", type="primary")
    values = {
        "ID Outlet": outlet_id.strip() or pd.NA,
        "Nama Outlet": outlet_name.strip() or pd.NA,
        "Alamat": address.strip() or pd.NA,
        "Zona": zone.strip() or pd.NA,
        "Sektor": sector.strip() or pd.NA,
        "Kategori Outlet": outlet_category.strip() or pd.NA,
        "Kategori Pekerjaan": job_category.strip() or pd.NA,
        "Brand": brand.strip() or pd.NA,
        "Nominal Kontrak": amount,
        "Periode Kontrak Awal": pd.Timestamp(contract_start) if contract_start else pd.NaT,
        "Periode Kontrak Akhir": pd.Timestamp(contract_end) if contract_end else pd.NaT,
    }
    return submitted, values


st.title("📊 Monitoring Kontrak Outlet")
st.caption("Pantau data outlet dan kontrak, lihat peringatan masa berlaku, serta kelola data melalui form dan Excel.")

if "contracts" not in st.session_state:
    initial, warning = load_initial_data()
    st.session_state.contracts = initial
    st.session_state.load_warning = warning
if st.session_state.load_warning:
    st.warning(st.session_state.load_warning)

with st.sidebar:
    st.header("Menu")
    active_menu = st.radio("Pilih halaman", ["Dashboard", "Edit data outlet", "Tambah outlet baru"], key="main_menu")
    st.divider()
    st.subheader("Upload / Download Excel")
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
        "Download data Excel", data=excel_bytes(st.session_state.contracts),
        file_name="Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
    )
    st.download_button(
        "Download template Excel", data=excel_bytes(pd.DataFrame(columns=COLUMNS)),
        file_name="Template_Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
    )
    st.caption("Penyimpanan lokal Streamlit Cloud dapat bersifat sementara. Unduh backup data secara berkala.")

contracts = add_contract_status(st.session_state.contracts)

if active_menu == "Dashboard":
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
    st.subheader("Daftar outlet dan status kontrak")
    if view.empty:
        st.info("Belum ada data yang cocok. Gunakan menu Upload Excel di sidebar untuk memasukkan data.")
    else:
        st.dataframe(view.sort_values("Sisa Hari", na_position="last"), use_container_width=True, hide_index=True)

elif active_menu == "Edit data outlet":
    st.header("Edit data outlet")
    if contracts.empty:
        st.info("Data master outlet belum ada. Gunakan menu **Tambah outlet baru** atau import workbook Excel dari sidebar.")
    else:
        record_options = []
        for row_index, row in contracts.iterrows():
            outlet_id = safe_text(row["ID Outlet"]) or "Tanpa ID"
            outlet_name = safe_text(row["Nama Outlet"]) or "Tanpa nama"
            record_options.append(f"{row_index} | {outlet_id} | {outlet_name}")
        selected_record = st.selectbox("Pilih outlet yang akan diedit", record_options, key="selected_outlet_record")
        selected_index = int(selected_record.split(" | ", 1)[0])
        current = contracts.loc[selected_index]
        submitted, updated_record = render_outlet_form(current, selected_index)
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
                    st.success("Data outlet berhasil diperbarui.")
                    st.rerun()
                except Exception as error:
                    st.error(f"Gagal menyimpan data: {error}")

elif active_menu == "Tambah outlet baru":
    st.header("Tambah outlet baru")
    st.caption("Isi form di bawah. Data master yang sudah ada tidak diubah.")
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
                st.success("Outlet baru berhasil ditambahkan.")
                st.rerun()
            except Exception as error:
                st.error(f"Gagal menyimpan outlet baru: {error}")
