from __future__ import annotations

import re

import pandas as pd

from services.contracts_service import (
    COLUMNS,
    DATE_COLUMNS,
    attach_search_blob,
    clean_headers,
    customer_id_key,
    normalize_data,
    parse_amount,
    safe_text,
)


def read_drp_crosscheck_workbook(source) -> tuple[pd.DataFrame, int]:
    """Membaca file DRP, mengembalikan DataFrame terdeduplikasi per ID Customer beserta jumlah baris duplikat."""
    if hasattr(source, "seek"):
        source.seek(0)
    excel_file = pd.ExcelFile(source, engine="openpyxl")
    candidate_sheets = ["CODPR03"] + [s for s in excel_file.sheet_names if s != "CODPR03"]
    parsed_df: pd.DataFrame | None = None

    for sheet in candidate_sheets:
        for hdr in (8, 0, 3, 9, 10):
            try:
                if hasattr(source, "seek"):
                    source.seek(0)
                raw = pd.read_excel(source, sheet_name=sheet, header=hdr, engine="openpyxl")
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
    result["_id_key"] = result["ID Customer"].map(customer_id_key)
    result = result[result["_id_key"].ne("")].reset_index(drop=True)

    dup_count = int(result.duplicated("_id_key", keep="last").sum())
    if dup_count > 0:
        result["_row_order"] = range(len(result))
        result = (
            result.sort_values(["Periode Kontrak Akhir", "_row_order"], kind="stable", na_position="first")
            .drop_duplicates("_id_key", keep="last")
            .sort_values("_row_order", kind="stable")
            .drop(columns=["_row_order"], errors="ignore")
            .reset_index(drop=True)
        )

    result = result.drop(columns=["_id_key"], errors="ignore")
    return result, dup_count


def fmt_val_cmp(col: str, val) -> str:
    if pd.isna(val) or str(val).strip() == "":
        return "-"
    if col == "Nominal Kontrak":
        num = parse_amount(val)
        return "-" if pd.isna(num) else f"Rp {float(num):,.0f}".replace(",", ".")
    if col in DATE_COLUMNS:
        ts = pd.to_datetime(val, errors="coerce")
        return "-" if pd.isna(ts) else ts.strftime("%d/%m/%Y")
    return str(val).strip()


def is_field_different(col: str, dash_val, drp_val) -> bool:
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


def is_drp_job_category(val) -> bool:
    txt = safe_text(val).casefold()
    if not txt:
        return False
    return bool(re.search(r"\bdrp\b", txt))


def build_drp_crosscheck_report(
    contracts: pd.DataFrame,
    drp_df: pd.DataFrame,
    compare_fields: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    dash_all = contracts.reset_index(drop=True)
    dash = dash_all[dash_all["Kategori Pekerjaan"].map(is_drp_job_category)].copy()
    dash_keys = dash["ID Outlet"].map(customer_id_key)
    key_to_dash_idx: dict[str, int] = {}
    for idx, k in dash_keys.items():
        if k and k not in key_to_dash_idx:
            key_to_dash_idx[k] = int(idx)

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
            if is_field_different(
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
            "Brand (Dash)": fmt_val_cmp("Brand", d_row.get("Brand")),
            "Brand (DRP)": fmt_val_cmp("Brand", drp_row.get("Brand")),
            "Nominal (Dash)": fmt_val_cmp("Nominal Kontrak", d_row.get("Nominal Kontrak")),
            "Nominal (DRP)": fmt_val_cmp("Nominal Kontrak", drp_row.get("Nominal Kontrak")),
            "Awal Kontrak (Dash)": fmt_val_cmp("Periode Kontrak Awal", d_row.get("Periode Kontrak Awal")),
            "Awal Kontrak (DRP)": fmt_val_cmp("Periode Kontrak Awal", drp_row.get("Periode Kontrak Awal")),
            "Akhir Kontrak (Dash)": fmt_val_cmp("Periode Kontrak Akhir", d_row.get("Periode Kontrak Akhir")),
            "Akhir Kontrak (DRP)": fmt_val_cmp("Periode Kontrak Akhir", drp_row.get("Periode Kontrak Akhir")),
            "No. Dokumen (Dash)": fmt_val_cmp("No. Dokumen", d_row.get("No. Dokumen")),
            "No. Dokumen (DRP)": fmt_val_cmp("No. Dokumen", drp_row.get("No. Dokumen")),
            "Zona (Dash)": fmt_val_cmp("Zona", d_row.get("Zona")),
            "Zona (DRP)": fmt_val_cmp("Zona", drp_row.get("Zona")),
            "Kategori Outlet (Dash)": fmt_val_cmp("Kategori Outlet", d_row.get("Kategori Outlet")),
            "Kategori Outlet (DRP)": fmt_val_cmp("Kategori Outlet", drp_row.get("Kategori Outlet")),
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


def build_dashboard_only_report(
    contracts: pd.DataFrame,
    drp_df: pd.DataFrame,
) -> pd.DataFrame:
    dash_all = contracts.reset_index(drop=True)
    drp_keys = {
        customer_id_key(val)
        for val in drp_df.get("ID Customer", pd.Series(dtype="object")).dropna()
        if customer_id_key(val)
    }

    rows = []
    for idx, d_row in dash_all.iterrows():
        if not is_drp_job_category(d_row.get("Kategori Pekerjaan")):
            continue
        raw_id = safe_text(d_row.get("ID Outlet"))
        k = customer_id_key(raw_id)
        if not k or k not in drp_keys:
            rows.append({
                "_dash_index": int(idx),
                "ID Outlet": raw_id or "-",
                "Nama Customer": safe_text(d_row.get("Nama Customer")) or "-",
                "Zona": safe_text(d_row.get("Zona")) or "-",
                "Sektor": safe_text(d_row.get("Sektor")) or "-",
                "Pos Budget": safe_text(d_row.get("Pos Budget")) or "-",
                "Kategori Pekerjaan": safe_text(d_row.get("Kategori Pekerjaan")) or "-",
                "Brand": safe_text(d_row.get("Brand")) or "-",
                "Nominal Kontrak": fmt_val_cmp("Nominal Kontrak", d_row.get("Nominal Kontrak")),
                "Periode Kontrak Awal": fmt_val_cmp("Periode Kontrak Awal", d_row.get("Periode Kontrak Awal")),
                "Periode Kontrak Akhir": fmt_val_cmp("Periode Kontrak Akhir", d_row.get("Periode Kontrak Akhir")),
                "No. Dokumen": safe_text(d_row.get("No. Dokumen")) or "-",
            })

    return pd.DataFrame(rows)


def apply_drp_updates_to_contracts(
    contracts: pd.DataFrame,
    rows_to_update: pd.DataFrame,
    fields_to_apply: list[str],
    already_normalized: bool = False,
) -> tuple[pd.DataFrame, int]:
    updated = contracts.copy() if already_normalized else normalize_data(contracts).copy()
    updated_counter = 0
    for _, r in rows_to_update.iterrows():
        target_idx = int(r["_dash_index"])
        changed = False
        for col in fields_to_apply:
            raw_val = r.get(f"_raw_{col}", pd.NA)
            if pd.isna(raw_val) or (isinstance(raw_val, str) and not raw_val.strip()):
                continue
            if is_field_different(col, updated.at[target_idx, col], raw_val):
                updated.at[target_idx, col] = raw_val
                changed = True
        if changed:
            updated_counter += 1
    return attach_search_blob(updated.reset_index(drop=True)), updated_counter


def append_missing_drp_to_contracts(
    contracts: pd.DataFrame,
    rows_to_insert_raw: pd.DataFrame,
    default_sector: str,
    default_budget: str,
    default_job_cat: str,
    already_normalized: bool = False,
) -> tuple[pd.DataFrame, int, int]:
    current = contracts if already_normalized else normalize_data(contracts)
    existing_keys = {
        customer_id_key(val)
        for val in current["ID Outlet"].dropna()
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
            "Ukuran Sticker / Vinyl Panel Toko": pd.NA,
            "Brand Panel": pd.NA,
            "Versi": pd.NA,
            "Tanggal Penggantian Sticker / Vinyl Panel": pd.NaT,
            "Nama Vendor": pd.NA,
            "Ukuran Layar": pd.NA,
            "Brand Layar": pd.NA,
        })

    if not new_records:
        return attach_search_blob(current[COLUMNS].reset_index(drop=True)), 0, skipped_existing

    combined = pd.concat([current[COLUMNS], pd.DataFrame(new_records)], ignore_index=True)
    return attach_search_blob(combined), len(new_records), skipped_existing
