from __future__ import annotations

import io
import json
import re
from datetime import date
from pathlib import Path

import matplotlib.patches as patches
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

from services.contracts_service import (
    DATA_DIR,
    attach_search_blob,
    clean_headers,
    customer_id_key,
    dataframe_fingerprint,
    extract_sizes,
    normalize_data,
    normalize_name_key,
    safe_text,
)

PROMO_STOCK_FILE = DATA_DIR / "PRCRG01_Gudang_Promosi.xlsx"
PEJABAT_EMAIL_FILE = DATA_DIR / "master_email_pejabat.xlsx"
STICKER_ORDER_LOG_FILE = DATA_DIR / "log_hasil_pemesanan_sticker.xlsx"

DEFAULT_PEJABAT_EMAILS = [
    "hengky.s.handoko@sumbercipta.com",
    "amin.prayogi@sumbercipta.com",
    "mochamad.basori@sumbercipta.com",
]

STICKER_ORDER_COLUMNS = [
    "ID Outlet",
    "Nama Outlet",
    "Alamat",
    "Material",
    "Brand",
    "Version",
    "Jumlah Sticker / Vinyl",
    "Tanggal OrderPenggantian Sticker / Vinyl",
    "Nama Vendor",
]

STICKER_LOG_COLUMNS = [
    "Waktu Input",
    "ID Outlet",
    "Nama Outlet",
    "Alamat",
    "Zona",
    "Material",
    "Ukuran",
    "Brand",
    "Version",
    "Jumlah Sticker / Vinyl",
    "Tanggal OrderPenggantian Sticker / Vinyl",
    "Nama Vendor",
    "Status Kirim Pejabat",
]


def _as_1d_series(obj, index: pd.Index) -> pd.Series:
    if isinstance(obj, pd.DataFrame):
        if obj.shape[1] == 0:
            return pd.Series(pd.NA, index=index)
        return obj.iloc[:, 0].reindex(index)
    if isinstance(obj, pd.Series):
        return obj.reindex(index)
    return pd.Series(obj, index=index)


def is_sticker_or_vinyl_panel_toko(value) -> bool:
    if pd.isna(value):
        return False
    return bool(
        re.search(
            r"\b(?:sticker|stiker|vinyl)\s+panel\s+toko\b",
            str(value).strip(),
            flags=re.IGNORECASE,
        )
    )


def normalize_sticker_order_records(records: list[dict] | pd.DataFrame) -> list[dict]:
    if isinstance(records, pd.DataFrame):
        raw_list = records.to_dict(orient="records") if not records.empty else []
    else:
        raw_list = list(records or [])

    now_str = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    cleaned: list[dict] = []
    for item in raw_list:
        if not isinstance(item, dict):
            continue
        qty_raw = item.get("Jumlah Sticker / Vinyl", 0)
        try:
            qty_val = int(float(qty_raw)) if pd.notna(qty_raw) and str(qty_raw).strip() != "" else 0
        except (TypeError, ValueError):
            qty_val = 0

        cleaned.append({
            "Waktu Input": safe_text(item.get("Waktu Input")) or now_str,
            "ID Outlet": safe_text(item.get("ID Outlet")) or "-",
            "Nama Outlet": safe_text(item.get("Nama Outlet")) or "-",
            "Alamat": safe_text(item.get("Alamat")) or "-",
            "Zona": safe_text(item.get("Zona")) or "-",
            "Material": safe_text(item.get("Material")) or "-",
            "Ukuran": safe_text(item.get("Ukuran")) or "-",
            "Brand": safe_text(item.get("Brand")) or "-",
            "Version": safe_text(item.get("Version")) or "-",
            "Jumlah Sticker / Vinyl": int(qty_val),
            "Tanggal OrderPenggantian Sticker / Vinyl": (
                safe_text(item.get("Tanggal OrderPenggantian Sticker / Vinyl"))
                or date.today().strftime("%d/%m/%Y")
            ),
            "Nama Vendor": safe_text(item.get("Nama Vendor")) or "-",
            "Status Kirim Pejabat": safe_text(item.get("Status Kirim Pejabat")) or "Siap Kirim",
        })
    return cleaned


def save_sticker_order_log(records: list[dict] | pd.DataFrame) -> list[dict]:
    cleaned = normalize_sticker_order_records(records)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    df_log = pd.DataFrame(cleaned, columns=STICKER_LOG_COLUMNS)
    df_log.to_excel(STICKER_ORDER_LOG_FILE, sheet_name="Log Pemesanan Sticker", index=False, engine="openpyxl")
    return cleaned


def load_sticker_order_log() -> list[dict]:
    if not STICKER_ORDER_LOG_FILE.exists():
        return []
    try:
        df_log = pd.read_excel(STICKER_ORDER_LOG_FILE, sheet_name=0, header=0, engine="openpyxl")
        if df_log.empty:
            return []
        return normalize_sticker_order_records(df_log)
    except Exception:
        return []


def append_sticker_order_log(new_records: list[dict], existing_records: list[dict]) -> list[dict]:
    now_str = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    stamped: list[dict] = []
    for rec in new_records:
        item = dict(rec)
        item.setdefault("Waktu Input", now_str)
        item.setdefault("Status Kirim Pejabat", "Siap Kirim")
        stamped.append(item)
    merged = normalize_sticker_order_records(stamped + list(existing_records or []))
    return save_sticker_order_log(merged)


def update_log_status_for_rows(
    all_records: list[dict],
    target_df: pd.DataFrame,
    new_status: str,
) -> list[dict]:
    normalized_all = normalize_sticker_order_records(all_records)
    if target_df is None or target_df.empty or not normalized_all:
        return normalized_all

    target_keys: set[tuple[str, str, str, str, str, str]] = set()
    for _, r in target_df.iterrows():
        target_keys.add((
            safe_text(r.get("Waktu Input")),
            safe_text(r.get("ID Outlet")),
            safe_text(r.get("Nama Outlet")),
            safe_text(r.get("Brand")),
            safe_text(r.get("Version")),
            safe_text(r.get("Tanggal OrderPenggantian Sticker / Vinyl")),
        ))

    updated: list[dict] = []
    for item in normalized_all:
        rec = dict(item)
        k = (
            safe_text(rec.get("Waktu Input")),
            safe_text(rec.get("ID Outlet")),
            safe_text(rec.get("Nama Outlet")),
            safe_text(rec.get("Brand")),
            safe_text(rec.get("Version")),
            safe_text(rec.get("Tanggal OrderPenggantian Sticker / Vinyl")),
        )
        if k in target_keys:
            rec["Status Kirim Pejabat"] = new_status
        updated.append(rec)

    return save_sticker_order_log(updated)


def load_pejabat_emails() -> list[str]:
    emails: list[str] = []
    if PEJABAT_EMAIL_FILE.exists():
        try:
            df_p = pd.read_excel(PEJABAT_EMAIL_FILE, sheet_name=0, engine="openpyxl")
            for col in df_p.columns:
                for val in df_p[col].dropna().astype(str):
                    for part in re.split(r"[;,]", val):
                        em = part.strip()
                        if "@" in em and em not in emails:
                            emails.append(em)
        except Exception:
            pass
    if not emails:
        try:
            sec_val = st.secrets.get("PEJABAT_EMAILS", "")
            raw_items = [str(x) for x in sec_val] if isinstance(sec_val, (list, tuple)) else re.split(r"[;,]", str(sec_val))
            for item in raw_items:
                em = item.strip()
                if "@" in em and em not in emails:
                    emails.append(em)
        except Exception:
            pass
    if not emails:
        emails = list(DEFAULT_PEJABAT_EMAILS)
    return emails


def save_pejabat_emails(emails: list[str]) -> None:
    cleaned = []
    for item in emails:
        for part in re.split(r"[;,]", str(item)):
            em = part.strip()
            if "@" in em and em not in cleaned:
                cleaned.append(em)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    pd.DataFrame({"Email Pejabat": cleaned}).to_excel(PEJABAT_EMAIL_FILE, index=False, engine="openpyxl")


def read_promo_stock_workbook(source) -> pd.DataFrame:
    empty_cols = [
        "Storage Location",
        "Material Group",
        "Material",
        "Brand",
        "Version",
        "Available",
        "UOM",
        "Status Material",
    ]
    if hasattr(source, "seek"):
        source.seek(0)

    excel_file = pd.ExcelFile(source, engine="openpyxl")
    sheet = "PRCRG01" if "PRCRG01" in excel_file.sheet_names else excel_file.sheet_names[0]

    if hasattr(source, "seek"):
        source.seek(0)
    raw = pd.read_excel(source, sheet_name=sheet, header=None, engine="openpyxl")
    if raw.empty:
        return pd.DataFrame(columns=empty_cols)

    header_row_idx = None
    for r_idx in range(min(len(raw), 15)):
        row_vals = [str(v).strip().casefold() for v in raw.iloc[r_idx].tolist() if pd.notna(v)]
        if "material" in row_vals and ("brand" in row_vals or "version" in row_vals):
            header_row_idx = r_idx
            break

    if header_row_idx is not None:
        h_top = raw.iloc[header_row_idx].ffill()
        h_sub = raw.iloc[header_row_idx + 1] if header_row_idx + 1 < len(raw) else pd.Series(index=raw.columns)
        col_names = []
        for top_v, sub_v in zip(h_top, h_sub):
            s_top = "" if pd.isna(top_v) else str(top_v).strip()
            s_sub = "" if pd.isna(sub_v) else str(sub_v).strip()
            if s_top and s_sub and s_top.casefold() != s_sub.casefold():
                col_names.append(f"{s_top} - {s_sub}")
            elif s_sub:
                col_names.append(s_sub)
            else:
                col_names.append(s_top)

        data_df = raw.iloc[header_row_idx + 2 :].copy()
        data_df.columns = col_names
        data_df = data_df.dropna(how="all").reset_index(drop=True)

        def pick_col(*candidates) -> pd.Series:
            for cand in candidates:
                c_low = cand.casefold()
                for idx_c, col_name in enumerate(data_df.columns):
                    if str(col_name).strip().casefold() == c_low:
                        return _as_1d_series(data_df.iloc[:, idx_c], data_df.index)
            return pd.Series(pd.NA, index=data_df.index)

        out = pd.DataFrame({
            "Storage Location": pick_col("Storage Location"),
            "Material Group": pick_col("Material Group"),
            "Material": pick_col("Material"),
            "Brand": pick_col("Brand"),
            "Version": pick_col("Version", "Versi"),
            "Available": pick_col("Stock Type - Available", "Available"),
            "UOM": pick_col("Stock Type - UOM", "UOM"),
            "Status Material": pick_col("Status Material"),
        })
    else:
        if hasattr(source, "seek"):
            source.seek(0)
        simple_df = clean_headers(
            pd.read_excel(source, sheet_name=sheet, header=0, engine="openpyxl")
        ).reset_index(drop=True)

        def pick_simple(*candidates) -> pd.Series:
            for cand in candidates:
                c_low = cand.casefold()
                for idx_c, col_name in enumerate(simple_df.columns):
                    if str(col_name).strip().casefold() == c_low:
                        return _as_1d_series(simple_df.iloc[:, idx_c], simple_df.index)
            return pd.Series(pd.NA, index=simple_df.index)

        out = pd.DataFrame({
            "Storage Location": pick_simple("Storage Location"),
            "Material Group": pick_simple("Material Group"),
            "Material": pick_simple("Material"),
            "Brand": pick_simple("Brand"),
            "Version": pick_simple("Version", "Versi"),
            "Available": pick_simple("Available", "Stock Type - Available", "Jumlah Sticker / Vinyl"),
            "UOM": pick_simple("UOM"),
            "Status Material": pick_simple("Status Material"),
        })

    for col in ("Storage Location", "Material Group", "Material", "Brand", "Version", "UOM", "Status Material"):
        s = _as_1d_series(out[col], out.index).astype("string").str.strip()
        out[col] = s.replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA, "None": pd.NA})

    avail_num = pd.to_numeric(_as_1d_series(out["Available"], out.index), errors="coerce").fillna(0)
    out["Available"] = avail_num
    mask_mat = out["Material"].map(is_sticker_or_vinyl_panel_toko).fillna(False).to_numpy(dtype=bool)
    mask_avail = (out["Available"] > 0).to_numpy(dtype=bool)
    return out.loc[mask_mat & mask_avail].reset_index(drop=True)


@st.cache_data(show_spinner=False)
def _cached_initial_promo_stock(file_mtime: float | None, file_path_str: str) -> pd.DataFrame:
    if PROMO_STOCK_FILE.exists():
        try:
            return read_promo_stock_workbook(PROMO_STOCK_FILE)
        except Exception:
            pass
    default_candidates = [
        Path("PRCRG01 - 2026-11-19T083353.650.xlsx"),
        Path(__file__).resolve().parent.parent / "PRCRG01 - 2026-11-19T083353.650.xlsx",
    ]
    for candidate in default_candidates:
        if candidate.exists():
            try:
                df = read_promo_stock_workbook(candidate)
                if not df.empty:
                    DATA_DIR.mkdir(parents=True, exist_ok=True)
                    df.to_excel(PROMO_STOCK_FILE, sheet_name="PRCRG01", index=False, engine="openpyxl")
                return df
            except Exception:
                continue
    return pd.DataFrame(
        columns=["Storage Location", "Material Group", "Material", "Brand", "Version", "Available", "UOM", "Status Material"]
    )


def load_initial_promo_stock() -> pd.DataFrame:
    try:
        mtime = PROMO_STOCK_FILE.stat().st_mtime if PROMO_STOCK_FILE.exists() else None
    except Exception:
        mtime = None
    return _cached_initial_promo_stock(mtime, str(PROMO_STOCK_FILE))


def clear_promo_stock_cache() -> None:
    _cached_initial_promo_stock.clear()


def get_warehouse_stock_options(promo_df: pd.DataFrame) -> tuple[list[str], dict[str, dict]]:
    default_label = "— Pilih dari List Materi Promosi Gudang (Available) —"
    if promo_df is None or promo_df.empty:
        return [default_label], {}

    fp = dataframe_fingerprint(promo_df)
    cache = st.session_state.get("_warehouse_stock_cache")
    if isinstance(cache, tuple) and len(cache) == 3 and cache[0] == fp:
        return cache[1], cache[2]

    labels: list[str] = [default_label]
    records: dict[str, dict] = {}

    mat_list = promo_df["Material"].fillna("").astype(str).tolist() if "Material" in promo_df else [""] * len(promo_df)
    brd_list = promo_df["Brand"].fillna("").astype(str).tolist() if "Brand" in promo_df else [""] * len(promo_df)
    ver_list = promo_df["Version"].fillna("").astype(str).tolist() if "Version" in promo_df else [""] * len(promo_df)
    avail_list = (
        pd.to_numeric(promo_df["Available"], errors="coerce").fillna(0).astype(int).tolist()
        if "Available" in promo_df
        else [0] * len(promo_df)
    )
    uom_list = promo_df["UOM"].fillna("").astype(str).tolist() if "UOM" in promo_df else [""] * len(promo_df)

    for idx in range(len(promo_df)):
        m_s = mat_list[idx].strip()
        b_s = brd_list[idx].strip()
        v_s = ver_list[idx].strip()
        a_i = int(avail_list[idx])
        u_s = uom_list[idx].strip()
        lbl = f"{idx + 1}. {m_s} | {b_s} | {v_s} (Available: {a_i} {u_s})".strip()
        labels.append(lbl)
        records[lbl] = {
            "Material": m_s,
            "Brand": b_s,
            "Version": v_s,
            "Available": a_i,
            "UOM": u_s,
        }

    st.session_state["_warehouse_stock_cache"] = (fp, labels, records)
    return labels, records


def get_master_lookup_maps(
    master_df: pd.DataFrame,
) -> tuple[dict[str, list[int]], dict[str, list[int]], dict[int, str]]:
    if master_df is None or master_df.empty:
        return {}, {}, {}

    fp = dataframe_fingerprint(master_df)
    cached = st.session_state.get("_master_lookup_cache")
    if isinstance(cached, tuple) and len(cached) == 4 and cached[0] == fp:
        return cached[1], cached[2], cached[3]

    id_map: dict[str, list[int]] = {}
    name_map: dict[str, list[int]] = {}
    addr_map: dict[int, str] = {}

    ids = master_df["ID Outlet"].tolist() if "ID Outlet" in master_df.columns else [pd.NA] * len(master_df)
    names = master_df["Nama Customer"].tolist() if "Nama Customer" in master_df.columns else [pd.NA] * len(master_df)
    addrs = master_df["Alamat"].tolist() if "Alamat" in master_df.columns else [pd.NA] * len(master_df)

    for pos, r_idx in enumerate(master_df.index):
        idx_int = int(r_idx)
        k_id = customer_id_key(ids[pos])
        k_name = normalize_name_key(names[pos])
        addr_map[idx_int] = safe_text(addrs[pos])
        if k_id:
            id_map.setdefault(k_id, []).append(idx_int)
        if k_name:
            name_map.setdefault(k_name, []).append(idx_int)

    st.session_state["_master_lookup_cache"] = (fp, id_map, name_map, addr_map)
    return id_map, name_map, addr_map


def lookup_master_indices(
    master_df: pd.DataFrame,
    id_outlet_val,
    nama_outlet_val,
) -> list[int]:
    if master_df is None or master_df.empty:
        return []
    id_map, name_map, _ = get_master_lookup_maps(master_df)

    k_id = customer_id_key(id_outlet_val)
    k_name = normalize_name_key(nama_outlet_val)

    if k_id:
        if k_id in id_map:
            return id_map[k_id]
        if k_name and k_name in name_map:
            return name_map[k_name]
        return []

    if k_name and k_name in name_map:
        return name_map[k_name]

    return []


def lookup_master_address_fast(
    master_df: pd.DataFrame,
    id_outlet_val,
    nama_outlet_val,
) -> tuple[bool, str]:
    matched = lookup_master_indices(master_df, id_outlet_val, nama_outlet_val)
    if not matched:
        return False, ""
    _, _, addr_map = get_master_lookup_maps(master_df)
    return True, addr_map.get(matched[0], "")


def apply_sticker_replacement_orders(
    contracts: pd.DataFrame,
    orders_df: pd.DataFrame,
    strict_master_lookup: bool = True,
    already_normalized: bool = False,
) -> tuple[pd.DataFrame, int, list[dict], list[str]]:
    target = contracts.copy() if already_normalized else normalize_data(contracts).copy()
    updated_rows_set: set[int] = set()
    matched_records: list[dict] = []
    unmatched_labels: list[str] = []

    id_map, name_map, _ = get_master_lookup_maps(target)

    for _, order in orders_df.iterrows():
        raw_id = order.get("ID Outlet", pd.NA)
        raw_name = order.get("Nama Outlet", pd.NA)
        raw_alamat = order.get("Alamat", pd.NA)
        raw_material = order.get("Material", order.get("Ukuran Sticker", pd.NA))
        raw_brand = order.get("Brand", order.get("Brand Sticker", pd.NA))
        raw_version = order.get("Version", order.get("Versi", pd.NA))
        raw_qty = order.get("Jumlah Sticker / Vinyl", 1)
        raw_vendor = order.get("Nama Vendor", pd.NA)
        raw_date = order.get(
            "Tanggal OrderPenggantian Sticker / Vinyl",
            order.get("Tanggal Penggantian Sticker Panel", pd.NaT),
        )

        id_str = safe_text(raw_id)
        name_str = safe_text(raw_name)
        if not id_str and not name_str:
            continue

        k_id = customer_id_key(id_str)
        k_name = normalize_name_key(name_str)
        if k_id:
            matched_indices = id_map.get(k_id) or (name_map.get(k_name, []) if k_name else [])
        elif k_name:
            matched_indices = name_map.get(k_name, [])
        else:
            matched_indices = []

        label = f"ID: {id_str or '-'} | Nama: {name_str or '-'}"

        if not matched_indices:
            unmatched_labels.append(label)
            if strict_master_lookup:
                continue

        size_val = extract_sizes(raw_material)
        mat_txt = safe_text(raw_material)
        alamat_val = safe_text(raw_alamat) or pd.NA
        brand_val = safe_text(raw_brand) or pd.NA
        version_val = safe_text(raw_version) or pd.NA
        vendor_val = safe_text(raw_vendor) or pd.NA
        qty_num = int(pd.to_numeric(raw_qty, errors="coerce") or 0)
        parsed_dt = pd.to_datetime(raw_date, errors="coerce", dayfirst=True)
        date_val = parsed_dt.tz_localize(None) if not pd.isna(parsed_dt) else pd.Timestamp(date.today()).normalize()
        now_stamp = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")

        for r_idx in matched_indices:
            if pd.notna(alamat_val) and str(alamat_val).strip() != "":
                target.at[r_idx, "Alamat"] = alamat_val
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

            m_row = target.loc[r_idx]
            matched_records.append({
                "Waktu Input": now_stamp,
                "ID Outlet": safe_text(m_row.get("ID Outlet")) or id_str or "-",
                "Nama Outlet": safe_text(m_row.get("Nama Customer")) or name_str or "-",
                "Alamat": safe_text(alamat_val) or safe_text(m_row.get("Alamat")) or "-",
                "Zona": safe_text(m_row.get("Zona")) or "-",
                "Material": mat_txt or "-",
                "Ukuran": safe_text(size_val) or "-",
                "Brand": safe_text(brand_val) or "-",
                "Version": safe_text(version_val) or "-",
                "Jumlah Sticker / Vinyl": int(qty_num),
                "Tanggal OrderPenggantian Sticker / Vinyl": date_val.strftime("%d/%m/%Y") if not pd.isna(date_val) else "-",
                "Nama Vendor": safe_text(vendor_val) or "-",
                "Status Kirim Pejabat": "Siap Kirim",
            })

    if strict_master_lookup and unmatched_labels:
        raise ValueError(
            "Data tidak ditemukan pada Master Data Dashboard sehingga pemesanan tidak dapat disimpan:\n• "
            + "\n• ".join(unmatched_labels)
        )

    return attach_search_blob(target.reset_index(drop=True)), len(updated_rows_set), matched_records, unmatched_labels


def read_sticker_order_excel(source) -> pd.DataFrame:
    if hasattr(source, "seek"):
        source.seek(0)
    raw = pd.read_excel(source, sheet_name=0, header=0, engine="openpyxl")
    raw = clean_headers(raw)
    if raw.empty:
        raise ValueError("File Excel kosong.")

    lowered = {c.casefold(): c for c in raw.columns}

    def pick(*aliases):
        for a in aliases:
            if a.casefold() in lowered:
                return raw[lowered[a.casefold()]]
        return pd.Series(pd.NA, index=raw.index)

    df = pd.DataFrame({
        "ID Outlet": pick("ID Outlet", "ID Customer", "Kode Outlet", "Outlet ID"),
        "Nama Outlet": pick("Nama Outlet", "Nama Customer", "Customer Name", "Nama Toko", "Outlet"),
        "Alamat": pick("Alamat", "Alamat Outlet", "Alamat Toko", "Address"),
        "Material": pick(
            "Material",
            "Ukuran Sticker",
            "Ukuran Sticker / Vinyl",
            "Ukuran Sticker Panel Toko",
            "Ukuran Sticker Panel",
            "Ukuran Stiker",
        ),
        "Brand": pick("Brand", "Brand Sticker", "Brand Panel", "Merek Sticker", "Merek Panel"),
        "Version": pick("Version", "Versi", "Versi Sticker", "Versi Panel"),
        "Jumlah Sticker / Vinyl": pick("Jumlah Sticker / Vinyl", "Jumlah", "Qty", "Jumlah Sticker"),
        "Tanggal OrderPenggantian Sticker / Vinyl": pick(
            "Tanggal OrderPenggantian Sticker / Vinyl",
            "Tanggal Order Penggantian Sticker / Vinyl",
            "Tanggal Penggantian Sticker Panel",
            "Tanggal Penggantian",
            "Tanggal Penggantian Terakhir Sticker Panel",
            "Tgl Penggantian",
        ),
        "Nama Vendor": pick("Nama Vendor", "Vendor", "Nama Vendor Sticker"),
    })

    for col in ("ID Outlet", "Nama Outlet", "Alamat", "Material", "Brand", "Version", "Nama Vendor"):
        df[col] = df[col].astype("string").str.strip()
        df[col] = df[col].replace({"": pd.NA, "<NA>": pd.NA, "nan": pd.NA, "None": pd.NA})

    df["Jumlah Sticker / Vinyl"] = pd.to_numeric(df["Jumlah Sticker / Vinyl"], errors="coerce").fillna(1)
    df["Tanggal OrderPenggantian Sticker / Vinyl"] = pd.to_datetime(
        df["Tanggal OrderPenggantian Sticker / Vinyl"], errors="coerce", dayfirst=True
    ).dt.tz_localize(None)

    has_any = df[["ID Outlet", "Nama Outlet", "Alamat", "Material", "Brand", "Version", "Nama Vendor"]].notna().any(axis=1)
    df = df.loc[has_any.to_numpy(dtype=bool)].reset_index(drop=True)

    missing_both_mask = df["ID Outlet"].isna() & df["Nama Outlet"].isna()
    if missing_both_mask.any():
        bad_rows = [str(i + 2) for i in df.index[missing_both_mask]]
        raise ValueError(
            f"Isi minimal 'ID Outlet' atau 'Nama Outlet' untuk pencocokan master data pada baris Excel: {', '.join(bad_rows)}."
        )

    return df


def _wrap_line_pure(text: str, width: int) -> list[str]:
    clean_text = "" if text is None else str(text)
    if clean_text == "":
        return [""]
    limit = max(int(width), 1)
    words = clean_text.split(" ")
    lines: list[str] = []
    current = ""

    for word in words:
        if len(word) > limit:
            if current:
                lines.append(current)
                current = ""
            for start in range(0, len(word), limit):
                chunk = word[start : start + limit]
                if len(chunk) == limit:
                    lines.append(chunk)
                else:
                    current = chunk
            continue

        if not current:
            current = word
        elif len(current) + 1 + len(word) <= limit:
            current = f"{current} {word}"
        else:
            lines.append(current)
            current = word

    if current or not lines:
        lines.append(current)
    return lines


@st.cache_data(show_spinner=False)
def _generate_sticker_order_png_cached(payload_json: str, tanggal_order: str, nama_vendor: str) -> bytes:
    disp = pd.DataFrame(json.loads(payload_json)) if payload_json else pd.DataFrame()
    if disp.empty:
        disp = pd.DataFrame(
            columns=[
                "Nama Outlet",
                "Ukuran",
                "Brand",
                "Version",
                "Jumlah Sticker / Vinyl",
                "Tanggal OrderPenggantian Sticker / Vinyl",
                "Nama Vendor",
            ]
        )

    col_names = [str(c) for c in disp.columns]
    col_max_chars: list[int] = []

    for col in disp.columns:
        header_w = len(str(col))
        raw_col_strings = ["" if pd.isna(v) else str(v) for v in disp[col].tolist()]
        line_lengths = [len(line) for s in raw_col_strings for line in s.splitlines()]
        token_lengths = [len(tok) for s in raw_col_strings for tok in s.split()]
        max_line_w = max(line_lengths) if line_lengths else 0
        longest_token = max(token_lengths) if token_lengths else 4
        natural_w = max(header_w, max_line_w, 8)
        min_needed = min(longest_token + 2, 16)
        col_max_chars.append(int(min(max(natural_w, min_needed, 8), 42)))

    total_chars = max(int(sum(col_max_chars)), 1)
    col_widths = [float(w) / float(total_chars) for w in col_max_chars]

    wrapped_rows: list[list[str]] = []
    row_line_counts: list[int] = [1]

    for _, row_series in disp.iterrows():
        wrapped_cell_lines: list[list[str]] = []
        max_lines_in_row = 1
        row_list = row_series.tolist()

        for c_idx in range(len(col_names)):
            val = row_list[c_idx] if c_idx < len(row_list) else ""
            txt = "" if pd.isna(val) else str(val)
            wrap_width = max(int(col_max_chars[c_idx]), 8)
            cell_lines: list[str] = []
            for para in (txt.splitlines() if txt != "" else [""]):
                cell_lines.extend(_wrap_line_pure(str(para), wrap_width))
            wrapped_cell_lines.append(cell_lines)
            if len(cell_lines) > max_lines_in_row:
                max_lines_in_row = len(cell_lines)

        row_line_counts.append(int(max_lines_in_row))
        wrapped_rows.append(["\n".join(wrapped_cell_lines[c_idx]) for c_idx in range(len(col_names))])

    total_lines = int(sum(row_line_counts))
    fig_width = float(max(14.0, min(28.0, total_chars * 0.135)))
    fig_height = float(max(3.6, min(36.0, total_lines * 0.34 + 1.8)))

    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.axis("off")
    ax.add_patch(patches.Rectangle((0.0, 0.88), 1.0, 0.10, facecolor="#0F172A", edgecolor="#000000", lw=1))
    ax.text(
        0.015,
        0.93,
        f"HASIL PEMESANAN PENGGANTIAN STICKER / VINYL PANEL TOKO   |   Tanggal: {tanggal_order}   |   Vendor: {nama_vendor or 'Semua'}",
        ha="left",
        va="center",
        fontsize=10.5,
        weight="bold",
        color="#FFFFFF",
    )

    cell_text = wrapped_rows if wrapped_rows else [["-"] * len(col_names)]
    table = ax.table(cellText=cell_text, colLabels=col_names, loc="center", bbox=[0.0, 0.02, 1.0, 0.83])
    table.auto_set_font_size(False)
    table.set_fontsize(8.5)

    for (row, col), cell in table.get_celld().items():
        cell.set_width(float(col_widths[col]))
        cell.set_edgecolor("#000000")
        cell.set_linewidth(0.8)
        if row == 0:
            cell.set_facecolor("#1E3A8A")
            cell.set_text_props(color="#FFFFFF", weight="bold", ha="center", va="center")
        else:
            cell.set_facecolor("#FFFFFF" if row % 2 == 1 else "#F8FAFC")
            is_left_col = col in (0, 1, 2, 3)
            cell.set_text_props(
                color="#0F172A",
                weight="bold" if col == 0 else "normal",
                ha="left" if is_left_col else "center",
                va="center",
            )

    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight", pad_inches=0.18, dpi=180)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def generate_sticker_order_png(df_table: pd.DataFrame, tanggal_order: str, nama_vendor: str) -> bytes:
    disp = df_table.drop(
        columns=["ID Outlet", "Material", "Waktu Input", "Alamat", "Zona", "Status Kirim Pejabat", "_search_blob"],
        errors="ignore",
    ).head(100)
    payload_json = disp.to_json(orient="records", date_format="iso", force_ascii=False)
    return _generate_sticker_order_png_cached(payload_json, tanggal_order, nama_vendor)
