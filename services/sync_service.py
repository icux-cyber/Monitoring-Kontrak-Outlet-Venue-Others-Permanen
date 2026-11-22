from __future__ import annotations

import json
from io import BytesIO

import pandas as pd
import requests
import streamlit as st

from services.contracts_service import COLUMNS, DATE_COLUMNS, MATERIAL_COLUMNS, SHEET_NAME, normalize_data, parse_tickmark
from services.sticker_service import STICKER_LOG_COLUMNS, normalize_sticker_order_records


@st.cache_data(show_spinner=False)
def _excel_bytes_cached(records_json: str, sheet_name: str, columns_tuple: tuple[str, ...]) -> bytes:
    df = pd.DataFrame(json.loads(records_json)) if records_json else pd.DataFrame(columns=list(columns_tuple))
    output = BytesIO()
    export = df.drop(columns=["_search_blob", "Status Kontrak", "Sisa Hari"], errors="ignore").copy()
    for column in MATERIAL_COLUMNS:
        if column in export:
            export[column] = export[column].map(lambda checked: "✓" if parse_tickmark(checked) else "")
    for column in (*DATE_COLUMNS, "Tanggal OrderPenggantian Sticker / Vinyl"):
        if column in export:
            export[column] = pd.to_datetime(export[column], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()


def excel_bytes(frame: pd.DataFrame, sheet_name: str = SHEET_NAME) -> bytes:
    clean_df = frame.drop(columns=["_search_blob"], errors="ignore")
    cols_tuple = tuple(str(c) for c in clean_df.columns)
    records_json = clean_df.to_json(orient="records", date_format="iso", force_ascii=False)
    return _excel_bytes_cached(records_json, sheet_name, cols_tuple)


def get_secret_str(*keys: str, default: str = "") -> str:
    for k in keys:
        try:
            val = st.secrets.get(k, "")
        except FileNotFoundError:
            val = ""
        if val is not None and str(val).strip() != "":
            return str(val).strip()
    return default


def backup_to_google_sheet(
    frame: pd.DataFrame,
    url: str,
    sticker_order_log: list[dict] | None = None,
    already_normalized: bool = False,
) -> int:
    clean_url = str(url or "").strip()
    if not clean_url.startswith("https://script.google.com/"):
        raise ValueError("URL backup harus menggunakan alamat HTTPS Google Apps Script.")
    normalized = frame.drop(columns=["_search_blob", "Status Kontrak", "Sisa Hari"], errors="ignore") if already_normalized else normalize_data(frame)
    records = json.loads(normalized.to_json(orient="records", date_format="iso", force_ascii=False))
    order_log_records = normalize_sticker_order_records(sticker_order_log or [])
    payload = {
        "action": "backup",
        "sheet_name": SHEET_NAME,
        "columns": COLUMNS,
        "records": records,
        "sticker_order_log": order_log_records,
        "sticker_order_log_columns": STICKER_LOG_COLUMNS,
        "backup_at": pd.Timestamp.now(tz="UTC").isoformat(),
    }
    response = requests.post(clean_url, json=payload, timeout=60)
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


def get_supabase_config() -> tuple[str, str, str]:
    url = get_secret_str("SUPABASE_URL", "SUPADATA_URL", default="").rstrip("/")
    key = get_secret_str(
        "SUPABASE_SERVICE_ROLE_KEY",
        "SUPABASE_ANON_KEY",
        "SUPABASE_KEY",
        "SUPADATA_KEY",
        "SUPADATA_SERVICE_ROLE_KEY",
        default="",
    )
    table = get_secret_str(
        "SUPABASE_TABLE",
        "SUPADATA_TABLE",
        default="monitoring_kontrak_outlet",
    )
    return url, key, table


def backup_to_supabase(
    frame: pd.DataFrame,
    url: str,
    api_key: str,
    table: str,
    sticker_order_log: list[dict] | None = None,
    already_normalized: bool = False,
) -> int:
    clean_url = str(url or "").strip().rstrip("/")
    clean_table = str(table or "").strip()
    if not clean_url.startswith("https://") or not clean_table:
        raise ValueError("SUPABASE_URL dan nama tabel Supabase harus valid.")
    normalized = frame.drop(columns=["_search_blob", "Status Kontrak", "Sisa Hari"], errors="ignore") if already_normalized else normalize_data(frame)
    records = json.loads(normalized.to_json(orient="records", date_format="iso", force_ascii=False))
    order_log_records = normalize_sticker_order_records(sticker_order_log or [])
    headers = _supabase_headers(api_key)
    endpoint = f"{clean_url}/rest/v1/{clean_table}"

    del_resp = requests.delete(f"{endpoint}?payload_json=not.is.null", headers=headers, timeout=60)
    if del_resp.status_code >= 400:
        raise ValueError(f"Gagal membersihkan data lama di Supabase ({del_resp.status_code}): {del_resp.text}")

    envelope = {
        "records": records,
        "sticker_order_log": order_log_records,
        "backup_at": pd.Timestamp.now(tz="UTC").isoformat(),
    }
    payload = [{"payload_json": envelope}]
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


def restore_from_supabase(url: str, api_key: str, table: str) -> tuple[pd.DataFrame, list[dict] | None]:
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
        return pd.DataFrame(columns=COLUMNS), None
    raw_payload = rows[0].get("payload_json")
    if isinstance(raw_payload, str):
        raw_payload = json.loads(raw_payload)

    if isinstance(raw_payload, dict):
        raw_records = raw_payload.get("records") or []
        restored_log = normalize_sticker_order_records(raw_payload.get("sticker_order_log") or [])
    elif isinstance(raw_payload, list):
        raw_records = raw_payload
        restored_log = None
    else:
        raw_records = []
        restored_log = None

    if not isinstance(raw_records, list) or not raw_records:
        return pd.DataFrame(columns=COLUMNS), restored_log
    return normalize_data(pd.DataFrame(raw_records)), restored_log


def sync_all_cloud_targets(
    contracts_df: pd.DataFrame,
    sticker_order_log: list[dict],
    already_normalized: bool = False,
) -> tuple[list[str], list[str]]:
    successes: list[str] = []
    warnings_list: list[str] = []

    gs_url = get_secret_str("GOOGLE_APPS_SCRIPT_URL", default="")
    if gs_url:
        try:
            row_cnt = backup_to_google_sheet(
                contracts_df,
                gs_url,
                sticker_order_log=sticker_order_log,
                already_normalized=already_normalized,
            )
            successes.append(f"Google Sheet ({row_cnt} kontrak)")
        except Exception as err:
            warnings_list.append(f"Google Sheet: {err}")
    else:
        warnings_list.append("Google Sheet belum dikonfigurasi (`GOOGLE_APPS_SCRIPT_URL`).")

    sb_url, sb_key, sb_table = get_supabase_config()
    if sb_url and sb_key:
        try:
            sb_cnt = backup_to_supabase(
                contracts_df,
                sb_url,
                sb_key,
                sb_table,
                sticker_order_log=sticker_order_log,
                already_normalized=already_normalized,
            )
            successes.append(f"Supabase ({sb_cnt} kontrak)")
        except Exception as err:
            warnings_list.append(f"Supabase: {err}")
    else:
        warnings_list.append("Supabase belum dikonfigurasi (`SUPABASE_URL` & API key).")

    return successes, warnings_list
