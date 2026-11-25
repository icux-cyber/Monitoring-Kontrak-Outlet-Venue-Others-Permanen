def restore_from_google_sheet(url: str) -> tuple[pd.DataFrame, list[dict] | None]:
    clean_url = str(url or "").strip()
    if not clean_url.startswith("https://script.google.com/"):
        raise ValueError("URL restore harus menggunakan alamat HTTPS Google Apps Script.")

    payload = {
        "action": "restore",
        "sheet_name": SHEET_NAME,
    }
    response = requests.post(clean_url, json=payload, timeout=60)
    response.raise_for_status()
    try:
        result = response.json()
    except ValueError as error:
        raise ValueError("Respons Google Apps Script bukan JSON yang valid.") from error

    if not isinstance(result, dict) or result.get("ok") is not True:
        message = (
            result.get("error", "Restore ditolak oleh Google Apps Script.")
            if isinstance(result, dict)
            else "Respons Google Apps Script tidak valid."
        )
        raise ValueError(message)

    raw_records = result.get("records") or []
    if not isinstance(raw_records, list) or not raw_records:
        raise ValueError("Data kontrak di Google Sheet kosong atau tidak ditemukan.")

    raw_log = result.get("sticker_order_log")
    restored_log = normalize_sticker_order_records(raw_log) if isinstance(raw_log, list) else None

    return normalize_data(pd.DataFrame(raw_records)), restored_log
