from __future__ import annotations

import re
import urllib.parse
from datetime import date

import pandas as pd
import streamlit as st

from services.contracts_service import (
    COLUMNS,
    DATE_COLUMNS,
    MATERIAL_COLUMNS,
    SHEET_NAME,
    apply_branding_updates,
    attach_search_blob,
    budget_group,
    compute_contract_status,
    customer_id_key,
    extract_sizes,
    format_outlet_table,
    load_initial_data,
    merge_panel_screen_by_id,
    normalize_data,
    parse_amount,
    parse_tickmark,
    read_branding_workbook,
    read_drp_sheet,
    read_workbook,
    safe_date,
    safe_text,
    save_data,
    validate_data,
)
from services.drp_service import (
    append_missing_drp_to_contracts,
    apply_drp_updates_to_contracts,
    build_dashboard_only_report,
    build_drp_crosscheck_report,
    fmt_val_cmp,
    read_drp_crosscheck_workbook,
)
from services.sticker_service import (
    DEFAULT_PEJABAT_EMAILS,
    PROMO_STOCK_FILE,
    STICKER_LOG_COLUMNS,
    STICKER_ORDER_COLUMNS,
    append_sticker_order_log,
    apply_sticker_replacement_orders,
    clear_promo_stock_cache,
    generate_sticker_order_png,
    get_warehouse_stock_options,
    load_initial_promo_stock,
    load_pejabat_emails,
    load_sticker_order_log,
    lookup_master_address_fast,
    lookup_master_indices,
    normalize_sticker_order_records,
    read_promo_stock_workbook,
    read_sticker_order_excel,
    save_pejabat_emails,
    save_sticker_order_log,
    update_log_status_for_rows,
)
from services.sync_service import (
    backup_to_google_sheet,
    backup_to_supabase,
    excel_bytes,
    get_secret_str,
    get_supabase_config,
    restore_from_google_sheet,
    restore_from_supabase,
    sync_all_cloud_targets,
)

APP_TITLE = "Monitoring Kontrak Outlet"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


def render_sticker_order_png_with_material(
    df: pd.DataFrame,
    tgl_ringkas: str,
    vendor_ringkas: str,
) -> bytes:
    """Generate gambar PNG tabel hasil pemesanan (tanpa ID Outlet) dengan kolom Material."""
    import io
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    if df is None or df.empty:
        return generate_sticker_order_png(df, tgl_ringkas, vendor_ringkas)

    df_sorted = df.copy()
    df_sorted["_sort_mat"] = (
        df_sorted.apply(
            lambda r: (safe_text(r.get("Material")) or safe_text(r.get("Ukuran")) or "").casefold(),
            axis=1,
        )
    )
    df_sorted["_sort_brand"] = df_sorted["Brand"].fillna("").astype(str).map(lambda x: x.strip().casefold())
    df_sorted = df_sorted.sort_values(by=["_sort_mat", "_sort_brand"], kind="mergesort").reset_index(drop=True)

    headers = [
        "No",
        "Nama Outlet",
        "Alamat",
        "Material",
        "Brand",
        "Version",
        "Qty",
        "Tgl Order",
        "Vendor",
    ]
    rows: list[list[str]] = []
    for idx, (_, r) in enumerate(df_sorted.iterrows(), start=1):
        mat_val = safe_text(r.get("Material")) or safe_text(r.get("Ukuran")) or "-"
        qty_raw = r.get("Jumlah Sticker / Vinyl", 0)
        try:
            qty_str = str(int(float(qty_raw)))
        except Exception:
            qty_str = safe_text(qty_raw) or "0"
        rows.append([
            str(idx),
            safe_text(r.get("Nama Outlet")) or "-",
            safe_text(r.get("Alamat")) or "-",
            mat_val,
            safe_text(r.get("Brand")) or "-",
            safe_text(r.get("Version")) or "-",
            qty_str,
            safe_text(r.get("Tanggal OrderPenggantian Sticker / Vinyl")) or "-",
            safe_text(r.get("Nama Vendor")) or "-",
        ])

    col_lens = [
        max(len(headers[c]), *(len(row[c]) for row in rows))
        for c in range(len(headers))
    ]
    min_widths = [4, 18, 22, 20, 14, 12, 5, 11, 12]
    max_widths = [5, 34, 42, 34, 24, 22, 6, 13, 20]
    adj_lens = [
        max(min_widths[i], min(max_widths[i], col_lens[i] + 3))
        for i in range(len(headers))
    ]
    total_units = sum(adj_lens)
    col_widths = [w / total_units for w in adj_lens]

    fig_w = max(13.5, min(24.0, total_units * 0.135))
    row_h = 0.44
    fig_h = max(2.8, 1.35 + (len(rows) + 1) * row_h)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=180)
    ax.axis("off")

    title_text = f"Hasil Pemesanan Penggantian Sticker / Vinyl Panel ({len(rows)} Item)"
    subtitle_text = f"Tanggal Order: {tgl_ringkas}   |   Vendor: {vendor_ringkas}"
    fig.text(0.015, 1.0 - (0.34 / fig_h), title_text, fontsize=13, fontweight="bold", color="#0f172a", va="top")
    fig.text(0.015, 1.0 - (0.72 / fig_h), subtitle_text, fontsize=9.5, color="#475569", va="top")

    table_top = 1.0 - (1.08 / fig_h)
    table_height = max(0.35, table_top - 0.03)
    table = ax.table(
        cellText=rows,
        colLabels=headers,
        colWidths=col_widths,
        loc="upper center",
        bbox=[0.0, 0.02, 1.0, table_height],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8.5)

    center_cols = {0, 6, 7}
    for (r_idx, c_idx), cell in table.get_celld().items():
        cell.set_edgecolor("#cbd5e1")
        cell.set_linewidth(0.6)
        if r_idx == 0:
            cell.set_facecolor("#1e3a8a")
            cell.get_text().set_color("#ffffff")
            cell.get_text().set_weight("bold")
            cell.get_text().set_ha("center")
        else:
            cell.set_facecolor("#f8fafc" if r_idx % 2 == 0 else "#ffffff")
            cell.get_text().set_color("#0f172a")
            cell.get_text().set_ha("center" if c_idx in center_cols else "left")

    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight", pad_inches=0.22, facecolor="#ffffff")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def _ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame(columns=COLUMNS)
    out = df.copy()
    for col in COLUMNS:
        if col not in out.columns:
            out[col] = pd.NA
    return out


def inject_custom_css() -> None:
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
            max-width: 96%;
        }
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

        section[data-testid="stSidebar"] {
            background-color: #f8fafc;
            border-right: 1px solid #e2e8f0;
        }
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

        @media (prefers-color-scheme: dark) {
            .kpi-card {
                background-color: #1e293b;
                border-color: #334155;
            }
            .kpi-value { color: #f8fafc; }
            .kpi-label, .kpi-sub { color: #94a3b8; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def reset_dashboard_filters() -> None:
    st.session_state["dash_search_query"] = ""
    st.session_state["dash_status"] = "Semua Status"
    st.session_state["dash_zone"] = "Semua"
    st.session_state["dash_sector"] = "Semua"
    st.session_state["dash_budget"] = "Semua"
    st.session_state["dash_job_cat"] = "Semua"
    st.session_state["dash_sticker_size"] = "Semua"
    st.session_state["dash_screen_size"] = "Semua"


def reset_multi_edit_state() -> None:
    st.session_state["multi_edit_quick_search"] = ""
    st.session_state["multi_edit_zone"] = "Semua Zona"
    st.session_state["multi_edit_status"] = "Semua Status"
    st.session_state["multi_select_all_state"] = False
    st.session_state.pop("confirm_delete_multi_cb", None)
    st.session_state.pop("multi_customer_data_editor", None)


def reset_sticker_single_form() -> None:
    for k in (
        "sticker_order_id",
        "sticker_order_name",
        "sticker_order_alamat",
        "sticker_order_material",
        "sticker_order_brand",
        "sticker_order_version",
        "sticker_order_qty",
        "sticker_order_vendor",
    ):
        st.session_state.pop(k, None)
    st.session_state["_last_autofill_alamat_key"] = ""
    st.session_state["_resolved_outlet_found"] = False
    st.session_state["_resolved_outlet_alamat"] = ""
    st.session_state["_pending_clear_sticker_outlet"] = False
    st.session_state["sticker_pick_from_warehouse"] = "— Pilih dari List Materi Promosi Gudang (Available) —"


def set_contracts_state(df: pd.DataFrame, already_normalized: bool = False) -> pd.DataFrame:
    ensured = _ensure_columns(df)
    normalized = ensured if already_normalized else normalize_data(ensured)
    attached = _ensure_columns(attach_search_blob(normalized))
    st.session_state.contracts = attached
    st.session_state.pop("_master_lookup_cache", None)
    st.session_state.pop("sticker_order_png_bytes", None)
    st.session_state["_last_autofill_alamat_key"] = ""
    return attached


def render_dashboard_table(title: str, frame: pd.DataFrame, download_filename: str, dl_key: str) -> None:
    col_head, col_dl = st.columns([3, 1])
    with col_head:
        st.markdown(f"#### {title} ({len(frame):,} outlet)".replace(",", "."))
    sorted_frame = frame.sort_values("Sisa Hari", na_position="last")
    with col_dl:
        st.download_button(
            "📥 Export Excel",
            data=excel_bytes(sorted_frame),
            file_name=download_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=dl_key,
            width="stretch",
        )

    if frame.empty:
        st.info(f"Tidak ada data customer pada kategori **{title}** yang sesuai dengan filter saat ini.")
        return

    formatted_df = format_outlet_table(sorted_frame).copy()

    panel_col = "Ukuran Panel Toko"
    sticker_col = next(
        (c for c in ("Ukuran Sticker Panel Toko", "Ukuran Sticker / Vinyl Panel") if c in formatted_df.columns),
        None,
    )
    if panel_col in formatted_df.columns and sticker_col and sticker_col != panel_col:
        cols = list(formatted_df.columns)
        cols.remove(panel_col)
        sticker_idx = cols.index(sticker_col)
        cols.insert(sticker_idx, panel_col)
        formatted_df = formatted_df[cols]

    st.dataframe(
        formatted_df,
        width="stretch",
        hide_index=True,
        height=440,
        column_config={
            "Status Kontrak": st.column_config.TextColumn("Status", width="medium"),
            "Sisa Hari": st.column_config.NumberColumn("Sisa Hari", format="%d hr", width="small"),
            "ID Outlet": st.column_config.TextColumn("ID Outlet", width="small"),
            "Nama Customer": st.column_config.TextColumn("Nama Customer", width="large"),
            "Nominal Kontrak": st.column_config.TextColumn("Nominal Kontrak", width="medium"),
            "Ukuran Panel Toko": st.column_config.TextColumn("Ukuran Panel Toko", width="medium"),
            "Ukuran Sticker Panel Toko": st.column_config.TextColumn("Ukuran Sticker / Vinyl Panel", width="medium"),
            "Brand Panel": st.column_config.TextColumn("Brand Sticker / Vinyl", width="medium"),
            "Versi": st.column_config.TextColumn("Versi", width="small"),
            "Tanggal Penggantian Sticker / Vinyl Panel": st.column_config.TextColumn("Tgl Ganti Sticker / Vinyl Terakhir", width="medium"),
            "Nama Vendor": st.column_config.TextColumn("Nama Vendor", width="medium"),
        },
    )


def render_outlet_form(current: pd.Series, row_index: int | None) -> tuple[bool, dict]:
    suffix = "new" if row_index is None else f"row_{row_index}"
    with st.form(f"outlet_form_{suffix}", border=True):
        st.markdown("##### 1. Identitas & Lokasi Outlet")
        c1, c2, c3 = st.columns([1, 2, 1])
        outlet_id = c1.text_input("ID Outlet *", value=safe_text(current.get("ID Outlet", pd.NA)), placeholder="Contoh: 501X1234", key=f"id_{suffix}")
        customer_name = c2.text_input("Nama Customer *", value=safe_text(current.get("Nama Customer", pd.NA)), placeholder="Nama Toko / Outlet", key=f"name_{suffix}")
        outlet_category = c3.text_input("Kategori Outlet", value=safe_text(current.get("Kategori Outlet", pd.NA)), placeholder="Warung / Minimarket", key=f"outlet_category_{suffix}")

        address = st.text_input("Alamat Lengkap", value=safe_text(current.get("Alamat", pd.NA)), placeholder="Jl. Raya...", key=f"address_{suffix}")

        z1, z2, z3, z4 = st.columns(4)
        zone = z1.text_input("Zona", value=safe_text(current.get("Zona", pd.NA)), key=f"zone_{suffix}")
        sector = z2.text_input("Sektor", value=safe_text(current.get("Sektor", pd.NA)), key=f"sector_{suffix}")
        budget_pos = z3.text_input("Pos Budget", value=safe_text(current.get("Pos Budget", pd.NA)), key=f"budget_{suffix}")
        job_category = z4.text_input("Kategori Pekerjaan", value=safe_text(current.get("Kategori Pekerjaan", pd.NA)), key=f"job_category_{suffix}")

        st.divider()
        st.markdown("##### 2. Detail Kontrak & Nilai")
        k1, k2, k3, k4, k5 = st.columns(5)
        brand = k1.text_input("Brand Utama", value=safe_text(current.get("Brand", pd.NA)), key=f"brand_{suffix}")
        parsed_amount = parse_amount(current.get("Nominal Kontrak", pd.NA))
        amount = k2.number_input(
            "Nominal Kontrak (Rp)",
            min_value=0.0,
            step=100000.0,
            value=0.0 if pd.isna(parsed_amount) else float(parsed_amount),
            key=f"amount_{suffix}",
        )
        contract_start = k3.date_input("Tanggal Mulai", value=safe_date(current.get("Periode Kontrak Awal", pd.NA)), key=f"start_{suffix}")
        contract_end = k4.date_input("Tanggal Berakhir", value=safe_date(current.get("Periode Kontrak Akhir", pd.NA)), key=f"end_{suffix}")
        doc_no = k5.text_input("No. Dokumen", value=safe_text(current.get("No. Dokumen", pd.NA)), key=f"doc_no_{suffix}")

        st.divider()
        st.markdown("##### 3. Materi Promosi & Spesifikasi Ukuran")
        material_cols = st.columns(4)
        material_values = {
            column: material_cols[index].checkbox(column, value=parse_tickmark(current.get(column, False)), key=f"material_{index}_{suffix}")
            for index, column in enumerate(MATERIAL_COLUMNS)
        }

        p1, p2, p3, p4, p5, p6 = st.columns(6)
        panel_size = p1.text_input("Ukuran Panel Toko", value=safe_text(current.get("Ukuran Panel Toko", pd.NA)), placeholder="1 x 3 m", key=f"panel_size_{suffix}")
        sticker_size = p2.text_input("Ukuran Sticker / Vinyl Panel", value=safe_text(current.get("Ukuran Sticker Panel Toko", pd.NA)), placeholder="1 x 2 m", key=f"sticker_size_{suffix}")
        panel_brand = p3.text_input("Brand Sticker / Vinyl", value=safe_text(current.get("Brand Panel", pd.NA)), key=f"panel_brand_{suffix}")
        sticker_version = p4.text_input("Versi (Opsional)", value=safe_text(current.get("Versi", pd.NA)), placeholder="V1 / V2", key=f"sticker_version_{suffix}")
        sticker_replace_date = p5.date_input(
            "Tgl Ganti Sticker / Vinyl Terakhir",
            value=safe_date(current.get("Tanggal Penggantian Sticker Panel", pd.NaT)),
            key=f"sticker_replace_date_{suffix}",
        )
        vendor_name = p6.text_input("Nama Vendor", value=safe_text(current.get("Nama Vendor", pd.NA)), placeholder="ARTISTIC", key=f"vendor_name_{suffix}")

        s1, s2 = st.columns(2)
        screen_size = s1.text_input("Ukuran Layar", value=safe_text(current.get("Ukuran Layar", pd.NA)), placeholder="2 x 3 m", key=f"screen_size_{suffix}")
        screen_brand = s2.text_input("Brand Layar", value=safe_text(current.get("Brand Layar", pd.NA)), key=f"screen_brand_{suffix}")

        st.markdown("")
        submitted = st.form_submit_button("💾 Simpan Data Customer", type="primary", width="stretch")

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
        "No. Dokumen": doc_no.strip() or pd.NA,
        **material_values,
        "Ukuran Panel Toko": panel_size.strip() or pd.NA,
        "Ukuran Sticker Panel Toko": extract_sizes(sticker_size) if sticker_size.strip() else pd.NA,
        "Brand Panel": panel_brand.strip() or pd.NA,
        "Versi": sticker_version.strip() or pd.NA,
        "Tanggal Penggantian Sticker Panel": pd.Timestamp(sticker_replace_date) if sticker_replace_date else pd.NaT,
        "Nama Vendor": vendor_name.strip() or pd.NA,
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
    st.session_state.contracts = _ensure_columns(initial)
    st.session_state.load_warning = warning
else:
    st.session_state.contracts = _ensure_columns(st.session_state.contracts)
    if "_search_blob" not in st.session_state.contracts.columns:
        st.session_state.contracts = _ensure_columns(
            attach_search_blob(normalize_data(st.session_state.contracts))
        )

if "promo_stock" not in st.session_state:
    st.session_state.promo_stock = load_initial_promo_stock()

if "pejabat_emails" not in st.session_state:
    st.session_state.pejabat_emails = load_pejabat_emails()

if "sticker_order_results" not in st.session_state:
    st.session_state.sticker_order_results = load_sticker_order_log()

if "multi_select_all_state" not in st.session_state:
    st.session_state.multi_select_all_state = False

contracts = _ensure_columns(compute_contract_status(st.session_state.contracts))
promo_stock_df = st.session_state.promo_stock
expired_count = int(contracts["Status Kontrak"].eq("🔴 Expired").sum())
critical_count = int(contracts["Status Kontrak"].eq("🟠 Kritis (≤ 30 hari)").sum())
warning_count = int(contracts["Status Kontrak"].eq("🟡 Warning (31–90 hari)").sum())
soon_count = critical_count + warning_count
total_value = float(pd.to_numeric(contracts["Nominal Kontrak"], errors="coerce").sum())

with st.sidebar:
    st.markdown("### 📊 DSO Surabaya Barat")
    st.caption("Sistem Monitoring Kontrak Outlet")
    st.divider()

    active_menu = st.radio(
        "Navigasi Utama",
        [
            "📊 Dashboard & Monitoring",
            "👥 Kelola Data Customer",
            "🏷️ Pemesanan Sticker / Vinyl Panel",
            "📦 List Materi Promosi",
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
    st.write(f"• **Stok Sticker/Vinyl Ready:** `{len(promo_stock_df)}`")
    st.write(f"• **Log Pemesanan Sticker / Vinyl:** `{len(st.session_state.sticker_order_results)}`")

    st.divider()
    st.download_button(
        "⬇️ Download Semua Data (.xlsx)",
        data=excel_bytes(st.session_state.contracts),
        file_name="Monitoring_Kontrak_Outlet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )

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

    if expired_count or critical_count:
        alert_col1, alert_col2 = st.columns(2)
        if expired_count:
            alert_col1.error(f"🚨 **{expired_count} kontrak** sudah melewati masa berlaku (Expired).")
        if critical_count:
            alert_col2.warning(f"⚠️ **{critical_count} kontrak** memasuki masa kritis (berakhir dalam ≤ 30 hari).")

    head_f, head_rst = st.columns([4, 1])
    with head_f:
        st.markdown("### 🔍 Pencarian & Filter Data")
    with head_rst:
        st.button("🔄 Reset Filter", width="stretch", key="btn_reset_dash_filter", on_click=reset_dashboard_filters)

    search_col, status_col = st.columns([2, 1])
    search_query = search_col.text_input(
        "Cari cepat (ID Outlet, Nama Customer, Alamat, No. Dokumen, atau Brand)",
        placeholder="Ketik kata kunci pencarian...",
        key="dash_search_query",
    )
    statuses = [
        "Semua Status",
        "🔴 Expired",
        "🟠 Kritis (≤ 30 hari)",
        "🟡 Warning (31–90 hari)",
        "🟢 Aktif (> 90 hari)",
        "⚪ Data belum lengkap",
    ]
    selected_status = status_col.selectbox("Status Masa Berlaku", statuses, key="dash_status")

    with st.expander("🎛️ Filter Lanjutan (Zona, Sektor, Pos Budget, Ukuran Panel/Layar)", expanded=False):
        f1, f2, f3, f4 = st.columns(4)
        zones = ["Semua"] + sorted(contracts["Zona"].dropna().astype(str).unique().tolist())
        sectors = ["Semua"] + sorted(contracts["Sektor"].dropna().astype(str).unique().tolist())
        budget_positions = ["Semua"] + sorted(contracts["Pos Budget"].dropna().astype(str).unique().tolist())
        job_categories = ["Semua"] + sorted(contracts["Kategori Pekerjaan"].dropna().astype(str).unique().tolist())

        selected_zone = f1.selectbox("Zona", zones, key="dash_zone")
        selected_sector = f2.selectbox("Sektor", sectors, key="dash_sector")
        selected_budget = f3.selectbox("Pos Budget", budget_positions, key="dash_budget")
        selected_job_category = f4.selectbox("Kategori Pekerjaan", job_categories, key="dash_job_cat")

        f5, f6 = st.columns(2)
        sticker_sizes = ["Semua"] + sorted(contracts["Ukuran Sticker Panel Toko"].dropna().astype(str).unique().tolist())
        screen_sizes = ["Semua"] + sorted(contracts["Ukuran Layar"].dropna().astype(str).unique().tolist())
        selected_sticker_size = f5.selectbox("Ukuran Sticker / Vinyl Panel", sticker_sizes, key="dash_sticker_size")
        selected_screen_size = f6.selectbox("Ukuran Layar", screen_sizes, key="dash_screen_size")

    view = contracts.copy()
    if search_query.strip():
        q = search_query.strip().casefold()
        blob = view["_search_blob"] if "_search_blob" in view.columns else view["Nama Customer"].fillna("").astype(str).str.casefold()
        view = view[blob.str.contains(q, regex=False)]

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
            render_dashboard_table("Semua Data Customer Terfilter", view, "Daftar_Semua_Customer_Terfilter.xlsx", "dl_all_filtered")
        with tab_retail:
            render_dashboard_table("Daftar Customer - Retail Supply", regular_view, "Daftar_Customer_Retail_Supply.xlsx", "dl_retail_filtered")
        with tab_venue:
            render_dashboard_table("Daftar Customer - Venue", venue_view, "Daftar_Customer_Venue.xlsx", "dl_venue_filtered")
        with tab_other:
            render_dashboard_table("Daftar Customer - Others Permanen", other_permanent_view, "Daftar_Customer_Others_Permanen.xlsx", "dl_other_filtered")


# ==========================================
# HALAMAN 2: KELOLA DATA CUSTOMER
# ==========================================
elif active_menu == "👥 Kelola Data Customer":
    tab_multi_edit, tab_single_edit, tab_add_manual, tab_add_batch = st.tabs([
        "⚡ Edit Multiple (Quick Search)",
        "✏️ Edit Satuan (Form)",
        "➕ Tambah Customer (Form)",
        "📑 Tambah Massal (Excel)",
    ])

    with tab_multi_edit:
        if contracts.empty:
            st.info("Belum ada data customer untuk diedit. Silakan tambah customer baru atau import Excel.")
        else:
            me_title_col, me_reset_col = st.columns([3.5, 1.5])
            with me_title_col:
                st.markdown("#### 🔍 Pencarian Instan & Edit Massal")
                st.caption("Cari beberapa outlet sekaligus, lalu edit langsung di tabel atau gunakan **Update Massal** untuk mengubah banyak baris dalam 1 klik.")
            with me_reset_col:
                st.button("🔄 Reset & Batalkan Edit", width="stretch", key="btn_reset_multi_edit", on_click=reset_multi_edit_state)

            qs_col1, qs_col2, qs_col3 = st.columns([2.2, 1, 1])
            multi_search = qs_col1.text_input(
                "Quick Search (ID Outlet, Nama Customer, Alamat, Brand, atau Sektor)",
                placeholder="Ketik kata kunci, misal: 'Warung', 'Surabaya', 'Super', atau kode ID...",
                key="multi_edit_quick_search",
            )
            multi_zones = ["Semua Zona"] + sorted(contracts["Zona"].dropna().astype(str).unique().tolist())
            multi_sel_zone = qs_col2.selectbox("Filter Zona", multi_zones, key="multi_edit_zone")

            multi_statuses = [
                "Semua Status",
                "🔴 Expired",
                "🟠 Kritis (≤ 30 hari)",
                "🟡 Warning (31–90 hari)",
                "🟢 Aktif (> 90 hari)",
                "⚪ Data belum lengkap",
            ]
            multi_sel_status = qs_col3.selectbox("Filter Status", multi_statuses, key="multi_edit_status")

            filtered_multi = contracts.copy()
            if multi_search.strip():
                q_multi = multi_search.strip().casefold()
                blob_m = (
                    filtered_multi["_search_blob"]
                    if "_search_blob" in filtered_multi.columns
                    else filtered_multi["Nama Customer"].fillna("").astype(str).str.casefold()
                )
                filtered_multi = filtered_multi[blob_m.str.contains(q_multi, regex=False)]

            if multi_sel_zone != "Semua Zona":
                filtered_multi = filtered_multi[filtered_multi["Zona"] == multi_sel_zone]
            if multi_sel_status != "Semua Status":
                filtered_multi = filtered_multi[filtered_multi["Status Kontrak"] == multi_sel_status]

            st.caption(f"Ditemukan **{len(filtered_multi):,}** baris customer sesuai pencarian.".replace(",", "."))

            if filtered_multi.empty:
                st.warning("Tidak ada customer yang cocok dengan kata kunci pencarian Anda.")
            else:
                sel_info_col, btn_all_col, btn_none_col = st.columns([3.5, 1, 1])
                sel_info_col.markdown("**Daftar Hasil Pencarian (Editable)**")
                if btn_all_col.button("☑️ Centang Semua", width="stretch", key="btn_check_all_multi"):
                    st.session_state.multi_select_all_state = True
                    st.session_state.pop("multi_customer_data_editor", None)
                    st.rerun()
                if btn_none_col.button("⬜ Batal Centang", width="stretch", key="btn_uncheck_all_multi"):
                    st.session_state.multi_select_all_state = False
                    st.session_state.pop("multi_customer_data_editor", None)
                    st.rerun()

                filtered_multi = _ensure_columns(filtered_multi)
                editor_df = filtered_multi[COLUMNS].copy()
                editor_df.insert(0, "Pilih", bool(st.session_state.multi_select_all_state))

                for col in COLUMNS:
                    if col not in MATERIAL_COLUMNS and col not in ("Nominal Kontrak", *DATE_COLUMNS):
                        editor_df[col] = editor_df[col].fillna("").astype(str)
                    elif col == "Nominal Kontrak":
                        editor_df[col] = pd.to_numeric(editor_df[col], errors="coerce").fillna(0.0)

                with st.expander("🪄 Aksi Cepat Massal (Ubah Kolom Secara Serentak untuk Banyak Outlet)", expanded=False):
                    ba1, ba2, ba3, ba4 = st.columns([1.3, 1.5, 1.2, 1])
                    bulk_target_col = ba1.selectbox(
                        "Kolom yang Diubah",
                        [
                            "Nama Customer", "Brand", "Nominal Kontrak", "Periode Kontrak Awal",
                            "Periode Kontrak Akhir", "Tanggal Penggantian Sticker Panel", "Nama Vendor",
                            "Versi", "No. Dokumen", "Zona", "Sektor", "Pos Budget", "Kategori Outlet",
                            "Kategori Pekerjaan", "Brand Panel", "Ukuran Panel Toko", "Ukuran Sticker Panel Toko",
                            "Brand Layar", "Ukuran Layar", "Panel Toko", "Showcase", "Layar Toko", "Shop Sign / Neon Box",
                        ],
                        key="bulk_target_col",
                    )

                    if bulk_target_col in DATE_COLUMNS:
                        bulk_new_val = ba2.date_input("Nilai Baru", value=date.today(), key="bulk_val_date")
                    elif bulk_target_col == "Nominal Kontrak":
                        bulk_new_val = ba2.number_input("Nilai Nominal Baru (Rp)", min_value=0.0, step=100000.0, key="bulk_val_num")
                    elif bulk_target_col in MATERIAL_COLUMNS:
                        bulk_new_val = ba2.selectbox(
                            "Status Centang Baru",
                            [True, False],
                            format_func=lambda x: "✅ Aktif / Terpasang" if x else "❌ Tidak Ada",
                            key="bulk_val_bool",
                        )
                    else:
                        bulk_new_val = ba2.text_input("Nilai Teks Baru", placeholder="Masukkan nilai baru...", key="bulk_val_text")

                    bulk_scope = ba3.selectbox(
                        "Target Baris",
                        ["Hanya yang dicentang (Pilih)", f"Semua hasil pencarian ({len(filtered_multi)} baris)"],
                        key="bulk_scope_choice",
                    )
                    ba4.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                    apply_bulk_clicked = ba4.button("Terapkan Massal", type="secondary", width="stretch", key="btn_apply_bulk")

                edited_df = st.data_editor(
                    editor_df,
                    width="stretch",
                    height=460,
                    hide_index=True,
                    key="multi_customer_data_editor",
                    column_config={
                        "Pilih": st.column_config.CheckboxColumn("Pilih", default=False, width="small"),
                        "ID Outlet": st.column_config.TextColumn("ID Outlet", width="small"),
                        "Nama Customer": st.column_config.TextColumn("Nama Customer", width="medium"),
                        "Nominal Kontrak": st.column_config.NumberColumn("Nominal (Rp)", min_value=0, step=50000, format="Rp %d"),
                        "Periode Kontrak Awal": st.column_config.DateColumn("Mulai Kontrak", format="DD/MM/YYYY"),
                        "Periode Kontrak Akhir": st.column_config.DateColumn("Akhir Kontrak", format="DD/MM/YYYY"),
                        "Tanggal Penggantian Sticker Panel": st.column_config.DateColumn("Tgl Ganti Sticker / Vinyl", format="DD/MM/YYYY"),
                        "Nama Vendor": st.column_config.TextColumn("Nama Vendor", width="medium"),
                        "Panel Toko": st.column_config.CheckboxColumn("Panel", width="small"),
                        "Showcase": st.column_config.CheckboxColumn("Showcase", width="small"),
                        "Layar Toko": st.column_config.CheckboxColumn("Layar", width="small"),
                        "Shop Sign / Neon Box": st.column_config.CheckboxColumn("Neon Box", width="small"),
                    },
                )

                if apply_bulk_clicked:
                    target_indices = (
                        edited_df.index[edited_df["Pilih"] == True].tolist()
                        if bulk_scope.startswith("Hanya")
                        else edited_df.index.tolist()
                    )
                    if not target_indices:
                        st.warning("⚠️ Klik **Centang Semua** atau centang minimal 1 baris pada kolom **Pilih** terlebih dahulu.")
                    else:
                        full_updated = _ensure_columns(st.session_state.contracts)[COLUMNS].copy()
                        for orig_idx in edited_df.index:
                            for col in COLUMNS:
                                if orig_idx in full_updated.index and col in edited_df.columns:
                                    full_updated.at[orig_idx, col] = edited_df.at[orig_idx, col]

                        val_to_set = pd.Timestamp(bulk_new_val) if bulk_target_col in DATE_COLUMNS else bulk_new_val
                        for orig_idx in target_indices:
                            if orig_idx in full_updated.index:
                                full_updated.at[orig_idx, bulk_target_col] = val_to_set

                        try:
                            saved_df = save_data(full_updated)
                            set_contracts_state(saved_df, already_normalized=True)
                            st.session_state.multi_select_all_state = False
                            st.session_state.pop("multi_customer_data_editor", None)
                            st.success(f"✅ Berhasil memperbarui kolom **{bulk_target_col}** untuk **{len(target_indices)}** customer sekaligus!")
                            st.rerun()
                        except Exception as error:
                            st.error(f"Gagal menyimpan pembaruan massal: {error}")

                selected_rows_count = int(edited_df["Pilih"].sum())
                act_col1, act_col2, act_col3 = st.columns([2, 1.3, 1.3])
                act_col1.info(f"💡 Terpilih: **{selected_rows_count}** dari **{len(edited_df)}** baris pada hasil pencarian ini.")

                if act_col2.button("💾 Simpan Perubahan Tabel", type="primary", width="stretch", key="btn_save_multi_table"):
                    full_updated = _ensure_columns(st.session_state.contracts)[COLUMNS].copy()
                    for orig_idx in edited_df.index:
                        for col in COLUMNS:
                            if orig_idx in full_updated.index and col in edited_df.columns:
                                full_updated.at[orig_idx, col] = edited_df.at[orig_idx, col]

                    try:
                        saved_df = save_data(full_updated)
                        set_contracts_state(saved_df, already_normalized=True)
                        st.session_state.multi_select_all_state = False
                        st.session_state.pop("multi_customer_data_editor", None)
                        st.success(f"✅ Perubahan pada {len(edited_df)} baris hasil pencarian berhasil disimpan!")
                        st.rerun()
                    except Exception as error:
                        st.error(f"Gagal menyimpan perubahan: {error}")

                confirm_delete_all = act_col3.checkbox(
                    "Konfirmasi hapus",
                    value=False,
                    disabled=(selected_rows_count == 0),
                    key="confirm_delete_multi_cb",
                )
                if st.button(
                    f"🗑️ Hapus Terpilih ({selected_rows_count} baris)",
                    type="secondary",
                    width="stretch",
                    disabled=(selected_rows_count == 0 or not confirm_delete_all),
                    key="btn_delete_multi_selected",
                ):
                    try:
                        indices_to_drop = edited_df.index[edited_df["Pilih"] == True].tolist()
                        remaining_df = st.session_state.contracts.drop(index=indices_to_drop).reset_index(drop=True)
                        saved_df = save_data(remaining_df, already_normalized=True)
                        set_contracts_state(saved_df, already_normalized=True)
                        st.session_state.multi_select_all_state = False
                        st.session_state.pop("confirm_delete_multi_cb", None)
                        st.session_state.pop("multi_customer_data_editor", None)
                        st.success(f"✅ {len(indices_to_drop)} customer terpilih berhasil dihapus.")
                        st.rerun()
                    except Exception as error:
                        st.error(f"Gagal menghapus customer terpilih: {error}")

    with tab_single_edit:
        if contracts.empty:
            st.info("Belum ada data customer. Silakan tambah customer baru atau import Excel.")
        else:
            record_options = [
                f"{i} | {safe_text(row.get('ID Outlet', pd.NA)) or 'Tanpa ID'} — {safe_text(row.get('Nama Customer', pd.NA)) or 'Tanpa Nama'} ({safe_text(row.get('Zona', pd.NA)) or '-'})"
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
                updated_data = _ensure_columns(contracts)[COLUMNS].copy()
                for column, value in updated_record.items():
                    updated_data.at[selected_index, column] = value
                try:
                    saved_df = save_data(updated_data)
                    set_contracts_state(saved_df, already_normalized=True)
                    st.success("✅ Data customer berhasil diperbarui!")
                    st.rerun()
                except Exception as error:
                    st.error(f"Gagal menyimpan data: {error}")

            with st.expander("🗑️ Zona Berbahaya: Hapus Customer Ini"):
                st.warning("Tindakan ini akan menghapus data customer secara permanen dari daftar aktif.")
                if st.button("Hapus Customer Terpilih", type="secondary", key="delete_customer"):
                    try:
                        updated_data = st.session_state.contracts.drop(index=selected_index).reset_index(drop=True)
                        saved_df = save_data(updated_data, already_normalized=True)
                        set_contracts_state(saved_df, already_normalized=True)
                        st.success("Customer berhasil dihapus.")
                        st.rerun()
                    except Exception as error:
                        st.error(f"Gagal menghapus customer: {error}")

    with tab_add_manual:
        blank_record = pd.Series({column: pd.NA for column in COLUMNS})
        submitted, new_record = render_outlet_form(blank_record, None)
        if submitted:
            base_contracts = _ensure_columns(contracts)[COLUMNS]
            updated_data = pd.concat([base_contracts, pd.DataFrame([new_record])], ignore_index=True)
            try:
                saved_df = save_data(updated_data)
                set_contracts_state(saved_df, already_normalized=True)
                st.success("✅ Customer baru berhasil ditambahkan!")
                st.rerun()
            except Exception as error:
                st.error(f"Gagal menyimpan customer baru: {error}")

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
                "No. Dokumen": "BRN-26/1/5SBB/999",
                "Panel Toko": True,
                "Showcase": False,
                "Layar Toko": False,
                "Shop Sign / Neon Box": False,
                "Ukuran Panel Toko": "1 x 3 m",
                "Ukuran Sticker Panel Toko": "1 x 2 m",
                "Brand Panel": "Djarum Super",
                "Versi": "V1",
                "Tanggal Penggantian Sticker / Vinyl Panel": pd.Timestamp(date.today()),
                "Nama Vendor": "ARTISTIC",
                "Ukuran Layar": pd.NA,
                "Brand Layar": pd.NA,
            }]
        )
        sample_row_ready = _ensure_columns(sample_row)[COLUMNS]
        dl1, dl2 = st.columns(2)
        dl1.download_button(
            "📄 Template Kosong (.xlsx)",
            data=excel_bytes(pd.DataFrame(columns=COLUMNS)),
            file_name="Template_Tambah_Customer_Baru.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        dl2.download_button(
            "📝 Template + Contoh Pengisian (.xlsx)",
            data=excel_bytes(sample_row_ready),
            file_name="Template_Tambah_Customer_Baru_Contoh.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
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
                        base_contracts = _ensure_columns(contracts)[COLUMNS]
                        new_rows_ensured = _ensure_columns(new_rows)
                        existing_keys = {
                            customer_id_key(val)
                            for val in base_contracts["ID Outlet"].dropna()
                            if customer_id_key(val)
                        }
                        skipped_ids: list[str] = []
                        rows_to_add: list[pd.Series] = []
                        seen_batch_keys: set[str] = set()

                        for _, r in new_rows_ensured.iterrows():
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
                            combined = pd.concat([base_contracts, pd.DataFrame(rows_to_add)], ignore_index=True)
                            saved_df = save_data(combined)
                            set_contracts_state(saved_df, already_normalized=True)
                            st.session_state.load_warning = None
                            st.success(f"✅ Berhasil menambahkan {len(rows_to_add)} customer baru!")
                            if skipped_ids:
                                st.info(f"Dilewati ({len(skipped_ids)} ID duplikat): {', '.join(skipped_ids[:15])}")
                            st.rerun()
                except Exception as error:
                    st.error(f"Gagal memproses file: {error}")


# ==========================================
# HALAMAN 3: PEMESANAN PENGGANTIAN STICKER / VINYL PANEL
# ==========================================
elif active_menu == "🏷️ Pemesanan Sticker / Vinyl Panel":
    st.markdown("### 🏷️ Pemesanan Penggantian Sticker / Vinyl Panel")
    st.caption(
        "**ID Outlet** bersifat **opsional** (boleh dikosongkan jika mengisi **Nama Outlet**). "
        "Setiap pemesanan akan di-*lookup* ke **Master Data Dashboard** — jika tidak ditemukan, pemesanan **tidak dapat disimpan**."
    )

    tab_sticker_single, tab_sticker_batch = st.tabs([
        "✏️ Form Pemesanan Satuan",
        "📤 Upload Penggantian Massal (Excel)",
    ])

    with tab_sticker_single:
        stock_labels, stock_records = get_warehouse_stock_options(promo_stock_df)

        def _on_warehouse_stock_change() -> None:
            sel_lbl = st.session_state.get("sticker_pick_from_warehouse", "")
            rec = stock_records.get(sel_lbl)
            if rec is not None:
                st.session_state["sticker_order_material"] = rec["Material"]
                st.session_state["sticker_order_brand"] = rec["Brand"]
                st.session_state["sticker_order_version"] = rec["Version"]
                st.session_state["sticker_order_qty"] = max(1, min(1, int(rec["Available"])))

        top_stock_col, top_rst_col = st.columns([4, 1])
        with top_stock_col:
            picked_stock_label = st.selectbox(
                "📦 Referensi List Materi Promosi di Gudang Promosi (Available > 0)",
                options=stock_labels,
                key="sticker_pick_from_warehouse",
                on_change=_on_warehouse_stock_change,
            )
        with top_rst_col:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            st.button("🔄 Reset Form", width="stretch", key="btn_reset_sticker_form", on_click=reset_sticker_single_form)

        picked_rec = stock_records.get(picked_stock_label)
        if picked_rec is not None:
            st.info(
                f"Stok tersedia di Gudang Promosi: **{picked_rec['Available']} {picked_rec['UOM']}** — "
                f"**{picked_rec['Material']}** ({picked_rec['Brand']} • {picked_rec['Version']})"
            )

        _submit_msg = st.session_state.pop("_sticker_submit_success_msg", None)
        if _submit_msg:
            st.success(_submit_msg)

        if st.session_state.pop("_pending_clear_sticker_outlet", False):
            st.session_state["sticker_order_id"] = ""
            st.session_state["sticker_order_name"] = ""
            st.session_state["sticker_order_alamat"] = ""
            st.session_state["_last_autofill_alamat_key"] = ""
            st.session_state["_resolved_outlet_found"] = False
            st.session_state["_resolved_outlet_alamat"] = ""

        for k_def, v_def in (
            ("sticker_order_id", ""),
            ("sticker_order_name", ""),
            ("sticker_order_alamat", ""),
            ("sticker_order_material", ""),
            ("sticker_order_brand", ""),
            ("sticker_order_version", ""),
            ("sticker_order_qty", 1),
            ("sticker_order_vendor", ""),
            ("_last_autofill_alamat_key", ""),
            ("_resolved_outlet_found", False),
            ("_resolved_outlet_alamat", ""),
        ):
            if k_def not in st.session_state:
                st.session_state[k_def] = v_def

        cur_id_val = str(st.session_state.get("sticker_order_id", "")).strip()
        cur_name_val = str(st.session_state.get("sticker_order_name", "")).strip()
        lookup_sig = f"{cur_id_val.casefold()}||{cur_name_val.casefold()}"

        if lookup_sig != st.session_state.get("_last_autofill_alamat_key", ""):
            st.session_state["_last_autofill_alamat_key"] = lookup_sig
            if cur_id_val or cur_name_val:
                found_m, auto_addr = lookup_master_address_fast(st.session_state.contracts, cur_id_val, cur_name_val)
                st.session_state["_resolved_outlet_found"] = found_m
                st.session_state["_resolved_outlet_alamat"] = auto_addr
                if found_m and auto_addr:
                    st.session_state["sticker_order_alamat"] = auto_addr
                elif not found_m:
                    st.session_state["sticker_order_alamat"] = ""
            else:
                st.session_state["_resolved_outlet_found"] = False
                st.session_state["_resolved_outlet_alamat"] = ""
                st.session_state["sticker_order_alamat"] = ""

        st.markdown("##### Detail Pemesanan Penggantian Sticker / Vinyl Panel")
        col_s1, col_s2 = st.columns([1.2, 2.3])
        in_id_outlet = col_s1.text_input("ID Outlet (Opsional)", placeholder="Opsional — contoh: 501X0001", key="sticker_order_id")
        in_nama_outlet = col_s2.text_input(
            "Nama Outlet",
            placeholder="Contoh: TOKO CONTOH 1 (wajib diisi jika ID Outlet kosong)",
            key="sticker_order_name",
        )

        has_outlet_input = bool(in_id_outlet.strip() or in_nama_outlet.strip())
        is_outlet_found = bool(st.session_state.get("_resolved_outlet_found", False)) if has_outlet_input else False
        resolved_alamat_val = str(st.session_state.get("_resolved_outlet_alamat", "")) if is_outlet_found else ""

        if has_outlet_input and is_outlet_found:
            st.caption(
                "✅ ID / Nama Outlet cocok dengan Master Data: kolom **Alamat** terisi otomatis "
                f"(`{resolved_alamat_val or 'Kosong di master — silakan lengkapi'}`)."
            )
        elif has_outlet_input and not is_outlet_found:
            st.caption(
                "⚠️ ID / Nama Outlet belum cocok atau tidak ditemukan di Master Data: "
                "silakan periksa kembali ID/Nama Outlet agar pemesanan dapat disimpan."
            )

        in_alamat_outlet = st.text_input(
            "Alamat",
            placeholder="Otomatis terisi jika ID / Nama Outlet cocok (dapat diedit)..." if is_outlet_found else "Isi Alamat...",
            key="sticker_order_alamat",
        )

        col_s3, col_s4, col_s5 = st.columns([1.6, 1.3, 1.4])
        in_material = col_s3.text_input("Material", placeholder="Contoh: Vinyl Panel Toko 1 x 2 m (BL)", key="sticker_order_material")
        in_brand = col_s4.text_input("Brand", placeholder="Contoh: DJARUM BLACK MATCHA KRETEK", key="sticker_order_brand")
        in_version = col_s5.text_input("Version", placeholder="Opsional — contoh: Matcha Powder - Rp 17.000/12 btg", key="sticker_order_version")

        col_s6, col_s7, col_s8 = st.columns([1.0, 1.2, 1.3])
        in_jumlah = col_s6.number_input("Jumlah Sticker / Vinyl", min_value=1, step=1, key="sticker_order_qty")
        in_tgl_order = col_s7.date_input("Tanggal OrderPenggantian Sticker / Vinyl", value=date.today(), key="sticker_order_date")
        in_nama_vendor = col_s8.text_input("Nama Vendor", placeholder="Contoh: ARTISTIC atau 77 ADVERTISING", key="sticker_order_vendor")

        if picked_rec is not None and int(in_jumlah) > int(picked_rec["Available"]):
            st.warning(
                f"⚠️ Jumlah pesanan (**{int(in_jumlah)}**) melebihi stok Available di Gudang Promosi "
                f"(**{picked_rec['Available']} {picked_rec['UOM']}**)."
            )

        st.markdown("")
        if st.button("💾 Simpan & Update Data Customer", type="primary", width="stretch", key="btn_submit_sticker_single"):
            if not in_id_outlet.strip() and not in_nama_outlet.strip():
                st.error("⚠️ Isi minimal **ID Outlet** atau **Nama Outlet** untuk melakukan lookup ke Master Data Dashboard.")
            elif not in_alamat_outlet.strip():
                st.error("⚠️ Kolom **Alamat** wajib diisi.")
            elif not lookup_master_indices(st.session_state.contracts, in_id_outlet.strip(), in_nama_outlet.strip()):
                st.error(
                    f"❌ Data (**ID Outlet**: `{in_id_outlet.strip() or '-'}` | **Nama Outlet**: `{in_nama_outlet.strip() or '-'}`) "
                    "tidak ditemukan pada **Master Data Dashboard**. Pemesanan **tidak dapat disimpan**."
                )
            else:
                order_row = pd.DataFrame([{
                    "ID Outlet": in_id_outlet.strip() or pd.NA,
                    "Nama Outlet": in_nama_outlet.strip() or pd.NA,
                    "Alamat": in_alamat_outlet.strip() or pd.NA,
                    "Material": in_material.strip() or pd.NA,
                    "Brand": in_brand.strip() or pd.NA,
                    "Version": in_version.strip() or pd.NA,
                    "Jumlah Sticker / Vinyl": int(in_jumlah),
                    "Tanggal OrderPenggantian Sticker / Vinyl": pd.Timestamp(in_tgl_order) if in_tgl_order else pd.Timestamp(date.today()),
                    "Nama Vendor": in_nama_vendor.strip() or pd.NA,
                }])
                try:
                    updated_df, upd_cnt, rec_list, _ = apply_sticker_replacement_orders(
                        st.session_state.contracts,
                        order_row,
                        strict_master_lookup=True,
                        already_normalized=True,
                    )
                    saved_df = save_data(updated_df, already_normalized=True)
                    set_contracts_state(saved_df, already_normalized=True)
                    st.session_state.load_warning = None
                    st.session_state.sticker_order_results = append_sticker_order_log(
                        rec_list, st.session_state.sticker_order_results
                    )
                    st.session_state["_pending_clear_sticker_outlet"] = True
                    st.session_state["_last_autofill_alamat_key"] = ""
                    st.session_state["_resolved_outlet_found"] = False
                    st.session_state["_resolved_outlet_alamat"] = ""
                    st.session_state["_sticker_submit_success_msg"] = (
                        f"✅ Lookup cocok dengan Master Data Dashboard: berhasil mengupdate **{upd_cnt}** baris customer "
                        f"dan mencatat hasil pemesanan sticker / vinyl panel ke log."
                    )
                    st.rerun()
                except Exception as error:
                    st.error(f"Gagal menyimpan pemesanan penggantian sticker / vinyl panel: {error}")

    with tab_sticker_batch:
        st.markdown("##### Langkah 1: Download Template Excel Penggantian Sticker / Vinyl Panel")
        st.caption("Kolom **ID Outlet** pada template bersifat opsional jika **Nama Outlet** diisi, dan keduanya tetap diverifikasi terhadap Master Data Dashboard.")
        sticker_sample = pd.DataFrame([
            {
                "ID Outlet": "501X0001",
                "Nama Outlet": "TOKO CONTOH 1",
                "Alamat": "JALAN RAYA KARANG TURI RT 24 RW 8",
                "Material": "Vinyl Panel Toko 1 x 2 m (BL)",
                "Brand": "DJARUM BLACK MATCHA KRETEK",
                "Version": "Matcha Powder - Rp 17.000/12 btg",
                "Jumlah Sticker / Vinyl": 2,
                "Tanggal OrderPenggantian Sticker / Vinyl": pd.Timestamp("2026-11-05"),
                "Nama Vendor": "ARTISTIC",
            },
            {
                "ID Outlet": "S50113D250917120644415",
                "Nama Outlet": "TOKO CONTOH 2",
                "Alamat": "JALAN RAYA INTAN P38 BLOK 12D",
                "Material": "Sticker / Vinyl Panel Toko 1 x 2 m",
                "Brand": "LA LIGHTS",
                "Version": "Cat Eyes",
                "Jumlah Sticker / Vinyl": 1,
                "Tanggal OrderPenggantian Sticker / Vinyl": pd.Timestamp("2026-11-05"),
                "Nama Vendor": "77 ADVERTISING",
            },
        ])[STICKER_ORDER_COLUMNS]

        tpl_col1, tpl_col2 = st.columns(2)
        tpl_col1.download_button(
            "📄 Download Template Kosong (.xlsx)",
            data=excel_bytes(pd.DataFrame(columns=STICKER_ORDER_COLUMNS), sheet_name="Penggantian Sticker - Vinyl"),
            file_name="Template_Penggantian_Sticker_Vinyl_Panel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        tpl_col2.download_button(
            "📝 Download Template + Contoh (.xlsx)",
            data=excel_bytes(sticker_sample, sheet_name="Penggantian Sticker - Vinyl"),
            file_name="Template_Penggantian_Sticker_Vinyl_Panel_Contoh.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

        st.divider()
        st.markdown("##### Langkah 2: Upload File Excel Penggantian Massal")
        sticker_batch_file = st.file_uploader(
            "Pilih file Excel (.xlsx) penggantian sticker / vinyl panel",
            type=["xlsx"],
            key="sticker_batch_uploader",
        )

        if st.button("🚀 Proses & Update Massal Sticker / Vinyl Panel", type="primary", key="btn_process_sticker_batch"):
            if sticker_batch_file is None:
                st.warning("Pilih file Excel terlebih dahulu.")
            else:
                try:
                    orders_df = read_sticker_order_excel(sticker_batch_file)
                    if orders_df.empty:
                        st.warning("Tidak ditemukan baris data pemesanan pada file Excel.")
                    else:
                        updated_df, upd_cnt, rec_list, _ = apply_sticker_replacement_orders(
                            st.session_state.contracts,
                            orders_df,
                            strict_master_lookup=True,
                            already_normalized=True,
                        )
                        saved_df = save_data(updated_df, already_normalized=True)
                        set_contracts_state(saved_df, already_normalized=True)
                        st.session_state.load_warning = None
                        st.session_state.sticker_order_results = append_sticker_order_log(
                            rec_list, st.session_state.sticker_order_results
                        )
                        st.success(
                            f"✅ Seluruh baris cocok dengan Master Data Dashboard: **{upd_cnt}** baris customer terupdate "
                            f"(dari **{len(orders_df)}** baris pemesanan) dan tercatat di log."
                        )
                        st.rerun()
                except Exception as error:
                    st.error(f"❌ Gagal memproses & menyimpan file Excel: {error}")

    st.divider()
    st.markdown("### 📋 Log Hasil Pemesanan & Kirim ke Pejabat")
    st.caption(
        "Menyimpan riwayat pemesanan penggantian sticker / vinyl panel secara persisten "
        "dan ikut disertakan pada proses **Backup & Restore**."
    )

    normalized_log = normalize_sticker_order_records(st.session_state.sticker_order_results)
    df_order_hasil = pd.DataFrame(normalized_log, columns=STICKER_LOG_COLUMNS) if normalized_log else pd.DataFrame(columns=STICKER_LOG_COLUMNS)

    if df_order_hasil.empty:
        st.info("Belum ada log hasil pemesanan penggantian sticker / vinyl panel. Silakan input form atau upload file Excel pemesanan terlebih dahulu.")
    else:
        log_q_col, log_t_col, log_v_col = st.columns([2, 1.1, 1.1])
        log_search = log_q_col.text_input(
            "Cari di Log Hasil Pemesanan (Outlet, Material, Brand, Version, Vendor, atau Tanggal)",
            placeholder="Ketik kata kunci...",
            key="sticker_log_search",
        )
        tgl_col_name = "Tanggal OrderPenggantian Sticker / Vinyl"
        log_dates = ["Semua Tanggal"] + sorted(
            [
                t
                for t in df_order_hasil[tgl_col_name].dropna().astype(str).map(str.strip).unique()
                if t and t != "-"
            ],
            reverse=True,
        )
        log_date_filter = log_t_col.selectbox(
            "Filter Tanggal Order Penggantian",
            log_dates,
            key="sticker_log_date",
        )
        log_vendors = ["Semua Vendor"] + sorted(
            [v for v in df_order_hasil["Nama Vendor"].dropna().astype(str).unique() if v and v != "-"]
        )
        log_vendor_filter = log_v_col.selectbox("Filter Vendor", log_vendors, key="sticker_log_vendor")

        df_log_view = df_order_hasil.copy()
        if log_search.strip():
            lq = log_search.strip().casefold()
            mat_series = (
                df_log_view["Material"].fillna("").astype(str)
                if "Material" in df_log_view.columns
                else pd.Series("", index=df_log_view.index)
            )
            l_mask = (
                df_log_view["Nama Outlet"].fillna("").astype(str).str.casefold().str.contains(lq, regex=False)
                | mat_series.str.casefold().str.contains(lq, regex=False)
                | df_log_view["Brand"].fillna("").astype(str).str.casefold().str.contains(lq, regex=False)
                | df_log_view["Version"].fillna("").astype(str).str.casefold().str.contains(lq, regex=False)
                | df_log_view["Nama Vendor"].fillna("").astype(str).str.casefold().str.contains(lq, regex=False)
                | df_log_view[tgl_col_name].fillna("").astype(str).str.casefold().str.contains(lq, regex=False)
            )
            df_log_view = df_log_view[l_mask]
        if log_date_filter != "Semua Tanggal":
            df_log_view = df_log_view[df_log_view[tgl_col_name].fillna("").astype(str).str.strip() == log_date_filter]
        if log_vendor_filter != "Semua Vendor":
            df_log_view = df_log_view[df_log_view["Nama Vendor"] == log_vendor_filter]

        st.dataframe(df_log_view, width="stretch", hide_index=True, height=320)

        target_export_df = (df_log_view if not df_log_view.empty else df_order_hasil).copy()
        if not target_export_df.empty:
            target_export_df["_sort_mat"] = target_export_df.apply(
                lambda r: (safe_text(r.get("Material")) or safe_text(r.get("Ukuran")) or "").casefold(),
                axis=1,
            )
            target_export_df["_sort_brand"] = target_export_df["Brand"].fillna("").astype(str).map(lambda x: x.strip().casefold())
            target_export_df = (
                target_export_df.sort_values(by=["_sort_mat", "_sort_brand"], kind="mergesort")
                .drop(columns=["_sort_mat", "_sort_brand"])
                .reset_index(drop=True)
            )
        first_row = target_export_df.iloc[0]
        tgl_ringkas = (
            log_date_filter
            if log_date_filter != "Semua Tanggal"
            else (safe_text(first_row.get(tgl_col_name)) or date.today().strftime("%d/%m/%Y"))
        )
        vendor_ringkas = (
            log_vendor_filter
            if log_vendor_filter != "Semua Vendor"
            else (safe_text(first_row.get("Nama Vendor")) or "Semua Vendor")
        )

        dl_ex_col, dl_png_col, clr_chk_col, clr_col = st.columns([1.3, 1.3, 0.9, 1.0])
        dl_ex_col.download_button(
            "📥 Download Log Hasil (.xlsx)",
            data=excel_bytes(target_export_df, sheet_name="Log Hasil Pemesanan"),
            file_name=f"Log_Hasil_Pemesanan_Sticker_Vinyl_Panel_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

        png_bytes_ready = st.session_state.get("sticker_order_png_bytes")
        if dl_png_col.button("🖼️ Generate Gambar PNG (Tanpa ID Outlet)", width="stretch", key="btn_prep_png"):
            with st.spinner("Menghasilkan gambar .PNG dengan kolom Material & auto-adjust kolom..."):
                png_bytes_ready = render_sticker_order_png_with_material(target_export_df, tgl_ringkas, vendor_ringkas)
                st.session_state["sticker_order_png_bytes"] = png_bytes_ready

        confirm_clear_log = clr_chk_col.checkbox("Konfirmasi hapus log", value=False, key="confirm_clear_sticker_log")
        if clr_col.button("🧹 Hapus Semua Log", width="stretch", disabled=not confirm_clear_log, key="btn_clear_order_session"):
            st.session_state.sticker_order_results = save_sticker_order_log([])
            st.session_state.pop("sticker_order_png_bytes", None)
            st.session_state.pop("confirm_clear_sticker_log", None)
            st.rerun()

        if png_bytes_ready:
            st.download_button(
                "⬇️ Download Gambar Hasil (.png)",
                data=png_bytes_ready,
                file_name=f"Hasil_Pemesanan_Sticker_Vinyl_Panel_{date.today().strftime('%Y%m%d')}.png",
                mime="image/png",
                key="btn_dl_png_ready",
            )
            with st.expander("🖼️ Preview Hasil Gambar Pemesanan (.PNG)", expanded=True):
                st.image(png_bytes_ready, caption="Hasil Pemesanan Penggantian Sticker / Vinyl Panel (ID Outlet disembunyikan)", width="stretch")

        st.markdown("#### ✉️ Kirim Hasil Pemesanan ke Pejabat & Simpan Data")
        daftar_pejabat = st.session_state.pejabat_emails
        p_col1, p_col2 = st.columns([2, 1.5])
        with p_col1:
            selected_cc_pejabat = st.multiselect(
                "Pilih Email Pejabat (Tujuan CC — To dikosongkan):",
                options=daftar_pejabat,
                default=[em for em in DEFAULT_PEJABAT_EMAILS if em in daftar_pejabat] or daftar_pejabat,
                key="pejabat_cc_multiselect",
            )
        with p_col2:
            manual_pejabat = st.text_input(
                "Tambah / Ketik Email CC Pejabat Manual (Opsional)",
                placeholder="nama@perusahaan.com",
                key="pejabat_manual_input",
            )
            if st.button("➕ Simpan ke Daftar Pejabat", key="btn_save_pejabat_email"):
                merged_list = list(daftar_pejabat)
                for em in re.split(r"[;,]", manual_pejabat):
                    clean_em = em.strip()
                    if "@" in clean_em and clean_em not in merged_list:
                        merged_list.append(clean_em)
                save_pejabat_emails(merged_list)
                st.session_state.pejabat_emails = merged_list
                st.success("✅ Daftar email pejabat berhasil diperbarui.")
                st.rerun()

        cc_list = list(selected_cc_pejabat)
        for em in re.split(r"[;,]", manual_pejabat):
            clean_em = em.strip()
            if "@" in clean_em and clean_em not in cc_list:
                cc_list.append(clean_em)

        cc_str = "; ".join(cc_list)
        table_cols = ["No", "Nama Outlet", "Ukuran", "Brand", "Version", "Qty", "Vendor"]
        max_email_rows = 15
        table_rows: list[list[str]] = []
        for idx_r, (_, r) in enumerate(target_export_df.head(max_email_rows).iterrows(), start=1):
            ukuran_val = safe_text(r.get("Ukuran")) or safe_text(r.get("Material")) or "-"
            table_rows.append([
                str(idx_r),
                safe_text(r.get("Nama Outlet")) or "-",
                ukuran_val,
                safe_text(r.get("Brand")) or "-",
                safe_text(r.get("Version")) or "-",
                str(r.get("Jumlah Sticker / Vinyl", 0)),
                safe_text(r.get("Nama Vendor")) or "-",
            ])

        col_widths = [max(len(col), *(len(row[c_i]) for row in table_rows)) for c_i, col in enumerate(table_cols)]

        def _fmt_email_row(cells: list[str]) -> str:
            return "| " + " | ".join(cells[i].ljust(col_widths[i]) for i in range(len(table_cols))) + " |"

        sep_line = "+-" + "-+-".join("-" * w for w in col_widths) + "-+"
        ringkasan_tabel = "\n".join([
            sep_line,
            _fmt_email_row(table_cols),
            sep_line,
            *(_fmt_email_row(row) for row in table_rows),
            sep_line,
        ])
        extra_note = (
            f"\n(Menampilkan {max_email_rows} dari {len(target_export_df)} item; detail lengkap terlampir pada file Excel/PNG)\n"
            if len(target_export_df) > max_email_rows
            else ""
        )

        email_subject = f"Pemesanan Penggantian Sticker / Vinyl Panel Toko - {tgl_ringkas}"
        email_body = (
            "Dengan Hormat,\n\n"
            f"Berikut Kami kirimkan pemesanan penggantian Sticker / Vinyl Panel Toko ({len(target_export_df)} item):\n\n"
            f"{ringkasan_tabel}\n"
            f"{extra_note}\n"
            "Terima kasih.\n\n"
            "Hormat kami,\n"
        )
        mailto_pejabat_url = (
            f"mailto:"
            f"?cc={urllib.parse.quote(cc_str)}"
            f"&subject={urllib.parse.quote(email_subject)}"
            f"&body={urllib.parse.quote(email_body)}"
        )

        act_p_col1, act_p_col2 = st.columns([1.3, 1.3])
        if act_p_col1.button("💾 Simpan Data (Tanpa Mengirim Email ke Pejabat)", type="secondary", width="stretch", key="btn_save_order_no_email"):
            with st.spinner("Menyimpan data ke dashboard, Google Sheet, dan Supadata/Supabase..."):
                try:
                    saved_contracts_now = save_data(st.session_state.contracts, already_normalized=True)
                    set_contracts_state(saved_contracts_now, already_normalized=True)
                    updated_log_now = update_log_status_for_rows(
                        st.session_state.sticker_order_results,
                        target_export_df,
                        "Tersimpan (Tanpa Kirim Email)",
                    )
                    st.session_state.sticker_order_results = updated_log_now
                    ok_targets, warn_targets = sync_all_cloud_targets(
                        saved_contracts_now,
                        updated_log_now,
                        already_normalized=True,
                    )
                    st.session_state["last_order_action_status"] = {
                        "type": "saved_no_email",
                        "ok": ok_targets,
                        "warn": warn_targets,
                        "mailto_url": None,
                        "cc_count": len(cc_list),
                        "cc_str": cc_str,
                    }
                    st.rerun()
                except Exception as err:
                    st.error(f"Gagal menyimpan data: {err}")

        if act_p_col2.button(
            f"✉️ Send Email To Pejabat (To: Kosong | CC: {len(cc_list)})",
            type="primary",
            width="stretch",
            key="btn_send_email_pejabat_and_save",
        ):
            with st.spinner("Menyimpan data ke dashboard, Google Sheet, dan Supadata/Supabase lalu menyiapkan email..."):
                try:
                    saved_contracts_now = save_data(st.session_state.contracts, already_normalized=True)
                    set_contracts_state(saved_contracts_now, already_normalized=True)
                    updated_log_now = update_log_status_for_rows(
                        st.session_state.sticker_order_results,
                        target_export_df,
                        "Disiapkan Kirim Email Pejabat",
                    )
                    st.session_state.sticker_order_results = updated_log_now
                    ok_targets, warn_targets = sync_all_cloud_targets(
                        saved_contracts_now,
                        updated_log_now,
                        already_normalized=True,
                    )
                    st.session_state["last_order_action_status"] = {
                        "type": "sent_email",
                        "ok": ok_targets,
                        "warn": warn_targets,
                        "mailto_url": mailto_pejabat_url,
                        "cc_count": len(cc_list),
                        "cc_str": cc_str,
                    }
                    st.rerun()
                except Exception as err:
                    st.error(f"Gagal menyimpan & menyiapkan email: {err}")

        last_act = st.session_state.get("last_order_action_status")
        if isinstance(last_act, dict):
            if last_act.get("ok"):
                st.success("✅ Data terupdate ke **Dashboard** dan tersimpan di: **" + ", ".join(last_act["ok"]) + "**.")
            else:
                st.info("✅ Data terupdate ke **Dashboard** & penyimpanan lokal.")

            for w_msg in last_act.get("warn") or []:
                st.warning(f"⚠️ {w_msg}")

            if last_act.get("type") == "sent_email" and last_act.get("mailto_url"):
                st.markdown(
                    f"""
                    <a href="{last_act['mailto_url']}" target="_blank" style="text-decoration: none;">
                        <div style="background-color: #1D4ED8; color: #FFFFFF; padding: 11px 18px; border-radius: 8px; text-align: center; font-weight: 700; font-size: 15px; margin-top: 8px;">
                            📤 Buka Draft Email ke Pejabat Sekarang (To: Kosong | CC: {last_act.get('cc_count', 0)} penerima)
                        </div>
                    </a>
                    """,
                    unsafe_allow_html=True,
                )
                st.caption(f"Tujuan **To**: *(dikosongkan)* • **CC**: `{last_act.get('cc_str') or '-'}`")


# ==========================================
# HALAMAN 4: LIST MATERI PROMOSI (GUDANG PROMOSI)
# ==========================================
elif active_menu == "📦 List Materi Promosi":
    st.markdown("### 📦 List Materi Promosi di Gudang Promosi")
    st.caption(
        "Menampilkan stok yang berstatus **Available > 0** dan hanya materi yang mengandung "
        "**`Sticker / Vinyl Panel Toko`**."
    )

    with st.expander("📤 Update / Upload File Laporan Gudang Promosi (PRCRG01 .xlsx)", expanded=promo_stock_df.empty):
        promo_upload = st.file_uploader("Upload File Stok Gudang Promosi (.xlsx)", type=["xlsx"], key="promo_stock_uploader")
        if st.button("🔄 Muat & Simpan List Materi Promosi", type="primary", key="btn_load_promo_stock"):
            if promo_upload is None:
                st.warning("Pilih file Excel stok gudang terlebih dahulu.")
            else:
                try:
                    parsed_stock = read_promo_stock_workbook(promo_upload)
                    PROMO_STOCK_FILE.parent.mkdir(parents=True, exist_ok=True)
                    parsed_stock.to_excel(PROMO_STOCK_FILE, sheet_name="PRCRG01", index=False, engine="openpyxl")
                    clear_promo_stock_cache()
                    st.session_state.pop("_warehouse_stock_cache", None)
                    st.session_state.promo_stock = parsed_stock
                    st.success(f"✅ Berhasil memuat **{len(parsed_stock)}** item materi Sticker / Vinyl Panel Toko yang Available.")
                    st.rerun()
                except Exception as error:
                    st.error(f"Gagal membaca file stok gudang: {error}")

    stock_view = promo_stock_df.copy()
    if stock_view.empty:
        st.info("Belum ada data materi **Sticker / Vinyl Panel Toko** yang berstatus **Available > 0**.")
    else:
        total_avail_qty = int(pd.to_numeric(stock_view["Available"], errors="coerce").sum())
        m_c1, m_c2, m_c3 = st.columns(3)
        m_c1.metric("Jumlah Varian Materi (Available)", len(stock_view))
        m_c2.metric("Total Stok Available", total_avail_qty)
        m_c3.metric("Jumlah Brand", int(stock_view["Brand"].dropna().nunique()))

        f_col1, f_col2 = st.columns([2, 1])
        promo_q = f_col1.text_input("Cari Materi / Brand / Version", placeholder="Ketik nama material, brand, atau versi...", key="promo_stock_search")
        brand_choices = ["Semua Brand"] + sorted(stock_view["Brand"].dropna().astype(str).unique().tolist())
        promo_brand = f_col2.selectbox("Filter Brand", brand_choices, key="promo_stock_brand")

        if promo_q.strip():
            qk = promo_q.strip().casefold()
            stock_view = stock_view[
                stock_view["Material"].fillna("").astype(str).str.casefold().str.contains(qk, regex=False)
                | stock_view["Brand"].fillna("").astype(str).str.casefold().str.contains(qk, regex=False)
                | stock_view["Version"].fillna("").astype(str).str.casefold().str.contains(qk, regex=False)
            ]
        if promo_brand != "Semua Brand":
            stock_view = stock_view[stock_view["Brand"] == promo_brand]

        st.dataframe(
            stock_view,
            width="stretch",
            hide_index=True,
            height=420,
            column_config={
                "Material": st.column_config.TextColumn("Material", width="large"),
                "Brand": st.column_config.TextColumn("Brand", width="medium"),
                "Version": st.column_config.TextColumn("Version", width="large"),
                "Available": st.column_config.NumberColumn("Available", format="%d", width="small"),
                "UOM": st.column_config.TextColumn("UOM", width="small"),
                "Status Material": st.column_config.TextColumn("Status Material", width="medium"),
            },
        )

        st.download_button(
            "📥 Export List Materi Promosi Available (.xlsx)",
            data=excel_bytes(stock_view, sheet_name="List Materi Promosi"),
            file_name="List_Materi_Promosi_Sticker_Vinyl_Available.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_promo_stock_available",
        )


# ==========================================
# HALAMAN 5: IMPORT & SINKRONISASI
# ==========================================
elif active_menu == "🔄 Import & Sinkronisasi":
    tab_main_import, tab_branding, tab_drp_crosscheck, tab_panel_sync = st.tabs([
        "📥 Replace Workbook Utama",
        "🏷️ Update SPKS Branding",
        "🔎 Crosscheck Data DRP",
        "📐 Sinkronisasi Ukuran Panel & Layar",
    ])

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
                        saved_df = save_data(imported, already_normalized=True)
                        set_contracts_state(saved_df, already_normalized=True)
                        st.session_state.load_warning = None
                        st.success(f"✅ Berhasil mengimpor {len(saved_df)} baris data.")
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
                width="stretch",
            )

    with tab_branding:
        st.caption("Mengupdate data customer yang cocok berdasarkan **ID Customer** (mengambil tanggal pekerjaan paling baru jika terdapat duplikat).")
        branding_upload = st.file_uploader("Upload Laporan SPKS Branding (.xlsx)", type=["xlsx"], key="branding_workbook_upload")
        if st.button("Proses & Update Data Customer", type="primary"):
            st.session_state.pop("branding_import_result", None)
            if branding_upload is None:
                st.warning("Pilih file laporan branding terlebih dahulu.")
            else:
                try:
                    branding_data = read_branding_workbook(branding_upload)
                    updated, unmatched, duplicates, updated_count = apply_branding_updates(
                        _ensure_columns(st.session_state.contracts),
                        branding_data,
                        already_normalized=True,
                    )
                    problems = validate_data(updated, already_normalized=True)
                    if problems:
                        st.error("\n".join(problems))
                    else:
                        if updated_count:
                            saved_df = save_data(updated, already_normalized=True)
                            set_contracts_state(saved_df, already_normalized=True)
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
                        width="stretch",
                        hide_index=True,
                    )
            missing_ids = import_result["unmatched"][import_result["unmatched"]["ID Customer"].map(customer_id_key).ne("")]
            if not missing_ids.empty:
                with st.expander(f"Lihat {len(missing_ids)} ID yang Tidak Ditemukan di Database"):
                    st.dataframe(missing_ids, width="stretch", hide_index=True)

    with tab_drp_crosscheck:
        st.markdown("#### 🔎 Crosscheck Data DRP dengan Dashboard & Monitoring")
        st.caption(
            "Unggah file Excel DRP (contoh: `CODPR03 (DRP).xlsx`) untuk mencocokkan **khusus outlet dengan Kategori Pekerjaan = DRP** "
            "(DRP vs DRP) berdasarkan **ID Customer**, memperbarui data yang tidak sesuai, menambahkan data yang belum ada, "
            "dan mengelola data DRP di Dashboard yang tidak ada di file upload."
        )

        drp_upload = st.file_uploader("Upload File Excel DRP (.xlsx)", type=["xlsx"], key="drp_crosscheck_upload")

        available_compare_fields = [
            "Nama Customer",
            "Brand",
            "Nominal Kontrak",
            "Periode Kontrak Awal",
            "Periode Kontrak Akhir",
            "No. Dokumen",
            "Zona",
            "Kategori Outlet",
        ]
        selected_compare_fields = st.multiselect(
            "Pilih Kolom yang Dicocokkan",
            options=available_compare_fields,
            default=["Nama Customer", "Brand", "Nominal Kontrak", "Periode Kontrak Awal", "Periode Kontrak Akhir", "No. Dokumen"],
            key="drp_compare_fields",
        )

        run_col, reset_drp_col = st.columns([3, 1])
        with run_col:
            run_drp_clicked = st.button("🚀 Jalankan Crosscheck DRP", type="primary", width="stretch", key="btn_run_drp_crosscheck")
        with reset_drp_col:
            if st.button("🧹 Reset Hasil DRP", width="stretch", key="btn_reset_drp_crosscheck"):
                st.session_state.pop("drp_crosscheck_source", None)
                st.session_state.pop("drp_crosscheck_fields", None)
                st.session_state.pop("drp_crosscheck_dup_count", None)
                st.rerun()

        if run_drp_clicked:
            if drp_upload is None:
                st.warning("Pilih file Excel DRP terlebih dahulu.")
            elif not selected_compare_fields:
                st.warning("Pilih minimal satu kolom untuk dicocokkan.")
            else:
                try:
                    drp_parsed, dup_cnt = read_drp_crosscheck_workbook(drp_upload)
                    st.session_state["drp_crosscheck_source"] = drp_parsed
                    st.session_state["drp_crosscheck_fields"] = list(selected_compare_fields)
                    st.session_state["drp_crosscheck_dup_count"] = dup_cnt
                except Exception as error:
                    st.error(f"Gagal membaca file DRP: {error}")

        drp_source_df = st.session_state.get("drp_crosscheck_source")
        active_fields = st.session_state.get("drp_crosscheck_fields", selected_compare_fields)
        dup_drp_count = int(st.session_state.get("drp_crosscheck_dup_count", 0))

        if drp_source_df is not None and not drp_source_df.empty:
            if dup_drp_count > 0:
                st.info(
                    f"ℹ️ Ditemukan **{dup_drp_count}** baris `ID Customer` duplikat pada file DRP upload — "
                    "otomatis diambil baris dengan tanggal akhir kontrak terbaru."
                )

            base_contracts_drp = _ensure_columns(st.session_state.contracts)
            mismatch_df, matched_ok_df, missing_df = build_drp_crosscheck_report(
                base_contracts_drp,
                drp_source_df,
                active_fields,
            )
            dash_only_df = build_dashboard_only_report(
                base_contracts_drp,
                drp_source_df,
            )

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Total Data DRP (Unik)", len(drp_source_df))
            m2.metric("Sesuai Semua", len(matched_ok_df))
            m3.metric("Tidak Sesuai", len(mismatch_df))
            m4.metric("ID Belum Ada di Dashboard", len(missing_df))
            m5.metric("Ada di Dash, Tidak di Upload", len(dash_only_df))

            sub_mismatch, sub_ok, sub_missing, sub_dash_only = st.tabs([
                f"⚠️ Tidak Sesuai ({len(mismatch_df)})",
                f"✅ Sudah Sesuai ({len(matched_ok_df)})",
                f"❓ ID Tidak di Dashboard ({len(missing_df)})",
                f"🗑️ Ada di Dash, Tidak di Upload ({len(dash_only_df)})",
            ])

            with sub_mismatch:
                if mismatch_df.empty:
                    st.success("✅ Seluruh data DRP yang cocok dengan ID di Dashboard sudah sesuai!")
                else:
                    st.markdown("**Data Tidak Sesuai — Pilih Baris & Kolom untuk Update ke Dashboard**")
                    fields_to_apply = st.multiselect(
                        "Kolom yang akan di-update ke Dashboard",
                        options=available_compare_fields,
                        default=active_fields,
                        key="drp_fields_to_apply",
                    )
                    select_all_mismatch = st.checkbox(
                        "Pilih semua baris tidak sesuai",
                        value=True,
                        key="drp_select_all_mismatch",
                    )

                    display_cols = [
                        "ID Customer", "Nama Customer (Dashboard)", "Nama Customer (DRP)", "Kolom Tidak Sesuai",
                        "Brand (Dash)", "Brand (DRP)", "Nominal (Dash)", "Nominal (DRP)",
                        "Awal Kontrak (Dash)", "Awal Kontrak (DRP)", "Akhir Kontrak (Dash)", "Akhir Kontrak (DRP)",
                        "No. Dokumen (Dash)", "No. Dokumen (DRP)", "Zona (Dash)", "Zona (DRP)",
                        "Kategori Outlet (Dash)", "Kategori Outlet (DRP)",
                    ]

                    editor_df = mismatch_df[display_cols].copy()
                    editor_df.insert(0, "Pilih Update", select_all_mismatch)

                    edited_mismatch = st.data_editor(
                        editor_df,
                        width="stretch",
                        hide_index=True,
                        height=420,
                        disabled=display_cols,
                        key="drp_mismatch_editor",
                    )

                    dl_col, upd_col = st.columns(2)
                    dl_col.download_button(
                        "📥 Download Laporan Selisih DRP (.xlsx)",
                        data=excel_bytes(mismatch_df[display_cols]),
                        file_name="Crosscheck_Selisih_DRP.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                    )

                    if upd_col.button("💾 Update Data Terpilih ke Dashboard", type="primary", width="stretch", key="btn_apply_drp_updates"):
                        if not fields_to_apply:
                            st.warning("Pilih minimal satu kolom yang akan di-update.")
                        else:
                            selected_mask = edited_mismatch["Pilih Update"].fillna(False).to_numpy(dtype=bool)
                            rows_to_update = mismatch_df.loc[selected_mask]
                            if rows_to_update.empty:
                                st.warning("Centang minimal satu baris data untuk di-update.")
                            else:
                                updated_contracts, updated_counter = apply_drp_updates_to_contracts(
                                    base_contracts_drp,
                                    rows_to_update,
                                    fields_to_apply,
                                    already_normalized=True,
                                )
                                try:
                                    saved_df = save_data(updated_contracts, already_normalized=True)
                                    set_contracts_state(saved_df, already_normalized=True)
                                    st.success(f"✅ Berhasil meng-update {updated_counter} outlet di Dashboard sesuai data DRP!")
                                    st.rerun()
                                except Exception as error:
                                    st.error(f"Gagal menyimpan update DRP: {error}")

            with sub_ok:
                if matched_ok_df.empty:
                    st.info("Belum ada data yang sesuai sepenuhnya pada kolom yang dipilih.")
                else:
                    ok_cols = [
                        "ID Customer", "Nama Customer (Dashboard)", "Brand (DRP)",
                        "Nominal (DRP)", "Awal Kontrak (DRP)", "Akhir Kontrak (DRP)", "No. Dokumen (DRP)",
                    ]
                    st.dataframe(matched_ok_df[ok_cols], width="stretch", hide_index=True)

            with sub_missing:
                if missing_df.empty:
                    st.info("Seluruh ID Customer pada file DRP sudah terdaftar di Dashboard.")
                else:
                    st.markdown("**Data DRP Belum Ada di Dashboard — Pilih Baris untuk Ditambahkan Langsung**")
                    st.caption("Centang baris pada kolom **Pilih Tambah**, atur informasi bawaan (opsional), lalu klik **Tambahkan Data Terpilih ke Dashboard**.")

                    m_cfg1, m_cfg2, m_cfg3 = st.columns(3)
                    default_sector = m_cfg1.text_input("Sektor Bawaan", value="Retail Supply", key="drp_add_default_sector")
                    default_budget = m_cfg2.text_input("Pos Budget Bawaan", value="Small Retailer", key="drp_add_default_budget")
                    default_job_cat = m_cfg3.text_input("Kategori Pekerjaan Bawaan", value="DRP", key="drp_add_default_job_cat")

                    select_all_missing = st.checkbox(
                        "Pilih semua baris yang belum ada di Dashboard",
                        value=False,
                        key="drp_select_all_missing",
                    )

                    missing_view_df = pd.DataFrame({
                        "ID Customer": missing_df["ID Customer"].map(safe_text),
                        "Nama Customer": missing_df["Customer"].map(safe_text),
                        "Zona": missing_df["Zona"].map(safe_text),
                        "Kategori Outlet": missing_df["Kategori Outlet"].map(safe_text),
                        "Kategori DRP": missing_df["Kategori DRP"].map(safe_text),
                        "Brand": missing_df["Brand"].map(safe_text),
                        "Nominal Kontrak": missing_df["Nominal Kontrak"].map(lambda v: fmt_val_cmp("Nominal Kontrak", v)),
                        "Periode Kontrak Awal": missing_df["Periode Kontrak Awal"].map(lambda v: fmt_val_cmp("Periode Kontrak Awal", v)),
                        "Periode Kontrak Akhir": missing_df["Periode Kontrak Akhir"].map(lambda v: fmt_val_cmp("Periode Kontrak Akhir", v)),
                        "No. Dokumen": missing_df["No. Dokumen"].map(safe_text),
                    })
                    missing_view_df.insert(0, "Pilih Tambah", select_all_missing)

                    edited_missing = st.data_editor(
                        missing_view_df,
                        width="stretch",
                        hide_index=True,
                        height=400,
                        disabled=list(missing_view_df.columns[1:]),
                        key="drp_missing_editor",
                    )

                    m_dl_col, m_add_col = st.columns(2)
                    m_dl_col.download_button(
                        "📥 Download Data Belum Ada di Dashboard (.xlsx)",
                        data=excel_bytes(missing_view_df.drop(columns=["Pilih Tambah"])),
                        file_name="DRP_Belum_Ada_Di_Dashboard.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                    )

                    if m_add_col.button("➕ Tambahkan Data Terpilih ke Dashboard", type="primary", width="stretch", key="btn_add_missing_drp_to_dash"):
                        add_mask = edited_missing["Pilih Tambah"].fillna(False).to_numpy(dtype=bool)
                        rows_to_insert_raw = missing_df.loc[add_mask]
                        if rows_to_insert_raw.empty:
                            st.warning("Centang minimal satu baris data untuk ditambahkan ke Dashboard.")
                        else:
                            combined_contracts, added_cnt, skipped_existing = append_missing_drp_to_contracts(
                                base_contracts_drp,
                                rows_to_insert_raw,
                                default_sector=default_sector,
                                default_budget=default_budget,
                                default_job_cat=default_job_cat,
                                already_normalized=True,
                            )
                            if added_cnt == 0:
                                st.info("Tidak ada baris baru ditambahkan karena ID sudah terdaftar di Dashboard.")
                            else:
                                try:
                                    saved_df = save_data(combined_contracts, already_normalized=True)
                                    set_contracts_state(saved_df, already_normalized=True)
                                    msg = f"✅ Berhasil menambahkan {added_cnt} outlet baru ke Dashboard!"
                                    if skipped_existing:
                                        msg += f" ({skipped_existing} ID dilewati karena sudah ada)."
                                    st.success(msg)
                                    st.rerun()
                                except Exception as error:
                                    st.error(f"Gagal menambahkan data DRP: {error}")

            with sub_dash_only:
                if dash_only_df.empty:
                    st.success("✅ Semua data di Dashboard & Monitoring sudah ada pada file DRP yang di-upload.")
                else:
                    st.markdown("**Data Kategori Pekerjaan DRP di Dashboard & Monitoring, tetapi Tidak Ada di File Upload DRP**")
                    st.caption(
                        "Hanya menampilkan outlet dengan **Kategori Pekerjaan = DRP** yang ada di Dashboard tetapi tidak ada di file upload DRP. "
                        "Anda dapat **menghapus sebagian (per row)** dengan mencentang **Pilih Hapus** atau **menghapus semua**."
                    )

                    only_no_id_filter = st.checkbox(
                        "Hanya tampilkan yang memiliki ID Outlet (sembunyikan baris tanpa ID)",
                        value=False,
                        key="drp_dash_only_has_id",
                    )
                    view_only_df = dash_only_df.copy()
                    if only_no_id_filter:
                        view_only_df = view_only_df[view_only_df["ID Outlet"] != "-"].reset_index(drop=True)

                    if view_only_df.empty:
                        st.info("Tidak ada data yang cocok dengan filter tersebut.")
                    else:
                        display_only_cols = [
                            "ID Outlet", "Nama Customer", "Zona", "Sektor", "Pos Budget",
                            "Kategori Pekerjaan", "Brand", "Nominal Kontrak",
                            "Periode Kontrak Awal", "Periode Kontrak Akhir", "No. Dokumen",
                        ]
                        only_editor_input = view_only_df[display_only_cols].copy()
                        only_editor_input.insert(0, "Pilih Hapus", False)

                        edited_only = st.data_editor(
                            only_editor_input,
                            width="stretch",
                            hide_index=True,
                            height=420,
                            disabled=display_only_cols,
                            key="drp_dash_only_editor",
                        )

                        selected_row_mask = edited_only["Pilih Hapus"].fillna(False).to_numpy(dtype=bool)
                        selected_row_count = int(selected_row_mask.sum())

                        st.download_button(
                            "📥 Download Daftar Ada di Dash (Tidak di Upload) (.xlsx)",
                            data=excel_bytes(view_only_df[display_only_cols]),
                            file_name="Dashboard_Tidak_Ada_Di_Upload_DRP.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_drp_dash_only",
                        )

                        d_col1, d_col2 = st.columns(2)
                        with d_col1:
                            confirm_row_del = st.checkbox(
                                f"Konfirmasi hapus {selected_row_count} baris terpilih",
                                value=False,
                                disabled=(selected_row_count == 0),
                                key="confirm_del_dash_only_rows",
                            )
                            if st.button(
                                f"🗑️ Hapus Sebagian / Per Row ({selected_row_count} baris)",
                                type="secondary",
                                width="stretch",
                                disabled=(selected_row_count == 0 or not confirm_row_del),
                                key="btn_del_dash_only_rows",
                            ):
                                indices_to_drop = view_only_df.loc[selected_row_mask, "_dash_index"].astype(int).tolist()
                                remaining_df = st.session_state.contracts.drop(index=indices_to_drop).reset_index(drop=True)
                                saved_df = save_data(remaining_df, already_normalized=True)
                                set_contracts_state(saved_df, already_normalized=True)
                                st.session_state.pop("confirm_del_dash_only_rows", None)
                                st.success(f"✅ Berhasil menghapus {len(indices_to_drop)} baris terpilih dari Dashboard!")
                                st.rerun()

                        with d_col2:
                            confirm_all_del = st.checkbox(
                                f"Konfirmasi hapus SEMUA ({len(view_only_df)} baris tampil)",
                                value=False,
                                key="confirm_del_dash_only_all",
                            )
                            if st.button(
                                f"🚨 Hapus Semua ({len(view_only_df)} baris)",
                                type="primary",
                                width="stretch",
                                disabled=not confirm_all_del,
                                key="btn_del_dash_only_all",
                            ):
                                indices_to_drop_all = view_only_df["_dash_index"].astype(int).tolist()
                                remaining_df = st.session_state.contracts.drop(index=indices_to_drop_all).reset_index(drop=True)
                                saved_df = save_data(remaining_df, already_normalized=True)
                                set_contracts_state(saved_df, already_normalized=True)
                                st.session_state.pop("confirm_del_dash_only_all", None)
                                st.success(f"✅ Berhasil menghapus semua ({len(indices_to_drop_all)} baris) data yang tidak ada di file upload DRP!")
                                st.rerun()

    with tab_panel_sync:
        st.caption("Mencocokkan ukuran sticker / vinyl panel, brand sticker / vinyl, versi, dan ukuran layar dari sheet DRP berdasarkan ID Customer.")
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
                        _ensure_columns(st.session_state.contracts),
                        source_frame,
                        overwrite_existing=overwrite_sync,
                        already_normalized=True,
                    )
                    saved_df = save_data(merged_df, already_normalized=True)
                    set_contracts_state(saved_df, already_normalized=True)
                    st.success(f"✅ Sinkronisasi selesai: {updated_cnt} baris customer diperbarui.")
                    if missing_list:
                        st.info(f"{len(missing_list)} ID tidak ditemukan: {', '.join(missing_list[:15])}")
                except Exception as error:
                    st.error(f"Gagal sinkronisasi: {error}")


# ==========================================
# HALAMAN 6: BACKUP & RESTORE CLOUD
# ==========================================
elif active_menu == "☁️ Backup & Restore":
    col_sb, col_gs = st.columns(2)

    with col_sb:
        with st.container(border=True):
            st.subheader("🗄️ Supabase / Supadata Cloud")
            st.caption("Simpan snapshot database & log pemesanan sticker / vinyl panel atau pulihkan data terakhir dari cloud.")
            sb_url, sb_key, sb_table = get_supabase_config()
            if not sb_url or not sb_key:
                st.warning("Konfigurasi URL dan API key Supabase/Supadata belum diatur di Streamlit Secrets.")
            else:
                st.info(
                    f"Tabel: `{sb_table}` • Data kontrak: **{len(st.session_state.contracts)}** outlet • "
                    f"Log pemesanan sticker / vinyl: **{len(st.session_state.sticker_order_results)}** baris"
                )
                b1, b2 = st.columns(2)
                if b1.button("⬆️ Backup Sekarang", type="primary", width="stretch", key="btn_backup_supabase"):
                    try:
                        saved_cnt = backup_to_supabase(
                            _ensure_columns(st.session_state.contracts),
                            sb_url,
                            sb_key,
                            sb_table,
                            sticker_order_log=st.session_state.sticker_order_results,
                            already_normalized=True,
                        )
                        st.success(
                            f"✅ {saved_cnt} baris kontrak & {len(st.session_state.sticker_order_results)} baris log pemesanan tersimpan."
                        )
                    except Exception as error:
                        st.error(f"Gagal backup: {error}")

                if b2.button("⬇️ Restore Data", type="secondary", width="stretch", key="btn_restore_supabase"):
                    try:
                        restored_df, restored_log = restore_from_supabase(sb_url, sb_key, sb_table)
                        saved_df = save_data(_ensure_columns(restored_df))
                        set_contracts_state(saved_df, already_normalized=True)
                        st.session_state.load_warning = None
                        if restored_log is not None:
                            st.session_state.sticker_order_results = save_sticker_order_log(restored_log)
                        st.success(f"✅ {len(saved_df)} baris kontrak dipulihkan.")
                        st.rerun()
                    except Exception as error:
                        st.error(f"Gagal restore: {error}")

    with col_gs:
        with st.container(border=True):
            st.subheader("📗 Google Sheets")
            st.caption(f"Simpan atau pulihkan data kontrak tab `{SHEET_NAME}` beserta `sticker_order_log` melalui Google Apps Script.")
            backup_url = get_secret_str("GOOGLE_APPS_SCRIPT_URL", default="")
            if not backup_url:
                st.warning("Konfigurasi `GOOGLE_APPS_SCRIPT_URL` belum diatur di Streamlit Secrets.")
            else:
                st.info(
                    f"Data kontrak: **{len(st.session_state.contracts)}** outlet • "
                    f"Log pemesanan sticker / vinyl: **{len(st.session_state.sticker_order_results)}** baris"
                )
                gs1, gs2 = st.columns(2)
                if gs1.button("⬆️ Backup ke Google Sheet", type="primary", width="stretch", key="btn_backup_gsheet"):
                    try:
                        row_count = backup_to_google_sheet(
                            _ensure_columns(st.session_state.contracts),
                            backup_url,
                            sticker_order_log=st.session_state.sticker_order_results,
                            already_normalized=True,
                        )
                        st.success(f"✅ {row_count} baris berhasil dikirim ke Google Sheet (termasuk log pemesanan).")
                    except Exception as error:
                        st.error(f"Gagal backup: {error}")

                if gs2.button("⬇️ Restore dari Google Sheet", type="secondary", width="stretch", key="btn_restore_gsheet"):
                    try:
                        restored_df, restored_log = restore_from_google_sheet(backup_url)
                        saved_df = save_data(_ensure_columns(restored_df), already_normalized=True)
                        set_contracts_state(saved_df, already_normalized=True)
                        st.session_state.load_warning = None
                        if restored_log is not None:
                            st.session_state.sticker_order_results = save_sticker_order_log(restored_log)
                        st.success(f"✅ {len(saved_df)} baris kontrak berhasil dipulihkan dari Google Sheet.")
                        st.rerun()
                    except Exception as error:
                        st.error(f"Gagal restore dari Google Sheet: {error}")
