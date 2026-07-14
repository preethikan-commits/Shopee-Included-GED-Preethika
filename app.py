import streamlit as st
import pandas as pd
from datetime import date
import traceback

from processor import (
    process_voucher_eligibility,
    process_shopee_voucher_eligibility,
    generate_excel_output,
    generate_shopee_excel_output,
    generate_eligible_export,
    load_ecom_tracker,
    load_consolidated,
    is_zip_upload,
    parse_price_tier_ref,
    get_exclusion_remarks,
    EcomTrackerSheetNotFoundError,
    WORKING_COLS,
)

# ─────────────────────────────────────────────
# Page Config & CSS
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Voucher Eligibility Automation",
    page_icon="🏷️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    html, body, [class*="css"] {
        font-family: 'Arial', sans-serif;
        background-color: #FFFFFF;
        color: #111111;
    }
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1280px;
    }
    [data-testid="stSidebar"] {
        background-color: #F5F5F5;
        border-right: 1px solid #E0E0E0;
    }
    .section-card {
        background: #FAFAFA;
        border: 1px solid #E2E2E2;
        border-radius: 8px;
        padding: 1.25rem 1.5rem;
        margin-bottom: 1.25rem;
    }
    .section-header {
        font-size: 0.82rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #555;
        border-bottom: 2px solid #2D4A6B;
        padding-bottom: 0.4rem;
        margin-bottom: 1rem;
    }
    .vc-card {
        background: #F0F4FF;
        border: 1px solid #C8D5F0;
        border-radius: 6px;
        padding: 0.9rem 1rem 0.6rem 1rem;
        margin-bottom: 0.75rem;
    }
    .vc-label {
        font-size: 0.75rem;
        font-weight: 700;
        color: #2D4A6B;
        text-transform: uppercase;
        letter-spacing: 0.07em;
        margin-bottom: 0.4rem;
    }
    .stDownloadButton > button {
        background-color: #1F6B3E !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border-radius: 6px !important;
        padding: 0.55rem 1.5rem !important;
        font-size: 0.95rem !important;
        border: none !important;
        width: 100%;
    }
    .stDownloadButton > button:hover { background-color: #175533 !important; }
    .stButton > button {
        background-color: #2D4A6B !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border-radius: 6px !important;
        padding: 0.55rem 1.5rem !important;
        font-size: 0.95rem !important;
        border: none !important;
        width: 100%;
    }
    .stButton > button:hover { background-color: #1E3450 !important; }
    [data-testid="stFileUploader"] {
        border: 1px solid #D0D0D0 !important;
        border-radius: 6px;
    }
    [data-testid="metric-container"] {
        background: #F8F9FA;
        border: 1px solid #E0E0E0;
        border-radius: 8px;
        padding: 0.5rem 1rem;
    }
    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Session State Init
# ─────────────────────────────────────────────

if "voucher_configs" not in st.session_state:
    st.session_state.voucher_configs = [{"name": "10% VC", "keywords": []}]
if "excl_remarks" not in st.session_state:
    st.session_state.excl_remarks = []


def section(title):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Title
# ─────────────────────────────────────────────

c1, c2 = st.columns([0.07, 0.93])
with c1:
    st.markdown("## 🏷️")
with c2:
    st.markdown("# Voucher Eligibility Automation")
    st.markdown(
        "<span style='color:#555; font-size:0.9rem;'>"
        "Lazada, Zalora & Shopee · SG / MY / PH · Independent per-voucher eligibility processing"
        "</span>",
        unsafe_allow_html=True,
    )

st.markdown("---")

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────

with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    st.markdown("---")

    section("Marketplace & Region")
    marketplace = st.selectbox("Select Marketplace", ["Lazada", "Zalora", "Shopee"], key="marketplace")
    region      = st.selectbox("Region", ["SG", "MY", "PH"], key="region")

    st.markdown("---")
    section("Price Tier Reference")
    price_tier_ref = st.text_input(
        "Ecom Tracker Price Tier Reference",
        value="AX-BA",
        help=(
            "Column range in the Ecom Tracker. Format: START-END (e.g. AX-BA).\n\n"
            "Positional mapping:\n"
            "  START   = RRP\n"
            "  START+1 = SRP\n"
            "  START+2 = DISC %\n"
            "  END     = Exclusion Remarks\n\n"
            "Columns are resolved by name first, then by position."
        ),
    )

    st.markdown("---")
    section("Launch Date Cutoff")
    cutoff_date = st.date_input(
        "Cutoff Date",
        value=date.today(),
        help=(
            "Products with a Launch Date after this date are treated as Future Launch and excluded. "
            "The Ecom Tracker Launch Dates column is expected in DD-MM-YYYY format; "
            "the tool handles this natively and normalises all date inputs to DD-MM-YYYY in the output. "
            "\"NA\" values in the selected Launch Date column are always treated as eligible."
        ),
    )

    st.markdown("---")
    st.markdown(
        "<span style='font-size:0.78rem; color:#888;'>"
        "Upload files in the main panel. Inclusion keywords will auto-populate "
        "from the Ecom Tracker's Exclusion column once uploaded."
        "</span>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────
# MAIN PANEL
# ─────────────────────────────────────────────

left_col, right_col = st.columns([1, 1], gap="large")

# ── LEFT: File Uploads ──
with left_col:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    section("📂 File Uploads")

    st.markdown(f"**{marketplace} SC Report** <span style='color:red'>*</span>", unsafe_allow_html=True)
    sc_file = st.file_uploader(
        "SC Report", type=["xlsx", "xls", "csv", "zip"], key="sc_report", label_visibility="collapsed"
    )

    st.markdown("**Ecom Tracker** <span style='color:red'>*</span>", unsafe_allow_html=True)
    ecom_file = st.file_uploader(
        "Ecom Tracker", type=["xlsx", "xls", "csv", "zip"], key="ecom_tracker", label_visibility="collapsed"
    )

    st.markdown("**Content File** <span style='color:red'>*</span>", unsafe_allow_html=True)
    content_file = st.file_uploader(
        "Content File", type=["xlsx", "xls", "csv", "zip"], key="content_file", label_visibility="collapsed"
    )

    st.markdown(
        f"**AM Exclusion Sheet** "
        f"<span style='color:#888; font-size:0.8rem;'>(Optional · {region} VC Exclusions tab)</span>",
        unsafe_allow_html=True,
    )
    am_excl_file = st.file_uploader(
        "AM Exclusion", type=["xlsx", "xls", "csv", "zip"], key="am_excl", label_visibility="collapsed"
    )
    st.markdown(
        "<span style='font-size:0.75rem; color:#888;'>"
        "📦 Any upload above can be a .zip — multiple files inside are extracted and "
        "consolidated automatically."
        "</span>",
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # Auto-load exclusion remarks + column list when Ecom Tracker uploaded
    st.session_state["ecom_tracker_error"] = None
    if ecom_file is not None:
        try:
            _ecom_preview = load_consolidated(ecom_file, load_ecom_tracker, region=region)
            _ecom_preview.columns = [str(c).strip() for c in _ecom_preview.columns]
            st.session_state["ecom_columns"] = list(_ecom_preview.columns)
            from processor import parse_price_tier_ref as _ptr, get_exclusion_remarks as _ger, find_col as _fc
            try:
                _tier = _ptr(price_tier_ref.strip(), _ecom_preview)
                _excl_col = _tier["exclusion_col"]
                _remarks = _ger(_ecom_preview, _excl_col)
                if _remarks != st.session_state.excl_remarks:
                    st.session_state.excl_remarks = _remarks
            except Exception:
                pass
            # Auto-detect a sensible default for the Launch Date column dropdown
            if "launch_date_col_default" not in st.session_state:
                _auto_ld = _fc(
                    _ecom_preview,
                    ["Launch Dates", "Launch Date", "LaunchDate", "Launch date",
                     "LAUNCH DATE", "LAUNCH DATES", "Ecom Launch Date"],
                )
                st.session_state["launch_date_col_default"] = _auto_ld
        except EcomTrackerSheetNotFoundError as e:
            # Clear validation error: required region tab is missing from the workbook.
            st.session_state["ecom_tracker_error"] = str(e)
            st.session_state["ecom_columns"] = []
        except Exception:
            # Non-blocking: other preview issues (e.g. price tier parse) don't stop the app,
            # they'll surface again with full detail when Process is clicked.
            pass
    else:
        st.session_state["ecom_columns"] = st.session_state.get("ecom_columns", [])

    if st.session_state.get("ecom_tracker_error"):
        st.error(f"❌ Ecom Tracker: {st.session_state['ecom_tracker_error']}")

    # ── Launch Date Column Selection (rendered in sidebar, populated dynamically) ──
    with st.sidebar:
        st.markdown("---")
        section("Launch Date Column Selection")
        _ecom_cols_opts = st.session_state.get("ecom_columns", [])
        if _ecom_cols_opts:
            _default_ld = st.session_state.get("launch_date_col_default")
            _default_idx = _ecom_cols_opts.index(_default_ld) if _default_ld in _ecom_cols_opts else 0
            launch_date_col_selected = st.selectbox(
                "Launch Date Column",
                options=_ecom_cols_opts,
                index=_default_idx,
                key="launch_date_col_select",
                help=(
                    "Select the Ecom Tracker column to use for Launch Date validation. "
                    "\"NA\" values in this column are treated as eligible."
                ),
            )
        else:
            st.markdown(
                "<span style='font-size:0.8rem; color:#856404;'>"
                "⚠️ Upload the Ecom Tracker to populate this dropdown."
                "</span>",
                unsafe_allow_html=True,
            )
            launch_date_col_selected = None

    # Upload Status
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    section("📋 Upload Status")
    for fname, fobj, req in [
        (f"{marketplace} SC Report", sc_file, True),
        ("Ecom Tracker", ecom_file, True),
        ("Content File", content_file, True),
        ("AM Exclusion", am_excl_file, False),
    ]:
        if fobj:
            zip_tag = " 📦 (zip — will be extracted & consolidated)" if is_zip_upload(fobj) else ""
            st.markdown(
                f'✅ <b>{fname}</b> — <span style="color:#1F6B3E; font-size:0.85rem;">{fobj.name}{zip_tag}</span>',
                unsafe_allow_html=True,
            )
        elif req:
            st.markdown(
                f'⚠️ <b>{fname}</b> — <span style="color:#856404; font-size:0.8rem;">Required</span>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f'⬜ <b>{fname}</b> — <span style="color:#888; font-size:0.8rem;">Not uploaded (optional)</span>',
                unsafe_allow_html=True,
            )

    if st.session_state.excl_remarks:
        st.markdown(
            f"<span style='font-size:0.8rem; color:#1F6B3E;'>"
            f"✅ {len(st.session_state.excl_remarks)} unique Exclusion Remarks loaded from Ecom Tracker"
            f"</span>",
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)


# ── RIGHT: Voucher Configurations ──
with right_col:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    section("🏷️ Voucher Configurations")

    st.markdown(
        "<span style='font-size:0.82rem; color:#555;'>"
        "Each voucher is evaluated <b>independently</b>. All filters (Launch Date, Ecom Status, "
        "RRP, SRP, Inclusion Keywords, AM Exclusion) run separately per voucher. "
        "Eligible SKUs are marked <b>Yes-Eligible</b>; ineligible SKUs are left blank."
        "</span>",
        unsafe_allow_html=True,
    )

    if st.session_state.excl_remarks:
        st.markdown(
            "<span style='font-size:0.8rem; color:#2D4A6B;'>"
            "💡 Inclusion Keywords are loaded from the Ecom Tracker's Exclusion Remarks column. "
            "Select one or more from the dropdown per voucher."
            "</span>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            "<span style='font-size:0.8rem; color:#856404;'>"
            "⚠️ Upload the Ecom Tracker to enable the Inclusion Keywords dropdown."
            "</span>",
            unsafe_allow_html=True,
        )

    st.markdown("")

    configs_to_remove = []
    for i, vc in enumerate(st.session_state.voucher_configs):
        st.markdown('<div class="vc-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="vc-label">Voucher {i + 1}</div>', unsafe_allow_html=True)

        vc_col1, vc_col_del = st.columns([11, 1])
        with vc_col1:
            st.session_state.voucher_configs[i]["name"] = st.text_input(
                "Voucher Name",
                value=vc.get("name", ""),
                key=f"vc_name_{i}",
                placeholder="e.g. 10% VC",
                help=(
                    "Used as the output column header. "
                    "If the name contains a % (e.g. '10% VC'), AM exclusion cascade logic applies."
                ),
            )
        with vc_col_del:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("✕", key=f"del_vc_{i}", help="Remove this voucher"):
                configs_to_remove.append(i)

        # Inclusion Keywords — multiselect dropdown from Ecom Tracker values
        current_kws = vc.get("keywords", [])
        if not isinstance(current_kws, list):
            current_kws = [kw.strip() for kw in str(current_kws).split(",") if kw.strip()]

        # Only keep currently selected values that still exist in the options
        options = st.session_state.excl_remarks
        valid_defaults = [kw for kw in current_kws if kw in options]

        if options:
            selected = st.multiselect(
                "Inclusion Keywords (Exclusion Remarks)",
                options=options,
                default=valid_defaults,
                key=f"vc_kw_{i}",
                help=(
                    "Select one or more Exclusion Remarks values from the Ecom Tracker. "
                    "A SKU must exactly match at least one selected keyword to be eligible. "
                    "Leave empty to skip keyword filtering for this voucher."
                ),
            )
        else:
            # Fallback text input when Ecom Tracker not yet loaded
            raw_text = st.text_input(
                "Inclusion Keywords (comma-separated — upload Ecom Tracker for dropdown)",
                value=", ".join(current_kws),
                key=f"vc_kw_{i}",
                placeholder="e.g. OPEN FOR ALL, OPEN FOR ALL (10days max)",
            )
            selected = [kw.strip() for kw in raw_text.split(",") if kw.strip()]

        st.session_state.voucher_configs[i]["keywords"] = selected
        st.markdown("</div>", unsafe_allow_html=True)

    for idx in sorted(configs_to_remove, reverse=True):
        st.session_state.voucher_configs.pop(idx)
    if configs_to_remove:
        st.rerun()

    if st.button("＋  Add Voucher", key="add_vc"):
        st.session_state.voucher_configs.append({"name": "", "keywords": []})
        st.rerun()

    valid_vcs = [vc for vc in st.session_state.voucher_configs if vc.get("name", "").strip()]
    st.markdown(
        f"<span style='font-size:0.8rem; color:#555;'>"
        f"{len(valid_vcs)} voucher(s) configured — each evaluated independently"
        f"</span>",
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # Run Summary
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    section("📌 Run Summary")
    c1s, c2s = st.columns(2)
    c1s.markdown(f"**Marketplace:** {marketplace}")
    c2s.markdown(f"**Region:** {region}")
    c1s.markdown(f"**Price Tier:** `{price_tier_ref}` → RRP | SRP | DISC% | Excl.")
    c2s.markdown(f"**Cutoff Date:** {cutoff_date.strftime('%d-%m-%Y')}")
    if valid_vcs:
        vc_names_str = ", ".join(f"`{v['name']}`" for v in valid_vcs)
        st.markdown(f"**Vouchers:** {vc_names_str}")
    st.markdown("</div>", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# PROCESS BUTTON
# ─────────────────────────────────────────────

st.markdown("---")
btn_col, _ = st.columns([0.3, 0.7])
with btn_col:
    process_btn = st.button("▶  Process Voucher Eligibility", key="process")

if process_btn:
    errors = []
    if not sc_file:         errors.append(f"{marketplace} SC Report is required.")
    if not ecom_file:       errors.append("Ecom Tracker is required.")
    if not content_file:    errors.append("Content File is required.")
    if not price_tier_ref.strip(): errors.append("Price Tier Reference is required.")
    if not valid_vcs:       errors.append("At least one named voucher configuration is required.")

    if errors:
        for e in errors:
            st.error(f"❌ {e}")
    else:
        with st.spinner("Processing voucher eligibility — please wait…"):
            try:
                if marketplace == "Shopee":
                    out_df, warnings, pid_flag_cols, dynamic_cols = process_shopee_voucher_eligibility(
                        region=region,
                        sc_report_file=sc_file,
                        ecom_file=ecom_file,
                        content_file=content_file,
                        am_excl_file=am_excl_file,
                        price_tier_ref=price_tier_ref.strip(),
                        cutoff_date=cutoff_date,
                        voucher_configs=valid_vcs,
                        launch_date_col_override=launch_date_col_selected,
                    )
                    excel_bytes = generate_shopee_excel_output(out_df, valid_vcs, pid_flag_cols, dynamic_cols)
                    st.session_state["pid_flag_cols"] = pid_flag_cols
                    st.session_state["dynamic_pid_cols"] = dynamic_cols
                else:
                    out_df, warnings = process_voucher_eligibility(
                        marketplace=marketplace,
                        region=region,
                        sc_report_file=sc_file,
                        ecom_file=ecom_file,
                        content_file=content_file,
                        am_excl_file=am_excl_file,
                        price_tier_ref=price_tier_ref.strip(),
                        cutoff_date=cutoff_date,
                        voucher_configs=valid_vcs,
                        launch_date_col_override=launch_date_col_selected,
                    )
                    excel_bytes = generate_excel_output(out_df, marketplace, valid_vcs)
                    st.session_state["pid_flag_cols"] = []
                    st.session_state["dynamic_pid_cols"] = []

                st.session_state["output_excel"]      = excel_bytes
                st.session_state["output_df"]         = out_df
                st.session_state["output_warnings"]   = warnings
                st.session_state["last_marketplace"]  = marketplace
                st.session_state["last_valid_vcs"]    = valid_vcs
                st.success("✅ Processing complete!")
            except EcomTrackerSheetNotFoundError as ex:
                st.error(f"❌ Ecom Tracker: {ex}")
            except Exception as ex:
                st.error(f"❌ Processing failed: {ex}")
                with st.expander("Error details"):
                    st.code(traceback.format_exc())


# ─────────────────────────────────────────────
# RESULTS
# ─────────────────────────────────────────────

if "output_df" in st.session_state:
    st.markdown("---")
    st.markdown("## 📊 Results")

    out_df     = st.session_state["output_df"]
    warnings   = st.session_state.get("output_warnings", [])
    mp         = st.session_state.get("last_marketplace", marketplace)
    result_vcs = st.session_state.get("last_valid_vcs", valid_vcs)

    info_msgs = [w for w in warnings if w.startswith("ℹ️")]
    warn_msgs = [w for w in warnings if not w.startswith("ℹ️")]

    if info_msgs:
        with st.expander("ℹ️ Column Mapping Info", expanded=True):
            for m in info_msgs:
                st.info(m.lstrip("ℹ️").strip())

    if warn_msgs:
        with st.expander(f"⚠️ {len(warn_msgs)} Warning(s)", expanded=True):
            for w in warn_msgs:
                st.warning(w)

    # KPIs
    vc_names   = [vc["name"] for vc in result_vcs if vc.get("name")]
    total_skus = len(out_df)
    ecom_active = (
        out_df["Ecom Status"].str.strip().str.upper().eq("YES").sum()
        if "Ecom Status" in out_df.columns else 0
    )

    kpi_cols = st.columns(min(len(vc_names) + 2, 6))
    kpi_cols[0].metric("Total SKUs", f"{total_skus:,}")
    kpi_cols[1].metric("Ecom Active", f"{ecom_active:,}")
    for i, vc_name in enumerate(vc_names[:4]):
        if vc_name in out_df.columns:
            count = (out_df[vc_name].str.strip().str.lower() == "yes-eligible").sum()
            kpi_cols[i + 2].metric(vc_name, f"{count:,} eligible")

    st.markdown("")

    # Preview
    pid_flag_cols   = st.session_state.get("pid_flag_cols", [])
    dynamic_pid_cols = st.session_state.get("dynamic_pid_cols", [])
    if mp == "Shopee":
        preview_col_order = WORKING_COLS + ["Product ID"] + pid_flag_cols + dynamic_pid_cols + vc_names
    else:
        preview_col_order = WORKING_COLS + vc_names
    preview_cols = [c for c in preview_col_order if c in out_df.columns]
    st.markdown("**Preview — Working Columns & Voucher Results**")
    st.dataframe(out_df[preview_cols].head(100), use_container_width=True, height=360)
    if len(out_df) > 100:
        st.caption(f"Showing first 100 of {len(out_df):,} rows. Download for full data.")

    st.markdown("")
    dl_col, exp_col, _ = st.columns([0.3, 0.3, 0.4])
    with dl_col:
        st.download_button(
            label=f"⬇️  Download Working Sheet ({mp})",
            data=st.session_state["output_excel"],
            file_name=f"Voucher_Eligibility_Output_{mp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_btn",
        )
    with exp_col:
        if st.button("📤  Export - Eligible List", key="export_eligible_btn"):
            try:
                exp_region = st.session_state.get("region", region)
                fname, fdata, is_zip = generate_eligible_export(
                    out_df, mp, exp_region, result_vcs
                )
                st.session_state["eligible_export_fname"] = fname
                st.session_state["eligible_export_data"]  = fdata
                st.session_state["eligible_export_is_zip"] = is_zip
            except Exception as ex:
                st.error(f"❌ Export failed: {ex}")

        if "eligible_export_data" in st.session_state:
            _mime = (
                "application/zip"
                if st.session_state.get("eligible_export_is_zip")
                else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.download_button(
                label=f"⬇️  Download {st.session_state['eligible_export_fname']}",
                data=st.session_state["eligible_export_data"],
                file_name=st.session_state["eligible_export_fname"],
                mime=_mime,
                key="download_eligible_export_btn",
            )
