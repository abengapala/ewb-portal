"""EWB Endorsement Automation Module (Streamlit tab)

Flow
  Step 1a  Upload worklist + DataGrid  ->  headers are scanned and auto-fixed (shown on screen, overridable)
  Step 1   Run matching                ->  CMS uploads (no Ch Code yet) + PULLOUT UPLOAD
  Step 1b  Download IndexOne input
  Step 2   Upload IndexOne Mapped      ->  Ch Codes filled -> FINAL uploads
"""
import copy
import io
import zipfile

import pandas as pd
import streamlit as st

import endo_automation as E

_CSS = """
<style>
  .step-header {
      background: linear-gradient(90deg, #1F4E78 0%, #2E75B6 100%);
      color: white; border-radius: 8px; padding: 10px 16px;
      font-size: 1.05rem; font-weight: 700; margin-bottom: 8px;
  }
</style>
"""
_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ----------------------------------------------------------------- helpers
def _step(title):
    st.markdown(f'<div class="step-header">{title}</div>', unsafe_allow_html=True)


def _csv(df):
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def _pack(res, final=False):
    """Build every download once (not on each Streamlit rerun)."""
    tag = "FINAL" if final else "RESULT"
    sfx = "_FINAL" if final else ""
    wb = io.BytesIO()
    E.write_workbook(res, wb)
    files = {
        "wb": wb.getvalue(),
        "bau": _csv(res["REGULAR CMS UPLOAD"]),
        "ffa": _csv(res["FFA CMS UPLOAD"]),
        "pull": _csv(res["PULLOUT UPLOAD"]),
    }
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"EWB_ENDORSEMENT_{tag}.xlsx", files["wb"])
        z.writestr(f"REGULAR_CMS_UPLOAD{sfx}.csv", files["bau"])
        z.writestr(f"FFA_CMS_UPLOAD{sfx}.csv", files["ffa"])
        z.writestr("PULLOUT_UPLOAD.csv", files["pull"])
    files["zip"] = zbuf.getvalue()
    if not final:
        ib = io.BytesIO()
        E.write_indexone_input(res, ib)
        idx = E.make_indexone_input(res)
        files["idx_wb"] = ib.getvalue()
        files["idx_bau"] = _csv(idx["BAU IndexOne Input"])
        files["idx_ffa"] = _csv(idx["FFA IndexOne Input"])
    return files


def _overrides():
    """Manual column choices the user made in the header-scan expanders."""
    wl_ov = {}
    for role in E.SHEET_PRIORITY:
        for canon in ("ACCTNO_2", "CHASSIS_SERIAL_NO"):
            v = st.session_state.get(f"ov_wl_{role}_{canon}")
            if v and v != "(auto)":
                wl_ov.setdefault(role, {})[canon] = v
    v = st.session_state.get("ov_dg_account")
    dg_ov = {"Account No.": v} if v and v != "(auto)" else {}
    return wl_ov, dg_ov


def _get_scan(wl, dg, wl_pw):
    """Decrypt + scan the two files; cached until a file, password or manual override changes."""
    fp = (wl.name, wl.size, wl_pw or "", dg.name, dg.size)
    if st.session_state.get("endo_fp") != fp:                      # new upload -> forget everything
        for k in [k for k in st.session_state if str(k).startswith("ov_")]:
            del st.session_state[k]
        for k in ("endo_scan", "res", "res_with_codes", "endo_dl", "endo_dl_final"):
            st.session_state.pop(k, None)
        st.session_state["endo_fp"] = fp
    wl_ov, dg_ov = _overrides()
    key = repr((wl_ov, dg_ov))
    sc = st.session_state.get("endo_scan")
    if sc is None or sc["key"] != key:
        if sc:
            wl_bytes, dg_bytes = sc["wl_bytes"], sc["dg_bytes"]
        else:
            wl_bytes = E.decrypt_bytes(wl.getvalue(), wl_pw or None)
            dg_bytes = E.decrypt_bytes(dg.getvalue(), wl_pw or None)   # DataGrid: same password if it has one
        with st.spinner("Scanning headers..."):
            wl_pre = E.scan_wl(wl_bytes, None, wl_ov)
            dg_pre = E.scan_dg(dg_bytes, None, dg_ov)
        sc = dict(key=key, wl_bytes=wl_bytes, dg_bytes=dg_bytes, wl=wl_pre, dg=dg_pre)
        st.session_state["endo_scan"] = sc
        st.session_state.pop("res", None)                          # results belong to the old mapping
        st.session_state.pop("res_with_codes", None)
    return sc


def _col_used(info, canon):
    return next((a for a, c, _ in info["mapped"] if c == canon), None)


def _show_scan(sc):
    """Show what the header scan fixed. Returns True if the run may continue."""
    _, rep = sc["wl"]
    _, drep = sc["dg"]
    _step("Step 1a: Header Scan - columns checked and auto-fixed")

    errors = rep["errors"] + drep["errors"]
    for e in errors:
        st.error(e)
    for w in rep["warnings"]:
        st.warning(w)
    if rep["ignored"]:
        st.caption("Sheets not used: " + ", ".join(rep["ignored"]))

    n_fixed = sum(1 for s in rep["sheets"] for a, c, _ in s["mapped"] if a != c) \
        + sum(1 for a, c, _ in drep["mapped"] if a != c)
    m1, m2, m3 = st.columns(3)
    m1.metric("Worklist sheets recognised", len(rep["sheets"]))
    m2.metric("Headers auto-fixed", n_fixed)
    m3.metric("Problems", len(errors))

    for s in rep["sheets"]:
        renamed = [(a, c, h) for a, c, h in s["mapped"] if a != c]
        a_used, c_used = _col_used(s, "ACCTNO_2"), _col_used(s, "CHASSIS_SERIAL_NO")
        label = (f"{s['sheet']}  →  {s['role']}   |   header on row {s['header_row']}   |   "
                 f"{s['rows']:,} rows   |   {len(renamed)} headers fixed")
        with st.expander(label, expanded=bool(s["missing"] and s["role"] == "BAU" and not a_used)):
            st.caption(f"Account no. column: **{a_used or 'NOT FOUND'}**   |   Chassis column: **{c_used or 'NOT FOUND'}**")
            if renamed:
                st.dataframe(pd.DataFrame(renamed, columns=["Your header", "Fixed to", "How"]),
                             hide_index=True, use_container_width=True)
            else:
                st.caption("Every header already matched - nothing needed fixing.")
            if s["unmapped"]:
                st.caption("Not used by the automation: " + ", ".join(s["unmapped"][:40])
                           + (" ..." if len(s["unmapped"]) > 40 else ""))
            opts = ["(auto)"] + s["original_columns"]
            c1, c2 = st.columns(2)
            c1.selectbox("Force account no. column", opts, key=f"ov_wl_{s['role']}_ACCTNO_2")
            c2.selectbox("Force chassis column", opts, key=f"ov_wl_{s['role']}_CHASSIS_SERIAL_NO")

    dg_renamed = [(a, c, h) for a, c, h in drep["mapped"] if a != c]
    with st.expander(f"DataGrid  |  header on row {drep['header_row']}  |  {drep['rows']:,} rows  |  "
                     f"{len(dg_renamed)} headers fixed", expanded=bool(drep["errors"])):
        st.caption(f"Account No. column: **{_col_used(drep, 'Account No.') or 'NOT FOUND'}**")
        if dg_renamed:
            st.dataframe(pd.DataFrame(dg_renamed, columns=["Your header", "Fixed to", "How"]),
                         hide_index=True, use_container_width=True)
        st.selectbox("Force DataGrid account column", ["(auto)"] + drep["original_columns"], key="ov_dg_account")
    return not errors


# ----------------------------------------------------------------- tab
def render_endo_tab():
    st.markdown(_CSS, unsafe_allow_html=True)
    st.subheader("📑 EWB Endorsement Automation")
    st.caption(
        "BAU key = ACCOUNT NO  |  FFA / 150 DPD FFA / SPM REPO AI key = CHASSIS  |  "
        "Priority: BAU > FFA > 150 DPD > SPM REPO AI"
    )

    # -- Step 1 inputs --
    _step("Step 1: Upload Input Workbooks")
    c1, c2, c3 = st.columns(3)
    with c1:
        wl = st.file_uploader("1. Bank worklist (xlsx)", type=["xlsx"], key="endo_wl")
        wl_pw = st.text_input("Worklist password (if any)", type="password", key="endo_wl_pw")
    with c2:
        dg = st.file_uploader("2. DataGrid / CMS export (.xlsx)", type=["xlsx"], key="endo_dg")
    with c3:
        tpl = st.file_uploader("3. Template (optional - copies exact upload headers)", type=["xlsx"], key="endo_tpl")

    sc, scan_ok = None, False
    if wl and dg:
        try:
            sc = _get_scan(wl, dg, wl_pw)
            scan_ok = _show_scan(sc)
        except Exception as ex:
            st.error(f"{type(ex).__name__}: {ex}")
    else:
        st.info("Upload the worklist and the DataGrid - the headers are scanned and fixed automatically.")

    go = st.button("🚀 Run Endorsement Matching", type="primary", disabled=not (sc and scan_ok), key="endo_run_btn")
    if go:
        try:
            with st.spinner("Matching worklist against DataGrid..."):
                res = E.run(sc["wl_bytes"], sc["dg_bytes"], None, None, tpl.getvalue() if tpl else None,
                            wl_pre=sc["wl"], dg_pre=sc["dg"])
                st.session_state["res"] = res
                st.session_state["endo_dl"] = _pack(res)
            st.session_state.pop("res_with_codes", None)
            st.session_state.pop("endo_dl_final", None)
            st.success("Matching complete! See results below.")
        except Exception as ex:
            st.error(f"{type(ex).__name__}: {ex}")

    res = st.session_state.get("res")
    dl = st.session_state.get("endo_dl")
    if not (res and dl):
        return

    # -- results --
    st.divider()
    st.subheader("📊 Step 1 Summary Metrics")
    cnt = res["_counts"]
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Worklist Accounts", f"{cnt.get('wl_accounts', 0):,}")
    m2.metric("DataGrid Rows", f"{cnt.get('datagrid_rows', 0):,}")
    m3.metric("New Endorsements", f"{cnt.get('new_endo', 0):,}")
    m4.metric("Active (Keep)", f"{cnt.get('active', 0):,}")
    m5.metric("Pullout Rows", f"{cnt.get('pullout', 0):,}")
    m6.metric("Key Changes", f"{cnt.get('key_changes', 0):,}")

    st.subheader("📥 Initial CMS Uploads (Before IndexOne)")
    u1, u2, u3 = st.columns(3)
    u1.download_button(f"📥 CMS Endo Upload BAU ({len(res['REGULAR CMS UPLOAD'])} rows)", dl["bau"],
                       "CMS_ENDO_UPLOAD_BAU.csv", "text/csv", use_container_width=True, key="dl_bau_cms_init")
    u2.download_button(f"📥 CMS Endo Upload FFA ({len(res['FFA CMS UPLOAD'])} rows)", dl["ffa"],
                       "CMS_ENDO_UPLOAD_FFA.csv", "text/csv", use_container_width=True, key="dl_ffa_cms_init")
    u3.download_button(f"📥 PULLOUT UPLOAD ({len(res['PULLOUT UPLOAD'])} rows)", dl["pull"],
                       "PULLOUT_UPLOAD.csv", "text/csv", use_container_width=True, key="dl_pullout_csv")
    d1, d2 = st.columns(2)
    d1.download_button("Download results workbook (.xlsx)", dl["wb"], "EWB_ENDORSEMENT_RESULT.xlsx", _XLSX, key="dl_res_wb")
    d2.download_button("Download everything (.zip)", dl["zip"], "EWB_ENDORSEMENT_RESULT.zip", key="dl_res_zip")

    tabs = [k for k in res if not k.startswith("_") and isinstance(res[k], pd.DataFrame) and len(res[k].columns)]
    for tab, name in zip(st.tabs(tabs), tabs):
        with tab:
            st.write(f"**{len(res[name])}** rows")
            st.dataframe(res[name].head(300).astype(str).replace("None", "").replace("nan", ""),
                         use_container_width=True)

    st.divider()

    # -- Step 1b - IndexOne input --
    _step("Step 1b: Download IndexOne Input")
    st.info(
        "Upload these files to IndexOne to get the UID Codes (Ch Codes) assigned.\n\n"
        "- **BAU IndexOne Input** - keyed by Account Number\n"
        "- **FFA IndexOne Input** - keyed by Chassis"
    )
    ib1, ib2, ib3 = st.columns(3)
    ib1.download_button("Download IndexOne Input (.xlsx)", dl["idx_wb"], "EWB_INDEXONE_INPUT.xlsx", _XLSX, key="dl_idx_wb")
    ib2.download_button("Download BAU IndexOne Input (.csv)", dl["idx_bau"], "BAU_INDEXONE_INPUT.csv", "text/csv", key="dl_idx_bau")
    ib3.download_button("Download FFA IndexOne Input (.csv)", dl["idx_ffa"], "FFA_INDEXONE_INPUT.csv", "text/csv", key="dl_idx_ffa")

    st.divider()

    # -- Step 2 - IndexOne Mapped -> Ch Codes --
    _step("Step 2: Fill Ch Codes from IndexOne Mapped Output")
    col_bau, col_ffa = st.columns(2)
    with col_bau:
        mapped_bau = st.file_uploader("BAU IndexOne Mapped (.xlsx)", type=["xlsx"], key="mapped_bau_tab")
    with col_ffa:
        mapped_ffa = st.file_uploader("FFA IndexOne Mapped (.xlsx)", type=["xlsx"], key="mapped_ffa_tab")

    fill_btn = st.button("Apply Ch Codes to CMS Uploads", type="primary",
                         disabled=not (mapped_bau or mapped_ffa), key="fill_btn_tab")
    if fill_btn:
        try:
            with st.spinner("Reading Mapped output and filling Ch Codes..."):
                res_coded = copy.deepcopy(res)
                for mf, target in ((mapped_bau, "BAU"), (mapped_ffa, "FFA")):   # each file fills its own upload
                    if mf:
                        res_coded = E.apply_chcodes(res_coded, mf.getvalue(), target)
                st.session_state["res_with_codes"] = res_coded
                st.session_state["endo_dl_final"] = _pack(res_coded, final=True)
            st.success("Successfully applied Ch Codes!")
        except Exception as ex:
            st.error(f"{type(ex).__name__}: {ex}")

    res_coded = st.session_state.get("res_with_codes")
    fdl = st.session_state.get("endo_dl_final")
    if not (res_coded and fdl):
        return

    stats = res_coded.get("_chcode_stats", {})
    matched, unmatched = stats.get("matched", 0), stats.get("unmatched", 0)
    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("Total CMS rows", matched + unmatched)
    mc2.metric("Ch Codes filled", matched)
    mc3.metric("Not matched", unmatched)
    um = stats.get("unmatched_rows")
    if unmatched and isinstance(um, pd.DataFrame) and len(um):
        with st.expander(f"{unmatched} rows without a Ch Code"):
            st.dataframe(um, hide_index=True, use_container_width=True)

    st.subheader("🎯 Final CMS Endo Upload Files (with Ch Codes)")
    f1, f2, f3 = st.columns(3)
    f1.download_button(f"📥 CMS Endo Upload BAU - FINAL ({len(res_coded['REGULAR CMS UPLOAD'])} rows)", fdl["bau"],
                       "CMS_ENDO_UPLOAD_BAU_FINAL.csv", "text/csv", use_container_width=True, key="dl_final_bau_csv")
    f2.download_button(f"📥 CMS Endo Upload FFA - FINAL ({len(res_coded['FFA CMS UPLOAD'])} rows)", fdl["ffa"],
                       "CMS_ENDO_UPLOAD_FFA_FINAL.csv", "text/csv", use_container_width=True, key="dl_final_ffa_csv")
    f3.download_button(f"📥 PULLOUT UPLOAD ({len(res_coded['PULLOUT UPLOAD'])} rows)", fdl["pull"],
                       "PULLOUT_UPLOAD.csv", "text/csv", use_container_width=True, key="dl_final_pullout_csv")
    g1, g2 = st.columns(2)
    g1.download_button("Download FINAL workbook (.xlsx)", fdl["wb"], "EWB_ENDORSEMENT_FINAL.xlsx", _XLSX, key="dl_final_wb")
    g2.download_button("Download FINAL everything (.zip)", fdl["zip"], "EWB_ENDORSEMENT_FINAL.zip", key="dl_final_zip")
