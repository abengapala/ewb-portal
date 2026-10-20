"""EWB Endorsement Automation Module
Encapsulates Worklist vs. DataGrid matching, IndexOne preparation, and final CMS upload generation.
"""
import copy
import io
import zipfile
import pandas as pd
import streamlit as st
import endo_automation as E

def render_endo_tab():
    st.markdown("""
    <style>
      .step-header {
          background: linear-gradient(90deg, #1F4E78 0%, #2E75B6 100%);
          color: white; border-radius: 8px; padding: 10px 16px;
          font-size: 1.05rem; font-weight: 700; margin-bottom: 8px;
      }
    </style>
    """, unsafe_allow_html=True)

    st.subheader("📑 EWB Endorsement Automation")
    st.caption(
        "BAU key = ACCOUNT NO  |  FFA / 150 DPD FFA / SPM REPO AI key = CHASSIS  |  "
        "Priority: BAU > FFA > 150 DPD > SPM REPO AI"
    )

    # -- Step 1 Inputs --
    st.markdown('<div class="step-header">Step 1: Upload Input Workbooks</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        wl = st.file_uploader("1. Bank worklist (EWB_x_SPM_WL xlsx)", type=["xlsx"], key="endo_wl")
        wl_pw = st.text_input("Worklist password (if any)", type="password", key="endo_wl_pw")
    with c2:
        dg = st.file_uploader("2. DataGrid / CMS export (.xlsx)", type=["xlsx"], key="endo_dg")
    with c3:
        tpl = st.file_uploader("3. Template (optional - copies exact upload headers)", type=["xlsx"], key="endo_tpl")
        st.write("")
        go = st.button("🚀 Run Endorsement Matching", type="primary", disabled=not (wl and dg), key="endo_run_btn")

    # -- Step 1 execution --
    if go:
        try:
            with st.spinner("Matching worklist against DataGrid..."):
                _wl_raw = E.read_wl(io.BytesIO(wl.getvalue()), wl_pw or None)
                st.session_state["wl_cols"] = {sh: list(df.columns) for sh, df in _wl_raw.items()}
                res = E.run(wl.getvalue(), dg.getvalue(), wl_pw or None, None,
                            tpl.getvalue() if tpl else None)
            st.session_state["res"] = res
            st.session_state.pop("res_with_codes", None)
            st.success("Matching complete! See results below.")
        except Exception as ex:
            st.error(f"{type(ex).__name__}: {ex}")

    res = st.session_state.get("res")

    if res:
        st.divider()
        st.subheader("📊 Step 1 Summary Metrics")
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("Worklist Accounts", f"{res['_counts'].get('wl_accounts', 0):,}")
        m2.metric("DataGrid Rows", f"{res['_counts'].get('datagrid_rows', 0):,}")
        m3.metric("New Endorsements", f"{res['_counts'].get('new_endo', 0):,}")
        m4.metric("Active (Keep)", f"{res['_counts'].get('active', 0):,}")
        m5.metric("Pullout Rows", f"{res['_counts'].get('pullout', 0):,}")
        m6.metric("Key Changes", f"{res['_counts'].get('key_changes', 0):,}")

        buf = io.BytesIO()
        E.write_workbook(res, buf)

        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("EWB_ENDORSEMENT_RESULT.xlsx", buf.getvalue())
            for n in ("REGULAR CMS UPLOAD", "FFA CMS UPLOAD"):
                z.writestr(
                    n.replace(" ", "_") + ".csv",
                    res[n].to_csv(index=False, encoding="utf-8-sig"),
                )

        st.subheader("📥 Initial CMS Uploads (Before IndexOne)")
        cms1, cms2 = st.columns(2)
        bau_rows = len(res["REGULAR CMS UPLOAD"])
        ffa_rows = len(res["FFA CMS UPLOAD"])
        bau_cms_csv = res["REGULAR CMS UPLOAD"].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        ffa_cms_csv = res["FFA CMS UPLOAD"].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

        cms1.download_button(
            f"📥 Download CMS Endo Upload BAU ({bau_rows} rows)",
            bau_cms_csv,
            "CMS_ENDO_UPLOAD_BAU.csv",
            "text/csv",
            use_container_width=True,
            key="dl_bau_cms_init"
        )
        cms2.download_button(
            f"📥 Download CMS Endo Upload FFA ({ffa_rows} rows)",
            ffa_cms_csv,
            "CMS_ENDO_UPLOAD_FFA.csv",
            "text/csv",
            use_container_width=True,
            key="dl_ffa_cms_init"
        )

        d1, d2, d3 = st.columns(3)
        pullout_df = res.get("PULLOUT UPLOAD (draft)", pd.DataFrame())
        pullout_count = len(pullout_df)
        pullout_csv = pullout_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

        d1.download_button(
            "Download results workbook (.xlsx)",
            buf.getvalue(),
            "EWB_ENDORSEMENT_RESULT.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_res_wb"
        )
        d2.download_button(
            "Download everything (.zip)",
            zbuf.getvalue(),
            "EWB_ENDORSEMENT_RESULT.zip",
            key="dl_res_zip"
        )
        d3.download_button(
            f"Download PULLOUT UPLOAD ({pullout_count} rows)",
            pullout_csv,
            "PULLOUT_UPLOAD.csv",
            "text/csv",
            key="dl_pullout_csv"
        )

        tabs = [k for k in res if not k.startswith("_") and len(res[k].columns)]
        for tab, name in zip(st.tabs(tabs), tabs):
            with tab:
                st.write(f"**{len(res[name])}** rows")
                st.dataframe(res[name].head(300).astype(str).replace("None", "").replace("nan", ""), use_container_width=True)

        st.divider()

        # -- Step 1b - IndexOne input download --
        st.markdown(
            '<div class="step-header">Step 1b: Download IndexOne Input</div>',
            unsafe_allow_html=True,
        )
        st.info(
            "Upload these files to IndexOne to get the UID Codes (Ch Codes) assigned.\n\n"
            "- **BAU IndexOne Input** - keyed by Account Number\n"
            "- **FFA IndexOne Input** - keyed by Chassis"
        )

        idx_buf = io.BytesIO()
        E.write_indexone_input(res, idx_buf)
        idx_inputs = E.make_indexone_input(res)
        bau_csv = idx_inputs["BAU IndexOne Input"].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        ffa_csv = idx_inputs["FFA IndexOne Input"].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

        ib1, ib2, ib3 = st.columns(3)
        ib1.download_button(
            "Download IndexOne Input (.xlsx)",
            idx_buf.getvalue(),
            "EWB_INDEXONE_INPUT.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_idx_wb"
        )
        ib2.download_button(
            "Download BAU IndexOne Input (.csv)",
            bau_csv,
            "BAU_INDEXONE_INPUT.csv",
            "text/csv",
            key="dl_idx_bau"
        )
        ib3.download_button(
            "Download FFA IndexOne Input (.csv)",
            ffa_csv,
            "FFA_INDEXONE_INPUT.csv",
            "text/csv",
            key="dl_idx_ffa"
        )

        st.divider()

        # -- Step 2 - Upload IndexOne Mapped output -> fill Ch Codes --
        st.markdown(
            '<div class="step-header">Step 2: Fill Ch Codes from IndexOne Mapped Output</div>',
            unsafe_allow_html=True,
        )
        col_bau, col_ffa = st.columns(2)
        with col_bau:
            mapped_bau = st.file_uploader("BAU IndexOne Mapped (.xlsx)", type=["xlsx"], key="mapped_bau_tab")
        with col_ffa:
            mapped_ffa = st.file_uploader("FFA IndexOne Mapped (.xlsx)", type=["xlsx"], key="mapped_ffa_tab")

        fill_btn = st.button(
            "Apply Ch Codes to CMS Uploads",
            type="primary",
            disabled=not (mapped_bau or mapped_ffa),
            key="fill_btn_tab",
        )

        if fill_btn and (mapped_bau or mapped_ffa):
            try:
                with st.spinner("Reading Mapped output and filling Ch Codes..."):
                    res_coded = copy.deepcopy({k: v for k, v in res.items()})
                    for mf in [mapped_bau, mapped_ffa]:
                        if mf:
                            res_coded = E.apply_chcodes(res_coded, mf.getvalue())
                st.session_state["res_with_codes"] = res_coded
                st.success("Successfully applied Ch Codes!")
            except Exception as ex:
                st.error(f"{type(ex).__name__}: {ex}")

        res_coded = st.session_state.get("res_with_codes")

        if res_coded:
            stats = res_coded.get("_chcode_stats", {})
            matched = stats.get("matched", 0)
            unmatched = stats.get("unmatched", 0)
            total = matched + unmatched

            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("Total CMS rows", total)
            mc2.metric("Ch Codes filled", matched)
            mc3.metric("Not matched", unmatched)

            final_buf = io.BytesIO()
            E.write_workbook(res_coded, final_buf)

            final_zbuf = io.BytesIO()
            with zipfile.ZipFile(final_zbuf, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("EWB_ENDORSEMENT_FINAL.xlsx", final_buf.getvalue())
                for n in ("REGULAR CMS UPLOAD", "FFA CMS UPLOAD"):
                    z.writestr(
                        n.replace(" ", "_") + "_FINAL.csv",
                        res_coded[n].to_csv(index=False, encoding="utf-8-sig"),
                    )

            st.subheader("🎯 Final CMS Endo Upload Files (with Ch Codes)")
            fcms1, fcms2 = st.columns(2)
            final_bau_csv = res_coded["REGULAR CMS UPLOAD"].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            final_ffa_csv = res_coded["FFA CMS UPLOAD"].to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

            fcms1.download_button(
                f"📥 CMS Endo Upload BAU — FINAL ({len(res_coded['REGULAR CMS UPLOAD'])} rows)",
                final_bau_csv,
                "CMS_ENDO_UPLOAD_BAU_FINAL.csv",
                "text/csv",
                use_container_width=True,
                key="dl_final_bau_csv"
            )
            fcms2.download_button(
                f"📥 CMS Endo Upload FFA — FINAL ({len(res_coded['FFA CMS UPLOAD'])} rows)",
                final_ffa_csv,
                "CMS_ENDO_UPLOAD_FFA_FINAL.csv",
                "text/csv",
                use_container_width=True,
                key="dl_final_ffa_csv"
            )

            fd1, fd2 = st.columns(2)
            fd1.download_button(
                "Download FINAL workbook (.xlsx)",
                final_buf.getvalue(),
                "EWB_ENDORSEMENT_FINAL.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_final_wb"
            )
            fd2.download_button(
                "Download FINAL everything (.zip)",
                final_zbuf.getvalue(),
                "EWB_ENDORSEMENT_FINAL.zip",
                key="dl_final_zip"
            )
