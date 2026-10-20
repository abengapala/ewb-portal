"""
EWB endorsement automation (S.P. Madrid) - core logic + command line.

Inputs : bank worklist (WL) xlsx (password ok), DataGrid (CMS) export xlsx, optional template xlsx.
Outputs: one results workbook + CSV copies of the two CMS uploads.

KEY RULE
  BAU                                  -> DataGrid "Account No." must be the ACCOUNT NUMBER
  FFA / 150 DPD FFA / FFA SPM REPO AI  -> DataGrid "Account No." must be the CHASSIS
  Priority when an account is on several sheets: BAU > FFA > 150 DPD FFA > FFA SPM REPO AI

INDEXONE FLOW
  Step 1 : run()                         -> CMS uploads with Ch Code = blank
  Step 1b: make_indexone_input(res)       -> dict of DataFrames to upload to IndexOne
             BAU input  : ACCOUNT_NUMBER + borrower/unit fields
             FFA input  : CHASSIS_SERIAL_NO used as ACCOUNT_NUMBER key + unit fields
  Step 2 : apply_chcodes(res, mapped_src) -> fills UID Code -> Ch Code in both CMS uploads
             Mapped "Account Number" matched against CMS "ACCOUNT_NUMBER" (normalised)
             For FFA, "Account Number" in Mapped = the chassis (IndexOne stores it that way)
"""
from __future__ import annotations
import argparse, io, os, re
from datetime import date
import pandas as pd

# ----------------------------------------------------------------- config
SHEET_PRIORITY = ["BAU", "FFA", "150 DPD FFA", "FFA SPM REPO AI"]
ACCT_COLS = ["ACCTNO_2", "ACCTNO_", "ACCTNO", "ACCOUNTNUMBER", "ACCOUNT_NUMBER"]
CHASSIS_COL = "CHASSIS_SERIAL_NO"
PLACEHOLDER_CHASSIS = {"0", "1"}          # not a real chassis
MIN_CHASSIS_LEN = 4                         # 6-digit chassis exist and are real
# PLACEMENT / CYCLE written on the FFA CMS upload, per WL sheet (CONFIRM 150 DPD / SPM REPO AI)
FFA_PLACEMENT = {"FFA": "FFA", "150 DPD FFA": "FFA", "FFA SPM REPO AI": "FFA"}
REGULAR_PLACEMENT = "REGULAR"
REPO_AI_CLIENT, REPO_AI_PRIORITY = "EA", "HIGH"
REPO_AI_ONLY_NEW_ENDO = True                # Repo AI sheets list NEW ENDO accounts only

REGULAR_UPLOAD_COLS = ["ACCOUNT_NUMBER", "ACCOUNT_NAME", "RESIDENCE_ADDRESS", "MOBILE_NO", "COMPANY_NAME",
    "OFFICE_ADDRESS", "NEW_DPD", "OB", "LAST_PAYMENT_AMOUNT", "LAST_PAYMENT_DATE", "UNIT_DETAILS",
    "PLATE_NUMBER", "CHASSIS_SERIAL_NO", "ENGINE_MOTOR_NUMBER", "FIN_PAST_DUE_AMOUNT", "EMAIL",
    "REFERENCE_PROVINCE", "Ch Code", "Agent", "PLACEMENT", "CYCLE"]
FFA_UPLOAD_COLS = ["ACCOUNT_NUMBER", "ACCOUNT_NAME", "RESIDENCE_ADDRESS", "MOBILE_NO", "COMPANY_NAME",
    "OFFICE_ADDRESS", "NEW_DPD", "OB", "LAST_PAYMENT_AMOUNT", "LAST_PAYMENT_DATE", "UNIT_DETAILS",
    "PLATE_NUMBER", "CHASSIS_SERIAL_NO", "ENGINE_MOTOR_NUMBER", "vehicle_color", "COLL_CONDUCTION\t",
    "FIN_PAST_DUE_AMOUNT", "EMAIL", "REFERENCE_PROVINCE", "Ch Code", "Agent"]
REPO_COLS = ["CH_CODE", "CRM_CODE", "PLATE", "CAR_MODEL", "ENDO_DATE", "CLIENT", "PRIORITY"]


# ----------------------------------------------------------------- helpers
def norm(v):
    """Normalise an account/chassis key: trim, upper, drop spaces/dashes/asterisks and trailing .0"""
    if v is None or (isinstance(v, float) and pd.isna(v)) or (not isinstance(v, str) and pd.isna(v)):
        return None
    s = re.sub(r"[\s\-\*]", "", str(v)).upper()
    s = re.sub(r"\.0$", "", s)
    return s or None


def valid_chassis(c):
    if not c or len(c) < MIN_CHASSIS_LEN or c in PLACEHOLDER_CHASSIS or len(set(c)) == 1:
        return None
    return c


def clean(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    # Remove characters that cause CMS upload errors (apostrophes, backticks, semicolons)
    s = re.sub(r"['`;]", "", s).strip()
    return s or None


def clean_vehicle_id(v):
    """Clean chassis or engine numbers:
    - strips illegal CMS characters (; ' `)
    - deduplicates exact repeated halves (e.g. ABCABC -> ABC)
    - handles multiple values separated by spaces or slashes (takes first valid)
    """
    s = clean(v)
    if not s:
        return None
    # Check for repeated half (e.g. KMHJWA37KBGU848417KMHJWA37KBGU848417)
    if len(s) >= 8 and len(s) % 2 == 0:
        half = len(s) // 2
        if s[:half] == s[half:]:
            s = s[:half]
    # Check for space / slash / comma / semicolon separated multiple values
    parts = re.split(r"[\s/;,]+", s)
    if len(parts) > 1:
        for p in parts:
            p_strip = p.strip()
            if len(p_strip) >= MIN_CHASSIS_LEN and p_strip not in PLACEHOLDER_CHASSIS:
                s = p_strip
                break
    return s


def open_xlsx(src, password=None):
    """Return a file-like object, decrypting if the workbook is password protected."""
    data = src if isinstance(src, (bytes, bytearray)) else (open(src, "rb").read() if isinstance(src, str) else src.read())
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":          # CDFV2 container = encrypted
        import msoffcrypto
        if not password:
            raise ValueError("The workbook is password protected - please give the password.")
        out = io.BytesIO()
        f = msoffcrypto.OfficeFile(io.BytesIO(data))
        f.load_key(password=password)
        f.decrypt(out)
        return io.BytesIO(out.getvalue())
    return io.BytesIO(data)


def read_wl(src, password=None):
    xl = pd.ExcelFile(open_xlsx(src, password))
    by_name = {s.strip().upper(): s for s in xl.sheet_names}
    wl = {}
    for sh in SHEET_PRIORITY:
        if sh.upper() in by_name:
            df = xl.parse(by_name[sh.upper()], dtype=object).dropna(how="all")
            wl[sh] = df
    missing = [s for s in ("BAU", "FFA") if s not in wl]
    if missing:
        raise ValueError(f"Worklist is missing sheet(s): {missing}. Found: {xl.sheet_names}")
    return wl


def read_datagrid(src, password=None):
    df = pd.read_excel(open_xlsx(src, password), dtype=str)
    if "Account No." not in df.columns:
        raise ValueError("DataGrid file has no 'Account No.' column.")
    df = df[df["Account No."].notna()].copy()
    df["_KEY"] = df["Account No."].map(norm)
    df = df[df["_KEY"].notna()]
    return df.reset_index(drop=True)


def template_headers(template):
    """Take the exact upload headers from the user's template when provided."""
    reg, ffa = REGULAR_UPLOAD_COLS, FFA_UPLOAD_COLS
    if template is None:
        return reg, ffa
    try:
        import openpyxl
        wb = openpyxl.load_workbook(open_xlsx(template), read_only=True)
        if "REGULAR CMS UPLOAD" in wb.sheetnames:
            r = [c for c in next(wb["REGULAR CMS UPLOAD"].iter_rows(min_row=2, max_row=2, values_only=True)) if c]
            if len(r) >= 15: reg = [str(c) for c in r]
        if "FFA CMS UPLOAD" in wb.sheetnames:
            f = [c for c in next(wb["FFA CMS UPLOAD"].iter_rows(min_row=1, max_row=1, values_only=True)) if c]
            if len(f) >= 15: ffa = [str(c) for c in f]
    except Exception:
        pass
    return reg, ffa


# ----------------------------------------------------------------- core
def build_owners(wl):
    """One record per unique account, honouring sheet priority. Returns owners, removed, dq."""
    seen, owners, removed, dq = set(), [], [], []
    for sh in SHEET_PRIORITY:
        if sh not in wl:
            continue
        df = wl[sh]
        acol = next((c for c in ACCT_COLS if c in df.columns), None)
        if acol is None:
            raise ValueError(f"No account number column found on sheet '{sh}' (looked for {ACCT_COLS}).")
        for row in df.to_dict("records"):
            a = norm(row.get(acol)); c_raw = norm(row.get(CHASSIS_COL)); c = valid_chassis(c_raw)
            rec = dict(sheet=sh, acct=a, acct_raw=clean(row.get(acol)), chassis=c,
                       chassis_raw=clean(row.get(CHASSIS_COL)), name=clean(row.get("FULLNAME")),
                       tag=clean(row.get("ENDO")) if sh == "BAU" else clean(row.get("ENDO TYPE")),
                       status=clean(row.get("STATUS")), row=row)
            if not a and not c:
                dq.append(dict(Issue="No account no. and no valid chassis - skipped", **_dq(rec))); continue
            if c_raw and not c:
                dq.append(dict(Issue="Chassis is blank/placeholder (0/1) - not usable", **_dq(rec)))
            if sh != "BAU" and not c:
                dq.append(dict(Issue="Non-BAU account without usable chassis - checked by account no. instead", **_dq(rec)))
            keys = [k for k in (a, c) if k]
            if any(k in seen for k in keys):
                rec["reason"] = ("Also on a higher-priority sheet (BAU/FFA) - counted there"
                                 if _on_other_sheet(rec, owners) else "Same account/chassis already on this sheet")
                removed.append(rec); continue
            seen.update(keys); owners.append(rec)
    return owners, removed, dq


def _on_other_sheet(rec, owners):
    ks = {rec["acct"], rec["chassis"]} - {None}
    return any(o["sheet"] != rec["sheet"] and ks & {o["acct"], o["chassis"]} for o in owners)


def _dq(r):
    return {"WL Sheet": r["sheet"], "ACCOUNT NO": r["acct_raw"], "CHASSIS": r["chassis_raw"], "FULLNAME": r["name"]}


def classify(owners, dg):
    dgrow = {r["_KEY"]: r for r in dg.to_dict("records")}
    endorse, pull, tochassis, toacct, stats = [], [], [], [], {}
    claimed = set()
    for o in owners:
        for k in (o["acct"], o["chassis"]):
            if k: claimed.add(k)
    for o in owners:
        sh = o["sheet"]; bau = sh == "BAU"
        exp = "acct" if (bau or not o["chassis"]) else "chassis"
        ra = dgrow.get(o["acct"]) if o["acct"] else None
        rc = dgrow.get(o["chassis"]) if o["chassis"] else None
        exp_row, oth_row = (ra, rc) if exp == "acct" else (rc, ra)
        expn = "ACCOUNT NO" if exp == "acct" else "CHASSIS"
        othn = "CHASSIS" if exp == "acct" else "ACCOUNT NO"
        key_raw = o["acct_raw"] if exp == "acct" else o["chassis_raw"]
        base = {"WL Sheet": sh, "ACCOUNT NO": o["acct_raw"], "CHASSIS": o["chassis_raw"], "FULLNAME": o["name"], "WL TAG": o["tag"]}

        def p(r, reason):
            pull.append({"DG No": r.get("No"), "DG Name": r.get("Name"), "Account No. (in DataGrid)": r.get("Account No."),
                         "New I.C. (CH code)": r.get("New I.C."), "Batch No.": r.get("Batch No."),
                         "Collector": r.get("Collector"), "Status Code": r.get("Status Code"), "Debtor ID": r.get("Debtor ID"),
                         "Reason": reason, "WL Sheet": sh, "WL ACCOUNT NO": o["acct_raw"], "WL FULLNAME": o["name"]})
        if exp_row is not None:
            stats[(sh, "ACTIVE - stays")] = stats.get((sh, "ACTIVE - stays"), 0) + 1
            if oth_row is not None:
                p(oth_row, f"DUPLICATE - DG has this account twice; keep the {expn} row, pull out the {othn} row")
                stats[(sh, "DUPLICATE row to pull out")] = stats.get((sh, "DUPLICATE row to pull out"), 0) + 1
        elif oth_row is not None:
            p(oth_row, f"WRONG KEY - loaded under {othn}, change to {expn}")
            e = {**base, "ENDORSE UNDER": expn, "KEY TO ENDORSE": key_raw, "REASON": f"CHANGE TO {expn}", "_o": o}
            endorse.append(e)
            (tochassis if exp == "chassis" else toacct).append(
                {**base, "Currently in DG as": othn, "DG Name": oth_row.get("Name"), "DG Batch": oth_row.get("Batch No."),
                 "Change to": expn, "KEY TO USE": key_raw})
            stats[(sh, "CHANGE KEY")] = stats.get((sh, "CHANGE KEY"), 0) + 1
        else:
            endorse.append({**base, "ENDORSE UNDER": expn, "KEY TO ENDORSE": key_raw, "REASON": "NEW ENDO", "_o": o})
            stats[(sh, "NEW ENDO")] = stats.get((sh, "NEW ENDO"), 0) + 1
    for r in dg.to_dict("records"):
        if r["_KEY"] not in claimed:
            pull.append({"DG No": r.get("No"), "DG Name": r.get("Name"), "Account No. (in DataGrid)": r.get("Account No."),
                         "New I.C. (CH code)": r.get("New I.C."), "Batch No.": r.get("Batch No."), "Collector": r.get("Collector"),
                         "Status Code": r.get("Status Code"), "Debtor ID": r.get("Debtor ID"),
                         "Reason": "PULLOUT - not in worklist", "WL Sheet": None, "WL ACCOUNT NO": None, "WL FULLNAME": None})
            stats[("DataGrid", "PULLOUT - not in worklist")] = stats.get(("DataGrid", "PULLOUT - not in worklist"), 0) + 1
    return endorse, pull, tochassis, toacct, stats


# ----------------------------------------------------------------- uploads
def _num(v):
    v = clean(v)
    if v is None or isinstance(v, (int, float)): return v
    try: return float(str(v).replace(",", ""))
    except ValueError: return v


def _dt(v):
    v = clean(v)
    if v is None: return None
    if isinstance(v, (int, float)):                       # Excel serial or 0 placeholder
        if v < 20000: return None
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))).date()
    try:
        d = pd.to_datetime(v).date()
    except Exception:
        return v
    return None if d.year < 1990 else d                  # 1970-01-01 = placeholder in the bank file


def _unit(row, sep):
    parts = [str(clean(row.get(k))).strip() for k in ("YEAR_MODEL", "ASSET_MAKE", "ASSET_MODEL") if clean(row.get(k)) is not None]
    return sep.join(p for p in parts if p and p != "0")


# Keywords that indicate the PLATE_NUMBER cell contains a text note, not a real plate
_NO_PLATE_KEYWORDS = (
    "NO CR", "NO PLATE", "NO CONDUCTION", "NOT AVAILABLE", "N/A",
    "NONE", "WITHOUT", "W/O", "NO STICKER", "NO REG", "ATTACHED",
    "AVAILABLE", "INDICATED", "PROVIDED",
)


def _plate(row):
    """Return the best plate/conduction value for a row.

    Logic:
      1. Use PLATE_NUMBER if it is a real plate (not blank, not a text note).
      2. Otherwise fall back to COLL_CONDUCTION.
      3. Return None if both are unusable.
    """
    raw = clean(row.get("PLATE_NUMBER"))
    if raw:
        upper = str(raw).strip().upper()
        is_text_note = (
            len(upper) > 15
            or any(kw in upper for kw in _NO_PLATE_KEYWORDS)
        )
        if not is_text_note:
            return raw          # genuine plate number
    # Fallback: try CS NO (conduction sticker number), then COLL_CONDUCTION
    return _cond(row.get("CS NO")) or _cond(row.get("COLL_CONDUCTION"))


def _cond(v):
    """Return the conduction sticker value, or None if blank / zero (int, float, or string).
    The worklist stores 0 or 0.0 when a unit has a plate number instead of a conduction sticker.
    """
    v = clean(v)
    if v is None:
        return None
    # Numeric zero check (covers int 0 and float 0.0)
    try:
        if float(v) == 0:
            return None
    except (ValueError, TypeError):
        pass
    s = str(v).strip()
    return None if s in ("", "0", "0.0") else s


def _cond_row(row):
    """Get conduction sticker from a worklist row.
    CS NO = actual sticker number across all sheets.
    COLL_CONDUCTION in worklist is often a 0/non-zero flag, not the number itself.
    CS/PLATE = FFA combined field — may contain CS or plate number.
    """
    return (
        _cond(row.get("CS NO")) or
        _cond(row.get("COLL_CONDUCTION")) or
        _cond(row.get("CS/PLATE"))
    )


def _color(row):
    """Get vehicle color from a worklist row.
    Tries vehicle_color first, then COLOR as fallback.
    """
    return clean(row.get("vehicle_color")) or clean(row.get("COLOR")) or None


def _date_fmt(v):
    """Format any date/datetime value as DD/MM/YYYY string (plain, no time).
    Returns None if the value is blank or unparseable.
    """
    if v is None:
        return None
    try:
        return pd.to_datetime(v).strftime("%d/%m/%Y")
    except Exception:
        s = str(v).strip()
        return s if s and s != "nan" else None



def regular_upload(endorse, cols):
    rows = []
    for e in endorse:
        o = e["_o"]
        if o["sheet"] != "BAU": continue
        r = o["row"]
        vals = {"ACCOUNT_NUMBER": o["acct_raw"], "ACCOUNT_NAME": clean(r.get("FULLNAME")),
                "RESIDENCE_ADDRESS": clean(r.get("RESIDENCE_ADDRESS")), "MOBILE_NO": clean(r.get("MOBILE_NO")),
                "COMPANY_NAME": clean(r.get("COMPANY_NAME")), "OFFICE_ADDRESS": clean(r.get("OFFICE_ADDRESS")),
                "NEW_DPD": _num(r.get("new_dpd2")), "OB": _num(r.get("ob")),
                "LAST_PAYMENT_AMOUNT": _num(r.get("LAST_PAYMENT_AMOUNT")), "LAST_PAYMENT_DATE": _date_fmt(r.get("LAST_PAYMENT_DATE")),
                "UNIT_DETAILS": clean(r.get("COLLATERAL_DESCRIPTION")), "PLATE_NUMBER": clean(_plate(r)),
                "CHASSIS_SERIAL_NO": clean_vehicle_id(r.get(CHASSIS_COL)), "ENGINE_MOTOR_NUMBER": clean_vehicle_id(r.get("ENGINE_MOTOR_NUMBER")),
                "vehicle_color": clean(_color(r)),
                "FIN_PAST_DUE_AMOUNT": _num(r.get("PAST_DUE_AMOUNT")), "EMAIL": clean(r.get("EMAIL")),
                "REFERENCE_PROVINCE": clean(r.get("REFERENCE_PROVINCE")), "Ch Code": None, "Agent": None,
                "PLACEMENT": REGULAR_PLACEMENT, "CYCLE": REGULAR_PLACEMENT}
        rows.append([vals.get(c) if c in vals else vals.get(c.strip()) for c in cols])
    return pd.DataFrame(rows, columns=cols)


def ffa_upload(endorse, cols):
    rows = []
    for e in endorse:
        o = e["_o"]
        if o["sheet"] == "BAU": continue
        r = o["row"]
        chassis_clean = clean_vehicle_id(o["chassis_raw"] if o["chassis"] else r.get(CHASSIS_COL))
        key = chassis_clean if chassis_clean else o["acct_raw"]
        # Per EWB FFA standard: demographic, DPD, payment, balance, email, and agent fields are blank
        vals = {"ACCOUNT_NUMBER": key, "ACCOUNT_NAME": "EWB AUTO", "OB": 1,
                "RESIDENCE_ADDRESS":   None,
                "MOBILE_NO":           None,
                "COMPANY_NAME":        None,
                "OFFICE_ADDRESS":      None,
                "NEW_DPD":             None,
                "LAST_PAYMENT_AMOUNT": None,
                "LAST_PAYMENT_DATE":   None,
                "FIN_PAST_DUE_AMOUNT": None,
                "EMAIL":               None,
                "REFERENCE_PROVINCE":  None,
                "UNIT_DETAILS": clean(_unit(r, " ")), "PLATE_NUMBER": clean(_plate(r)),
                "CHASSIS_SERIAL_NO": chassis_clean,
                "ENGINE_MOTOR_NUMBER": clean_vehicle_id(r.get("ENGINE_MOTOR_NUMBER")),
                "vehicle_color": clean(_color(r)),
                "COLL_CONDUCTION": clean(_cond_row(r)),
                "COLL_CONDUCTION\t": clean(_cond_row(r)),
                "Ch Code": None, "Agent": None}
        rows.append([vals.get(c) if c in vals else vals.get(c.strip()) for c in cols])
    return pd.DataFrame(rows, columns=cols)


def repo_ai(endorse):
    out = {"REPO AI BAU": [], "REPO AI FFA": [], "REPO AI SPECIAL": []}
    today = date.today().isoformat()
    for e in endorse:
        if REPO_AI_ONLY_NEW_ENDO and e["REASON"] != "NEW ENDO": continue
        o = e["_o"]; r = o["row"]
        plate = _plate(r)
        if o["sheet"] == "FFA SPM REPO AI":              # special units: also try conduction if plate still blank
            plate = plate or clean(r.get("COLL_CONDUCTION"))
        if not plate or str(plate).strip() in ("", "0"): continue   # Repo AI only lists units with a plate
        if o["sheet"] == "BAU": tab, model = "REPO AI BAU", _unit(r, "-").replace(" ", "-")
        elif o["sheet"] == "FFA SPM REPO AI": tab, model = "REPO AI SPECIAL", _unit(r, " ")
        else: tab, model = "REPO AI FFA", _unit(r, "-").replace(" ", "-")
        out[tab].append([None, None, plate, model, today, REPO_AI_CLIENT, REPO_AI_PRIORITY])
    return {k: pd.DataFrame(v, columns=REPO_COLS) for k, v in out.items()}


# ----------------------------------------------------------------- pullout upload
def make_pullout_upload(pull_df: "pd.DataFrame") -> "pd.DataFrame":
    """Build the clean PULLOUT UPLOAD from the full PULLOUT detail table.

    Covers all three pullout cases:
      1. In DataGrid but NOT in the worklist
      2. In DataGrid as CHASSIS  but now on BAU worklist -> pull chassis, re-endorse as acct no
      3. In DataGrid as ACCT NO  but now on FFA worklist -> pull acct no, re-endorse as chassis
      4. Both keys exist in DataGrid (duplicate) -> pull the wrong-key row

    Output columns: ACCOUNT_NUMBER | CH_CODE | REASON
    ACCOUNT_NUMBER = the value that is currently sitting in DataGrid (Account No. column).
    CH_CODE        = the collector's CH code for that DataGrid row (New I.C. column).
    """
    _COLS = ["ACCOUNT_NUMBER", "CH_CODE", "REASON"]
    if pull_df is None or not len(pull_df):
        return pd.DataFrame(columns=_COLS)

    rows = []
    for _, r in pull_df.iterrows():
        raw = str(r.get("Reason") or "")
        raw_up = raw.upper()

        if "NOT IN WORKLIST" in raw_up:
            reason = "NOT IN WORKLIST"
        elif "DUPLICATE" in raw_up and "CHASSIS" in raw_up:
            # keep the expected key row, pull the other
            if "PULL OUT THE CHASSIS" in raw_up or "KEEP THE ACCOUNT NO" in raw_up:
                reason = "DUPLICATE - pull out CHASSIS row (keep ACCOUNT NO row)"
            else:
                reason = "DUPLICATE - pull out ACCOUNT NO row (keep CHASSIS row)"
        elif "WRONG KEY" in raw_up and "CHASSIS" in raw_up and "ACCOUNT NO" in raw_up:
            if "CHANGE TO ACCOUNT NO" in raw_up:
                reason = "WRONG KEY - in DG as CHASSIS, re-endorse as BAU (acct no.)"
            else:
                reason = "WRONG KEY - in DG as ACCT NO, re-endorse as FFA (chassis)"
        else:
            reason = raw.split(" - ")[0] if " - " in raw else raw

        rows.append({
            "ACCOUNT_NUMBER": r.get("Account No. (in DataGrid)"),
            "CH_CODE":        r.get("New I.C. (CH code)"),
            "REASON":         reason,
        })

    return pd.DataFrame(rows, columns=_COLS)


# ----------------------------------------------------------------- orchestration
def run(wl_src, dg_src, wl_password=None, dg_password=None, template=None):
    wl = read_wl(wl_src, wl_password)
    dg = read_datagrid(dg_src, dg_password)
    owners, removed, dq = build_owners(wl)
    endorse, pull, tochassis, toacct, stats = classify(owners, dg)
    reg_cols, ffa_cols = template_headers(template)
    order = {s: i for i, s in enumerate(SHEET_PRIORITY)}
    endorse_df = pd.DataFrame([{k: v for k, v in e.items() if k != "_o"} for e in endorse])
    if len(endorse_df):
        endorse_df["_s"] = endorse_df["WL Sheet"].map(order)
        endorse_df = endorse_df.sort_values(["REASON", "_s"], kind="stable").drop(columns="_s")
    rem_df = pd.DataFrame([{"WL Sheet": r["sheet"], "ACCOUNT NO": r["acct_raw"], "CHASSIS": r["chassis_raw"],
                            "FULLNAME": r["name"], "WL TAG": r["tag"], "STATUS": r["status"], "REASON": r["reason"]} for r in removed])
    res = {
        "REGULAR CMS UPLOAD": regular_upload(endorse, reg_cols),
        "FFA CMS UPLOAD": ffa_upload(endorse, ffa_cols),
        **repo_ai(endorse),
        "ENDORSE": endorse_df,
        "PULLOUT": pd.DataFrame(pull).sort_values("Reason", kind="stable") if pull else pd.DataFrame(),
        "CHANGE_TO_CHASSIS": pd.DataFrame(tochassis), "CHANGE_TO_ACCTNO_BAU": pd.DataFrame(toacct),
        "REMOVED_FROM_SPECIAL": rem_df, "DATA_QUALITY": pd.DataFrame(dq),
    }
    res["_endorse_raw"] = endorse   # keep raw rows for IndexOne input builder
    pu = res["PULLOUT"]
    res["PULLOUT UPLOAD"] = make_pullout_upload(pu)
    summ = []
    for sh in SHEET_PRIORITY:
        if sh in wl:
            g = lambda k: stats.get((sh, k), 0)
            summ.append({"WL Sheet": sh, "WL rows": len(wl[sh]), "ACTIVE - stays": g("ACTIVE - stays"), "NEW ENDO": g("NEW ENDO"),
                         "CHANGE KEY": g("CHANGE KEY"), "DUPLICATE row to pull out": g("DUPLICATE row to pull out"),
                         "Removed (dup of higher sheet)": int((rem_df["WL Sheet"] == sh).sum()) if len(rem_df) else 0})
    res["_summary"] = pd.DataFrame(summ)
    res["_pullout_counts"] = (pu["Reason"].str.split(" - ").str[0].value_counts().rename_axis("Reason").reset_index(name="DataGrid rows")
                              if len(pu) else pd.DataFrame())
    return res


# ----------------------------------------------------------------- output
def write_workbook(res, path_or_buf):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook(); ws = wb.active; ws.title = "SUMMARY"
    hf, fill, bold = Font(name="Arial", bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78"), Font(name="Arial", bold=True)
    body = Font(name="Arial")

    def put(ws, df, r0=1):
        for j, c in enumerate(df.columns, 1):
            x = ws.cell(r0, j, c.rstrip("\t") if False else c); x.font, x.fill = hf, fill
            x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            w = max([len(str(c))] + [len(str(v)) for v in df[c].head(200).tolist()]) + 3
            ws.column_dimensions[get_column_letter(j)].width = max(10, min(w, 55))
        for i, row in enumerate(df.itertuples(index=False), r0 + 1):
            for j, v in enumerate(row, 1):
                v = None if (v is None or (not isinstance(v, str) and pd.isna(v))) else v
                ws.cell(i, j, v).font = body
        ws.freeze_panes = ws.cell(r0 + 1, 1)
    summ = res["_summary"]
    ws["A1"] = "EWB ENDORSEMENT AUTOMATION - SUMMARY"; ws["A1"].font = bold
    put(ws, summ, 3)
    r = 3 + len(summ) + 3
    ws.cell(r - 1, 1, "DataGrid rows to PULL OUT").font = bold
    if len(res["_pullout_counts"]): put(ws, res["_pullout_counts"], r)
    n = r + len(res["_pullout_counts"]) + 3
    notes = ["RULES: BAU key = ACCOUNT NO; FFA / 150 DPD FFA / FFA SPM REPO AI key = CHASSIS. Priority BAU > FFA > 150 DPD FFA > SPM REPO AI.",
             "Uploads: REGULAR CMS UPLOAD = BAU, FFA CMS UPLOAD = FFA + 150 DPD FFA + SPM REPO AI (new endo + change-key accounts).",
             "IndexOne flow: download IndexOne input files -> upload to IndexOne -> download Mapped output -> upload here to fill Ch Code.",
             "PULLOUT UPLOAD (draft) mirrors the 'EWB RECOV (PULLOUT)' layout - confirm the real pullout format before using.",
             "Check DATA_QUALITY before uploading."]
    for k, t in enumerate(notes): ws.cell(n + k, 1, t).font = body
    ws.column_dimensions["A"].width = 34
    for name, df in res.items():
        if name.startswith("_") or df is None or not len(df.columns): continue
        put(wb.create_sheet(name[:31]), df)
    wb.save(path_or_buf)


def save_all(res, outdir):
    os.makedirs(outdir, exist_ok=True)
    stamp = date.today().strftime("%Y%m%d")
    xlsx = os.path.join(outdir, f"EWB_ENDORSEMENT_RESULT_{stamp}.xlsx")
    write_workbook(res, xlsx)
    for name in ("REGULAR CMS UPLOAD", "FFA CMS UPLOAD"):
        res[name].to_csv(os.path.join(outdir, f"{name.replace(' ', '_')}_{stamp}.csv"), index=False, encoding="utf-8-sig")
    return xlsx


# ----------------------------------------------------------------- IndexOne helpers

# Columns IndexOne expects for a BAU upload (matches ewb-auto-new BAU.xlsx exactly)
_IDX_BAU_COLS = [
    "ENDORSEMENT DATE", "ENDORSING TEAM", "accountnumber", "FULLNAME", "MIDDLENAME",
    "GUARANTOR_CO_BORROWER", "RESIDENCE_ADDRESS", "RESIDENCE_PHONE_NO", "MOBILE_NO",
    "COMPANY_NAME", "OFFICE_ADDRESS", "OFFICE_PHONE_NO", "OTHER_CONTACT_NO",
    "new_dpd2", "nmi", "NEXT_PAYMENT_DATE", "loan_amount", "ob", "term",
    "LAST_PAYMENT_AMOUNT", "LAST_PAYMENT_DATE", "NUMBER_OF_PAYMENTS_PAID", "PAY_MODE_",
    "YEAR_MODEL", "ASSET_MAKE", "ASSET_MODEL", "PLATE_NUMBER", "CHASSIS_SERIAL_NO",
    "ENGINE_MOTOR_NUMBER", "vehicle_color", "COLL_CONDUCTION",
    "PBCOM_", "Program_Availed", "adjusted_maturity_date", "fin_past_due_amount",
    "EWB TEAM HEAD", "EWB AGENCY COORDINATOR", "TEAM LEADER", "ICONNECT ID",
    "FIELD HANDLER", "OUTLOOK EMAIL ID", "AGENCY", "GEO", "REGION", "DOB", "EMAIL",
    "COLLATERAL_DESCRIPTION",
]

# Columns IndexOne expects for an FFA upload (mirrors the sample the user already sends)
_IDX_FFA_COLS = [
    "YEAR_MODEL", "ASSET_MAKE", "ASSET_MODEL", "PLATE_NUMBER", "CHASSIS_SERIAL_NO",
    "ENGINE_MOTOR_NUMBER", "vehicle_color", "COLL_CONDUCTION", "UNIT DESCRIPTION",
]


def make_indexone_input(res: dict) -> dict:
    """Return a dict of DataFrames ready to upload to IndexOne.

    Keys:
      'BAU IndexOne Input'  - one row per REGULAR CMS UPLOAD row
      'FFA IndexOne Input'  - one row per FFA CMS UPLOAD row (chassis = account key)

    The DataFrames contain the exact columns IndexOne already receives from SPM,
    derived from the CMS upload DataFrames produced by run().
    """
    reg = res["REGULAR CMS UPLOAD"].copy()
    ffa = res["FFA CMS UPLOAD"].copy()

    # --- BAU input: build from raw worklist rows, matching the 47-col IndexOne BAU format ---
    bau_rows = []
    for e in res.get("_endorse_raw", []):
        o = e["_o"]
        if o["sheet"] != "BAU":
            continue
        r = o["row"]
        bau_rows.append({
            "ENDORSEMENT DATE":        _date_fmt(r.get("ENDORSEMENT DATE")),
            "ENDORSING TEAM":          clean(r.get("ENDORSING TEAM")),
            "accountnumber":           o["acct_raw"],
            "FULLNAME":                clean(r.get("FULLNAME")),
            "MIDDLENAME":              clean(r.get("MIDDLENAME")),
            "GUARANTOR_CO_BORROWER":   clean(r.get("GUARANTOR_CO_BORROWER")),
            "RESIDENCE_ADDRESS":       clean(r.get("RESIDENCE_ADDRESS")),
            "RESIDENCE_PHONE_NO":      clean(r.get("RESIDENCE_PHONE_NO")),
            "MOBILE_NO":               clean(r.get("MOBILE_NO")),
            "COMPANY_NAME":            clean(r.get("COMPANY_NAME")),
            "OFFICE_ADDRESS":          clean(r.get("OFFICE_ADDRESS")),
            "OFFICE_PHONE_NO":         clean(r.get("OFFICE_PHONE_NO")),
            "OTHER_CONTACT_NO":        clean(r.get("OTHER_CONTACT_NO")),
            "new_dpd2":                clean(r.get("new_dpd2")),
            "nmi":                     clean(r.get("nmi")),
            "NEXT_PAYMENT_DATE":       _date_fmt(r.get("NEXT_PAYMENT_DATE")),
            "loan_amount":             clean(r.get("loan_amount")),
            "ob":                      clean(r.get("ob")),
            "term":                    clean(r.get("term")),
            "LAST_PAYMENT_AMOUNT":     clean(r.get("LAST_PAYMENT_AMOUNT")),
            "LAST_PAYMENT_DATE":       _date_fmt(r.get("LAST_PAYMENT_DATE")),
            "NUMBER_OF_PAYMENTS_PAID": clean(r.get("NUMBER_OF_PAYMENTS_PAID")),
            "PAY_MODE_":               clean(r.get("PAY_MODE_")),
            "YEAR_MODEL":              clean(r.get("YEAR_MODEL")),
            "ASSET_MAKE":              clean(r.get("ASSET_MAKE")),
            "ASSET_MODEL":             clean(r.get("ASSET_MODEL")),
            "PLATE_NUMBER":            _plate(r),
            "CHASSIS_SERIAL_NO":       clean(r.get("CHASSIS_SERIAL_NO")),
            "ENGINE_MOTOR_NUMBER":     clean(r.get("ENGINE_MOTOR_NUMBER")),
            "vehicle_color":           _color(r),
            "COLL_CONDUCTION":         _cond_row(r),
            "PBCOM_":                  clean(r.get("PBCOM_")),
            "Program_Availed":         clean(r.get("Program_Availed")),
            "adjusted_maturity_date":  _date_fmt(r.get("adjusted_maturity_date")),
            "fin_past_due_amount":     clean(r.get("fin_past_due_amount")) or clean(r.get("PAST_DUE_AMOUNT")),
            "EWB TEAM HEAD":           clean(r.get("EWB TEAM HEAD")),
            "EWB AGENCY COORDINATOR":  clean(r.get("EWB AGENCY COORDINATOR")),
            "TEAM LEADER":             clean(r.get("TEAM LEADER")),
            "ICONNECT ID":             clean(r.get("ICONNECT ID")),
            "FIELD HANDLER":           clean(r.get("FIELD HANDLER")),
            "OUTLOOK EMAIL ID":        clean(r.get("OUTLOOK EMAIL ID")),
            "AGENCY":                  clean(r.get("AGENCY")),
            "GEO":                     clean(r.get("GEO")),
            "REGION":                  clean(r.get("REGION")),
            "DOB":                     _date_fmt(r.get("DOB")),
            "EMAIL":                   clean(r.get("EMAIL")),
            "COLLATERAL_DESCRIPTION":  clean(r.get("COLLATERAL_DESCRIPTION")),
        })
    bau_df = pd.DataFrame(bau_rows, columns=_IDX_BAU_COLS)

    # --- FFA input: build from raw worklist rows (same as BAU) ---
    ffa_rows = []
    for e in res.get("_endorse_raw", []):
        o = e["_o"]
        if o["sheet"] == "BAU":
            continue
        r = o["row"]
        ffa_rows.append({
            "YEAR_MODEL":          clean(r.get("YEAR_MODEL")),
            "ASSET_MAKE":          clean(r.get("ASSET_MAKE")),
            "ASSET_MODEL":         clean(r.get("ASSET_MODEL")),
            "PLATE_NUMBER":        _plate(r),
            "CHASSIS_SERIAL_NO":   o["chassis_raw"] if o["chassis"] else clean(r.get("CHASSIS_SERIAL_NO")),
            "ENGINE_MOTOR_NUMBER": clean(r.get("ENGINE_MOTOR_NUMBER")),
            "vehicle_color":       _color(r),
            "COLL_CONDUCTION":     _cond_row(r),
            "UNIT DESCRIPTION":    _unit(r, " "),
        })
    ffa_df = pd.DataFrame(ffa_rows, columns=_IDX_FFA_COLS)


    return {"BAU IndexOne Input": bau_df, "FFA IndexOne Input": ffa_df}


def apply_chcodes(res: dict, mapped_src) -> dict:
    """Read the IndexOne 'Mapped' output and fill Ch Code (UID Code) into both CMS uploads.

    Parameters
    ----------
    res        : the dict returned by run()
    mapped_src : path, bytes, or file-like for the IndexOne output workbook
                 (must have a 'Mapped' sheet with columns 'Account Number' and 'UID Code')

    Returns
    -------
    Updated res dict (the CMS upload DataFrames are updated in-place and also returned).
    """
    # Read the Mapped sheet
    mapped = pd.read_excel(open_xlsx(mapped_src), sheet_name="Mapped", dtype=str)
    cols = set(mapped.columns)

    # Auto-detect format:
    #   BAU mapped  → key column = "Account Number"  → apply to REGULAR CMS UPLOAD
    #   FFA mapped  → key column = "Chassis"          → apply to FFA CMS UPLOAD
    if "Account Number" in cols:
        key_col   = "Account Number"
        targets   = ["REGULAR CMS UPLOAD"]
    elif "Chassis" in cols:
        key_col   = "Chassis"
        targets   = ["FFA CMS UPLOAD"]
    else:
        raise ValueError(
            f"IndexOne Mapped sheet has neither 'Account Number' nor 'Chassis' column. "
            f"Found: {list(mapped.columns)}"
        )

    if "UID Code" not in cols:
        raise ValueError(
            f"IndexOne Mapped sheet is missing 'UID Code' column. Found: {list(mapped.columns)}"
        )

    # Build normalised lookup: normed_key → UID Code
    mapped = mapped[mapped[key_col].notna() & mapped["UID Code"].notna()].copy()
    mapped["_KEY"] = mapped[key_col].map(norm)
    chcode_map = dict(zip(mapped["_KEY"], mapped["UID Code"]))

    matched_total, unmatched = 0, []

    for sheet_name in targets:
        df = res[sheet_name].copy()
        if "Ch Code" not in df.columns or "ACCOUNT_NUMBER" not in df.columns:
            continue
        ch_codes = []
        for acct in df["ACCOUNT_NUMBER"]:
            key = norm(acct)
            code = chcode_map.get(key)
            ch_codes.append(code)
            if code:
                matched_total += 1
            elif key:
                unmatched.append({"Sheet": sheet_name, "ACCOUNT_NUMBER": acct})
        df["Ch Code"] = ch_codes
        res[sheet_name] = df

    # Accumulate stats across multiple calls (BAU then FFA)
    prev = res.get("_chcode_stats", {"matched": 0, "unmatched": 0, "unmatched_rows": pd.DataFrame()})
    all_unmatched = pd.concat(
        [prev["unmatched_rows"], pd.DataFrame(unmatched)], ignore_index=True
    ) if len(prev["unmatched_rows"]) or unmatched else pd.DataFrame(unmatched)
    res["_chcode_stats"] = {
        "matched":        prev["matched"] + matched_total,
        "unmatched":      len(all_unmatched),
        "unmatched_rows": all_unmatched,
    }
    return res



def write_indexone_input(res: dict, path_or_buf):
    """Write the IndexOne input workbook (BAU + FFA sheets) to path_or_buf."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    idx = make_indexone_input(res)
    wb = Workbook()
    hf = Font(name="Arial", bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F4E78")
    body = Font(name="Arial")
    first = True
    for sheet_name, df in idx.items():
        ws = wb.active if first else wb.create_sheet()
        ws.title = sheet_name[:31]
        first = False
        for j, col in enumerate(df.columns, 1):
            cell = ws.cell(1, j, col)
            cell.font, cell.fill = hf, fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            w = max([len(str(col))] + [len(str(v)) for v in df[col].head(200).tolist()]) + 3
            ws.column_dimensions[get_column_letter(j)].width = max(12, min(w, 55))
        for i, row in enumerate(df.itertuples(index=False), 2):
            for j, v in enumerate(row, 1):
                v = None if (v is None or (not isinstance(v, str) and pd.isna(v))) else v
                ws.cell(i, j, v).font = body
        ws.freeze_panes = ws.cell(2, 1)
    wb.save(path_or_buf)


def main():
    ap = argparse.ArgumentParser(description="EWB endorsement automation")
    ap.add_argument("--wl", required=True); ap.add_argument("--wl-password")
    ap.add_argument("--datagrid", required=True); ap.add_argument("--datagrid-password")
    ap.add_argument("--template"); ap.add_argument("--out", default="output")
    a = ap.parse_args()
    res = run(a.wl, a.datagrid, a.wl_password, a.datagrid_password, a.template)
    print(res["_summary"].to_string(index=False)); print(); print(res["_pullout_counts"].to_string(index=False))
    for k in ("REGULAR CMS UPLOAD", "FFA CMS UPLOAD", "ENDORSE", "PULLOUT", "DATA_QUALITY"):
        print(f"{k}: {len(res[k])} rows")
    print("Saved:", save_all(res, a.out))


if __name__ == "__main__":
    main()
