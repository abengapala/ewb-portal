"""
EWB endorsement automation (S.P. Madrid) - core logic + command line.

Inputs : bank worklist (WL) xlsx (password ok), DataGrid (CMS) export xlsx, optional template xlsx.
Outputs: one results workbook + CSV copies of the two CMS uploads.

HEADER SCAN (new)
  Before any matching, every sheet / column is scanned and auto-fixed:
    * sheet names are recognised by alias   ("FFA 150" == "150 DPD FFA", "SPM REPO AI" == "FFA SPM REPO AI")
    * the header row is found automatically (not always row 1)
    * column names are matched by alias / punctuation-insensitive / fuzzy match
      ("Account Number", "ACCTNO_", "ACCOUNT NO." -> ACCTNO_2 ; "Chassis No" -> CHASSIS_SERIAL_NO ...)
    * a scan report says exactly what was renamed, what was not recognised, and what is missing
    * anything the scan cannot decide can be forced with an `overrides` dict

KEY RULE
  BAU                                  -> DataGrid "Account No." must be the ACCOUNT NUMBER
  FFA / 150 DPD FFA / FFA SPM REPO AI  -> DataGrid "Account No." must be the CHASSIS
  Priority when an account is on several sheets: BAU > FFA > 150 DPD FFA > FFA SPM REPO AI

INDEXONE FLOW
  Step 1 : run()                         -> CMS uploads with Ch Code = blank
  Step 1b: make_indexone_input(res)       -> dict of DataFrames to upload to IndexOne
  Step 2 : apply_chcodes(res, mapped_src) -> fills UID Code -> Ch Code in both CMS uploads
"""
from __future__ import annotations
import argparse, difflib, io, os, re
from datetime import date
import pandas as pd

# ----------------------------------------------------------------- config
SHEET_PRIORITY = ["BAU", "FFA", "150 DPD FFA", "FFA SPM REPO AI"]
ACCT_COL = "ACCTNO_2"                       # canonical account column after the header fix
CHASSIS_COL = "CHASSIS_SERIAL_NO"
PLACEHOLDER_CHASSIS = {"0", "1"}          # not a real chassis
MIN_CHASSIS_LEN = 4                         # 6-digit chassis exist and are real
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

# ----------------------------------------------------------------- header aliases
# canonical name -> list of other spellings.  Canonical names are the ones the rest of this file uses.
SHEET_ALIASES = {
    "BAU": ["BAU", "BAU WORKLIST", "REGULAR"],
    "FFA": ["FFA", "FFA WORKLIST"],
    "150 DPD FFA": ["150 DPD FFA", "FFA 150", "150 DPD", "FFA 150 DPD", "150DPD FFA", "FFA150"],
    "FFA SPM REPO AI": ["FFA SPM REPO AI", "SPM REPO AI", "REPO AI", "SPM REPO", "FFA REPO AI"],
}

WL_COLUMN_ALIASES = {
    "ACCTNO_2": ["ACCTNO_", "ACCTNO", "ACCT NO", "ACCOUNT NO", "ACCOUNT NO.", "ACCOUNT NUMBER", "ACCOUNT_NUMBER",
                 "ACCOUNTNUMBER", "ACCT NUMBER", "ACCT_NO"],
    "CHASSIS_SERIAL_NO": ["CHASSIS", "CHASSIS NO", "CHASSIS NO.", "CHASSIS NUMBER", "CHASSIS_NO", "CHASSIS SERIAL NO",
                          "CHASSIS SERIAL NUMBER", "VIN", "VIN NO"],
    "ENGINE_MOTOR_NUMBER": ["ENGINE NO", "ENGINE NUMBER", "ENGINE_NO", "MOTOR NO", "MOTOR NUMBER",
                            "ENGINE/MOTOR NUMBER", "ENGINE MOTOR NO"],
    "PLATE_NUMBER": ["PLATE NO", "PLATE NO.", "PLATE", "PLATE_NO", "PLATE NUMBER"],
    "CS NO": ["CS_NO", "CS NUMBER", "CS NO.", "CONDUCTION STICKER", "CONDUCTION STICKER NO"],
    "COLL_CONDUCTION": ["COLL CONDUCTION", "CONDUCTION"],
    "vehicle_color": ["VEHICLE COLOR", "VEHICLE COLOUR", "VEHICLE_COLOUR"],
    "COLOR": ["COLOUR"],
    "FULLNAME": ["FULL NAME", "ACCOUNT NAME", "CUSTOMER NAME", "BORROWER NAME", "BORROWER"],
    "MIDDLENAME": ["MIDDLE NAME"],
    "GUARANTOR_CO_BORROWER": ["GUARANTOR", "CO BORROWER", "GUARANTOR/CO-BORROWER", "GUARANTOR CO BORROWER"],
    "RESIDENCE_ADDRESS": ["RESIDENCE ADDRESS", "HOME ADDRESS"],
    "RESIDENCE_PHONE_NO": ["RESIDENCE PHONE NO", "RESIDENCE PHONE", "HOME PHONE"],
    "MOBILE_NO": ["MOBILE NO", "MOBILE NUMBER", "MOBILE", "CELLPHONE NO"],
    "COMPANY_NAME": ["COMPANY NAME", "EMPLOYER"],
    "OFFICE_ADDRESS": ["OFFICE ADDRESS"],
    "OFFICE_PHONE_NO": ["OFFICE PHONE NO", "OFFICE PHONE"],
    "OTHER_CONTACT_NO": ["OTHER CONTACT NO", "OTHER CONTACT"],
    "new_dpd2": ["NEW_DPD", "NEW DPD", "NEW DPD2", "DPD", "DPD2", "DAYS PAST DUE"],
    "nmi": ["NMI", "MONTHLY AMORTIZATION"],
    "PAST_DUE_AMOUNT": ["PAST DUE AMOUNT", "PAST DUE", "PAST_DUE"],
    "fin_past_due_amount": ["FIN PAST DUE AMOUNT", "FIN_PAST_DUE_AMOUNT"],
    "ob": ["OB", "OUTSTANDING BALANCE"],
    "loan_amount": ["LOAN AMOUNT"],
    "term": ["LOAN TERM"],
    "LAST_PAYMENT_AMOUNT": ["LAST PAYMENT AMOUNT"],
    "LAST_PAYMENT_DATE": ["LAST PAYMENT DATE"],
    "NEXT_PAYMENT_DATE": ["NEXT PAYMENT DATE"],
    "NUMBER_OF_PAYMENTS_PAID": ["NUMBER OF PAYMENTS PAID", "PAYMENTS PAID"],
    "PAY_MODE_": ["PAY MODE", "PAY_MODE", "PAYMENT MODE"],
    "YEAR_MODEL": ["YEAR MODEL", "YEAR"],
    "ASSET_MAKE": ["ASSET MAKE", "MAKE", "BRAND"],
    "ASSET_MODEL": ["ASSET MODEL", "MODEL"],
    "COLLATERAL_DESCRIPTION": ["COLLATERAL DESCRIPTION", "UNIT DESCRIPTION", "UNIT_DETAILS", "UNIT DETAILS"],
    "EMAIL": ["EMAIL ADDRESS", "E-MAIL"],
    "REFERENCE_PROVINCE": ["REFERENCE PROVINCE", "PROVINCE"],
    "REFERENCE_MUNICIPALITY": ["REFERENCE MUNICIPALITY", "MUNICIPALITY"],
    "STATUS": ["ACCOUNT STATUS"],
    "ENDO": ["ENDO TAG"],
    "ENDORSEMENT DATE": ["ENDO DATE", "ENDORSEMENT_DATE"],
    "ENDORSING TEAM": ["ENDORSING_TEAM"],
    "FIELD HANDLER": ["FIELD_HANDLER"],
}
# columns that only need punctuation/case-insensitive matching (no extra aliases)
for _c in ["ENDO TYPE", "CS/PLATE", "ENDORSEMENT MONTH", "PBCOM_", "Program_Availed", "adjusted_maturity_date",
           "EWB TEAM HEAD", "EWB AGENCY COORDINATOR", "TEAM LEADER", "ICONNECT ID", "OUTLOOK EMAIL ID", "AGENCY",
           "GEO", "REGION", "DOB", "CH CODE", "MATURITY_DATE", "remaining_terms", "PLATE_NUMBER2"]:
    WL_COLUMN_ALIASES.setdefault(_c, [])

DG_COLUMN_ALIASES = {
    "Account No.": ["ACCOUNT NO", "ACCOUNT NUMBER", "ACCOUNT_NUMBER", "ACCT NO", "ACCTNO", "ACCOUNTNO"],
    "Name": ["ACCOUNT NAME", "DEBTOR NAME", "FULL NAME", "FULLNAME"],
    "New I.C.": ["NEW IC", "NEW I C", "CH CODE", "CHCODE", "UID CODE"],
    "Batch No.": ["BATCH NO", "BATCH", "BATCH NUMBER"],
    "Collector": ["COLLECTOR NAME", "AGENT"],
    "Status Code": ["STATUS CODE", "STATUS"],
    "Debtor ID": ["DEBTORID", "DEBTOR NO"],
    "No": ["NO.", "ROW NO"],
    "Account has PTP?": ["HAS PTP", "ACCOUNT HAS PTP"],
    "PTP Date": [],
    "PTP Amount": [],
}

MAPPED_UID_COLS = ["UID CODE", "UID", "CH CODE", "CHCODE"]
MAPPED_KEY_COLS = ["ACCOUNT NUMBER", "ACCOUNT NO", "ACCOUNT_NUMBER", "CHASSIS", "CHASSIS NO", "CHASSIS SERIAL NO",
                   "CHASSIS_SERIAL_NO"]


# ----------------------------------------------------------------- header scan helpers
def _k(s):
    """Punctuation / case / space-insensitive comparison key."""
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def _blank(v):
    return v is None or (not isinstance(v, str) and pd.isna(v))


def detect_role(sheet_name):
    """Map a sheet name to a worklist role. Returns (role, how) or (None, None)."""
    n = _k(sheet_name)
    for role, al in SHEET_ALIASES.items():
        if n in {_k(a) for a in al}:
            return role, "name"
    if "REPO" in n: return "FFA SPM REPO AI", "keyword"
    if "150" in n: return "150 DPD FFA", "keyword"
    if "FFA" in n: return "FFA", "keyword"
    if "BAU" in n: return "BAU", "keyword"
    return None, None


def _lookups(aliases):
    canon = {_k(c): c for c in aliases}
    alias = {}
    for c, al in aliases.items():
        for a in al:
            alias.setdefault(_k(a), c)
    return canon, alias


def find_header_row(raw, aliases, scan_rows=15):
    """Return (row_index, hits) of the row that looks most like a header."""
    canon, alias = _lookups(aliases)
    known = set(canon) | set(alias)
    best, best_i = -1, 0
    for i in range(min(scan_rows, len(raw))):
        hits = sum(1 for v in raw.iloc[i].tolist() if not _blank(v) and _k(v) in known)
        if hits > best:
            best, best_i = hits, i
    return (best_i if best >= 2 else 0), max(best, 0)


def map_columns(cols, aliases, overrides=None):
    """Match actual column names to canonical names.
    Order: manual override -> exact canonical -> alias -> fuzzy (>=0.9). A canonical name is used once only.
    Returns (mapping actual->canonical, how actual->method)."""
    canon_n, alias_n = _lookups(aliases)
    mapping, how, claimed = {}, {}, set()
    for canon, act in (overrides or {}).items():
        if act in cols and canon in aliases:
            mapping[act], how[act] = canon, "manual"; claimed.add(canon)
    norm_of = {c: _k(c) for c in cols}
    for c in cols:                                            # pass 1: exact canonical
        if c in mapping: continue
        canon = canon_n.get(norm_of[c])
        if canon and canon not in claimed:
            mapping[c], how[c] = canon, "exact"; claimed.add(canon)
    for c in cols:                                            # pass 2: alias
        if c in mapping: continue
        canon = alias_n.get(norm_of[c])
        if canon and canon not in claimed:
            mapping[c], how[c] = canon, "alias"; claimed.add(canon)
    pool = list(set(canon_n) | set(alias_n))
    for c in cols:                                            # pass 3: fuzzy
        if c in mapping or len(norm_of[c]) < 5: continue
        m = difflib.get_close_matches(norm_of[c], pool, n=1, cutoff=0.9)
        if m:
            canon = canon_n.get(m[0]) or alias_n.get(m[0])
            if canon and canon not in claimed:
                mapping[c], how[c] = canon, "fuzzy"; claimed.add(canon)
    return mapping, how


def _frame_from_raw(raw, hi):
    hdr, seen = [], {}
    for j, v in enumerate(raw.iloc[hi].tolist()):
        name = "" if _blank(v) else str(v).strip()
        if not name: name = f"Unnamed: {j}"
        if name in seen:
            seen[name] += 1; name = f"{name}.{seen[name]}"
        else:
            seen[name] = 0
        hdr.append(name)
    df = raw.iloc[hi + 1:].copy()
    df.columns = hdr
    return df.reset_index(drop=True)


def _scan_df(raw, aliases, overrides=None):
    """Find the header row, rename recognised columns to canonical names. Returns (df, info)."""
    hi, hits = find_header_row(raw, aliases)
    df = _frame_from_raw(raw, hi).dropna(how="all")
    orig = list(df.columns)
    mapping, how = map_columns(orig, aliases, overrides)
    taken = set(mapping.values())
    ren = {}
    for c in orig:
        if c in mapping: ren[c] = mapping[c]
        elif c in taken: ren[c] = f"{c}__orig"               # never let two columns share a name
    df = df.rename(columns=ren).reset_index(drop=True)
    info = dict(header_row=hi + 1, header_hits=hits, original_columns=orig,
                mapped=[(c, mapping[c], how[c]) for c in orig if c in mapping],
                unmapped=[c for c in orig if c not in mapping])
    return df, info


def decrypt_bytes(src, password=None):
    """Return plain .xlsx bytes (decrypts a password protected workbook once, so it need not be repeated)."""
    return open_xlsx(src, password).getvalue()


def scan_wl(src, password=None, overrides=None):
    """Scan the bank worklist. Returns (wl, report) - wl = {role: DataFrame with canonical headers}.
    overrides = {role: {canonical: actual column name}}"""
    overrides = overrides or {}
    xl = pd.ExcelFile(open_xlsx(src, password))
    found = [(n,) + detect_role(n) for n in xl.sheet_names]
    found.sort(key=lambda t: 0 if t[2] == "name" else 1)      # exact sheet names win over keyword guesses
    role_of, ignored = {}, []
    for name, role, how in found:
        if role is None: ignored.append(name)
        elif role in role_of: ignored.append(f"{name} (duplicate of '{role_of[role]}')")
        else: role_of[role] = name
    wl, sheets, errors, warnings = {}, [], [], []
    for role in SHEET_PRIORITY:
        if role not in role_of: continue
        name = role_of[role]
        raw = xl.parse(name, header=None, dtype=object)
        df, info = _scan_df(raw, WL_COLUMN_ALIASES, overrides.get(role))
        wl[role] = df
        cols = set(df.columns)
        has_a, has_c = ACCT_COL in cols, CHASSIS_COL in cols
        miss = []
        if not has_a: miss.append(ACCT_COL)
        if not has_c: miss.append(CHASSIS_COL)
        if role == "BAU" and not has_a:
            errors.append(f"Sheet '{name}' ({role}): no ACCOUNT NUMBER column found.")
        if role != "BAU" and not has_a and not has_c:
            errors.append(f"Sheet '{name}' ({role}): neither a CHASSIS nor an ACCOUNT NUMBER column found.")
        elif role != "BAU" and not has_c:
            warnings.append(f"Sheet '{name}' ({role}): no CHASSIS column - matching by account no. instead.")
        sheets.append(dict(sheet=name, role=role, role_how=next(t[2] for t in found if t[0] == name),
                           rows=len(df), missing=miss, **info))
    for r in ("BAU", "FFA"):
        if r not in wl:
            errors.append(f"Worklist has no '{r}' sheet. Sheets found: {xl.sheet_names}")
    return wl, dict(sheets=sheets, ignored=ignored, errors=errors, warnings=warnings)


def scan_dg(src, password=None, overrides=None):
    """Scan the DataGrid export. Returns (df, report). Adds the _KEY column used for matching."""
    xl = pd.ExcelFile(open_xlsx(src, password))
    best = None
    for name in xl.sheet_names:
        raw = xl.parse(name, header=None, dtype=str)
        hi, hits = find_header_row(raw, DG_COLUMN_ALIASES)
        if best is None or hits > best[0]:
            best = (hits, name, raw)
    _, name, raw = best
    df, info = _scan_df(raw, DG_COLUMN_ALIASES, overrides)
    errors = []
    if "Account No." not in df.columns:
        errors.append("DataGrid: no 'Account No.' column found.")
    else:
        df = df[df["Account No."].notna()].copy()
        df["_KEY"] = df["Account No."].map(norm)
        df = df[df["_KEY"].notna()].reset_index(drop=True)
    report = dict(sheet=name, rows=len(df), errors=errors, **info)
    return df, report


def scan_table(wl_rep=None, dg_rep=None):
    """Flat table of what the header scan changed (goes into the results workbook / UI)."""
    rows = []
    for s in (wl_rep or {}).get("sheets", []):
        for a, c, h in s["mapped"]:
            rows.append(["WORKLIST", s["sheet"], s["role"], s["header_row"], a, c, h, "RENAMED" if a != c else "same"])
        for a in s["unmapped"]:
            rows.append(["WORKLIST", s["sheet"], s["role"], s["header_row"], a, None, "not used", ""])
    if dg_rep:
        for a, c, h in dg_rep["mapped"]:
            rows.append(["DATAGRID", dg_rep["sheet"], "DataGrid", dg_rep["header_row"], a, c, h, "RENAMED" if a != c else "same"])
        for a in dg_rep["unmapped"]:
            rows.append(["DATAGRID", dg_rep["sheet"], "DataGrid", dg_rep["header_row"], a, None, "not used", ""])
    return pd.DataFrame(rows, columns=["File", "Sheet", "Role", "Header row", "Original column", "Mapped to", "Method", "Change"])


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
    s = re.sub(r"['`;]", "", s).strip()      # characters that break the CMS upload
    return s or None


def clean_vehicle_id(v):
    """Clean chassis / engine numbers: strip illegal chars, collapse repeated halves, take first of many."""
    s = clean(v)
    if not s:
        return None
    if len(s) >= 8 and len(s) % 2 == 0:
        half = len(s) // 2
        if s[:half] == s[half:]:
            s = s[:half]
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
        try:
            f.load_key(password=password)
            f.decrypt(out)
        except Exception as ex:
            raise ValueError(f"Could not open the workbook - wrong password? ({ex})")
        return io.BytesIO(out.getvalue())
    return io.BytesIO(data)


def read_wl(src, password=None, overrides=None):
    wl, rep = scan_wl(src, password, overrides)
    if rep["errors"]:
        raise ValueError(" | ".join(rep["errors"]))
    return wl


def read_datagrid(src, password=None, overrides=None):
    df, rep = scan_dg(src, password, overrides)
    if rep["errors"]:
        raise ValueError(" | ".join(rep["errors"]))
    return df


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
        acol = ACCT_COL if ACCT_COL in df.columns else None
        if acol is None and sh == "BAU":
            raise ValueError("No account number column found on the BAU sheet.")
        for row in df.to_dict("records"):
            a = norm(row.get(acol)) if acol else None
            c_raw = norm(row.get(CHASSIS_COL)); c = valid_chassis(c_raw)
            rec = dict(sheet=sh, acct=a, acct_raw=clean(row.get(acol)) if acol else None, chassis=c,
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


def _yn(v):
    v = clean(v)
    return {"1": "Yes", "1.0": "Yes", "0": "No", "0.0": "No"}.get(v, v)


def _pull_row(r, reason, sh=None, o=None):
    return {"DG No": r.get("No"), "DG Name": r.get("Name"), "Account No. (in DataGrid)": r.get("Account No."),
            "New I.C. (CH code)": r.get("New I.C."), "Batch No.": r.get("Batch No."),
            "Collector": r.get("Collector"), "Status Code": r.get("Status Code"), "Debtor ID": r.get("Debtor ID"),
            "Has PTP?": _yn(r.get("Account has PTP?")), "PTP Date": r.get("PTP Date"), "PTP Amount": r.get("PTP Amount"),
            "Reason": reason, "WL Sheet": sh, "WL ACCOUNT NO": o["acct_raw"] if o else None,
            "WL FULLNAME": o["name"] if o else None}


def classify(owners, dg):
    dgrow, dg_dups = {}, []
    for r in dg.to_dict("records"):
        if r["_KEY"] in dgrow: dg_dups.append(r)
        else: dgrow[r["_KEY"]] = r
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

        if exp_row is not None:
            stats[(sh, "ACTIVE - stays")] = stats.get((sh, "ACTIVE - stays"), 0) + 1
            if oth_row is not None:
                pull.append(_pull_row(oth_row, f"DUPLICATE - DG has this account twice; keep the {expn} row, pull out the {othn} row", sh, o))
                stats[(sh, "DUPLICATE row to pull out")] = stats.get((sh, "DUPLICATE row to pull out"), 0) + 1
        elif oth_row is not None:
            pull.append(_pull_row(oth_row, f"WRONG KEY - loaded under {othn}, change to {expn}", sh, o))
            endorse.append({**base, "ENDORSE UNDER": expn, "KEY TO ENDORSE": key_raw, "REASON": f"CHANGE TO {expn}", "_o": o})
            (tochassis if exp == "chassis" else toacct).append(
                {**base, "Currently in DG as": othn, "DG Name": oth_row.get("Name"), "DG Batch": oth_row.get("Batch No."),
                 "Change to": expn, "KEY TO USE": key_raw})
            stats[(sh, "CHANGE KEY")] = stats.get((sh, "CHANGE KEY"), 0) + 1
        else:
            endorse.append({**base, "ENDORSE UNDER": expn, "KEY TO ENDORSE": key_raw, "REASON": "NEW ENDO", "_o": o})
            stats[(sh, "NEW ENDO")] = stats.get((sh, "NEW ENDO"), 0) + 1
    for r in dg_dups:                                                   # same key twice in DataGrid
        if r["_KEY"] in claimed:
            pull.append(_pull_row(r, "DUPLICATE ROW - same key appears more than once in DataGrid (extra row)"))
            stats[("DataGrid", "DUPLICATE row in DataGrid")] = stats.get(("DataGrid", "DUPLICATE row in DataGrid"), 0) + 1
    for r in dg.to_dict("records"):
        if r["_KEY"] not in claimed:
            pull.append(_pull_row(r, "PULLOUT - not in worklist"))
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
    if isinstance(v, (int, float)):
        if v < 20000: return None
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))).date()
    try:
        d = pd.to_datetime(v).date()
    except Exception:
        return v
    return None if d.year < 1990 else d


def _unit(row, sep):
    parts = [str(clean(row.get(k))).strip() for k in ("YEAR_MODEL", "ASSET_MAKE", "ASSET_MODEL") if clean(row.get(k)) is not None]
    return sep.join(p for p in parts if p and p != "0")


_NO_PLATE_KEYWORDS = ("NO CR", "NO PLATE", "NO CONDUCTION", "NOT AVAILABLE", "N/A", "NONE", "WITHOUT", "W/O",
                      "NO STICKER", "NO REG", "ATTACHED", "AVAILABLE", "INDICATED", "PROVIDED")


def _plate(row):
    """Best plate/conduction value: PLATE_NUMBER if real, else CS NO, else COLL_CONDUCTION."""
    raw = clean(row.get("PLATE_NUMBER"))
    if raw:
        upper = str(raw).strip().upper()
        if not (len(upper) > 15 or any(kw in upper for kw in _NO_PLATE_KEYWORDS)):
            return raw
    return _cond(row.get("CS NO")) or _cond(row.get("COLL_CONDUCTION"))


def _cond(v):
    """Conduction sticker value, or None if blank / zero."""
    v = clean(v)
    if v is None:
        return None
    try:
        if float(v) == 0:
            return None
    except (ValueError, TypeError):
        pass
    s = str(v).strip()
    return None if s in ("", "0", "0.0") else s


def _cond_row(row):
    return _cond(row.get("CS NO")) or _cond(row.get("COLL_CONDUCTION")) or _cond(row.get("CS/PLATE"))


def _color(row):
    return clean(row.get("vehicle_color")) or clean(row.get("COLOR")) or None


def _date_fmt(v):
    """Any date value -> DD/MM/YYYY string. None if blank / unparseable."""
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
        vals = {"ACCOUNT_NUMBER": key, "ACCOUNT_NAME": "EWB AUTO", "OB": 1,
                "RESIDENCE_ADDRESS": None, "MOBILE_NO": None, "COMPANY_NAME": None, "OFFICE_ADDRESS": None,
                "NEW_DPD": None, "LAST_PAYMENT_AMOUNT": None, "LAST_PAYMENT_DATE": None,
                "FIN_PAST_DUE_AMOUNT": None, "EMAIL": None, "REFERENCE_PROVINCE": None,
                "UNIT_DETAILS": clean(_unit(r, " ")), "PLATE_NUMBER": clean(_plate(r)),
                "CHASSIS_SERIAL_NO": chassis_clean,
                "ENGINE_MOTOR_NUMBER": clean_vehicle_id(r.get("ENGINE_MOTOR_NUMBER")),
                "vehicle_color": clean(_color(r)),
                "COLL_CONDUCTION": clean(_cond_row(r)), "COLL_CONDUCTION\t": clean(_cond_row(r)),
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
        if o["sheet"] == "FFA SPM REPO AI":
            plate = plate or clean(r.get("COLL_CONDUCTION"))
        if not plate or str(plate).strip() in ("", "0"): continue
        if o["sheet"] == "BAU": tab, model = "REPO AI BAU", _unit(r, "-").replace(" ", "-")
        elif o["sheet"] == "FFA SPM REPO AI": tab, model = "REPO AI SPECIAL", _unit(r, " ")
        else: tab, model = "REPO AI FFA", _unit(r, "-").replace(" ", "-")
        out[tab].append([None, None, plate, model, today, REPO_AI_CLIENT, REPO_AI_PRIORITY])
    return {k: pd.DataFrame(v, columns=REPO_COLS) for k, v in out.items()}


# ----------------------------------------------------------------- pullout upload
def make_pullout_upload(pull_df: "pd.DataFrame") -> "pd.DataFrame":
    """Clean PULLOUT UPLOAD from the full PULLOUT detail table. Columns: ACCOUNT_NUMBER | CH_CODE | REASON
    ACCOUNT_NUMBER = the value currently in DataGrid (Account No.), CH_CODE = DataGrid New I.C."""
    _COLS = ["ACCOUNT_NUMBER", "CH_CODE", "REASON"]
    if pull_df is None or not len(pull_df):
        return pd.DataFrame(columns=_COLS)
    rows = []
    for _, r in pull_df.iterrows():
        raw = str(r.get("Reason") or "")
        raw_up = raw.upper()
        if "NOT IN WORKLIST" in raw_up:
            reason = "NOT IN WORKLIST"
        elif "DUPLICATE ROW" in raw_up:
            reason = "DUPLICATE ROW IN DATAGRID"
        elif "DUPLICATE" in raw_up and "CHASSIS" in raw_up:
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
        rows.append({"ACCOUNT_NUMBER": r.get("Account No. (in DataGrid)"),
                     "CH_CODE": r.get("New I.C. (CH code)"), "REASON": reason})
    return pd.DataFrame(rows, columns=_COLS)


# ----------------------------------------------------------------- orchestration
def run(wl_src, dg_src, wl_password=None, dg_password=None, template=None,
        wl_overrides=None, dg_overrides=None, wl_pre=None, dg_pre=None):
    """wl_pre / dg_pre = the (frame(s), report) already returned by scan_wl / scan_dg (skips re-reading)."""
    wl, wl_rep = wl_pre if wl_pre is not None else scan_wl(wl_src, wl_password, wl_overrides)
    dg, dg_rep = dg_pre if dg_pre is not None else scan_dg(dg_src, dg_password, dg_overrides)
    errs = wl_rep["errors"] + dg_rep["errors"]
    if errs:
        raise ValueError(" | ".join(errs))
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
        "HEADER SCAN": scan_table(wl_rep, dg_rep),
    }
    res["_endorse_raw"] = endorse
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
    tot = lambda key: sum(v for (s, k), v in stats.items() if k == key)
    res["_counts"] = {"wl_accounts": len(owners), "datagrid_rows": len(dg), "new_endo": tot("NEW ENDO"),
                      "active": tot("ACTIVE - stays"), "pullout": len(pull), "key_changes": tot("CHANGE KEY")}
    res["_scan"] = {"wl": wl_rep, "dg": dg_rep}
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
            x = ws.cell(r0, j, c); x.font, x.fill = hf, fill
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
             "PULLOUT UPLOAD mirrors the 'EWB RECOV (PULLOUT)' layout - confirm the real pullout format before using.",
             "HEADER SCAN sheet lists every column that was renamed / not recognised. Check DATA_QUALITY before uploading."]
    for k, t in enumerate(notes): ws.cell(n + k, 1, t).font = body
    ws.column_dimensions["A"].width = 34
    for name, df in res.items():
        if name.startswith("_") or not isinstance(df, pd.DataFrame) or not len(df.columns): continue
        put(wb.create_sheet(name[:31]), df)
    wb.save(path_or_buf)


def save_all(res, outdir):
    os.makedirs(outdir, exist_ok=True)
    stamp = date.today().strftime("%Y%m%d")
    xlsx = os.path.join(outdir, f"EWB_ENDORSEMENT_RESULT_{stamp}.xlsx")
    write_workbook(res, xlsx)
    for name in ("REGULAR CMS UPLOAD", "FFA CMS UPLOAD", "PULLOUT UPLOAD"):
        res[name].to_csv(os.path.join(outdir, f"{name.replace(' ', '_')}_{stamp}.csv"), index=False, encoding="utf-8-sig")
    return xlsx


# ----------------------------------------------------------------- IndexOne helpers
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
_IDX_FFA_COLS = [
    "YEAR_MODEL", "ASSET_MAKE", "ASSET_MODEL", "PLATE_NUMBER", "CHASSIS_SERIAL_NO",
    "ENGINE_MOTOR_NUMBER", "vehicle_color", "COLL_CONDUCTION", "UNIT DESCRIPTION",
]


def make_indexone_input(res: dict) -> dict:
    """DataFrames ready to upload to IndexOne: 'BAU IndexOne Input' (account key) and 'FFA IndexOne Input' (chassis key)."""
    bau_rows = []
    for e in res.get("_endorse_raw", []):
        o = e["_o"]
        if o["sheet"] != "BAU":
            continue
        r = o["row"]
        bau_rows.append({
            "ENDORSEMENT DATE": _date_fmt(r.get("ENDORSEMENT DATE")), "ENDORSING TEAM": clean(r.get("ENDORSING TEAM")),
            "accountnumber": o["acct_raw"], "FULLNAME": clean(r.get("FULLNAME")), "MIDDLENAME": clean(r.get("MIDDLENAME")),
            "GUARANTOR_CO_BORROWER": clean(r.get("GUARANTOR_CO_BORROWER")),
            "RESIDENCE_ADDRESS": clean(r.get("RESIDENCE_ADDRESS")), "RESIDENCE_PHONE_NO": clean(r.get("RESIDENCE_PHONE_NO")),
            "MOBILE_NO": clean(r.get("MOBILE_NO")), "COMPANY_NAME": clean(r.get("COMPANY_NAME")),
            "OFFICE_ADDRESS": clean(r.get("OFFICE_ADDRESS")), "OFFICE_PHONE_NO": clean(r.get("OFFICE_PHONE_NO")),
            "OTHER_CONTACT_NO": clean(r.get("OTHER_CONTACT_NO")), "new_dpd2": clean(r.get("new_dpd2")),
            "nmi": clean(r.get("nmi")), "NEXT_PAYMENT_DATE": _date_fmt(r.get("NEXT_PAYMENT_DATE")),
            "loan_amount": clean(r.get("loan_amount")), "ob": clean(r.get("ob")), "term": clean(r.get("term")),
            "LAST_PAYMENT_AMOUNT": clean(r.get("LAST_PAYMENT_AMOUNT")), "LAST_PAYMENT_DATE": _date_fmt(r.get("LAST_PAYMENT_DATE")),
            "NUMBER_OF_PAYMENTS_PAID": clean(r.get("NUMBER_OF_PAYMENTS_PAID")), "PAY_MODE_": clean(r.get("PAY_MODE_")),
            "YEAR_MODEL": clean(r.get("YEAR_MODEL")), "ASSET_MAKE": clean(r.get("ASSET_MAKE")),
            "ASSET_MODEL": clean(r.get("ASSET_MODEL")), "PLATE_NUMBER": _plate(r),
            "CHASSIS_SERIAL_NO": clean(r.get("CHASSIS_SERIAL_NO")), "ENGINE_MOTOR_NUMBER": clean(r.get("ENGINE_MOTOR_NUMBER")),
            "vehicle_color": _color(r), "COLL_CONDUCTION": _cond_row(r), "PBCOM_": clean(r.get("PBCOM_")),
            "Program_Availed": clean(r.get("Program_Availed")), "adjusted_maturity_date": _date_fmt(r.get("adjusted_maturity_date")),
            "fin_past_due_amount": clean(r.get("fin_past_due_amount")) or clean(r.get("PAST_DUE_AMOUNT")),
            "EWB TEAM HEAD": clean(r.get("EWB TEAM HEAD")), "EWB AGENCY COORDINATOR": clean(r.get("EWB AGENCY COORDINATOR")),
            "TEAM LEADER": clean(r.get("TEAM LEADER")), "ICONNECT ID": clean(r.get("ICONNECT ID")),
            "FIELD HANDLER": clean(r.get("FIELD HANDLER")), "OUTLOOK EMAIL ID": clean(r.get("OUTLOOK EMAIL ID")),
            "AGENCY": clean(r.get("AGENCY")), "GEO": clean(r.get("GEO")), "REGION": clean(r.get("REGION")),
            "DOB": _date_fmt(r.get("DOB")), "EMAIL": clean(r.get("EMAIL")),
            "COLLATERAL_DESCRIPTION": clean(r.get("COLLATERAL_DESCRIPTION")),
        })
    bau_df = pd.DataFrame(bau_rows, columns=_IDX_BAU_COLS)
    ffa_rows = []
    for e in res.get("_endorse_raw", []):
        o = e["_o"]
        if o["sheet"] == "BAU":
            continue
        r = o["row"]
        ffa_rows.append({
            "YEAR_MODEL": clean(r.get("YEAR_MODEL")), "ASSET_MAKE": clean(r.get("ASSET_MAKE")),
            "ASSET_MODEL": clean(r.get("ASSET_MODEL")), "PLATE_NUMBER": _plate(r),
            "CHASSIS_SERIAL_NO": o["chassis_raw"] if o["chassis"] else clean(r.get("CHASSIS_SERIAL_NO")),
            "ENGINE_MOTOR_NUMBER": clean(r.get("ENGINE_MOTOR_NUMBER")), "vehicle_color": _color(r),
            "COLL_CONDUCTION": _cond_row(r), "UNIT DESCRIPTION": _unit(r, " "),
        })
    ffa_df = pd.DataFrame(ffa_rows, columns=_IDX_FFA_COLS)
    return {"BAU IndexOne Input": bau_df, "FFA IndexOne Input": ffa_df}


def apply_chcodes(res: dict, mapped_src, target=None) -> dict:
    """Fill Ch Code (UID Code) from an IndexOne 'Mapped' workbook into the CMS uploads.

    The Mapped file may key on 'Account Number' (BAU) and/or 'Chassis' (FFA). Every key column found is
    put into one lookup (account numbers and chassis never collide), so the codes land on the right
    upload whatever the column is called. `target` ('BAU' / 'FFA') only restricts which upload is filled.
    Stats are recomputed from the uploads themselves, so calling this once per Mapped file is safe."""
    xl = pd.ExcelFile(open_xlsx(mapped_src))
    uid_n = {_k(c) for c in MAPPED_UID_COLS}
    key_n = {_k(c) for c in MAPPED_KEY_COLS}
    sheet = next((s for s in xl.sheet_names if s.strip().lower() == "mapped"), None)
    if sheet is None:                                              # fall back to any sheet that has a UID column
        for s in xl.sheet_names:
            if {_k(c) for c in xl.parse(s, nrows=0).columns} & uid_n:
                sheet = s; break
    if sheet is None:
        raise ValueError(f"No 'Mapped' sheet (or any sheet with a UID Code column) found. Sheets: {xl.sheet_names}")
    mapped = xl.parse(sheet, dtype=str)
    uid_col = next((c for c in mapped.columns if _k(c) in uid_n), None)
    key_cols = [c for c in mapped.columns if _k(c) in key_n]
    if uid_col is None:
        raise ValueError(f"Mapped sheet is missing the 'UID Code' column. Found: {list(mapped.columns)}")
    if not key_cols:
        raise ValueError(f"Mapped sheet has neither an 'Account Number' nor a 'Chassis' column. Found: {list(mapped.columns)}")
    chcode_map = {}
    for kc in key_cols:
        sub = mapped[mapped[kc].notna() & mapped[uid_col].notna()]
        for k, u in zip(sub[kc].map(norm), sub[uid_col]):
            if k: chcode_map.setdefault(k, u)
    targets = {"BAU": ["REGULAR CMS UPLOAD"], "FFA": ["FFA CMS UPLOAD"]}.get(target, ["REGULAR CMS UPLOAD", "FFA CMS UPLOAD"])
    for sheet_name in targets:
        df = res[sheet_name].copy()
        if "Ch Code" not in df.columns or "ACCOUNT_NUMBER" not in df.columns:
            continue
        df["Ch Code"] = [chcode_map.get(norm(a)) or old for a, old in zip(df["ACCOUNT_NUMBER"], df["Ch Code"])]
        res[sheet_name] = df
    matched, unmatched = 0, []
    for sheet_name in ("REGULAR CMS UPLOAD", "FFA CMS UPLOAD"):
        df = res[sheet_name]
        for acct, code in zip(df["ACCOUNT_NUMBER"], df["Ch Code"]):
            if not _blank(code) and str(code).strip(): matched += 1
            elif norm(acct): unmatched.append({"Sheet": sheet_name, "ACCOUNT_NUMBER": acct})
    res["_chcode_stats"] = {"matched": matched, "unmatched": len(unmatched), "unmatched_rows": pd.DataFrame(unmatched)}
    return res


def write_indexone_input(res: dict, path_or_buf):
    """Write the IndexOne input workbook (BAU + FFA sheets)."""
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
    sc = res["HEADER SCAN"]
    print("Header scan - renamed columns:"); print(sc[sc["Change"] == "RENAMED"].to_string(index=False)); print()
    print(res["_summary"].to_string(index=False)); print(); print(res["_pullout_counts"].to_string(index=False))
    for k in ("REGULAR CMS UPLOAD", "FFA CMS UPLOAD", "ENDORSE", "PULLOUT", "PULLOUT UPLOAD", "DATA_QUALITY"):
        print(f"{k}: {len(res[k])} rows")
    print("Saved:", save_all(res, a.out))


if __name__ == "__main__":
    main()
