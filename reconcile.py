"""
reconcile.py
Core reconciliation logic for Cigna-style insurance premium reconciliation.
Kept separate from the Streamlit UI so it can be tested and reused
(e.g. from a notebook or a scheduled script) without a browser.
"""

import re
import pandas as pd
import numpy as np


# ---------------------------------------------------------------------------
# Loading & column detection
# ---------------------------------------------------------------------------

def _tokens(s):
    s = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", str(s))
    return re.findall(r"[a-z0-9]+", s.lower())


def find_header_row(raw_df, must_contain="unique identifier", max_scan_rows=20):
    """Scan the first N rows of a headerless read for the row that looks like
    the real header (contains a cell matching `must_contain`)."""
    target_tokens = set(_tokens(must_contain))
    for i in range(min(max_scan_rows, len(raw_df))):
        for v in raw_df.iloc[i].tolist():
            if target_tokens.issubset(set(_tokens(v))):
                return i
    raise ValueError(
        f"Could not find a header row containing '{must_contain}' in the "
        f"first {max_scan_rows} rows. Check the file format."
    )


def _pick_column(columns, *keyword_sets, exclude=None):
    """Return the column whose tokens are a superset of one of the keyword
    sets, preferring the most specific match (fewest extra tokens) and
    skipping any column whose tokens include `exclude`."""
    col_tokens = {c: set(_tokens(c)) for c in columns}
    for keywords in keyword_sets:
        kw_set = set(k.lower() for k in keywords)
        candidates = [
            c for c, toks in col_tokens.items()
            if kw_set.issubset(toks) and not (exclude and exclude in toks)
        ]
        if candidates:
            return min(candidates, key=lambda c: len(col_tokens[c]))
    return None


def load_billing(file) -> pd.DataFrame:
    """Load a billing/rectification-style extract. Returns a standardized
    DataFrame: UID, NameFamily, FirstNameFamily, DOB, AgeNum, Category,
    AgeRange, Code, Days, Country, PremiumBilled. Name/DOB/AgeNum are
    optional — left blank if the source file doesn't have them."""
    raw = pd.read_excel(file, header=None, sheet_name=0)
    hdr = find_header_row(raw)
    df = pd.read_excel(file, header=hdr, sheet_name=0)
    df = df.dropna(how="all")

    cols = list(df.columns)
    uid_col = _pick_column(cols, ["unique", "identifier"])
    name_col = _pick_column(cols, ["name", "family"], ["name"])
    fname_col = _pick_column(cols, ["first", "name", "family"], ["first", "name"])
    dob_col = _pick_column(cols, ["dob"], ["date", "birth"])
    age_num_col = _pick_column(cols, ["age"])
    cat_col = _pick_column(cols, ["category", "family"], ["category"])
    age_col = _pick_column(cols, ["age", "range"])
    days_col = _pick_column(cols, ["days", "insured"])
    country_col = _pick_column(cols, ["country", "name", "station"], ["country", "station"])
    code_col = _pick_column(cols, ["code"])
    # "Premium medical" but not "taxes medical"
    premium_col = _pick_column(cols, ["premium", "medical"], exclude="taxes")

    missing = [n for n, v in [
        ("Unique Identifier", uid_col), ("Age range", age_col),
        ("Days insured", days_col), ("Premium medical", premium_col),
    ] if v is None]
    if missing:
        raise ValueError(f"Billing file is missing expected column(s): {', '.join(missing)}")

    out = pd.DataFrame({
        "UID": df[uid_col].astype(str),
        "NameFamily": (df[name_col].fillna("").astype(str) if name_col else ""),
        "FirstNameFamily": (df[fname_col].fillna("").astype(str) if fname_col else ""),
        "DOB": (pd.to_datetime(df[dob_col], errors="coerce") if dob_col else pd.NaT),
        "AgeNum": (pd.to_numeric(df[age_num_col], errors="coerce") if age_num_col else np.nan),
        "Category": df[cat_col].astype(str).str.strip() if cat_col else "",
        "AgeRange": df[age_col].astype(str).str.strip(),
        "Days": pd.to_numeric(df[days_col], errors="coerce").fillna(0),
        "Country": (df[country_col].fillna("").astype(str) if country_col else ""),
        "PremiumBilled": pd.to_numeric(df[premium_col], errors="coerce").fillna(0),
    })
    if code_col:
        out["Code"] = df[code_col].astype(str).str.strip()
    else:
        out["Code"] = (out["Category"] + " " + out["AgeRange"]).str.strip()
    out = out[out["UID"].notna() & (out["UID"].str.strip().str.lower() != "nan") & (out["UID"].str.strip() != "")]
    return out.reset_index(drop=True)


def load_rates(file) -> pd.DataFrame:
    """Load a rate table. Tries a clean 'Code' + 'Rate' header format first.
    Falls back to scanning the whole sheet for CODE / AGE-RANGE style values
    (e.g. "NA1 0-15") next to a numeric rate — this handles rate cards that
    are laid out as stacked region blocks with repeated sub-headers, rather
    than one clean table. A 'Region' column is included where it can be
    determined (from a header, or from a region-label row like "Orion
    Region 1" appearing above a block of codes) — left blank otherwise."""
    raw = pd.read_excel(file, header=None, sheet_name=0)

    try:
        hdr = find_header_row(raw, must_contain="code")
        df = pd.read_excel(file, header=hdr, sheet_name=0)
        df = df.dropna(how="all")
        cols = list(df.columns)
        code_col = _pick_column(cols, ["code"])
        rate_col = _pick_column(cols, ["premium"], ["rate"], ["amount"])
        region_col = _pick_column(cols, ["region"])
        if code_col is not None and rate_col is not None:
            out = pd.DataFrame({
                "Code": df[code_col].astype(str).str.strip(),
                "Rate": pd.to_numeric(df[rate_col], errors="coerce"),
                "Region": (df[region_col].astype(str).str.strip() if region_col else ""),
            }).dropna(subset=["Rate"])
            if len(out) > 0:
                return out.reset_index(drop=True)
    except ValueError:
        pass

    # --- Fallback: pattern-scan for "CODE AGE-RANGE" style values anywhere,
    # tracking the most recent "Region" label seen above each block ---
    code_pattern = re.compile(r"^[A-Za-z]{2,4}\d?\s+(\d{1,3}-\d{1,3}|\d{1,3}\+)$")
    region_pattern = re.compile(r"region\s*\d+|region\s*[a-z]\b", re.IGNORECASE)
    codes, rates, regions = [], [], []
    n_rows, n_cols = raw.shape
    current_region = ""
    for i in range(n_rows):
        row_has_code = False
        for j in range(n_cols):
            val = raw.iat[i, j]
            if not isinstance(val, str):
                continue
            m = region_pattern.search(val)
            if m and not code_pattern.match(val.strip()):
                current_region = m.group(0).strip().title()
            if not code_pattern.match(val.strip()):
                continue
            row_has_code = True
            rate_val = None
            for k in range(j + 1, n_cols):
                v2 = raw.iat[i, k]
                if pd.notna(v2):
                    try:
                        rate_val = float(str(v2).replace(",", "").replace("$", "").strip())
                        break
                    except ValueError:
                        continue
            if rate_val is not None:
                codes.append(val.strip())
                rates.append(rate_val)
                regions.append(current_region)
        if row_has_code:
            continue

    if codes:
        out = pd.DataFrame({"Code": codes, "Rate": rates, "Region": regions}).drop_duplicates(subset=["Code"])
        return out.reset_index(drop=True)

    raise ValueError(
        "Rate table needs either a 'Code' + rate/premium/amount column, or "
        "rows shaped like 'NA1 0-15   1,236.19' somewhere in the sheet. "
        "Neither pattern was found."
    )


def load_rectification(file) -> pd.DataFrame:
    """Load a rectification file. Returns UID, PremiumRect, PreviousDue, Diff
    (Diff is always recomputed, never trusted from the source file)."""
    raw = pd.read_excel(file, header=None, sheet_name=0)
    hdr = find_header_row(raw, must_contain="unique identifier")
    df = pd.read_excel(file, header=hdr, sheet_name=0)
    df = df.dropna(how="all")
    cols = list(df.columns)
    uid_col = _pick_column(cols, ["unique", "identifier"])
    premium_col = _pick_column(cols, ["premium", "medical"], exclude="taxes")
    prevdue_col = _pick_column(cols, ["previous", "due"])
    if premium_col is None or prevdue_col is None:
        raise ValueError("Rectification file needs a 'Premium medical' column and a 'Previous Due' column.")

    if uid_col is not None:
        uid = df[uid_col].astype(str)
    else:
        uid = pd.Series([f"__row_{i}" for i in range(len(df))])

    out = pd.DataFrame({
        "UID": uid,
        "PremiumRect": pd.to_numeric(df[premium_col], errors="coerce").fillna(0),
        "PreviousDue": pd.to_numeric(df[prevdue_col], errors="coerce").fillna(0),
    })
    out["Diff"] = out["PremiumRect"] - out["PreviousDue"]
    return out.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Reconciliation
# ---------------------------------------------------------------------------

def run_reconciliation(billing_df, rates_df, tolerance_pct=10.0, rate_is_annual=True):
    """Join billing to rates by Code, compute ExpectedUSD, flag CHECK/OK.

    Returns (validation_df, sanity_warning) where sanity_warning is a string
    (empty if nothing looks wrong) describing a likely systemic issue such as
    a rate-table unit mismatch.
    """
    df = billing_df.merge(rates_df, on="Code", how="left")
    df["RateFound"] = df["Rate"].notna()
    df["Rate"] = df["Rate"].fillna(0)
    if "Region" not in df.columns:
        df["Region"] = ""
    df["Region"] = df["Region"].fillna("")
    df = df.rename(columns={"Region": "DerivedRegion"})

    df["MonthlyRate"] = df["Rate"] / 12 if rate_is_annual else df["Rate"]
    monthly_rate = df["MonthlyRate"]
    df["ExpectedUSD"] = np.where(df["Days"] == 0, 0, monthly_rate * df["Days"] / 30)
    df["DifferenceUSD"] = df["PremiumBilled"] - df["ExpectedUSD"]
    df["PercentDiff"] = np.where(
        df["ExpectedUSD"] == 0, np.nan,
        (df["DifferenceUSD"] / df["ExpectedUSD"] * 100).round(2)
    )

    def status(row):
        if not row["RateFound"] and row["Days"] > 0:
            return "No Rate"
        if row["Days"] == 0:
            return "OK"
        if abs(row["PercentDiff"]) > tolerance_pct:
            return "CHECK"
        return "OK"

    df["Status"] = df.apply(status, axis=1)
    df["Reason"] = df.apply(
        lambda r: f"Outside tolerance ({r['PercentDiff']:.2f}%)" if r["Status"] == "CHECK"
        else ("Rate not found for code" if r["Status"] == "No Rate" else ""),
        axis=1,
    )

    # --- Sanity check: does this look like a systemic rate-scale issue? ---
    warning = ""
    checks = df[df["Status"] == "CHECK"]
    if len(df) > 0:
        check_rate = len(checks) / len(df)
        if check_rate > 0.5 and len(checks) >= 5:
            pct_values = checks["PercentDiff"].dropna()
            if len(pct_values) >= 5:
                spread = pct_values.std()
                avg = pct_values.mean()
                if spread < 3 and abs(avg) > 40:
                    factor_hint = ""
                    ratio = abs(avg) / 100
                    if 0.85 < ratio < 0.95:
                        factor_hint = " This is close to what you'd see if an annual rate were applied as monthly (~12x too high) — check the 'rates are monthly/annual' setting."
                    warning = (
                        f"{check_rate:.0%} of records are flagged CHECK, and they all show "
                        f"nearly the same deviation (avg {avg:.1f}%, spread {spread:.1f}pp). "
                        f"This pattern usually means one systemic issue (a rate scale/unit "
                        f"mismatch), not many individual billing errors."
                        f"{factor_hint}"
                    )

    return df, warning


def build_summary(validation_df):
    ok = int((validation_df["Status"] == "OK").sum())
    check = int((validation_df["Status"] == "CHECK").sum())
    norate = int((validation_df["Status"] == "No Rate").sum())
    total_billed = float(validation_df["PremiumBilled"].sum())
    total_expected = float(validation_df["ExpectedUSD"].sum())
    total_diff = float(validation_df["DifferenceUSD"].sum())
    ok_pct = validation_df.loc[validation_df["Status"] == "OK", "PercentDiff"]
    avg_ok_pct = float(ok_pct.dropna().mean()) if ok_pct.notna().any() else None
    return {
        "total_rows": len(validation_df),
        "ok": ok, "check": check, "no_rate": norate,
        "total_billed": total_billed, "total_expected": total_expected,
        "total_diff": total_diff, "avg_ok_pct": avg_ok_pct,
    }


# ---------------------------------------------------------------------------
# Rollups, data quality, and pre-run preview
# ---------------------------------------------------------------------------

def category_rollup(validation_df, group_col="DerivedRegion"):
    """Group results by region (or category) so a systemic issue affecting
    one group is visible, not just the grand total. Falls back to 'Category'
    if DerivedRegion is entirely blank (no region could be determined)."""
    df = validation_df.copy()
    if group_col == "DerivedRegion" and df["DerivedRegion"].replace("", np.nan).isna().all():
        group_col = "Category"

    df["_group"] = df[group_col].replace("", "(unspecified)")
    grouped = df.groupby("_group").agg(
        Records=("UID", "count"),
        OK=("Status", lambda s: (s == "OK").sum()),
        CHECK=("Status", lambda s: (s == "CHECK").sum()),
        NoRate=("Status", lambda s: (s == "No Rate").sum()),
        TotalBilled=("PremiumBilled", "sum"),
        TotalExpected=("ExpectedUSD", "sum"),
        TotalDiff=("DifferenceUSD", "sum"),
        AvgPctDiff=("PercentDiff", "mean"),
    ).reset_index().rename(columns={"_group": group_col})
    return grouped.sort_values("Records", ascending=False).reset_index(drop=True)


def data_quality_checks(billing_df):
    """Best-effort checks for issues that would silently distort totals.
    Returns a list of human-readable warning strings (empty if clean)."""
    issues = []
    dup = billing_df["UID"].value_counts()
    dup = dup[dup > 1]
    if len(dup) > 0:
        issues.append(
            f"{len(dup)} Unique Identifier(s) appear more than once in the billing file "
            f"(e.g. {dup.index[0]} appears {int(dup.iloc[0])} times) — this will double-count "
            f"their premium in the totals."
        )
    neg_days = (billing_df["Days"] < 0).sum()
    if neg_days > 0:
        issues.append(f"{int(neg_days)} record(s) have negative insured days.")
    neg_prem = (billing_df["PremiumBilled"] < 0).sum()
    if neg_prem > 0:
        issues.append(f"{int(neg_prem)} record(s) have a negative billed premium — confirm these are intentional credits, not data errors.")
    blank_code = (billing_df["Code"].str.strip() == "").sum()
    if blank_code > 0:
        issues.append(f"{int(blank_code)} record(s) have no category/age code and can't be matched to a rate.")
    return issues


def preview_billing_columns(file):
    """Detect which source column maps to each expected field, without
    building the full validation. Lets the UI show 'here's what I matched'
    before the person commits to running the reconciliation."""
    raw = pd.read_excel(file, header=None, sheet_name=0)
    hdr = find_header_row(raw)
    df = pd.read_excel(file, header=hdr, sheet_name=0)
    cols = list(df.columns)
    mapping = {
        "Unique Identifier": _pick_column(cols, ["unique", "identifier"]),
        "Name Family": _pick_column(cols, ["name", "family"], ["name"]),
        "First Name Family": _pick_column(cols, ["first", "name", "family"], ["first", "name"]),
        "DOB": _pick_column(cols, ["dob"], ["date", "birth"]),
        "Age": _pick_column(cols, ["age"]),
        "Category": _pick_column(cols, ["category", "family"], ["category"]),
        "Age range": _pick_column(cols, ["age", "range"]),
        "Days insured": _pick_column(cols, ["days", "insured"]),
        "Country name station": _pick_column(cols, ["country", "name", "station"], ["country", "station"]),
        "Code": _pick_column(cols, ["code"]),
        "Premium medical": _pick_column(cols, ["premium", "medical"], exclude="taxes"),
    }
    return mapping, hdr