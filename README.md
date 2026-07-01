# Insurance Premium Reconciliation Tool

Internal tool for VHW finance: upload a billing file and a rate table (and
optionally a rectification file), get back a validated reconciliation
workbook and a summary report — no manual spreadsheet work.

Built from and tested against the June–August 2026 Cigna reconciliation.

## What it does

1. Reads the billing file and finds each member's category + age code.
2. Looks up the expected premium from the rate table (handles annual or
   monthly rate tables — you tell it which, via a toggle).
3. Flags any member whose billed premium differs from expected by more than
   your chosen tolerance.
4. **Automatically flags a likely rate-table scale/unit mismatch** if most
   records show the same large deviation — this is the exact issue we ran
   into with the June–August file, where annual rates were mistakenly
   treated as monthly. The app checks for that pattern automatically now.
5. If given, checks each member's premium against their previous due amount
   and reports new charges (no previous due on file) separately.
6. Produces a downloadable `.xlsx` (Validation / Summary / Exceptions /
   Rectification) and a `.docx` summary report.

## File requirements

**Billing file**: any `.xlsx` with a header row (anywhere in the first ~20
rows) containing columns whose names include "Unique Identifier",
"Category", "Age range", "days insured", and "Premium medical". A "Code"
column (category + age band) is used if present; otherwise it's built from
Category + Age range.

**Rate table**: a `.xlsx` with a "Code" column and a rate/premium/amount
column, one row per code (e.g. `NA1 45-49`, rate).

**Rectification file** (optional): a `.xlsx` with "Unique Identifier",
"Premium medical", and "Previous Due" columns.

Column matching is name-based (not position-based), so reasonable header
variations are fine. If a file fails to load, the error message names which
expected column wasn't found — check the source file's headers.

## Run it locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

This opens in your browser at `http://localhost:8501`. Only you can access
it unless you're on a shared network and someone else browses to your
machine's IP.

## Deploy for the finance department (free, no server to maintain)

1. Create a GitHub repo (can be private) and push this folder to it.
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click "New app", point it at the repo and `app.py`.
3. You get a permanent URL (e.g. `yourapp.streamlit.app`). Share that with
   finance — they open it in a browser, upload files, download results.
   No installs needed on their end.
4. If the repo is private, Streamlit Community Cloud still works — it asks
   for GitHub access when you deploy, and the app URL itself is only shared
   with people you give it to (no public listing unless you make the repo
   public).

If VHW prefers not to use a third-party host for financial data, the same
`app.py` runs identically on any internal server with Python installed —
just run `streamlit run app.py --server.port 8501` there instead and share
the internal URL.

## Known limitations (worth knowing before relying on this for real)

- Column matching is header-name based. If a source system changes its
  header wording drastically, matching may fail — the error message will
  say which column it couldn't find.
- The sanity-check warning is a heuristic (uniform deviation across most
  records), not a guarantee. Always read the warning if it appears, but
  don't assume its absence means the numbers are definitely right.
- This has been tested against one real dataset (June–August 2026,
  94 members). Test it against a couple more periods before fully trusting
  it for a live approval decision — treat this first version as a strong
  starting point, not a finished audited tool.
