# Bookkeeping Operations Review

A Streamlit app that turns a general-ledger journal export into a bookkeeper's operations and
controls dashboard: posting volume and workload, timeliness, off-hours activity, journal-entry tests,
AR/AP settlement and a data-quality report. Upload a file, or explore the built-in sample.

**Live demo:** _add your Streamlit link here_

## What it reads

| Input | How it is recognised |
|---|---|
| QuickBooks Online **Journal** report exported to Excel or CSV | header row with Date, Transaction Type, Num, Name, Account, Debit, Credit |
| The `GL_Bookkeeping_Operations.xlsx` workbook | sheet named `FactJournalLines` (and `DimAccount` for account classes) |
| CSV with the app's own column names | `PostingID, PostingDate, AccountName, Debit, Credit, ...` |
| Any other CSV / Excel with one row per journal line | a column-mapping screen appears in the sidebar |

Required: an entry ID (or transaction type + number), a date, an account, and debit/credit (or one signed amount).
Optional columns switch on more pages: prepared-by (workload), entered date-time (timeliness, timing red flags),
document date, customer/supplier name (duplicate test, settlement matching), cleared-against reference (exact settlement).

Uploaded files stay in the visitor's browser session and are never stored.

## Files

```
app.py                         the whole app
requirements.txt               Python packages Streamlit Cloud installs
.streamlit/config.toml         colours and fonts
data/sample_journal.csv.gz     built-in sample (synthetic garden-tools retailer, 2023-2024)
data/sample_accounts.csv       chart of accounts for the sample
data/example_quickbooks_journal.xlsx   a QuickBooks-style export to try the upload with
data/example_generic_export.csv        an odd-format CSV to try the mapping screen with
```

Sample data source: JGro2026/General-Ledger-Data on GitHub (synthetic journal published with a research paper).

## Deploy for free on Streamlit Community Cloud

1. Create a free account at **github.com** (skip if you have one).
2. On GitHub, click **New repository**, name it `bookkeeping-review`, keep it Public, tick *Add a README*, click **Create repository**.
3. Click **Add file → Upload files** and drag in every file and folder from this package (`app.py`, `requirements.txt`, the `.streamlit` folder, the `data` folder). Click **Commit changes**.
   - If the `.streamlit` folder does not upload, create it by hand: **Add file → Create new file**, type `.streamlit/config.toml` as the name and paste the contents.
4. Go to **share.streamlit.io** and sign in with your GitHub account.
5. Click **Create app → Deploy a public app from GitHub**. Choose the `bookkeeping-review` repository, branch `main`, main file `app.py`. Pick a short app URL. Click **Deploy**.
6. Wait two to three minutes. Your app is live at `https://<your-url>.streamlit.app`.

To update the app later, edit `app.py` on GitHub (or upload a new copy) and commit — the live app rebuilds itself within a minute.

## Run on your own computer (optional)

```
pip install -r requirements.txt
streamlit run app.py
```

## Settings (sidebar)

On-time threshold (default 7 days), working hours (08:00-17:00), month-end window (last 5 days),
cash-payment limit (1,500), rarely-used-account threshold (5 postings a year). Every page recalculates
when a setting changes.
