"""Bookkeeping Operations Review — a Streamlit app that reads a general-ledger journal export
and produces a bookkeeper's operations and controls dashboard.

Supported inputs: the GL_Bookkeeping_Operations workbook (FactJournalLines sheet), a QuickBooks Online
Journal report export (xlsx/csv), a CSV with the app's own column names, or any CSV/xlsx via the column
mapping screen. Uploaded files stay in the visitor's session and are never stored.
"""
import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------------------- setup
st.set_page_config(page_title="Bookkeeping Operations Review", page_icon="📒", layout="wide")

PALETTE = ["#1F3864", "#2E75B6", "#9DC3E6", "#7F7F7F", "#C9A227", "#C0504D", "#70AD47", "#BF9000"]
FLAG_COLOR = "#C0504D"
px.defaults.color_discrete_sequence = PALETTE
px.defaults.template = "simple_white"

DATA_DIR = Path(__file__).parent / "data"
CANON = ["PostingID", "PostingDate", "DocumentDate", "EnteredDateTime", "AccountNo", "AccountName",
         "Debit", "Credit", "UserID", "Counterparty", "RefPostingID", "LineText"]
SYNONYMS = {
    "PostingID": ["postingid", "entryid", "entry id", "journal id", "journalid", "transaction id", "txnid", "num", "reference", "entry no", "je number", "voucher"],
    "PostingDate": ["postingdate", "posting date", "date", "txndate", "transaction date", "gl date"],
    "DocumentDate": ["documentdate", "document date", "invoice date", "doc date", "bill date"],
    "EnteredDateTime": ["entereddatetime", "entered", "created", "create date", "created at", "entry time", "timestamp", "date entered"],
    "AccountNo": ["accountno", "account no", "account number", "account code", "gl account", "acct"],
    "AccountName": ["accountname", "account name", "account", "gl account name", "ledger account"],
    "Debit": ["debit", "dr", "debit amount"],
    "Credit": ["credit", "cr", "credit amount"],
    "UserID": ["userid", "user", "created by", "prepared by", "entered by", "preparer", "posted by"],
    "Counterparty": ["counterparty", "name", "customer", "vendor", "supplier", "contact", "payee"],
    "RefPostingID": ["refpostingid", "ref_id", "cleared against", "applied to", "reference id"],
    "LineText": ["linetext", "memo", "description", "memo/description", "narration", "text", "details"],
}
LIQUID_KEYWORDS = ["cash", "bank", "petty", "checking", "savings", "current account"]
BUCKETS = ["0-30", "31-60", "61-90", "90+"]


# --------------------------------------------------------------------------------------- reading files
def _norm(s):
    return re.sub(r"[^a-z0-9/]", "", str(s).lower())


def _guess(columns, field):
    cols = list(columns)
    normed = {_norm(c): c for c in cols}
    for syn in SYNONYMS[field]:
        n = _norm(syn)
        if n in normed:
            return normed[n]
    for syn in SYNONYMS[field]:
        n = _norm(syn)
        for k, c in normed.items():
            if len(n) >= 4 and n in k:
                return c
    return None


def _find_header_row(raw: pd.DataFrame):
    """Locate the header row in exports that start with title lines (QuickBooks, Xero)."""
    for i in range(min(25, len(raw))):
        vals = [_norm(v) for v in raw.iloc[i].tolist()]
        hits = sum(1 for v in vals if v in {"debit", "credit", "account", "date", "transactiontype", "accountno", "postingid", "amount"})
        if hits >= 3:
            return i
    return 0


def read_any(file_bytes: bytes, name: str):
    """Return (frames dict, kind). kind in {'workbook','canonical','qbo','unknown'}."""
    lower = name.lower()
    if lower.endswith((".xlsx", ".xlsm", ".xls")):
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        if "FactJournalLines" in xl.sheet_names:
            frames = {"lines": xl.parse("FactJournalLines")}
            if "DimAccount" in xl.sheet_names:
                frames["accounts"] = xl.parse("DimAccount")
            return frames, "workbook"
        raw = xl.parse(xl.sheet_names[0], header=None)
    else:
        text = file_bytes.decode("utf-8-sig", errors="replace")
        sep = ";" if text[:2000].count(";") > text[:2000].count(",") else ","
        raw = pd.read_csv(io.StringIO(text), sep=sep, header=None, dtype=str, keep_default_na=False, engine="python")
    h = _find_header_row(raw)
    df = raw.iloc[h + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[h].tolist()]
    df = df.loc[:, [c for c in df.columns if c and c.lower() != "nan"]]
    df = df.dropna(how="all")
    cols = {_norm(c) for c in df.columns}
    if {"transactiontype", "debit", "credit", "account"} <= cols:
        return {"lines": df}, "qbo"
    if {"postingid", "accountname", "debit", "credit"} <= cols:
        return {"lines": df}, "canonical"
    return {"lines": df}, "unknown"


def _parse_dates(s: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    s = s.astype(str).str.strip().replace({"": None, "nan": None, "None": None})
    a = pd.to_datetime(s, errors="coerce", dayfirst=False)
    b = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return b if b.notna().sum() > a.notna().sum() else a


def _parse_amount(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return s.fillna(0.0).astype(float)
    t = s.astype(str).str.strip()
    neg = t.str.startswith("(") & t.str.endswith(")")
    t = t.str.replace(r"[^0-9.\-]", "", regex=True).replace({"": "0", "-": "0", ".": "0"})
    v = pd.to_numeric(t, errors="coerce").fillna(0.0)
    return np.where(neg, -v, v)


def to_canonical(df: pd.DataFrame, mapping: dict, amount_mode: str) -> pd.DataFrame:
    """mapping: canonical field -> source column (or None). amount_mode: 'split' or 'signed'."""
    out = pd.DataFrame(index=df.index)
    for f in CANON:
        src = mapping.get(f)
        out[f] = df[src] if src in df.columns else None
    if amount_mode == "signed":
        amt = _parse_amount(df[mapping["Amount"]])
        out["Debit"] = np.where(amt > 0, amt, 0.0)
        out["Credit"] = np.where(amt < 0, -amt, 0.0)
    else:
        out["Debit"] = _parse_amount(out["Debit"]) if out["Debit"].notna().any() else 0.0
        out["Credit"] = _parse_amount(out["Credit"]) if out["Credit"].notna().any() else 0.0
    for c in ["PostingDate", "DocumentDate", "EnteredDateTime"]:
        out[c] = _parse_dates(out[c]) if out[c].notna().any() else pd.NaT
    for c in ["PostingID", "AccountNo", "AccountName", "UserID", "Counterparty", "RefPostingID", "LineText"]:
        out[c] = out[c].astype(str).str.strip().replace({"nan": None, "None": None, "": None}) if out[c].notna().any() else None
    if out["AccountName"].isna().all() and out["AccountNo"].notna().any():
        out["AccountName"] = out["AccountNo"]
    if out["AccountNo"].isna().all():
        out["AccountNo"] = out["AccountName"]
    # drop rows with no account and no amount (titles, totals, separators)
    out = out[~(out["AccountName"].isna() & (out["Debit"].abs() + out["Credit"].abs() == 0))]
    out = out[~out["AccountName"].astype(str).str.lower().str.startswith("total")]
    if out["PostingID"].isna().all():
        key = out["PostingDate"].dt.strftime("%Y%m%d").fillna("nodate") + "-" + out["Counterparty"].fillna("").astype(str)
        out["PostingID"] = "E" + pd.factorize(key)[0].astype(str)
    out["PostingID"] = out["PostingID"].ffill()
    return out.reset_index(drop=True)


def qbo_to_canonical(df: pd.DataFrame) -> pd.DataFrame:
    cols = {_norm(c): c for c in df.columns}
    d = df.copy()
    has_account = d[cols["account"]].astype(str).str.strip().replace({"nan": ""}) != ""
    for k in ["date", "transactiontype", "num", "name"]:
        if k in cols:
            d[cols[k]] = d[cols[k]].replace({"": None, "nan": None}).ffill()
    d = d[has_account]
    typ = d[cols["transactiontype"]].astype(str).str.strip()
    num = d[cols["num"]].astype(str).str.strip().replace({"nan": "", "None": ""}) if "num" in cols else ""
    dt = _parse_dates(d[cols["date"]]).dt.strftime("%Y-%m-%d")
    d["__pid"] = np.where(num != "", typ + " " + num, typ + " " + dt)
    mapping = {"PostingID": "__pid", "PostingDate": cols["date"], "AccountName": cols["account"], "Debit": cols["debit"], "Credit": cols["credit"],
               "Counterparty": cols.get("name"), "LineText": cols.get("memo/description", cols.get("memo", cols.get("description")))}
    return to_canonical(d, mapping, "split")


# --------------------------------------------------------------------------------------- classification
def classify_accounts(lines: pd.DataFrame, provided: pd.DataFrame | None):
    acc = lines.groupby("AccountNo").agg(AccountName=("AccountName", "first"), Lines=("PostingID", "size"),
                                         Debits=("Debit", "sum"), Credits=("Credit", "sum")).reset_index()
    if provided is not None and {"AccountNo", "AccountClass", "AccountGroup"} <= set(provided.columns):
        p = provided.copy()
        p["AccountNo"] = p["AccountNo"].astype(str).str.strip()
        acc = acc.merge(p[["AccountNo", "AccountClass", "AccountGroup"]], on="AccountNo", how="left")
        acc["Source"] = np.where(acc["AccountClass"].notna(), "chart of accounts", "keyword guess")
    else:
        acc["AccountClass"] = None
        acc["AccountGroup"] = None
        acc["Source"] = "keyword guess"
    rules = [
        (["receivable", "debtors"], "Asset", "Receivables"),
        (["vat", "gst", "sales tax", "input tax", "tax payable", "hst"], "Liability", "VAT / sales tax"),
        (["payable", "creditors"], "Liability", "Payables"),
        (["cash", "bank", "petty", "checking", "savings", "current account"], "Asset", "Cash & bank"),
        (["cost of", "purchase", "cogs", "merchandise", "goods sold"], "Expense", "Cost of goods"),
        (["salar", "wage", "payroll", "social security", "pension"], "Expense", "Personnel"),
        (["revenue", "sales", "income", "fees earned", "turnover"], "Revenue", "Revenue"),
        (["loan", "credit card", "mortgage", "accrued"], "Liability", "Other liabilities"),
        (["equity", "capital", "retained", "drawings", "owner"], "Equity", "Equity"),
        (["inventory", "stock", "equipment", "furniture", "deposit", "prepaid", "fixed asset"], "Asset", "Other assets"),
    ]
    for i, row in acc.iterrows():
        if pd.notna(row["AccountClass"]):
            continue
        name = str(row["AccountName"]).lower()
        name = re.sub(r"\(.*?\)", " ", name)                                   # drop bracketed qualifiers
        name = re.sub(r",?\s*\d+(\.\d+)?\s*%\s*(vat|input tax|tax|gst)?", " ", name)  # drop "19% VAT"-style suffixes
        cls, grp = "Expense", "Other operating"
        for kws, c, g in rules:
            if any(k in name for k in kws):
                cls, grp = c, g
                break
        if grp == "VAT / sales tax" and any(k in name for k in ["input", "deductible", "receivable", "recoverable"]):
            cls = "Asset"
        acc.at[i, "AccountClass"], acc.at[i, "AccountGroup"] = cls, grp
    acc["IsLiquid"] = acc["AccountGroup"].eq("Cash & bank")
    return acc


def build_postings(lines: pd.DataFrame, acc: pd.DataFrame, cfg: dict):
    ln = lines.merge(acc[["AccountNo", "AccountClass", "AccountGroup", "IsLiquid"]], on="AccountNo", how="left")
    g = ln.groupby("PostingID", sort=False)
    post = g.agg(PostingDate=("PostingDate", "first"), DocumentDate=("DocumentDate", "first"), EnteredDateTime=("EnteredDateTime", "first"),
                 UserID=("UserID", "first"), Counterparty=("Counterparty", "first"), RefPostingID=("RefPostingID", "first"),
                 LineCount=("AccountNo", "size"), Debits=("Debit", "sum"), Credits=("Credit", "sum")).reset_index()
    post["PostingAmount"] = post["Debits"].round(2)
    post["Imbalance"] = (post["Debits"] - post["Credits"]).round(2)
    post["IsBalanced"] = post["Imbalance"].abs() < 0.01

    is_cash_name = ln.AccountName.astype(str).str.lower().str.contains("cash|petty", regex=True)
    dr, cr = ln.Debit > 0, ln.Credit > 0
    feats = pd.DataFrame({
        "PostingID": ln.PostingID,
        "DrAP": dr & (ln.AccountGroup == "Payables"), "CrAP": cr & (ln.AccountGroup == "Payables"),
        "DrAR": dr & (ln.AccountGroup == "Receivables"), "CrAR": cr & (ln.AccountGroup == "Receivables"),
        "DrLiquid": dr & ln.IsLiquid.fillna(False), "CrLiquid": cr & ln.IsLiquid.fillna(False),
        "CrRevenue": cr & (ln.AccountClass == "Revenue"), "DrExpense": dr & (ln.AccountClass == "Expense"),
        "Personnel": ln.AccountGroup == "Personnel",
        "LiquidAmount": np.where(ln.IsLiquid.fillna(False), ln.Debit + ln.Credit, 0.0),
        "TouchesCash": is_cash_name,
        "CashAmount": np.where(is_cash_name, ln.Debit + ln.Credit, 0.0),
    })
    f = feats.groupby("PostingID", sort=False).agg({"DrAP": "any", "CrAP": "any", "DrAR": "any", "CrAR": "any", "DrLiquid": "any", "CrLiquid": "any",
                                                    "CrRevenue": "any", "DrExpense": "any", "Personnel": "any", "LiquidAmount": "sum", "TouchesCash": "any", "CashAmount": "sum"}).reset_index()
    post = post.merge(f, on="PostingID", how="left")
    post["PostingType"] = np.select(
        [post.CrAP & post.DrExpense, post.CrRevenue & post.DrAR, post.CrRevenue, post.DrAP & post.CrLiquid, post.CrAR & post.DrLiquid, post.Personnel, post.DrExpense],
        ["Purchase invoice", "Sales invoice", "Cash sale", "Supplier payment", "Customer receipt", "Payroll", "Operating expense"], default="Other")

    # timeliness
    base = post["DocumentDate"].where(post["DocumentDate"].notna(), post["PostingDate"])
    entered_day = post["EnteredDateTime"].dt.normalize()
    post["LagDays"] = (entered_day - base).dt.days
    post["HasLag"] = post["LagDays"].notna()
    post["LagBand"] = pd.cut(post["LagDays"], bins=[-10000, 3, 7, 14, 30, 100000], labels=["0-3 days", "4-7 days", "8-14 days", "15-30 days", "31+ days"]).astype(object)
    post["IsOnTime"] = post["LagDays"] <= cfg["on_time_days"]
    post["EnteredHour"] = post["EnteredDateTime"].dt.hour
    post["EnteredWeekday"] = post["EnteredDateTime"].dt.day_name()
    post["IsWeekend"] = post["EnteredDateTime"].dt.dayofweek >= 5
    post["IsAfterHours"] = post["EnteredHour"].notna() & ((post["EnteredHour"] < cfg["work_start"]) | (post["EnteredHour"] >= cfg["work_end"]))
    post["Year"] = post["PostingDate"].dt.year
    post["YearMonth"] = post["PostingDate"].dt.to_period("M").astype(str)
    post["IsMonthEndWeek"] = (post["PostingDate"] + pd.offsets.MonthEnd(0) - post["PostingDate"]).dt.days <= cfg["month_end_days"] - 1
    return post, ln


# --------------------------------------------------------------------------------------- journal-entry tests
def run_tests(post: pd.DataFrame, ln: pd.DataFrame, acc: pd.DataFrame, cfg: dict):
    out = []

    def add(mask, test, detail):
        sub = post.loc[mask, ["PostingID", "PostingDate", "PostingType", "UserID", "PostingAmount"]].copy()
        sub["Test"] = test
        sub["Detail"] = detail(sub) if callable(detail) else detail
        out.append(sub)

    add(~post.IsBalanced, "Unbalanced entry", lambda s: "Debits differ from credits by " + post.set_index("PostingID").loc[s.PostingID, "Imbalance"].abs().round(2).astype(str).values)
    add(post.TouchesCash & (post.CashAmount > cfg["cash_limit"]), "Large amount through cash", f"Cash-account movement above {cfg['cash_limit']:,.0f}")
    # duplicates: same date, amount, type and counterparty (if known) with different ids
    if post.Counterparty.notna().any():
        dup = post.duplicated(subset=["PostingDate", "PostingAmount", "PostingType", "Counterparty"], keep=False) & post.PostingAmount.gt(0) & post.Counterparty.notna()
        add(dup, "Possible duplicate", "Same date, amount, type and counterparty as another entry")
    if post.DocumentDate.notna().any():
        add(post.DocumentDate.notna() & (post.PostingDate < post.DocumentDate), "Back-dated entry", "Posting date is earlier than the document date")
    # dormant accounts
    yr = ln.PostingDate.dt.year
    per_year = ln.groupby([ln.AccountNo, yr]).PostingID.nunique()
    dormant_keys = set(per_year[per_year <= cfg["dormant_max"]].index)
    dormant_mask = pd.Series(list(zip(ln.AccountNo, yr)), index=ln.index).isin(dormant_keys)
    dormant_ids = set(ln.loc[dormant_mask, "PostingID"])
    if dormant_ids:
        names = ln[ln.PostingID.isin(dormant_ids)].groupby("PostingID").AccountName.first()
        add(post.PostingID.isin(dormant_ids), "Rarely used account", lambda s: "Account with few postings in the year: " + names.reindex(s.PostingID).fillna("").values)
    if post.EnteredDateTime.notna().any():
        add((post.IsWeekend | post.IsAfterHours) & (post.LiquidAmount > cfg["cash_limit"]), "Weekend / after-hours on cash or bank",
            lambda s: np.where(post.set_index("PostingID").loc[s.PostingID, "IsWeekend"].values, "Entered on a weekend", "Entered outside working hours"))
        add(post.HasLag & (post.LagDays > 30), "Very late entry", lambda s: "Keyed " + post.set_index("PostingID").loc[s.PostingID, "LagDays"].astype(int).astype(str).values + " days after the document date")
    if post.RefPostingID.notna().any():
        typ = post.set_index("PostingID").PostingType
        ref_type = post.RefPostingID.map(typ)
        wrong = post.RefPostingID.notna() & (((post.PostingType == "Supplier payment") & (ref_type == "Sales invoice")) | ((post.PostingType == "Customer receipt") & (ref_type == "Purchase invoice")))
        add(wrong, "Payment cleared against wrong invoice", "Payment references an invoice of the opposite type")
    flags = pd.concat(out, ignore_index=True) if out else pd.DataFrame(columns=["PostingID", "PostingDate", "PostingType", "UserID", "PostingAmount", "Test", "Detail"])
    return flags


# --------------------------------------------------------------------------------------- settlement
def build_settlements(post: pd.DataFrame):
    inv = post[post.PostingType.isin(["Purchase invoice", "Sales invoice"])][["PostingID", "PostingType", "PostingDate", "DocumentDate", "PostingAmount", "Counterparty"]].copy()
    inv["InvoiceDate"] = inv.DocumentDate.where(inv.DocumentDate.notna(), inv.PostingDate)
    inv["Counterparty type"] = np.where(inv.PostingType == "Purchase invoice", "Supplier (AP)", "Customer (AR)")
    pay = post[post.PostingType.isin(["Supplier payment", "Customer receipt"])][["PostingID", "PostingType", "PostingDate", "PostingAmount", "Counterparty", "RefPostingID", "TouchesCash"]].copy()
    pay["PaymentDate"] = pay.PostingDate
    if post.RefPostingID.notna().any():
        p = pay.dropna(subset=["RefPostingID"]).sort_values("PaymentDate").drop_duplicates("RefPostingID")
        m = inv.merge(p[["RefPostingID", "PostingID", "PaymentDate", "TouchesCash"]].rename(columns={"PostingID": "PaymentID", "RefPostingID": "PostingID"}), on="PostingID", how="left")
        method = "reference column"
    elif post.Counterparty.notna().any():
        # FIFO match on counterparty + exact amount
        inv_s = inv.sort_values("InvoiceDate").copy()
        pay_s = pay.sort_values("PaymentDate").copy()
        used = set()
        match = {}
        pay_idx = pay_s.groupby(["Counterparty", "PostingAmount"]).apply(lambda d: list(d.index), include_groups=False).to_dict()
        for i, r in inv_s.iterrows():
            cands = pay_idx.get((r.Counterparty, r.PostingAmount), [])
            for j in cands:
                if j in used: continue
                if pay_s.loc[j, "PaymentDate"] >= r.InvoiceDate - pd.Timedelta(days=5):
                    used.add(j); match[r.PostingID] = j; break
        inv["PaymentID"] = inv.PostingID.map({k: pay_s.loc[v, "PostingID"] for k, v in match.items()})
        inv["PaymentDate"] = pd.to_datetime(inv.PostingID.map({k: pay_s.loc[v, "PaymentDate"] for k, v in match.items()}))
        inv["TouchesCash"] = inv.PostingID.map({k: pay_s.loc[v, "TouchesCash"] for k, v in match.items()})
        m = inv
        method = "counterparty and amount matching"
    else:
        return None, None
    m["IsSettled"] = m.PaymentID.notna()
    m["DaysToSettle"] = (m.PaymentDate - m.InvoiceDate).dt.days
    period_end = post.PostingDate.max()
    m["AgeAtPeriodEnd"] = (period_end - m.InvoiceDate).dt.days
    m["AgeBucket"] = pd.cut(m.AgeAtPeriodEnd, bins=[-1, 30, 60, 90, 100000], labels=BUCKETS).astype(object)
    m["PaidVia"] = np.where(m.IsSettled, np.where(m.TouchesCash.fillna(False), "Cash", "Bank"), "Open")
    return m, method


# --------------------------------------------------------------------------------------- data quality
def data_quality(lines: pd.DataFrame, post: pd.DataFrame, acc: pd.DataFrame):
    issues = []
    def item(name, df, note):
        issues.append({"Check": name, "Count": len(df), "Note": note, "Rows": df})
    item("Unbalanced entries", post[~post.IsBalanced][["PostingID", "PostingDate", "Imbalance"]], "Debits and credits differ within one entry")
    item("Lines with no date", lines[lines.PostingDate.isna()][["PostingID", "AccountName", "Debit", "Credit"]], "Date missing or unreadable")
    item("Lines with no account", lines[lines.AccountName.isna()][["PostingID", "PostingDate", "Debit", "Credit"]], "Account blank")
    item("Lines with both debit and credit", lines[(lines.Debit > 0) & (lines.Credit > 0)][["PostingID", "AccountName", "Debit", "Credit"]], "A line should carry one side only")
    item("Accounts classified by keyword guess", acc[acc.Source == "keyword guess"][["AccountNo", "AccountName", "AccountClass", "AccountGroup"]], "Review these classes in the expander below")
    item("Entries with no preparer", post[post.UserID.isna()][["PostingID", "PostingDate", "PostingAmount"]], "Workload and user-based tests skip these")
    dup_lines = lines[lines.duplicated(subset=["PostingID", "AccountNo", "Debit", "Credit"], keep=False)]
    item("Duplicate lines inside an entry", dup_lines[["PostingID", "AccountName", "Debit", "Credit"]], "Same account and amount repeated within one entry")
    return issues


# --------------------------------------------------------------------------------------- helpers
def pct(x):
    return "—" if x is None or (isinstance(x, float) and np.isnan(x)) else f"{x:.1%}"


def money(x):
    return f"{x:,.0f}"


def kpi_row(items):
    cols = st.columns(len(items))
    for c, (label, value, help_) in zip(cols, items):
        c.metric(label, value, help=help_)


def bar(df, x, y, title, color=None, orientation="v", **kw):
    fig = px.bar(df, x=x, y=y, color=color, title=title, orientation=orientation, **kw)
    fig.update_layout(margin=dict(l=10, r=10, t=40, b=10), legend_title_text="", height=360)
    return fig


def findings_text(post, flags, sett, cfg):
    lines_ = []
    n = len(post)
    if post.HasLag.any():
        late = 1 - post.loc[post.HasLag, "IsOnTime"].mean()
        lines_.append(f"{late:.0%} of entries were keyed more than {cfg['on_time_days']} days after the document date (average lag {post.LagDays.mean():.1f} days).")
    if len(flags):
        top = flags.Test.value_counts().head(3)
        lines_.append("Journal-entry tests flagged " + f"{flags.PostingID.nunique():,} entries ({flags.PostingID.nunique()/n:.1%}) worth {money(flags.drop_duplicates('PostingID').PostingAmount.sum())}; most common: " + ", ".join(f"{k} ({v})" for k, v in top.items()) + ".")
    if post.UserID.notna().any() and len(flags):
        by_user = flags.drop_duplicates("PostingID").groupby("UserID").size() / post.groupby("UserID").size()
        by_user = by_user.dropna().sort_values(ascending=False)
        if len(by_user):
            u = by_user.index[0]
            lines_.append(f"{u} has the highest flag rate ({by_user.iloc[0]:.1%} of their entries) on {post.UserID.eq(u).mean():.0%} of total volume — review access and approval for this user first.")
    if sett is not None:
        open_ = sett[~sett.IsSettled]
        if len(open_):
            lines_.append(f"{len(open_)} invoices ({money(open_.PostingAmount.sum())}) were still open at period end; {(open_.AgeBucket == '90+').sum()} are over 90 days old.")
        cash = (sett.PaidVia == "Cash").mean()
        if cash > 0.1:
            lines_.append(f"{cash:.0%} of settled invoices were paid through cash rather than bank — consider a cash-payment limit.")
    if not post.IsBalanced.all():
        lines_.insert(0, f"{(~post.IsBalanced).sum()} entries do not balance — fix these before relying on any balance.")
    return lines_[:5]


def summary_workbook(post, flags, sett, acc, issues):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        yr = post.groupby("Year").agg(Postings=("PostingID", "size"), TotalDebits=("PostingAmount", "sum"), AvgLagDays=("LagDays", "mean"),
                                      OnTimeRate=("IsOnTime", "mean"), WeekendRate=("IsWeekend", "mean"), AfterHoursRate=("IsAfterHours", "mean"))
        yr.to_excel(xw, sheet_name="By year")
        if post.UserID.notna().any():
            u = post.groupby("UserID").agg(Postings=("PostingID", "size"), Share=("PostingID", lambda s: len(s) / len(post)), AvgLagDays=("LagDays", "mean"), OnTimeRate=("IsOnTime", "mean"))
            u["Flagged"] = flags.drop_duplicates("PostingID").groupby("UserID").size()
            u.to_excel(xw, sheet_name="By preparer")
        post.groupby("YearMonth").agg(Postings=("PostingID", "size"), AvgLagDays=("LagDays", "mean"), OnTimeRate=("IsOnTime", "mean")).to_excel(xw, sheet_name="By month")
        flags.to_excel(xw, sheet_name="Flags", index=False)
        if sett is not None:
            sett.drop(columns=["TouchesCash"], errors="ignore").to_excel(xw, sheet_name="Settlements", index=False)
        acc.to_excel(xw, sheet_name="Accounts", index=False)
        pd.DataFrame([{"Check": i["Check"], "Count": i["Count"], "Note": i["Note"]} for i in issues]).to_excel(xw, sheet_name="Data quality", index=False)
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def load_sample():
    lines = pd.read_csv(DATA_DIR / "sample_journal.csv.gz", dtype={"PostingID": str, "AccountNo": str, "RefPostingID": str})
    acc = pd.read_csv(DATA_DIR / "sample_accounts.csv", dtype={"AccountNo": str})
    lines = to_canonical(lines, {c: c for c in CANON}, "split")
    return lines, acc


@st.cache_data(show_spinner=False)
def analyse(lines: pd.DataFrame, acc_provided: pd.DataFrame | None, cfg: dict, overrides: dict):
    acc = classify_accounts(lines, acc_provided)
    for k, (cls, grp) in overrides.items():
        acc.loc[acc.AccountNo == k, ["AccountClass", "AccountGroup", "Source"]] = [cls, grp, "your edit"]
    acc["IsLiquid"] = acc.AccountGroup.eq("Cash & bank")
    post, ln = build_postings(lines, acc, cfg)
    flags = run_tests(post, ln, acc, cfg)
    sett, method = build_settlements(post)
    issues = data_quality(lines, post, acc)
    return acc, post, ln, flags, sett, method, issues


# --------------------------------------------------------------------------------------- sidebar: data in
st.sidebar.title("Bookkeeping operations review")
uploaded = st.sidebar.file_uploader("Upload a journal export", type=["xlsx", "xlsm", "xls", "csv"],
                                    help="QuickBooks Online Journal report, the GL_Bookkeeping_Operations workbook, or any CSV/Excel with one row per journal line. Files stay in this session and are never stored.")
st.sidebar.caption("Nothing uploaded yet? The dashboard below runs on a public sample: a synthetic two-year journal of a garden-tools retailer.")

with st.sidebar.expander("Settings", expanded=False):
    cfg = {
        "on_time_days": st.number_input("On-time threshold (days after document)", 1, 60, 7),
        "work_start": st.number_input("Working hours start", 0, 23, 8),
        "work_end": st.number_input("Working hours end", 1, 24, 17),
        "month_end_days": st.number_input("Month-end window (last N days)", 1, 10, 5),
        "cash_limit": st.number_input("Cash-payment limit", 0, 1_000_000, 1500, step=100),
        "dormant_max": st.number_input("Rarely used account: max postings per year", 1, 50, 5),
    }

lines = None
acc_provided = None
source_label = "Sample data"
if uploaded is not None:
    frames, kind = read_any(uploaded.getvalue(), uploaded.name)
    raw = frames["lines"]
    if kind == "workbook":
        lines = to_canonical(raw, {c: c for c in CANON}, "split")
        acc_provided = frames.get("accounts")
        source_label = f"{uploaded.name} (workbook format)"
    elif kind == "qbo":
        lines = qbo_to_canonical(raw)
        source_label = f"{uploaded.name} (QuickBooks Online journal)"
    elif kind == "canonical":
        lines = to_canonical(raw, {c: _guess(raw.columns, c) for c in CANON}, "split")
        source_label = f"{uploaded.name}"
    else:
        st.sidebar.warning("Layout not recognised — map the columns below.")
        with st.sidebar.form("mapping"):
            st.write("Which column holds each item?")
            opts = ["(none)"] + list(raw.columns)
            def pick(label, field, required=False):
                g = _guess(raw.columns, field)
                return st.selectbox(label + (" *" if required else ""), opts, index=opts.index(g) if g in opts else 0)
            m = {"PostingID": pick("Entry ID", "PostingID"), "PostingDate": pick("Transaction date", "PostingDate", True),
                 "AccountName": pick("Account", "AccountName", True), "AccountNo": pick("Account number", "AccountNo")}
            amount_mode = st.radio("Amounts are", ["Debit and credit columns", "One signed amount column"])
            m["Debit"] = pick("Debit", "Debit"); m["Credit"] = pick("Credit", "Credit")
            m["Amount"] = st.selectbox("Amount (if one column)", opts, index=0)
            m["UserID"] = pick("Prepared by / user", "UserID"); m["EnteredDateTime"] = pick("Entered date-time", "EnteredDateTime")
            m["DocumentDate"] = pick("Document date", "DocumentDate"); m["Counterparty"] = pick("Customer / supplier name", "Counterparty")
            m["RefPostingID"] = pick("Cleared-against reference", "RefPostingID"); m["LineText"] = pick("Memo / description", "LineText")
            ok = st.form_submit_button("Apply mapping")
        if ok:
            m = {k: (None if v == "(none)" else v) for k, v in m.items()}
            missing = [k for k in ["PostingDate", "AccountName"] if m.get(k) is None]
            if amount_mode.startswith("One") and m.get("Amount") is None:
                missing.append("Amount")
            if amount_mode.startswith("Debit") and (m.get("Debit") is None or m.get("Credit") is None):
                missing.append("Debit / Credit")
            if missing:
                st.sidebar.error("Still needed: " + ", ".join(missing))
            else:
                lines = to_canonical(raw, m, "signed" if amount_mode.startswith("One") else "split")
                source_label = f"{uploaded.name} (mapped)"
        if lines is None:
            st.info("Map the required columns in the sidebar to analyse this file. The sample data is shown meanwhile.")
if lines is None:
    lines, acc_provided = load_sample()

# account class overrides live in session state
if "acc_overrides" not in st.session_state:
    st.session_state.acc_overrides = {}
if st.session_state.get("acc_source") != source_label:
    st.session_state.acc_overrides = {}
    st.session_state.acc_source = source_label

acc, post_all, ln_all, flags_all, sett_all, sett_method, issues = analyse(lines, acc_provided, cfg, st.session_state.acc_overrides)

# --------------------------------------------------------------------------------------- sidebar: filters
st.sidebar.markdown("---")
dmin, dmax = post_all.PostingDate.min(), post_all.PostingDate.max()
if pd.isna(dmin):
    st.error("No readable dates in this file. Check the date column mapping.")
    st.stop()
date_range = st.sidebar.date_input("Period", (dmin.date(), dmax.date()), min_value=dmin.date(), max_value=dmax.date())
if isinstance(date_range, tuple) and len(date_range) == 2:
    d0, d1 = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
else:
    d0, d1 = dmin, dmax
users = sorted(post_all.UserID.dropna().unique().tolist())
sel_users = st.sidebar.multiselect("Preparer", users, default=users) if users else []
types = sorted(post_all.PostingType.unique().tolist())
sel_types = st.sidebar.multiselect("Posting type", types, default=types)

mask = post_all.PostingDate.between(d0, d1) & post_all.PostingType.isin(sel_types)
if users:
    mask &= post_all.UserID.isin(sel_users) | post_all.UserID.isna()
post = post_all[mask].copy()
flags = flags_all[flags_all.PostingID.isin(post.PostingID)].copy()
sett = sett_all[sett_all.PostingID.isin(post.PostingID)].copy() if sett_all is not None else None
ln = ln_all[ln_all.PostingID.isin(post.PostingID)].copy()

has_time = post.EnteredDateTime.notna().any()
has_lag = post.HasLag.any()
has_user = post.UserID.notna().any()

pages = ["Overview", "Volume & workload", "Timeliness", "Timing red flags", "Journal-entry tests", "AR / AP settlement", "Data quality"]
page = st.sidebar.radio("Page", pages)
st.sidebar.caption(f"Source: {source_label}")
st.sidebar.download_button("Download summary (Excel)", summary_workbook(post, flags, sett, acc, issues), file_name="bookkeeping_review_summary.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if len(post) == 0:
    st.warning("No entries match the current filters.")
    st.stop()

# --------------------------------------------------------------------------------------- pages
if page == "Overview":
    st.title("Overview")
    st.caption(f"{len(post):,} entries from {post.PostingDate.min():%d %b %Y} to {post.PostingDate.max():%d %b %Y}")
    unbalanced = int((~post.IsBalanced).sum())
    kpi_row([
        ("Entries", f"{len(post):,}", "Journal entries (postings) in the selected period"),
        ("Journal lines", f"{len(ln):,}", "Debit and credit lines"),
        ("Total debits", money(post.PostingAmount.sum()), "Equals total credits when every entry balances"),
        ("Balance check", "All balanced" if unbalanced == 0 else f"{unbalanced} unbalanced", "Entries whose debits and credits differ"),
    ])
    kpi_row([
        ("Average lag", f"{post.LagDays.mean():.1f} days" if has_lag else "n/a", "Days from document date to the day the entry was keyed"),
        ("Posted on time", pct(post.loc[post.HasLag, 'IsOnTime'].mean()) if has_lag else "n/a", f"Within {cfg['on_time_days']} days"),
        ("Entries flagged", pct(flags.PostingID.nunique() / len(post)), "By the journal-entry tests"),
        ("Open AR/AP", money(sett[~sett.IsSettled].PostingAmount.sum()) if sett is not None else "n/a", "Invoices without a matched payment at period end"),
    ])
    st.markdown("#### Key findings")
    for i, t in enumerate(findings_text(post, flags, sett, cfg), 1):
        st.write(f"{i}. {t}")
    c1, c2 = st.columns([3, 2])
    m = post.groupby("YearMonth").agg(Entries=("PostingID", "size")).reset_index()
    fl = flags.drop_duplicates("PostingID").assign(YearMonth=lambda d: d.PostingDate.dt.to_period("M").astype(str)).groupby("YearMonth").size().reindex(m.YearMonth).fillna(0).values
    fig = go.Figure()
    fig.add_bar(x=m.YearMonth, y=m.Entries, name="Entries", marker_color=PALETTE[0])
    fig.add_scatter(x=m.YearMonth, y=fl, name="Flagged", mode="lines+markers", line=dict(color=FLAG_COLOR), yaxis="y2")
    fig.update_layout(title="Entries per month, with flagged entries", yaxis2=dict(overlaying="y", side="right", showgrid=False), margin=dict(l=10, r=10, t=40, b=10), height=360, legend=dict(orientation="h", y=1.1))
    c1.plotly_chart(fig, width="stretch")
    t = post.PostingType.value_counts().reset_index()
    t.columns = ["PostingType", "Entries"]
    c2.plotly_chart(bar(t, "Entries", "PostingType", "Entries by type", orientation="h"), width="stretch")

elif page == "Volume & workload":
    st.title("Volume & workload")
    m = post.groupby(["YearMonth", "PostingType"]).size().reset_index(name="Entries")
    st.plotly_chart(bar(m, "YearMonth", "Entries", "Entries per month by type", color="PostingType"), width="stretch")
    c1, c2 = st.columns(2)
    if has_user:
        u = post.groupby("UserID").agg(Entries=("PostingID", "size")).reset_index().sort_values("Entries")
        u["Share"] = u.Entries / u.Entries.sum()
        fig = bar(u, "Entries", "UserID", "Entries by preparer", orientation="h", text=u.Share.map(lambda v: f"{v:.0%}"))
        c1.plotly_chart(fig, width="stretch")
    else:
        c1.info("Preparer column not provided — workload by person is not available for this file.")
    me = post.groupby("YearMonth").IsMonthEndWeek.mean().reset_index()
    fig = px.line(me, x="YearMonth", y="IsMonthEndWeek", title=f"Share of entries dated in the last {cfg['month_end_days']} days of the month", markers=True)
    fig.update_layout(yaxis_tickformat=".0%", margin=dict(l=10, r=10, t=40, b=10), height=360)
    c2.plotly_chart(fig, width="stretch")
    lc = post.LineCount.value_counts().sort_index().reset_index()
    lc.columns = ["Lines per entry", "Entries"]
    st.plotly_chart(bar(lc, "Lines per entry", "Entries", "Lines per entry"), width="stretch")

elif page == "Timeliness":
    st.title("Timeliness")
    if not has_lag:
        st.info("This page needs an entered date-time (or a document date and posting date). The uploaded file has neither, so timeliness can't be measured. QuickBooks users: the Audit Log export carries the created time.")
    else:
        pl = post[post.HasLag]
        kpi_row([("Average lag", f"{pl.LagDays.mean():.1f} days", ""), ("Median lag", f"{pl.LagDays.median():.0f} days", ""),
                 ("Posted on time", pct(pl.IsOnTime.mean()), f"Within {cfg['on_time_days']} days"), ("Over 30 days", pct((pl.LagDays > 30).mean()), "")])
        m = pl.groupby("YearMonth").LagDays.agg(Average="mean", Median="median").reset_index()
        fig = px.line(m, x="YearMonth", y=["Average", "Median"], title="Lag by month (days)", markers=True)
        fig.update_layout(margin=dict(l=10, r=10, t=40, b=10), height=360, legend_title_text="")
        st.plotly_chart(fig, width="stretch")
        c1, c2 = st.columns(2)
        order = ["0-3 days", "4-7 days", "8-14 days", "15-30 days", "31+ days"]
        if has_user:
            b = pl.groupby(["UserID", "LagBand"]).size().reset_index(name="Entries")
            b["Share"] = b.Entries / b.groupby("UserID").Entries.transform("sum")
            fig = bar(b, "UserID", "Share", "Lag bands by preparer", color="LagBand", category_orders={"LagBand": order})
            fig.update_layout(yaxis_tickformat=".0%")
            c1.plotly_chart(fig, width="stretch")
        else:
            b = pl.LagBand.value_counts().reindex(order).fillna(0).reset_index()
            b.columns = ["LagBand", "Entries"]
            c1.plotly_chart(bar(b, "LagBand", "Entries", "Lag bands"), width="stretch")
        ot = pl.groupby("PostingType").IsOnTime.mean().reset_index().sort_values("IsOnTime")
        fig = bar(ot, "IsOnTime", "PostingType", "On-time rate by type", orientation="h")
        fig.update_layout(xaxis_tickformat=".0%")
        c2.plotly_chart(fig, width="stretch")
        late = pl[pl.LagDays > 30].sort_values("LagDays", ascending=False)[["PostingID", "PostingType", "DocumentDate", "EnteredDateTime", "UserID", "LagDays", "PostingAmount"]]
        st.markdown(f"#### Backlog: {len(late)} entries keyed more than 30 days late")
        st.dataframe(late, width="stretch", hide_index=True)

elif page == "Timing red flags":
    st.title("Timing red flags")
    if not has_time:
        st.info("This page needs an entered date-time. The uploaded file doesn't have one, so weekend and after-hours activity can't be measured.")
    else:
        pt = post[post.EnteredDateTime.notna()]
        kpi_row([("Weekend entries", pct(pt.IsWeekend.mean()), "Keyed on Saturday or Sunday"),
                 ("After-hours entries", pct(pt.IsAfterHours.mean()), f"Outside {cfg['work_start']:02d}:00–{cfg['work_end']:02d}:00"),
                 ("Large, on cash or bank", f"{int(((pt.IsWeekend | pt.IsAfterHours) & (pt.LiquidAmount > cfg['cash_limit'])).sum()):,}", f"Weekend or after-hours entries moving more than {cfg['cash_limit']:,.0f} through cash or bank")])
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        hm = pt.groupby(["EnteredWeekday", "EnteredHour"]).size().unstack(fill_value=0).reindex(days).fillna(0)
        fig = px.imshow(hm, aspect="auto", color_continuous_scale=["#F3F4F6", "#1F3864"], title="When entries are keyed (weekday × hour)", labels=dict(x="Hour of day", y="", color="Entries"))
        fig.update_layout(margin=dict(l=10, r=10, t=40, b=10), height=380)
        st.plotly_chart(fig, width="stretch")
        if has_user:
            u = pt.groupby("UserID").agg(Weekend=("IsWeekend", "mean"), After_hours=("IsAfterHours", "mean")).reset_index()
            fig = px.bar(u, x="UserID", y=["Weekend", "After_hours"], barmode="group", title="Weekend and after-hours rate by preparer")
            fig.update_layout(yaxis_tickformat=".0%", margin=dict(l=10, r=10, t=40, b=10), height=360, legend_title_text="")
            st.plotly_chart(fig, width="stretch")
        risky = pt[(pt.IsWeekend | pt.IsAfterHours) & (pt.LiquidAmount > cfg["cash_limit"])].sort_values("PostingAmount", ascending=False)
        st.markdown(f"#### Weekend / after-hours entries moving more than {cfg['cash_limit']:,.0f} through cash or bank ({len(risky)})")
        st.dataframe(risky[["PostingID", "PostingType", "EnteredDateTime", "UserID", "PostingAmount"]], width="stretch", hide_index=True)

elif page == "Journal-entry tests":
    st.title("Journal-entry tests")
    st.caption("Each test is a standard audit-style check run on every entry. An entry can trip more than one test. Thresholds are in Settings.")
    skipped = []
    if not post.Counterparty.notna().any(): skipped.append("Possible duplicate (needs a customer/supplier name column)")
    if not post.DocumentDate.notna().any(): skipped.append("Back-dated entry (needs a document date)")
    if not has_time: skipped.append("Weekend / after-hours on cash or bank and Very late entry (need an entered date-time)")
    if not post.RefPostingID.notna().any(): skipped.append("Payment cleared against wrong invoice (needs a cleared-against reference)")
    if skipped: st.info("Not run for this file: " + "; ".join(skipped) + ".")
    if flags.empty:
        st.success("No entries flagged in the selected period.")
    else:
        fu = flags.drop_duplicates("PostingID")
        kpi_row([("Entries flagged", f"{len(fu):,}", ""), ("Flag rate", pct(len(fu) / len(post)), ""), ("Amount flagged", money(fu.PostingAmount.sum()), "Sum of flagged entries"),
                 ("Tests tripped", f"{flags.Test.nunique()} of 8", "")])
        c1, c2 = st.columns(2)
        t = flags.groupby("Test").agg(Entries=("PostingID", "nunique"), Amount=("PostingAmount", "sum")).reset_index().sort_values("Entries")
        c1.plotly_chart(bar(t, "Entries", "Test", "Flags by test", orientation="h"), width="stretch")
        if has_user:
            mx = flags.drop_duplicates(["PostingID", "Test"]).groupby(["UserID", "Test"]).size().unstack(fill_value=0)
            fig = px.imshow(mx, aspect="auto", color_continuous_scale=["#F3F4F6", FLAG_COLOR], title="Preparer × test", labels=dict(color="Entries"), text_auto=True)
            fig.update_layout(margin=dict(l=10, r=10, t=40, b=10), height=360)
            c2.plotly_chart(fig, width="stretch")
        tr = flags.drop_duplicates(["PostingID", "Test"]).assign(YearMonth=lambda d: d.PostingDate.dt.to_period("M").astype(str)).groupby(["YearMonth", "Test"]).size().reset_index(name="Entries")
        st.plotly_chart(bar(tr, "YearMonth", "Entries", "Flags by month", color="Test"), width="stretch")
        st.markdown("#### Flagged entries")
        sel = st.multiselect("Show tests", sorted(flags.Test.unique()), default=sorted(flags.Test.unique()))
        show = flags[flags.Test.isin(sel)].sort_values(["Test", "PostingAmount"], ascending=[True, False])
        st.dataframe(show, width="stretch", hide_index=True)
        st.download_button("Download flagged entries (CSV)", show.to_csv(index=False).encode(), "flagged_entries.csv", "text/csv")

elif page == "AR / AP settlement":
    st.title("AR / AP settlement")
    if sett is None or sett.empty:
        st.info("This page needs a reference column linking payments to invoices, or a customer/supplier name column so payments can be matched by name and amount. Neither is present in this file.")
    else:
        st.caption(f"Invoices matched to payments by {sett_method}. Period end for aging: {post.PostingDate.max():%d %b %Y}.")
        s = sett
        settled = s[s.IsSettled]
        kpi_row([("Invoices", f"{len(s):,}", ""), ("Settled", pct(s.IsSettled.mean()), ""),
                 ("Average days to settle", f"{settled.DaysToSettle.mean():.1f}" if len(settled) else "n/a", ""),
                 ("Open at period end", money(s[~s.IsSettled].PostingAmount.sum()), f"{(~s.IsSettled).sum()} invoices")])
        c1, c2 = st.columns(2)
        if len(settled):
            fig = px.histogram(settled, x="DaysToSettle", color="Counterparty type", nbins=40, barmode="overlay", title="Days from invoice to payment")
            fig.update_layout(margin=dict(l=10, r=10, t=40, b=10), height=360, legend_title_text="")
            c1.plotly_chart(fig, width="stretch")
        open_ = s[~s.IsSettled].groupby(["Counterparty type", "AgeBucket"]).PostingAmount.sum().reset_index()
        if len(open_):
            c2.plotly_chart(bar(open_, "AgeBucket", "PostingAmount", "Open items by age at period end", color="Counterparty type", category_orders={"AgeBucket": BUCKETS}), width="stretch")
        else:
            c2.success("No open invoices at period end.")
        by = s.groupby("Counterparty type").agg(Invoices=("PostingID", "size"), Settled=("IsSettled", "sum"), AvgDays=("DaysToSettle", "mean"),
                                                Within30=("DaysToSettle", lambda d: (d <= 30).mean()), PaidViaCash=("PaidVia", lambda d: (d == "Cash").sum())).reset_index()
        st.dataframe(by.style.format({"AvgDays": "{:.1f}", "Within30": "{:.0%}"}), width="stretch", hide_index=True)
        st.markdown("#### Open invoices")
        st.dataframe(s[~s.IsSettled][["PostingID", "Counterparty type", "InvoiceDate", "PostingAmount", "AgeAtPeriodEnd", "AgeBucket"]].sort_values("AgeAtPeriodEnd", ascending=False), width="stretch", hide_index=True)

elif page == "Data quality":
    st.title("Data quality")
    st.caption("Checks run on every file before anything is charted.")
    dq = pd.DataFrame([{"Check": i["Check"], "Count": i["Count"], "Note": i["Note"]} for i in issues])
    st.dataframe(dq, width="stretch", hide_index=True)
    for i in issues:
        if i["Count"]:
            with st.expander(f"{i['Check']} ({i['Count']})"):
                st.dataframe(i["Rows"].head(500), width="stretch", hide_index=True)
    st.markdown("#### Account classes")
    st.caption("Change a class or group here and every page recalculates. Cash & bank drives the cash tests; Payables and Receivables drive settlement.")
    edited = st.data_editor(acc[["AccountNo", "AccountName", "AccountClass", "AccountGroup", "Source", "Lines"]], hide_index=True, width="stretch",
                            disabled=["AccountNo", "AccountName", "Source", "Lines"],
                            column_config={"AccountClass": st.column_config.SelectboxColumn(options=["Asset", "Liability", "Equity", "Revenue", "Expense"]),
                                           "AccountGroup": st.column_config.SelectboxColumn(options=["Cash & bank", "Receivables", "Payables", "VAT / sales tax", "Other assets", "Other liabilities", "Equity", "Revenue", "Cost of goods", "Personnel", "Occupancy", "Marketing", "Professional fees", "Other operating"])},
                            key="acc_editor")
    changed = {}
    for (_, a), (_, b) in zip(acc.iterrows(), edited.iterrows()):
        if a.AccountClass != b.AccountClass or a.AccountGroup != b.AccountGroup:
            changed[a.AccountNo] = (b.AccountClass, b.AccountGroup)
    if changed and st.button("Apply account changes"):
        st.session_state.acc_overrides.update(changed)
        st.rerun()
