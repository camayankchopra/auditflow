#!/usr/bin/env python3
"""
Contra Elimination
==================
Takes a journal entry (JE) listing and removes contra lines, so the population that is sampled and tested is cleaner.

What counts as contra (both are applied, and every elimination is listed so it can be reviewed):
  1. Within an entry   - lines in the SAME journal, on the SAME account, that offset each other
                         (equal and opposite amounts, or a group of lines that nets to nil).
  2. Across entries    - a line in one journal and an equal and opposite line in ANOTHER journal on the SAME account
                         (typically an entry and its reversal / cancellation), within a date window.
                         "Reversal evidence" mode only pairs them when the narration, reference or journal numbers
                         point to a reversal; "Any equal and opposite" mode pairs them on amount alone.

What remains is split into a DEBIT population and a CREDIT population. Because contra pairs always sit on one account
and net to nil, the net of what remains equals the net of the original listing for every account, and it is
tied to the trial balance:  opening balance + net of remaining lines = TB closing balance.

Output workbook: Summary | Debit Population | Credit Population | Contra Eliminated | TB Tie-out | All Lines
"""
import csv
import datetime as dt
import io
import re
from bisect import bisect_left
from copy import copy
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

import lead_schedule as ls
import tb_mapping as tm
from lead_schedule import FSLI, band, col_headers, put, widths

# statuses
S_KEEP = "Retained"
S_W = "Eliminated - within entry"
S_X = "Eliminated - across entries"
S_NIL = "Nil value line (excluded)"

NUM2 = '#,##0.00;(#,##0.00);"-"'
NUM0 = ls.NUM
DATEF = ls.DATEF
EXTRA_WRAP_LIMIT = 250000          # max candidate pairs examined per account/amount before falling back to nearest-date search
REV_RX = re.compile(r"revers|\brev\b|cancel|contra|correct|reclass|rectif|wrong|error|\bvoid|write ?back|auto ?rev|reinstat|offset")


# --------------------------------------------------------------------------- #
# READING FILES
# --------------------------------------------------------------------------- #
def hn(s):
    return re.sub(r"[^a-z0-9/]+", " ", str(s or "").casefold()).strip()


JE_SYN = {
    "je": ["journal no", "journal number", "je no", "je number", "je", "jv no", "jv number", "voucher no", "voucher number", "voucher",
           "entry no", "entry number", "document no", "document number", "doc no", "journal id", "entry id", "journal",
           "batch no", "batch", "transaction id", "txn id", "trans no", "doc number", "voucher id"],
    "date": ["date", "posting date", "voucher date", "entry date", "document date", "transaction date", "journal date", "txn date",
             "trans date", "posted date", "effective date"],
    "code": ["account code", "a/c code", "ac code", "gl code", "gl account", "ledger code", "account no", "account number", "account id",
             "code", "ledger id", "gl account code"],
    "name": ["account name", "account", "ledger name", "ledger", "gl description", "account description", "gl account name",
             "particulars", "account title", "ledger account", "gl name"],
    "narr": ["narration", "description", "remarks", "memo", "line description", "text", "comments", "header text", "line text", "note",
             "notes", "details", "journal description", "item text"],
    "ref": ["reference", "ref", "ref no", "reference no", "document reference", "cross reference", "ref number", "reference number",
            "external reference", "doc ref"],
    "user": ["posted by", "user", "created by", "prepared by", "entered by", "user id", "approved by", "username", "posted user"],
    "dr": ["debit", "dr", "debit amount", "amount dr", "debit inr", "dr amount", "debits", "debit amt", "amount debit"],
    "cr": ["credit", "cr", "credit amount", "amount cr", "credit inr", "cr amount", "credits", "credit amt", "amount credit"],
    "amt": ["amount", "net amount", "signed amount", "amount dr/cr", "net", "value", "amount inr", "line amount", "amt"],
    "drcr": ["dr/cr", "dr cr", "drcr", "debit/credit", "debit credit", "type", "dr or cr", "indicator", "entry type", "side", "d/c"],
}
JE_ROLE_SYN = {r: {hn(x) for x in v} for r, v in JE_SYN.items()}
JE_ORDER = ["dr", "cr", "drcr", "je", "date", "code", "name", "ref", "user", "narr", "amt"]


def _grid_pick(data, filename, prefs, exact=False):
    ext = Path(filename).suffix.lower()
    if ext == ".csv":
        try:
            text = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = data.decode("latin-1")
        return [list(r) for r in csv.reader(io.StringIO(text))]
    wb = load_workbook(io.BytesIO(data), data_only=True)
    names = wb.sheetnames
    pick = None
    for n in names:
        c = n.casefold().strip()
        if (c in prefs) if exact else any(p in c for p in prefs):
            pick = n
            break
    if pick is None:
        rest = [n for n in names if n.casefold().strip() not in ("instructions", "summary", "readme", "notes")]
        pick = (rest or names)[0]
    return [list(r) for r in wb[pick].iter_rows(values_only=True)]


def _find_header(grid, need):
    for ri, row in enumerate(grid[:30]):
        found = {}
        for ci, v in enumerate(row):
            val = hn(v)
            if not val:
                continue
            for role in JE_ORDER:
                if val in JE_ROLE_SYN[role] and role not in found.values():
                    found[ci] = role
                    break
        roles = set(found.values())
        if need(roles):
            return ri, {role: ci for ci, role in found.items()}
    return None, None


def _s(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


DATE_FMTS = ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y", "%d/%m/%y", "%d-%b-%y", "%Y/%m/%d", "%d.%m.%Y", "%b %d, %Y")


def _date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = _s(v)
    if not s:
        return None
    for f in DATE_FMTS:
        try:
            return dt.datetime.strptime(s[:19] if " " in s and ":" in s else s, f).date()
        except ValueError:
            pass
    try:
        return dt.datetime.fromisoformat(s).date()
    except ValueError:
        return None


class Line:
    __slots__ = ("n", "row", "je", "date", "code", "name", "narr", "ref", "user", "dr", "cr", "net", "key",
                 "status", "pair", "basis", "flag", "tbi")

    def __init__(self, **kw):
        self.status, self.pair, self.basis, self.flag, self.tbi = S_KEEP, "", "", "", None
        for k, v in kw.items():
            setattr(self, k, v)


def read_journals(data, filename, convention="dr_pos"):
    """Return (lines, info). Net = debit - credit. Raises ValueError with a readable message."""
    grid = _grid_pick(data, filename, ("journal", "je", "gl", "ledger", "entries", "jv", "voucher"))
    hi, cols = _find_header(grid, lambda r: (("dr" in r and "cr" in r) or "amt" in r) and ("code" in r or "name" in r))
    if hi is None:
        raise ValueError("Could not find the header row of the journal listing. It needs an account (code or name) and either "
                         "Debit and Credit columns or one Amount column. Suggested headers: Journal No, Date, Account Code, "
                         "Account Name, Narration, Debit, Credit.")
    layout = "separate Debit / Credit columns" if ("dr" in cols and "cr" in cols) else (
        "amount column with Dr/Cr indicator" if "drcr" in cols else "single signed amount column")
    lines, n, undated = [], 0, 0
    for ri in range(hi + 1, len(grid)):
        row = grid[ri]
        if all(c in (None, "") for c in row):
            continue
        g = lambda role: row[cols[role]] if role in cols and cols[role] < len(row) else None
        code, name = _s(g("code")), _s(g("name"))
        je = _s(g("je"))
        if not code and not name:
            continue
        if re.match(r"\s*(grand\s+)?total", f"{je} {name or code}".casefold()):
            continue
        try:
            if "dr" in cols and "cr" in cols:
                net = ls.num(g("dr")) - ls.num(g("cr"))
            else:
                a = ls.num(g("amt"))
                ind = hn(g("drcr"))
                if ind in ("cr", "credit", "c"):
                    net = -abs(a)
                elif ind in ("dr", "debit", "d"):
                    net = abs(a)
                else:
                    net = -a if convention == "cr_pos" else a
        except ValueError as e:
            raise ValueError(f"Row {ri + 1} ({name or code}): {e}")
        d = _date(g("date"))
        if "date" in cols and d is None and _s(g("date")):
            undated += 1
        n += 1
        lines.append(Line(n=n, row=ri + 1, je=je or f"Line {n}", date=d, code=code, name=name, narr=_s(g("narr")), ref=_s(g("ref")),
                          user=_s(g("user")), dr=net if net > 0 else 0.0, cr=-net if net < 0 else 0.0, net=net, key=""))
    if not lines:
        raise ValueError("No journal lines found under the header row.")
    info = dict(layout=layout, has_je="je" in cols, has_date="date" in cols, undated=undated, has_user="user" in cols,
                has_ref="ref" in cols)
    return lines, info


def read_tb(data, filename):
    """Mapped TB (or any TB with a balance column). Returns list of dict(code, name, fsli, py, cy)."""
    grid = _grid_pick(data, filename, ("tb", "trial balance", "mapping"), exact=True)
    hi, cols = tm._find_header(grid, lambda r: "name" in r and ("cy" in r or {"dr", "cr"} <= r))
    if hi is None:
        raise ValueError("Could not find the header row in the trial balance. It needs an Account Name column and a CY Balance "
                         "column (or Debit and Credit columns). The output of the TB Mapping tool works as is.")
    rows = []
    for ri in range(hi + 1, len(grid)):
        row = grid[ri]
        g = lambda role: row[cols[role]] if role in cols and cols[role] < len(row) else None
        code, name = _s(g("code")), _s(g("name"))
        if not name:
            continue
        if not code and re.search(r"\b(grand )?total\b|^difference$|^net (profit|loss)$|^closing balance$", ls.norm(name)):
            continue
        try:
            cy = ls.num(g("cy")) if "cy" in cols else ls.num(g("dr")) - ls.num(g("cr"))
            py = ls.num(g("py")) if "py" in cols else 0.0
        except ValueError as e:
            raise ValueError(f"Trial balance row {ri + 1} ({name}): {e}")
        rows.append(dict(code=code, name=name, fsli=ls.canon(g("fsli")) if "fsli" in cols else None, py=py, cy=cy))
    if not rows:
        raise ValueError("No account rows found in the trial balance.")
    return rows


# --------------------------------------------------------------------------- #
# ACCOUNT KEYS (link journal lines to TB accounts)
# --------------------------------------------------------------------------- #
def _clean_key(s):
    return re.sub(r"[*?~<>=]", "", s).strip()


def assign_keys(lines, tb):
    """Set line.key (text used to group an account) and line.tbi (index into tb or None)."""
    tb_keys, by_code, by_name, used = [], {}, {}, set()
    for i, t in enumerate(tb or []):
        k = _clean_key(t["code"] or "N:" + t["name"])
        base, j = k, 2
        while k in used:
            k, j = f"{base}#{j}", j + 1
        used.add(k)
        tb_keys.append(k)
        if t["code"] and t["code"] not in by_code:
            by_code[t["code"]] = i
        by_name.setdefault(tm.nm(t["name"]), i)
    for l in lines:
        i = by_code.get(l.code) if l.code else None
        if i is None and l.name:
            i = by_name.get(tm.nm(l.name))
        l.tbi = i
        l.key = tb_keys[i] if i is not None else _clean_key(l.code or "N:" + l.name)
    return tb_keys


# --------------------------------------------------------------------------- #
# ELIMINATION ENGINE
# --------------------------------------------------------------------------- #
def cents(x):
    return int(round(abs(x) * 100))


def _link_tier(p, n):
    """0 = explicit link (same reference, cross-referenced journal number, identical narration), 1 = reversal wording, 2 = none."""
    if p.ref and p.ref == n.ref:
        return 0
    pt, nt = f"{p.narr} {p.ref}".casefold(), f"{n.narr} {n.ref}".casefold()
    for a, b in ((p, nt), (n, pt)):
        if len(a.je) >= 3 and a.je.casefold() in b:
            return 0
    if p.narr and tm.nm(p.narr) == tm.nm(n.narr):
        return 0
    if REV_RX.search(pt) or REV_RX.search(nt):
        return 1
    return 2


def _rev_is_later(p, n):
    """True when the line carrying the reversal wording is the later one (the natural original -> reversal order)."""
    pk, nk = bool(REV_RX.search(p.narr.casefold())), bool(REV_RX.search(n.narr.casefold()))
    if pk == nk or p.date is None or n.date is None:
        return True
    rev, orig = (p, n) if pk else (n, p)
    return rev.date >= orig.date


def eliminate(lines, window_days=31, strict=True):
    """Mark contra lines in place. Returns stats dict."""
    for l in lines:
        if cents(l.net) == 0:
            l.status = S_NIL
    live = [l for l in lines if l.status == S_KEEP]
    wid, xid = 0, 0
    # ---- 1. within an entry
    grp = defaultdict(list)
    for l in live:
        grp[(l.je, l.key)].append(l)
    for g in grp.values():
        pos, neg = [l for l in g if l.net > 0], [l for l in g if l.net < 0]
        if not pos or not neg:
            continue
        nb = defaultdict(list)
        for l in neg:
            nb[cents(l.net)].append(l)
        for p in pos:
            lst = nb.get(cents(p.net))
            if lst:
                m = lst.pop(0)
                wid += 1
                for x in (p, m):
                    x.status, x.pair = S_W, f"W-{wid:04d}"
                    x.basis = "Same journal, same account, equal and opposite"
        rem = [l for l in g if l.status == S_KEEP]
        if len(rem) >= 2 and any(l.net > 0 for l in rem) and any(l.net < 0 for l in rem):
            tot = sum(l.net for l in rem)
            if cents(tot) == 0:
                wid += 1
                for x in rem:
                    x.status, x.pair, x.basis = S_W, f"W-{wid:04d}", f"Same journal, same account, {len(rem)} lines net to nil"
            else:
                side = "Dr" if tot > 0 else "Cr"
                for x in rem:
                    x.flag = f"Offsetting lines in the same journal on this account (journal net {side} {abs(tot):,.2f}); only part is contra"
    # ---- 2. across entries
    by_acct = defaultdict(lambda: defaultdict(lambda: ([], [])))
    for l in live:
        if l.status == S_KEEP:
            by_acct[l.key][cents(l.net)][0 if l.net > 0 else 1].append(l)
    use_je = True
    for amt_buckets in by_acct.values():
        for P, N in amt_buckets.values():
            if not P or not N:
                continue
            cands = _candidates(P, N, window_days, strict, use_je)
            cands.sort(key=lambda t: t[0])
            used = set()
            for _, p, n, tier, gap in cands:
                if p.n in used or n.n in used:
                    continue
                used.update((p.n, n.n))
                xid += 1
                basis = ("Reference / journal cross-reference / identical narration" if tier == 0 else
                         "Reversal wording in narration or reference" if tier == 1 else "Equal and opposite amount on the same account")
                if gap is not None and tier == 2:
                    basis += f" ({gap} day{'s' if gap != 1 else ''} apart)"
                for x in (p, n):
                    x.status, x.pair, x.basis = S_X, f"X-{xid:04d}", basis
    # ---- 3. tell the reviewer about equal-and-opposite lines that were NOT eliminated
    by_acct2 = defaultdict(lambda: defaultdict(lambda: ([], [])))
    for l in live:
        if l.status == S_KEEP:
            by_acct2[l.key][cents(l.net)][0 if l.net > 0 else 1].append(l)
    for amt_buckets in by_acct2.values():
        for P, N in amt_buckets.values():
            if not P or not N or len(P) * len(N) > 20000:
                continue
            for p in P:
                for n in N:
                    if p.je == n.je or (window_days and p.date and n.date and abs((p.date - n.date).days) > window_days):
                        continue
                    if not p.flag:
                        p.flag = f"Equal and opposite line in journal {n.je} (not eliminated: no reversal evidence)"
                    if not n.flag:
                        n.flag = f"Equal and opposite line in journal {p.je} (not eliminated: no reversal evidence)"
    elim = [l for l in lines if l.status in (S_W, S_X)]
    return dict(within=wid, across=xid, lines=len(elim))


def _candidates(P, N, window_days, strict, use_je):
    out = []

    def consider(p, n):
        if use_je and p.je == n.je:
            return
        gap = abs((p.date - n.date).days) if p.date and n.date else None
        if window_days and gap is not None and gap > window_days:
            return
        tier = _link_tier(p, n)
        if strict and tier == 2:
            return
        order = 0 if (tier == 2 or _rev_is_later(p, n)) else 1
        out.append(((tier, order, gap if gap is not None else 10 ** 6, p.n, n.n), p, n, tier, gap))

    if len(P) * len(N) <= EXTRA_WRAP_LIMIT:
        for p in P:
            for n in N:
                consider(p, n)
    else:                                    # very large buckets: look at the nearest dates only
        Ns = sorted(N, key=lambda x: x.date or dt.date.min)
        ds = [x.date or dt.date.min for x in Ns]
        for p in P:
            i = bisect_left(ds, p.date or dt.date.min)
            for n in Ns[max(0, i - 40):i + 40]:
                consider(p, n)
    return out


# --------------------------------------------------------------------------- #
# OUTPUT WORKBOOK
# --------------------------------------------------------------------------- #
OPEN_TXT = {"bs_py": "Opening = PY balance for balance sheet accounts, nil for P&L accounts",
            "all_py": "Opening = PY balance for every account",
            "none": "Journal listing includes opening balances (no opening added)"}
HEAD_FILL = "D9E2F3"
YELLOW = "FFFF99"
FONTO = Font(name=ls.FONT, size=10)


_PROTO = {}


def _protos(wb):
    """Pre-built cell styles (house font + number format). Copying a style array is far faster than setting font/number_format per cell."""
    tmp = wb.create_sheet("_tmp")
    for f in (None, DATEF, NUM2):
        c = tmp.cell(row=1, column=1)
        c.font = FONTO
        c.number_format = f or "General"
        _PROTO[f] = copy(c._style)
    wb.remove(tmp)


def _bulk(ws, r0, rows, fmts):
    """Write rows (list of lists) from row r0 with the house font; fmts maps column index (0-based) -> number format."""
    base = _PROTO[None]
    for i, vals in enumerate(rows):
        r = r0 + i
        for j, v in enumerate(vals):
            c = ws.cell(row=r, column=j + 1)
            if v is not None:
                c.value = v
            c._style = copy(_PROTO[fmts[j]] if j in fmts else base)


def _title(ws, text, note=None):
    put(ws, 1, 1, text, bold=True, size=14, color=ls.NAVY)
    if note:
        put(ws, 4, 1, note, italic=True, color=ls.GREY)


def build(lines, tb, tb_keys, meta, out_path, info, opts):
    n = len(lines)
    has_tb = bool(tb)
    wb = Workbook()
    _protos(wb)
    sm = wb.active
    sm.title = "Summary"
    dp = wb.create_sheet("Debit Population")
    cp = wb.create_sheet("Credit Population")
    ce = wb.create_sheet("Contra Eliminated")
    ti = wb.create_sheet("TB Tie-out")
    al = wb.create_sheet("All Lines")

    # ---------------- All Lines
    AL0, ALN = 5, 4 + n
    A = lambda col: f"'All Lines'!${col}${AL0}:${col}${ALN}"
    _title(al, "All journal lines with status", "Every line of the uploaded listing. Filter the Status column. Amounts: debit and credit are positive, Net = Debit - Credit.")
    col_headers(al, 4, ["Line #", "Source row", "Journal no", "Date", "Account code", "Account name", "Narration", "Reference", "Posted by",
                        "Debit", "Credit", "Net (Dr +)", "Account key", "Status", "Pair ID", "Match basis", "Review note",
                        "Helper: entry (by status)", "Helper: entry (by group)", "Helper: entry (all)"])
    seen_s, seen_g, seen_a, rows = set(), set(), set(), []
    for i, l in enumerate(lines):
        grp = "keep" if l.status == S_KEEP else "nil" if l.status == S_NIL else "elim"
        f1 = 0 if (l.je, l.status) in seen_s else 1
        f2 = 0 if (l.je, grp) in seen_g else 1
        f3 = 0 if l.je in seen_a else 1
        seen_s.add((l.je, l.status)); seen_g.add((l.je, grp)); seen_a.add(l.je)
        r = AL0 + i
        rows.append([l.n, l.row, l.je, l.date, l.code or None, l.name or None, l.narr or None, l.ref or None, l.user or None,
                     l.dr, l.cr, f"=J{r}-K{r}", l.key, l.status, l.pair or None, l.basis or None, l.flag or None, f1, f2, f3])
    _bulk(al, AL0, rows, {3: DATEF, 9: NUM2, 10: NUM2, 11: NUM2})
    widths(al, {"A": 8, "B": 9, "C": 14, "D": 12, "E": 12, "F": 34, "G": 44, "H": 14, "I": 12, "J": 15, "K": 15, "L": 15, "M": 16,
                "N": 27, "O": 10, "P": 46, "Q": 60, "R": 12, "S": 12, "T": 12})
    al.freeze_panes = "A5"
    al.auto_filter.ref = f"A4:T{ALN}"

    # ---------------- Debit / Credit populations
    def population(ws, title, items, sign):
        _title(ws, title, "Contra-free population for testing and sampling. Values are static; the Amount column is positive. "
                          "To sample, copy this sheet's rows into the Audit Sampling tool.")
        first, last = 6, max(6, 5 + len(items))
        put(ws, 2, 1, "Lines in population", bold=True)
        put(ws, 2, 3, f"=COUNT(A{first}:A{last})", fmt="#,##0", bold=True)
        put(ws, 3, 1, "Total amount", bold=True)
        put(ws, 3, 3, f"=SUM(I{first}:I{last})", fmt=NUM2, bold=True)
        col_headers(ws, 5, ["Item ID", "Journal no", "Date", "Account code", "Account name", "Narration", "Reference", "Posted by",
                            "Amount", "Review note"])
        _bulk(ws, first, [[l.n, l.je, l.date, l.code or None, l.name or None, l.narr or None, l.ref or None, l.user or None,
                           abs(l.net), l.flag or None] for l in items], {2: DATEF, 8: NUM2})
        widths(ws, {"A": 9, "B": 14, "C": 12, "D": 13, "E": 36, "F": 46, "G": 14, "H": 12, "I": 16, "J": 64})
        ws.freeze_panes = "A6"
        ws.auto_filter.ref = f"A5:J{last}"
        return last

    keep = [l for l in lines if l.status == S_KEEP]
    dr_items = [l for l in keep if l.net > 0]
    cr_items = [l for l in keep if l.net < 0]
    population(dp, "Debit population (remaining after contra elimination)", dr_items, 1)
    population(cp, "Credit population (remaining after contra elimination)", cr_items, -1)

    # ---------------- Contra eliminated
    elim = sorted([l for l in lines if l.status in (S_W, S_X)], key=lambda l: (l.pair, l.n))
    _title(ce, "Contra lines eliminated", "Every eliminated line, grouped by pair. Debit and Credit totals must be equal. "
                                          "Use the yellow column for your review comments; if you disagree with a match, re-run with stricter settings.")
    e0, eN = 6, max(6, 5 + len(elim))
    put(ce, 2, 1, "Debit total", bold=True); put(ce, 2, 3, f"=SUM(J{e0}:J{eN})", fmt=NUM2, bold=True)
    put(ce, 2, 5, "Credit total", bold=True); put(ce, 2, 6, f"=SUM(K{e0}:K{eN})", fmt=NUM2, bold=True)
    put(ce, 3, 1, "Pairs / groups", bold=True); put(ce, 3, 3, len({l.pair for l in elim}), fmt="#,##0", bold=True)
    put(ce, 3, 5, "Difference", bold=True); put(ce, 3, 6, "=C2-F2", fmt=NUM2, bold=True)
    col_headers(ce, 5, ["Pair ID", "Type", "Match basis", "Line #", "Journal no", "Date", "Account code", "Account name", "Narration",
                        "Debit", "Credit", "Your review comment"])
    _bulk(ce, e0, [[l.pair, "Within entry" if l.status == S_W else "Across entries", l.basis, l.n, l.je, l.date, l.code or None,
                    l.name or None, l.narr or None, l.dr, l.cr, None] for l in elim], {5: DATEF, 9: NUM2, 10: NUM2})
    for r in range(e0, e0 + len(elim)):
        ce.cell(row=r, column=12).fill = PatternFill("solid", start_color=YELLOW, end_color=YELLOW)
    widths(ce, {"A": 10, "B": 16, "C": 52, "D": 8, "E": 14, "F": 12, "G": 13, "H": 34, "I": 44, "J": 15, "K": 15, "L": 36})
    ce.freeze_panes = "A6"
    ce.auto_filter.ref = f"A5:L{eN}"

    # ---------------- TB tie-out
    _title(ti, "Tie-out of the remaining population to the trial balance",
           "Expected closing = Opening + net of the REMAINING (contra-free) lines. Difference = Expected closing - TB closing. "
           "Contra pairs sit on one account and net to nil, so the tie-out holds before and after elimination.")
    put(ti, 2, 1, OPEN_TXT[opts["opening"]] if has_tb else "No trial balance uploaded: tie-out not performed (movement only).", italic=True, color=ls.GREY)
    heads = ["Account key", "Account code", "Account name", "FSLI category", "Stmt", "Opening balance", "Journal debits", "Journal credits",
             "Journal net", "Contra lines eliminated", "Contra amount eliminated (Dr side)", "Net of remaining lines", "Expected closing",
             "TB closing (CY)", "Difference", "Status", "Note"]
    col_headers(ti, 5, heads)
    ties, seen = [], set()
    for i, t in enumerate(tb or []):
        stmt = FSLI[t["fsli"]]["stmt"] if t["fsli"] else ""
        opening = {"bs_py": 0.0 if stmt == "PL" else t["py"], "all_py": t["py"], "none": 0.0}[opts["opening"]]
        ties.append(dict(key=tb_keys[i], code=t["code"], name=t["name"], fsli=t["fsli"] or "", stmt={"BS": "BS", "PL": "P&L"}.get(stmt, "?"),
                         opening=opening, cy=t["cy"], note=""))
        seen.add(tb_keys[i])
    extra = {}
    for l in lines:
        if l.key not in seen and l.key not in extra:
            extra[l.key] = l
    for k in sorted(extra):
        l = extra[k]
        ties.append(dict(key=k, code=l.code, name=l.name, fsli="", stmt="", opening=0.0, cy=0.0,
                         note="Account in the journal listing but not in the trial balance" if has_tb else ""))
    t0 = 6
    tN = 5 + len(ties)
    tot = tN + 1
    for i, t in enumerate(ties):
        r = t0 + i
        put(ti, r, 1, t["key"])
        put(ti, r, 2, t["code"] or None)
        put(ti, r, 3, t["name"] or None)
        put(ti, r, 4, t["fsli"] or None)
        put(ti, r, 5, t["stmt"] or None, h="center")
        put(ti, r, 6, t["opening"] if has_tb else None, fmt=NUM2, color=ls.BLUE)
        put(ti, r, 7, f"=SUMIFS({A('J')},{A('M')},$A{r})", fmt=NUM2)
        put(ti, r, 8, f"=SUMIFS({A('K')},{A('M')},$A{r})", fmt=NUM2)
        put(ti, r, 9, f"=G{r}-H{r}", fmt=NUM2)
        put(ti, r, 10, f"=COUNTIFS({A('M')},$A{r},{A('N')},\"Eliminated*\")", fmt="#,##0", h="center")
        put(ti, r, 11, f"=SUMIFS({A('J')},{A('M')},$A{r},{A('N')},\"Eliminated*\")", fmt=NUM2)
        put(ti, r, 12, f"=SUMIFS({A('L')},{A('M')},$A{r},{A('N')},\"{S_KEEP}\")", fmt=NUM2)
        if has_tb:
            put(ti, r, 13, f"=F{r}+L{r}", fmt=NUM2)
            put(ti, r, 14, t["cy"], fmt=NUM2, color=ls.BLUE)
            put(ti, r, 15, f"=M{r}-N{r}", fmt=NUM2)
            put(ti, r, 16, f'=IF(ABS(O{r})<0.5,"OK","CHECK")', h="center")
        put(ti, r, 17, t["note"] or None)
    put(ti, tot, 3, "Total", bold=True)
    for c in (6, 7, 8, 9, 10, 11, 12) + ((13, 14, 15) if has_tb else ()):
        L = "ABCDEFGHIJKLMNOPQ"[c - 1]
        put(ti, tot, c, f"=SUM({L}{t0}:{L}{tN})", fmt=NUM2 if c != 10 else "#,##0", bold=True, top=True, h="center" if c == 10 else None)
    if has_tb:
        ls.status_formats(ti, f"P{t0}:P{tN}", f"P{t0}")
    widths(ti, {"A": 14, "B": 12, "C": 36, "D": 28, "E": 6, "F": 15, "G": 15, "H": 15, "I": 15, "J": 11, "K": 17, "L": 16, "M": 16,
                "N": 16, "O": 14, "P": 9, "Q": 44})
    ti.freeze_panes = "D6"
    ti.auto_filter.ref = f"A5:Q{tN}"
    TI = lambda col: f"'TB Tie-out'!${col}${t0}:${col}${tN}"

    # ---------------- Summary
    put(sm, 1, 2, "Contra Elimination - Summary", bold=True, size=14, color=ls.NAVY)
    inf = [("Client", meta.get("client") or "Client"), ("Year ended", meta.get("year_end")), ("Journal listing", info.get("layout", "")),
           ("Within-entry rule", "Equal and opposite lines on the same account in the same journal, or lines that net to nil"),
           ("Across-entry rule", (f"Equal and opposite on the same account, {opts['window']} days apart at most" if opts["window"] else
                                  "Equal and opposite on the same account, any date gap") +
            ("; only with reversal evidence (narration, reference, journal cross-reference)" if opts["strict"] else "; on amount alone")),
           ("TB tie-out", OPEN_TXT[opts["opening"]] if has_tb else "Trial balance not provided")]
    for k, (lab, v) in enumerate(inf, 2):
        put(sm, k, 2, lab, bold=True)
        put(sm, k, 3, v, fmt=DATEF if isinstance(v, (dt.date, dt.datetime)) else None, h="left")

    band(sm, 9, 2, 7, "Population reconciliation")
    col_headers(sm, 10, ["", "", "Lines", "Entries", "Debit", "Credit", "Net"], start=1)
    sm.cell(row=10, column=2).value = "Measure"
    sm.cell(row=10, column=1).value = None

    def srow(r, label, lines_f, ent_f, dr_f, cr_f, bold=False, top=False, indent=0):
        put(sm, r, 2, label, bold=bold, top=top, indent=indent)
        put(sm, r, 3, lines_f, fmt="#,##0", bold=bold, top=top, h="right")
        put(sm, r, 4, ent_f, fmt="#,##0", bold=bold, top=top, h="right")
        put(sm, r, 5, dr_f, fmt=NUM2, bold=bold, top=top)
        put(sm, r, 6, cr_f, fmt=NUM2, bold=bold, top=top)
        put(sm, r, 7, f"=E{r}-F{r}", fmt=NUM2, bold=bold, top=top)

    st = A("N")
    srow(11, "Original journal population", f"=COUNT({A('A')})", f"=SUM({A('T')})", f"=SUM({A('J')})", f"=SUM({A('K')})", bold=True)
    srow(12, "Less: nil-value lines excluded", f'=COUNTIF({st},"{S_NIL}")', f'=SUMIFS({A("R")},{st},"{S_NIL}")',
         f'=SUMIFS({A("J")},{st},"{S_NIL}")', f'=SUMIFS({A("K")},{st},"{S_NIL}")')
    srow(13, "Less: contra eliminated - within entry", f'=COUNTIF({st},"{S_W}")', f'=SUMIFS({A("R")},{st},"{S_W}")',
         f'=SUMIFS({A("J")},{st},"{S_W}")', f'=SUMIFS({A("K")},{st},"{S_W}")')
    srow(14, "Less: contra eliminated - across entries", f'=COUNTIF({st},"{S_X}")', f'=SUMIFS({A("R")},{st},"{S_X}")',
         f'=SUMIFS({A("J")},{st},"{S_X}")', f'=SUMIFS({A("K")},{st},"{S_X}")')
    srow(15, "Total contra eliminated", "=C13+C14", f'=SUMIFS({A("S")},{st},"Eliminated*")', "=E13+E14", "=F13+F14", bold=True, top=True)
    srow(16, "Remaining population to test", f'=COUNTIF({st},"{S_KEEP}")', f'=SUMIFS({A("S")},{st},"{S_KEEP}")',
         f'=SUMIFS({A("J")},{st},"{S_KEEP}")', f'=SUMIFS({A("K")},{st},"{S_KEEP}")', bold=True, top=True)
    srow(17, "of which: debit population", f'=COUNTIFS({st},"{S_KEEP}",{A("L")},">0")', "-",
         f'=SUMIFS({A("L")},{st},"{S_KEEP}",{A("L")},">0")', 0, indent=2)
    srow(18, "of which: credit population", f'=COUNTIFS({st},"{S_KEEP}",{A("L")},"<0")', "-", 0,
         f'=-SUMIFS({A("L")},{st},"{S_KEEP}",{A("L")},"<0")', indent=2)
    put(sm, 19, 2, "Contra amount eliminated (each side)", bold=True)
    put(sm, 19, 5, "=E15", fmt=NUM2, bold=True)
    j_all = len(seen_a)
    kept_j = {l.je for l in lines if l.status == S_KEEP}
    elim_j = {l.je for l in lines if l.status in (S_W, S_X)}
    nonnil_j = {l.je for l in lines if l.status != S_NIL}
    put(sm, 20, 2, "Entries fully eliminated")
    put(sm, 20, 3, len(nonnil_j - kept_j), fmt="#,##0", h="right")
    put(sm, 21, 2, "Entries partly eliminated (some lines remain)")
    put(sm, 21, 3, len(elim_j & kept_j), fmt="#,##0", h="right")
    put(sm, 20, 4, "static count", italic=True, color=ls.GREY)

    band(sm, 23, 2, 7, "Checks")
    put(sm, 24, 3, "Value", bold=True, h="right")
    put(sm, 24, 4, "Status", bold=True, h="center")
    checks = [("Original debits equal credits (journals balance)", "=E11-F11"),
              ("Original net = net of contra + nil lines + remaining", "=G11-G12-G15-G16"),
              ("Contra eliminated nets to nil", "=G15"),
              ("Line count: original = nil + contra + remaining", "=C11-C12-C15-C16"),
              ("Debit Population sheet: lines agree to summary", f"='Debit Population'!C2-C17"),
              ("Debit Population sheet: amount agrees to summary", f"='Debit Population'!C3-E17"),
              ("Credit Population sheet: lines agree to summary", f"='Credit Population'!C2-C18"),
              ("Credit Population sheet: amount agrees to summary", f"='Credit Population'!C3-F18"),
              ("Contra Eliminated sheet: debits equal credits", "='Contra Eliminated'!F3")]
    if has_tb:
        checks += [("Remaining population net ties to TB (sum of absolute differences)", f"=SUMPRODUCT(ABS({TI('O')}))")]
    for k, (lab, f) in enumerate(checks, 25):
        put(sm, k, 2, lab)
        put(sm, k, 3, f, fmt=NUM2)
        put(sm, k, 4, f'=IF(ABS(C{k})<0.5,"OK","CHECK")', h="center")
    last_chk = 24 + len(checks)
    k = last_chk + 1
    if has_tb:
        put(sm, k, 2, "Accounts agreeing to the TB")
        put(sm, k, 3, f'=COUNTIF({TI("P")},"OK")&" of "&COUNTA({TI("A")})', h="right")
        put(sm, k, 4, f'=IF(COUNTIF({TI("P")},"CHECK")=0,"OK","CHECK")', h="center")
        last_chk = k
    else:
        put(sm, k, 2, "TB tie-out")
        put(sm, k, 3, "Not performed", h="right")
        put(sm, k, 4, "-", h="center")
        last_chk = k
    ls.status_formats(sm, f"D25:D{last_chk}", "D25")

    nr = last_chk + 2
    band(sm, nr, 2, 7, "How to use this")
    notes = ["1. Sample from the Debit Population and Credit Population sheets: they hold only lines that survived contra elimination.",
             "2. Review the Contra Eliminated sheet. Matches shown as 'Equal and opposite amount' rest on amount only, so look at them with care.",
             "3. Review notes on retained lines point to equal and opposite lines that were NOT eliminated, and to journals that are only partly contra.",
             "4. Bank, cash and control accounts often carry genuine equal and opposite entries. Keep 'reversal evidence' mode for them.",
             "5. On the TB Tie-out sheet any CHECK means the listing and the trial balance disagree for that account (missing or extra journals, wrong opening).",
             "The rules are mechanical and are not a substitute for professional judgement: the auditor decides what is excluded from testing."]
    for j, t in enumerate(notes):
        put(sm, nr + 1 + j, 2, t, italic=(j == len(notes) - 1), color=ls.GREY if j == len(notes) - 1 else None)
    widths(sm, {"A": 3, "B": 58, "C": 18, "D": 12, "E": 18, "F": 18, "G": 18})
    sm.sheet_view.showGridLines = False
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    return dict(n=n, within=sum(1 for l in lines if l.status == S_W), across=sum(1 for l in lines if l.status == S_X),
                nil=sum(1 for l in lines if l.status == S_NIL), kept=len(keep), debit=len(dr_items), credit=len(cr_items),
                accounts=len(ties))


# --------------------------------------------------------------------------- #
# ENTRY POINT USED BY THE WEB LAYER
# --------------------------------------------------------------------------- #
def process(je_data, je_name, tb_data, tb_name, out_path, client="Client", year_end=None, convention="dr_pos",
            window=31, strict=True, opening="bs_py"):
    lines, info = read_journals(je_data, je_name, convention)
    tb = read_tb(tb_data, tb_name) if tb_data else None
    tb_keys = assign_keys(lines, tb)
    if not info["has_je"]:
        info["layout"] += " (no journal number column: within-entry matching skipped)"
    eliminate(lines, window_days=window, strict=strict)
    opts = dict(window=window, strict=strict, opening=opening)
    res = build(lines, tb, tb_keys, dict(client=client, year_end=year_end), out_path, info, opts)
    tie = None
    if tb:
        bal = defaultdict(float)
        for l in lines:
            if l.status == S_KEEP:
                bal[l.key] += l.net
        diffs, ok = {}, 0
        for i, t in enumerate(tb):
            stmt = FSLI[t["fsli"]]["stmt"] if t["fsli"] else ""
            op = {"bs_py": 0.0 if stmt == "PL" else t["py"], "all_py": t["py"], "none": 0.0}[opening]
            d = op + bal.get(tb_keys[i], 0.0) - t["cy"]
            ok += abs(d) < 0.5
        extra = {l.key for l in lines if l.tbi is None}
        tie = (ok, len(tb) + len(extra))
    res["tie"] = tie
    res["info"] = info
    return res


# --------------------------------------------------------------------------- #
# TEMPLATE (sample listing + matching TB, a fictitious company)
# --------------------------------------------------------------------------- #
ACCTS = [  # code, name, FSLI, PY balance (balance sheet only)
    ("1004", "Plant & Machinery", tm.PPE, 210000), ("1005", "Accumulated Depreciation - Plant & Machinery", tm.PPE, -114700),
    ("1200", "Domestic Customers", tm.TR, 62400), ("1300", "HDFC Bank - Current A/c", tm.CB, 18400),
    ("1502", "Prepaid Insurance", tm.OCA, 2700), ("1504", "Staff Advances", tm.OCA, 950),
    ("2001", "Equity Share Capital", tm.SC, -50000), ("2300", "Sundry Creditors - Materials", tm.TP, -44800),
    ("2400", "GST Payable", tm.OCL, -6200), ("2401", "Salary Payable", tm.OCL, -5100), ("2402", "Advance from Customers", tm.OCL, -3800),
    ("2403", "Accrued Expenses", tm.OCL, -4500), ("2003", "Retained Earnings", tm.RS, None),
    ("3001", "Sales - Domestic", tm.REV, 0), ("4001", "Purchases - Raw Materials", tm.COM, 0), ("5001", "Salaries & Bonus", tm.EMP, 0),
    ("5002", "Rent - Office", tm.OE, 0), ("5004", "Insurance", tm.OE, 0), ("5005", "Depreciation", tm.DEP, 0),
    ("9999", "Suspense A/c", None, 0)]
_N = {a[0]: a[1] for a in ACCTS}
D = dt.date
SAMPLE_JE = [  # journal, date, narration, reference, user, [(code, dr, cr)]
    ("JE-001", D(2025, 4, 5), "Salary accrual April", "", "ravi", [("5001", 52000, 0), ("2401", 0, 52000)]),
    ("JE-002", D(2025, 4, 30), "Salary paid April", "", "ravi", [("2401", 52000, 0), ("1300", 0, 52000)]),
    ("JE-003", D(2025, 4, 30), "Accrue April rent", "", "meena", [("5002", 12000, 0), ("2403", 0, 12000)]),
    ("JE-004", D(2025, 5, 1), "Reversal of April rent accrual", "", "meena", [("2403", 12000, 0), ("5002", 0, 12000)]),
    ("JE-005", D(2025, 5, 2), "April rent paid", "", "meena", [("5002", 12000, 0), ("1300", 0, 12000)]),
    ("JE-006", D(2025, 6, 10), "Reclass of customer receipt", "", "ravi",
     [("1300", 25000, 0), ("1300", 0, 25000), ("1200", 25000, 0), ("2402", 0, 25000)]),
    ("JE-007", D(2025, 6, 15), "Receipt from customer ABC", "", "ravi", [("1300", 40000, 0), ("1200", 0, 40000)]),
    ("JE-008", D(2025, 6, 20), "Payment to supplier XYZ", "", "ravi", [("2300", 40000, 0), ("1300", 0, 40000)]),
    ("JE-009", D(2025, 6, 30), "Insurance premium paid", "", "meena", [("1502", 9000, 0), ("1300", 0, 9000)]),
    ("JE-010", D(2025, 9, 30), "Amortise prepaid insurance", "", "meena", [("5004", 4500, 0), ("1502", 0, 4500)]),
    ("JE-011", D(2025, 10, 15), "Purchase invoice P-881", "P-881", "ravi", [("4001", 30000, 0), ("2300", 0, 30000)]),
    ("JE-012", D(2025, 10, 18), "Cancel purchase invoice - supplier error", "P-881", "ravi", [("2300", 30000, 0), ("4001", 0, 30000)]),
    ("JE-013", D(2025, 10, 20), "Corrected invoice", "P-881A", "ravi", [("4001", 31500, 0), ("2300", 0, 31500)]),
    ("JE-014", D(2025, 10, 31), "Depreciation H1", "", "meena", [("5005", 19600, 0), ("1005", 0, 19600)]),
    ("JE-015", D(2026, 3, 31), "GST paid", "", "meena", [("2400", 6200, 0), ("1300", 0, 6200)]),
    ("JE-016", D(2026, 3, 31), "Sales summary FY", "", "ravi", [("1200", 415000, 0), ("3001", 0, 415000)]),
    ("JE-017", D(2026, 3, 31), "Sales net of credit note", "", "ravi", [("1200", 10000, 0), ("1200", 0, 6000), ("3001", 0, 4000)]),
]


def _sample():
    movement = defaultdict(float)
    for _, _, _, _, _, ls_ in SAMPLE_JE:
        for c, d, cr in ls_:
            movement[c] += d - cr
    assert abs(sum(movement.values())) < 0.01
    py = {c: p for c, _, _, p in ACCTS if p is not None}
    plug = -sum(py.values())
    tb = []
    for c, name, fsli, p in ACCTS:
        p = plug if p is None else p
        cy = (p if c not in ("3001", "4001", "5001", "5002", "5004", "5005") else 0) + movement.get(c, 0.0)
        if c == "1504":
            cy += 2500          # deliberately differs from the journals
        if c == "9999":
            cy = -2500
        tb.append((c, name, fsli, p, cy))
    assert abs(sum(t[4] for t in tb)) < 0.01 and abs(sum(t[3] for t in tb)) < 0.01
    return tb


def make_template(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Instructions"
    lines = [("JOURNAL LISTING + TRIAL BALANCE - Contra Elimination tool", True), ("", False),
             ("1.  'Journals' tab: replace the sample with the client's journal listing. Keep the header row.", False),
             ("2.  Needed: an account (code or name) and EITHER Debit and Credit columns OR one Amount column (optionally with a Dr/Cr column).", False),
             ("3.  Helpful: Journal No (required for within-entry matching), Date, Narration, Reference, Posted By.", False),
             ("4.  'TB' tab: the trial balance with CY Balance (and PY Balance). The output of the TB Mapping tool works as is.", False),
             ("5.  Upload this file in BOTH slots (journals and trial balance) - the tool picks the right tab - or upload each file separately.", False),
             ("", False),
             ("The sample is a fictitious company. It contains a rent accrual and its reversal, a cancelled and re-posted purchase invoice, a journal with an", False),
             ("offsetting bank line, equal bank receipt and payment that are NOT linked, and a journal that is only partly contra.", False),
             ("In the sample TB, Staff Advances and Suspense A/c deliberately do not agree to the journals, to show a CHECK on the tie-out.", False)]
    for i, (t, b) in enumerate(lines, 1):
        put(ws, i, 1, t, bold=b, size=14 if b else 10, color=ls.NAVY if b else None)
    ws.column_dimensions["A"].width = 140
    je = wb.create_sheet("Journals")
    col_headers(je, 1, ["Journal No", "Date", "Account Code", "Account Name", "Narration", "Reference", "Debit", "Credit", "Posted By"])
    r = 2
    for j, d, narr, ref, user, ls_ in SAMPLE_JE:
        for c, dr, cr in ls_:
            put(je, r, 1, j); put(je, r, 2, d, fmt=DATEF); put(je, r, 3, c); put(je, r, 4, _N[c]); put(je, r, 5, narr)
            put(je, r, 6, ref or None); put(je, r, 7, dr or None, fmt=NUM2, color=ls.BLUE); put(je, r, 8, cr or None, fmt=NUM2, color=ls.BLUE)
            put(je, r, 9, user)
            r += 1
    widths(je, {"A": 12, "B": 12, "C": 13, "D": 44, "E": 42, "F": 12, "G": 15, "H": 15, "I": 11})
    je.freeze_panes = "A2"
    tbs = wb.create_sheet("TB")
    col_headers(tbs, 1, ["Account Code", "Account Name", "FSLI Category", "PY Balance", "CY Balance"])
    for i, (c, name, fsli, p, cy) in enumerate(_sample(), 2):
        put(tbs, i, 1, c); put(tbs, i, 2, name); put(tbs, i, 3, fsli)
        put(tbs, i, 4, p, fmt=NUM2, color=ls.BLUE); put(tbs, i, 5, cy, fmt=NUM2, color=ls.BLUE)
    widths(tbs, {"A": 14, "B": 46, "C": 30, "D": 16, "E": 16})
    tbs.freeze_panes = "A2"
    wb.save(path)
