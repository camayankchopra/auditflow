"""Fixed Asset Workpaper Tool - builds a fixed asset audit workpaper (Excel) from a fixed asset register.

Register = one row per asset (or per addition). The workpaper contains:
  Summary       engagement details, thresholds, integrity checks, depreciation result, additions/disposals summary, TB tie-out
  Roll-forward  gross block, accumulated depreciation, NBV and depreciation recalculation by asset class
  Depn Recalc   depreciation recomputed per asset (class method / life / residual from the Parameters tab)
  Additions     additions listing with vouching columns (key additions flagged)
  Disposals     disposals / write-offs with profit/loss and accumulated depreciation released
  Checks        row-level integrity flags
  Register      the register as loaded
  Parameters    class policy table (method, life, residual %) - ILLUSTRATIVE defaults, edit to your policy
"""
import csv
import datetime as dt
import random
import re
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

MAX_ROWS = 5000
METHODS = ["SLM", "WDV", "None"]

# ------------------------------------------------------------------ default class policy (illustrative, Companies Act Sch II style)
POLICY_DEFAULTS = [  # (keyword, method, life, note)
    ("land", "None", 0, "Land - not depreciated"),
    ("capital work", "None", 0, "CWIP - not depreciated until put to use"),
    ("cwip", "None", 0, "CWIP - not depreciated until put to use"),
    ("building", "SLM", 60, "Sch II: RCC building (other than factory) 60 yrs"),
    ("plant", "SLM", 15, "Sch II: plant & machinery (general) 15 yrs"),
    ("machiner", "SLM", 15, "Sch II: plant & machinery (general) 15 yrs"),
    ("furniture", "SLM", 10, "Sch II: furniture & fittings 10 yrs"),
    ("office", "SLM", 5, "Sch II: office equipment 5 yrs"),
    ("computer", "SLM", 3, "Sch II: computers (end-user) 3 yrs"),
    ("vehicle", "SLM", 8, "Sch II: motor vehicles 8 yrs"),
    ("motor", "SLM", 8, "Sch II: motor vehicles 8 yrs"),
    ("electrical", "SLM", 10, "Sch II: electrical installations 10 yrs"),
    ("software", "SLM", 5, "Intangible - confirm policy"),
]


def default_policy(cls, default_method):
    low = cls.lower()
    for kw, method, life, note in POLICY_DEFAULTS:
        if kw in low:
            if method == "None":
                return "None", 0, note
            return default_method, life, note
    return default_method, 10, "Not recognised - policy to be confirmed"


# ------------------------------------------------------------------ reading the register
def _norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


COLS = [  # key, label, synonyms (normalised), required
    ("id", "Asset ID", ["assetid", "assetno", "assetnumber", "assettag", "tag", "ref", "reference", "id", "itemid", "faid", "fano"], False),
    ("desc", "Description", ["description", "assetdescription", "desc", "name", "assetname", "particulars", "details"], False),
    ("cls", "Asset class", ["assetclass", "class", "category", "assetcategory", "block", "assetblock", "group", "assetgroup", "classofasset"], True),
    ("pdate", "Date of purchase", ["dateofpurchase", "purchasedate", "dateofcapitalisation", "capitalisationdate", "datecapitalised", "putousedate", "dateofacquisition", "acquisitiondate", "purchaseddate"], False),
    ("sdate", "Date of sale / write-off", ["dateofsale", "saledate", "dateofdisposal", "disposaldate", "dateofsalewriteoff", "dateofsaledisposalwriteoff", "dateofwriteoff", "dateofsaledisposal", "datesold"], False),
    ("og", "Opening gross value", ["openinggrossvalue", "openinggross", "openingcost", "grossopening", "openingassets", "grossvalueopening", "openinggrossblock", "grossblockopening", "grossvalue", "opening"], True),
    ("add", "Additions", ["additions", "addition", "additionsduringtheyear", "additionsfortheyear", "purchases"], True),
    ("disp", "Disposals", ["disposalsgross", "disposals", "disposal", "sales", "deletions", "saledisposals", "disposalsduringtheyear"], False),
    ("wo", "Write-off", ["writeoffgross", "writeoff", "writeoffs", "writtenoff", "writeoffduringtheyear"], False),
    ("cg", "Closing gross value", ["closinggrossvalue", "closinggross", "closingcost", "grossclosing", "closinggrossblock", "grossblockclosing"], True),
    ("oad", "Opening accumulated depreciation", ["openingaccumulateddepreciation", "openingaccdep", "openingaccumulateddep", "accumulateddepreciationopening", "openingdepreciation", "accdepopening", "openingad"], True),
    ("dep", "Depreciation for the year", ["depreciationfortheyear", "depreciationduringtheyear", "currentyeardepreciation", "depreciationcharge", "depreciation", "depfortheyear", "depreciationforyear", "depreciationchargefortheyear"], True),
    ("adrel", "Acc. depreciation on disposals / write-off", ["accdepreciationondisposalswriteoff", "accumulateddepreciationondisposals", "depreciationondisposals", "accdepondisposals", "depondisposals", "accumulateddepreciationondisposalwriteoff", "depreciationondisposalandwriteoff", "accdepwrittenback", "depreciationwrittenback", "accumulateddepreciationondisposalsandwriteoff", "depreciationondeletions"], False),
    ("cad", "Closing accumulated depreciation", ["closingaccumulateddepreciation", "closingaccdep", "closingaccumulateddep", "accumulateddepreciationclosing", "closingdepreciation", "accdepclosing", "closingad"], True),
    ("nbvcy", "Net book value - current year", ["netbookvaluecurrentyear", "nbvcy", "closingnbv", "nbvcurrentyear", "netbookvalue", "writtendownvalue", "wdv", "netblockclosing", "nbv"], False),
    ("nbvpy", "Net book value - prior year", ["netbookvalueprioryear", "nbvpy", "openingnbv", "nbvprioryear", "netbookvalueprevioussyear", "netbookvaluepreviousyear", "nbvpreviousyear", "netblockopening"], False),
    ("proc", "Sale proceeds", ["saleproceeds", "proceeds", "salevalue", "saleconsideration", "salesprice", "saleprice", "consideration"], False),
]
NUM_KEYS = ["og", "add", "disp", "wo", "cg", "oad", "dep", "adrel", "cad", "nbvcy", "nbvpy", "proc"]


def _num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("₹", "").replace("$", "")
    if s in ("", "-"):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    val = float(s.strip("()"))
    return -val if neg else val


def _date(v):
    if v in (None, ""):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    s = str(v).strip()
    if s in ("", "-", "NA", "N/A"):
        return None
    for f in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d-%b-%y", "%d/%m/%y", "%d.%m.%Y"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise ValueError(f"unrecognised date '{s}'")


def read_register(path):
    """Return (rows, notes). rows are dicts keyed by COLS keys. Raises ValueError with readable messages."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as f:
            raw = [r for r in csv.reader(f)]
    else:
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb["Register"] if "Register" in wb.sheetnames else wb[wb.sheetnames[0]]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    hdr_i, cmap = None, {}
    for i, r in enumerate(raw[:25]):
        norm = [_norm(c) for c in r]
        m = {}
        for key, label, syn, _req in COLS:
            for j, h in enumerate(norm):
                if (h in syn or h == _norm(label)) and j not in m.values():
                    m[key] = j
                    break
        if "cls" in m and "og" in m and "dep" in m:
            hdr_i, cmap = i, m
            break
    if hdr_i is None:
        raise ValueError("Could not find the register header row. Use the template column names "
                         "(Asset class, Opening gross value, Additions, Depreciation for the year ...).")
    missing = [lab for key, lab, _, req in COLS if req and key not in cmap]
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))
    rows, notes, seen = [], [], {}
    for k, r in enumerate(raw[hdr_i + 1:], start=hdr_i + 2):
        r = list(r) + [None] * (max(cmap.values()) + 1 - len(r))
        if all(c in (None, "") for c in r):
            continue
        get = lambda key: r[cmap[key]] if key in cmap else None
        txt = f"{get('id') or ''} {get('desc') or ''} {get('cls') or ''}".lower()
        if re.match(r"\s*(grand\s+)?total", txt):
            continue
        row = {}
        try:
            for key in NUM_KEYS:
                row[key] = _num(get(key))
            row["pdate"] = _date(get("pdate"))
            row["sdate"] = _date(get("sdate"))
        except ValueError as e:
            raise ValueError(f"Row {k} of the file: {e}")
        idv = str(get("id")).strip() if get("id") not in (None, "") else f"Asset {len(rows) + 1}"
        if idv.endswith(".0") and isinstance(get("id"), float):
            idv = idv[:-2]
        row["id"] = idv
        row["desc"] = str(get("desc")).strip() if get("desc") not in (None, "") else ""
        row["cls"] = str(get("cls")).strip() if get("cls") not in (None, "") else "Unclassified"
        rows.append(row)
    if not rows:
        raise ValueError("No assets found in the register.")
    if len(rows) > MAX_ROWS:
        raise ValueError(f"Register has {len(rows):,} rows; the limit is {MAX_ROWS:,}.")
    for key, lab in (("nbvcy", "Net book value - current year"), ("nbvpy", "Net book value - prior year")):
        if key not in cmap:
            for row in rows:
                row[key] = (row["cg"] - row["cad"]) if key == "nbvcy" else (row["og"] - row["oad"])
            notes.append(f"'{lab}' column not found - derived from the register's gross and accumulated depreciation.")
    for opt, lab in (("pdate", "Date of purchase"), ("sdate", "Date of sale / write-off")):
        if opt not in cmap:
            notes.append(f"'{lab}' column not found - date-based checks and pro-rata depreciation will be incomplete.")
    return rows, notes


# ------------------------------------------------------------------ calculation (mirrors the Excel formulas)
def recalc_asset(r, pol, fy0, fy1):
    method, life, res = pol
    cost = r["og"] + r["add"]
    days_year = (fy1 - fy0).days + 1
    start = fy0 if not r["pdate"] else max(r["pdate"], fy0)
    disposed = (r["disp"] + r["wo"]) > 0
    end = fy1
    if disposed and r["sdate"]:
        end = min(r["sdate"], fy1)
    days = max(0, (end - start).days + 1)
    rate = (1 - res ** (1 / life)) if (life > 0 and res < 1) else 0
    if method == "SLM":
        annual = cost * (1 - res) / life if life > 0 else 0
    elif method == "WDV":
        annual = (r["og"] - r["oad"] if r["og"] > 0 else cost) * rate
    else:
        annual = 0
    pro = annual * days / days_year
    cap = max(0, cost * (1 - res) - r["oad"])
    dep = 0 if method in ("None", "Not set") else min(pro, cap)
    exp_rel = ((r["disp"] + r["wo"]) / cost * (r["oad"] + dep)) if (disposed and cost > 0) else 0
    return dep, exp_rel


def analyse(rows, policy, fy0, fy1, tol):
    """policy: {class: (method, life, residual)}. Returns dict of counts and totals."""
    out = dict(recalc=0.0, reg_dep=0.0, dep_flags=0, rel_flags=0, exceptions=0, flagged_ids=[])
    ids = {}
    for r in rows:
        ids[r["id"]] = ids.get(r["id"], 0) + 1
    for r in rows:
        pol = policy.get(r["cls"], ("Not set", 0, 0))
        dep, exp_rel = recalc_asset(r, pol, fy0, fy1)
        out["recalc"] += dep
        out["reg_dep"] += r["dep"]
        dflag = abs(r["dep"] - dep) > tol
        rflag = abs(r["adrel"] - exp_rel) > tol
        out["dep_flags"] += dflag
        out["rel_flags"] += rflag
        disposed = (r["disp"] + r["wo"]) > 0
        f = 0
        f += abs(r["cg"] - (r["og"] + r["add"] - r["disp"] - r["wo"])) > tol
        f += abs(r["cad"] - (r["oad"] + r["dep"] - r["adrel"])) > tol
        f += abs(r["nbvcy"] - (r["cg"] - r["cad"])) > tol
        f += abs(r["nbvpy"] - (r["og"] - r["oad"])) > tol
        pd_, sd = r["pdate"], r["sdate"]
        f += bool((pd_ is None and r["og"] + r["add"] > 0) or (pd_ and pd_ > fy1))
        f += bool((r["add"] > 0 and pd_ and pd_ < fy0) or (r["add"] == 0 and pd_ and fy0 <= pd_ <= fy1))
        f += bool((disposed and not sd) or (sd and not disposed) or (sd and (sd < fy0 or sd > fy1)))
        f += bool(r["og"] > 0 and r["add"] > 0)
        f += bool(min(r[k] for k in ("og", "add", "disp", "wo", "cg", "oad", "dep", "adrel", "cad", "nbvcy", "nbvpy", "proc")) < 0)
        f += ids[r["id"]] > 1
        f += (r["cad"] - r["cg"]) > tol
        f += bool(pol[0] == "None" and r["dep"] > 0)
        f += pol[0] == "Not set"
        f += dflag + rflag
        if f:
            out["exceptions"] += 1
            out["flagged_ids"].append(r["id"])
    return out


# ------------------------------------------------------------------ styling
FONT = "Arial"
NAVY, WHITE, GREY = "1F3864", "FFFFFF", "595959"
YELLOW, LBLUE, GREEN_F, RED_F, AMBER_F = "FFFFCC", "DCE6F1", "C6EFCE", "FFC7CE", "FFEB9C"
BLUE, GREENTXT, REDTXT, AMBERTXT = "0000FF", "006100", "9C0006", "9C5700"
NUM = '#,##0;(#,##0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
PCT = "0.0%"
DATE = "dd-mmm-yyyy"
thin = Side(style="thin", color="BFBFBF")


def put(ws, ref, v=None, fmt=None, bold=False, italic=False, color=None, fill=None, size=10, wrap=False, align=None, border=False):
    c = ws[ref]
    if v is not None:
        c.value = v
    c.font = Font(name=FONT, size=size, bold=bold, italic=italic, color=color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = PatternFill("solid", start_color=fill)
    if wrap or align:
        c.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal=align)
    if border:
        c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
    return c


def band(ws, r, text, c1="B", c2="F"):
    for col in range(ord(c1), ord(c2) + 1):
        put(ws, f"{chr(col)}{r}", None, bold=True, color=WHITE, fill=NAVY)
    put(ws, f"{c1}{r}", text, bold=True, color=WHITE, fill=NAVY)


def inp(ws, ref, v, fmt=None, wrap=False):
    return put(ws, ref, v, fmt=fmt, color=BLUE, fill=YELLOW, wrap=wrap, border=True)


def head_row(ws, r, labels, start=1, height=44):
    for j, t in enumerate(labels):
        c = ws.cell(r, start + j, t)
        c.font = Font(name=FONT, size=10, bold=True, color=WHITE)
        c.fill = PatternFill("solid", start_color=NAVY)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[r].height = height


def landscape(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def status_cf(ws, rng, first):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first},2)="OK"'], fill=PatternFill("solid", bgColor=GREEN_F), font=Font(name=FONT, color=GREENTXT, bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first},5)="CHECK"'], fill=PatternFill("solid", bgColor=RED_F), font=Font(name=FONT, color=REDTXT, bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first},6)="REVIEW"'], fill=PatternFill("solid", bgColor=AMBER_F), font=Font(name=FONT, color=AMBERTXT, bold=True)))


def top_block(ws, title):
    put(ws, "A1", title, bold=True, size=14, color=NAVY)
    for r, lab, f, fmt in ((2, "Client", "=Summary!$C$6", None), (3, "Year ended", "=Summary!$C$7", DATE)):
        put(ws, f"A{r}", lab, bold=True)
        put(ws, f"B{r}", f, fmt=fmt, color="008000", align="left")


# ------------------------------------------------------------------ workbook
REG_COLS = [("A", "id"), ("B", "desc"), ("C", "cls"), ("D", "pdate"), ("E", "sdate"), ("F", "og"), ("G", "add"), ("H", "disp"),
            ("I", "wo"), ("J", "cg"), ("K", "oad"), ("L", "dep"), ("M", "adrel"), ("N", "cad"), ("O", "nbvcy"), ("P", "nbvpy"), ("Q", "proc")]
REG_HEAD = ["Asset ID", "Description", "Asset class", "Date of purchase", "Date of sale / write-off", "Opening gross value", "Additions",
            "Disposals (gross)", "Write-off (gross)", "Closing gross value", "Opening accumulated depreciation", "Depreciation for the year",
            "Acc. depreciation on disposals / write-off", "Closing accumulated depreciation", "Net book value - current year",
            "Net book value - prior year", "Sale proceeds"]


def build(rows, P, out_path, notes=None):
    N = len(rows)
    fy1 = P["year_end"]
    fy0 = P.get("fy_start") or (dt.date(fy1.year - 1, fy1.month, fy1.day) + dt.timedelta(days=1) if not (fy1.month == 2 and fy1.day == 29)
                               else dt.date(fy1.year - 1, 3, 1))
    tol = float(P.get("tol", 1))
    classes = []
    for r in rows:
        if r["cls"] not in classes:
            classes.append(r["cls"])
    NC = len(classes)
    policy, pnotes = {}, {}
    for c in classes:
        m, life, note = default_policy(c, P.get("default_method", "SLM"))
        policy[c] = (m, life, P.get("residual", 0.05) if m != "None" else 0.0)
        pnotes[c] = note
    om = float(P["om"])
    pm = om * P["pm_pct"]
    thr = float(P["add_thr"]) if P.get("add_thr") not in (None, "") else pm
    ana = analyse(rows, policy, fy0, fy1, tol)

    R1, RL = 8, 7 + N
    rg = lambda col: f"Register!${col}${R1}:${col}${RL}"
    adds = [i for i, r in enumerate(rows) if r["add"] > 0]
    disps = [i for i, r in enumerate(rows) if (r["disp"] + r["wo"]) > 0]
    A1_, AL = 8, max(8, 7 + len(adds))
    D1_, DL = 8, max(8, 7 + len(disps))
    PR1, PRL = 7, 6 + NC + 10          # parameter table rows (10 spare)
    pcls, pm_, pl_, pr_, pt_ = (f"Parameters!${c}${PR1}:${c}${PRL}" for c in "BCDEF")

    wb = Workbook()
    wsS = wb.active
    wsS.title = "Summary"
    wsF = wb.create_sheet("Roll-forward")
    wsD = wb.create_sheet("Depn Recalc")
    wsA = wb.create_sheet("Additions")
    wsX = wb.create_sheet("Disposals")
    wsC = wb.create_sheet("Checks")
    wsR = wb.create_sheet("Register")
    wsP = wb.create_sheet("Parameters")

    # ---------------- Parameters
    ws = wsP
    put(ws, "B2", "PARAMETERS - depreciation policy by asset class", bold=True, size=14, color=NAVY)
    put(ws, "B3", "Defaults are ILLUSTRATIVE (Companies Act Sch II style lives, 5% residual). Confirm against the client's accounting policy. "
                  "Blue cells are editable; spare rows at the bottom can hold extra classes.", italic=True, color=GREY)
    head_row(ws, 6, ["Asset class (must match register)", "Method (SLM / WDV / None)", "Useful life (years)", "Residual value %", "WDV rate (calculated)", "Basis / note"], start=2, height=32)
    for k in range(NC + 10):
        r = PR1 + k
        if k < NC:
            c = classes[k]
            m, life, res = policy[c]
            inp(ws, f"B{r}", c)
            inp(ws, f"C{r}", m)
            inp(ws, f"D{r}", life, fmt="0")
            inp(ws, f"E{r}", res, fmt="0.0%")
            put(ws, f"G{r}", pnotes[c], italic=True, color=GREY)
        else:
            for col in "BCDE":
                inp(ws, f"{col}{r}", None, fmt="0.0%" if col == "E" else None)
        put(ws, f"F{r}", f'=IF(B{r}="","",IF(AND(D{r}>0,E{r}<1),1-E{r}^(1/D{r}),0))', fmt="0.00%", border=True)
    dv = DataValidation(type="list", formula1='"SLM,WDV,None"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"C{PR1}:C{PRL}")
    put(ws, f"B{PRL + 2}", "Notes", bold=True)
    for i, t in enumerate(["SLM: annual depreciation = cost x (1 - residual %) / useful life.",
                           "WDV: annual depreciation = opening net book value (cost for new additions) x rate; rate = 1 - residual%^(1/life) unless you overwrite column F.",
                           "Pro-rata by days from the later of purchase date and year start, to the earlier of sale date and year end.",
                           "Depreciation is capped so accumulated depreciation never goes below the residual value.",
                           "Method 'None' (land, CWIP): no depreciation recalculated."]):
        put(ws, f"B{PRL + 3 + i}", t, color=GREY)
    for c, w in zip("ABCDEFG", (3, 34, 22, 16, 16, 18, 52)):
        ws.column_dimensions[c].width = w

    # ---------------- Register
    ws = wsR
    top_block(ws, "FIXED ASSET REGISTER (as provided)")
    put(ws, "A5", "Figures are as per the client's register. All checks recompute from these cells. " + " ".join(notes or []), italic=True, color=GREY)
    head_row(ws, 7, REG_HEAD)
    for k, r in enumerate(rows):
        rr = R1 + k
        for col, key in REG_COLS:
            c = ws[f"{col}{rr}"]
            c.value = r[key] if r[key] not in ("",) else None
            c.font = Font(name=FONT, size=10, color=BLUE)
            if key in ("pdate", "sdate"):
                c.number_format = DATE
            elif key in NUM_KEYS:
                c.number_format = NUM
    tr = RL + 1
    put(ws, f"C{tr}", "Total", bold=True)
    for col, key in REG_COLS:
        if key in NUM_KEYS:
            put(ws, f"{col}{tr}", f"=SUM({col}{R1}:{col}{RL})", fmt=NUM, bold=True)
    for col, w in zip("ABCDEFGHIJKLMNOPQ", (14, 32, 22, 14, 14, 16, 14, 14, 14, 16, 18, 16, 18, 18, 16, 16, 14)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "D8"
    ws.auto_filter.ref = f"A7:Q{RL}"

    # ---------------- Depn Recalc
    ws = wsD
    top_block(ws, "DEPRECIATION RECALCULATION - per asset")
    put(ws, "A5", "Method, life and residual come from the Parameters tab. Differences above the per-asset tolerance (Summary) are flagged.", italic=True, color=GREY)
    head_row(ws, 7, ["Asset ID", "Asset class", "Method", "Life (yrs)", "Residual %", "WDV rate", "Cost (opening + additions)", "Depreciation from", "Depreciation to",
                     "Days", "Annual depreciation", "Pro-rata depreciation", "Cap (to residual)", "Recalculated depreciation", "Depreciation per register",
                     "Difference (register - recalc)", "Flag", "Expected acc dep released", "Acc dep released per register", "Difference", "Flag"], height=56)
    for k in range(N):
        r = R1 + k
        g = lambda col, r=r: f"Register!{col}{r}"
        f = {
            "A": f"={g('A')}",
            "B": f"={g('C')}",
            "C": f'=IFERROR(IF(INDEX({pm_},MATCH(B{r},{pcls},0))="","Not set",INDEX({pm_},MATCH(B{r},{pcls},0))),"Not set")',
            "D": f"=IFERROR(INDEX({pl_},MATCH(B{r},{pcls},0)),0)",
            "E": f"=IFERROR(INDEX({pr_},MATCH(B{r},{pcls},0)),0)",
            "F": f"=IFERROR(INDEX({pt_},MATCH(B{r},{pcls},0)),0)",
            "G": f"={g('F')}+{g('G')}",
            "H": f'=IF({g("D")}="",Summary!$C$8,MAX({g("D")},Summary!$C$8))',
            "I": f'=IF({g("H")}+{g("I")}>0,IF({g("E")}="",Summary!$C$7,MIN({g("E")},Summary!$C$7)),Summary!$C$7)',
            "J": f"=MAX(0,I{r}-H{r}+1)",
            "K": f'=IF(C{r}="SLM",IF(D{r}>0,G{r}*(1-E{r})/D{r},0),IF(C{r}="WDV",IF({g("F")}>0,{g("F")}-{g("K")},G{r})*F{r},0))',
            "L": f"=K{r}*J{r}/Summary!$C$9",
            "M": f"=MAX(0,G{r}*(1-E{r})-{g('K')})",
            "N": f'=IF(OR(C{r}="None",C{r}="Not set"),0,MIN(L{r},M{r}))',
            "O": f"={g('L')}",
            "P": f"=O{r}-N{r}",
            "Q": f"=IF(ABS(P{r})>Summary!$C$22,1,0)",
            "R": f"=IF(AND({g('H')}+{g('I')}>0,G{r}>0),({g('H')}+{g('I')})/G{r}*({g('K')}+N{r}),0)",
            "S": f"={g('M')}",
            "T": f"=S{r}-R{r}",
            "U": f"=IF(ABS(T{r})>Summary!$C$22,1,0)",
        }
        for col, v in f.items():
            c = ws[f"{col}{r}"]
            c.value = v
            c.font = Font(name=FONT, size=10, color="008000" if col in "ABGOS" else "000000")
            c.number_format = {"H": DATE, "I": DATE, "D": "0", "E": "0.0%", "F": "0.00%", "J": "0", "Q": "0", "U": "0"}.get(col, NUM)
    tr = RL + 1
    put(ws, f"A{tr}", "Total", bold=True)
    for col in "GKLNOPQRSTU":
        put(ws, f"{col}{tr}", f"=SUM({col}{R1}:{col}{RL})", fmt=NUM if col not in "QU" else "0", bold=True)
    ws.conditional_formatting.add(f"Q{R1}:Q{RL}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", bgColor=RED_F)))
    ws.conditional_formatting.add(f"U{R1}:U{RL}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", bgColor=RED_F)))
    for col, w in zip("ABCDEFGHIJKLMNOPQRSTU", (14, 22, 10, 9, 10, 9, 16, 13, 13, 7, 14, 14, 14, 15, 15, 15, 7, 15, 15, 13, 7)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C8"

    # ---------------- Checks
    ws = wsC
    top_block(ws, "REGISTER INTEGRITY CHECKS - per asset")
    put(ws, "A5", "Columns C-F show differences (should be nil). Columns G-Q are flags (1 = exception). Column R is information only.", italic=True, color=GREY)
    head_row(ws, 7, ["Asset ID", "Asset class", "Closing gross vs roll-forward", "Closing acc dep vs roll-forward", "NBV CY vs gross - acc dep", "NBV PY vs opening gross - acc dep",
                     "Purchase date missing / after year end", "Additions vs purchase date", "Disposal vs sale date", "Mixed row (opening + additions)", "Negative value",
                     "Duplicate asset ID", "Acc dep exceeds gross", "Depreciation on non-depreciable class", "Class not in Parameters", "Depreciation differs from recalc",
                     "Acc dep released differs", "Info: fully depreciated, still on register", "Exceptions", "Status"], height=70)
    for k in range(N):
        r = R1 + k
        g = lambda col, r=r: f"Register!{col}{r}"
        dd = lambda col, r=r: f"'Depn Recalc'!{col}{r}"
        f = {
            "A": f"={g('A')}", "B": f"={g('C')}",
            "C": f"={g('J')}-({g('F')}+{g('G')}-{g('H')}-{g('I')})",
            "D": f"={g('N')}-({g('K')}+{g('L')}-{g('M')})",
            "E": f"={g('O')}-({g('J')}-{g('N')})",
            "F": f"={g('P')}-({g('F')}-{g('K')})",
            "G": f'=IF(OR(AND({g("D")}="",{g("F")}+{g("G")}>0),AND({g("D")}<>"",{g("D")}>Summary!$C$7)),1,0)',
            "H": f'=IF(OR(AND({g("G")}>0,{g("D")}<>"",{g("D")}<Summary!$C$8),AND({g("G")}=0,{g("D")}<>"",{g("D")}>=Summary!$C$8,{g("D")}<=Summary!$C$7)),1,0)',
            "I": f'=IF(OR(AND({g("H")}+{g("I")}>0,{g("E")}=""),AND({g("E")}<>"",{g("H")}+{g("I")}=0),AND({g("E")}<>"",OR({g("E")}<Summary!$C$8,{g("E")}>Summary!$C$7))),1,0)',
            "J": f"=IF(AND({g('F')}>0,{g('G')}>0),1,0)",
            "K": f"=IF(OR(MIN(Register!F{r}:P{r})<0,{g('Q')}<0),1,0)",
            "L": f"=IF(COUNTIF({rg('A')},{g('A')})>1,1,0)",
            "M": f"=IF({g('N')}-{g('J')}>Summary!$C$22,1,0)",
            "N": f'=IF(AND({dd("C")}="None",{g("L")}>0),1,0)',
            "O": f'=IF({dd("C")}="Not set",1,0)',
            "P": f"={dd('Q')}",
            "Q": f"={dd('U')}",
            "R": f'=IF(AND({g("E")}="",{g("J")}>0,{g("J")}-{g("N")}<={dd("G")}*{dd("E")}+Summary!$C$22,{dd("C")}<>"None"),1,0)',
            "S": f"=(ABS(C{r})>Summary!$C$22)+(ABS(D{r})>Summary!$C$22)+(ABS(E{r})>Summary!$C$22)+(ABS(F{r})>Summary!$C$22)+SUM(G{r}:Q{r})",
            "T": f'=IF(S{r}>0,"CHECK","OK")',
        }
        for col, v in f.items():
            c = ws[f"{col}{r}"]
            c.value = v
            c.font = Font(name=FONT, size=10, color="008000" if col in "AB" else "000000")
            if col in "CDEF":
                c.number_format = NUM
    for col in "GHIJKLMNOPQR":
        ws.conditional_formatting.add(f"{col}{R1}:{col}{RL}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", bgColor=RED_F if col != "R" else AMBER_F)))
    for col in "CDEF":
        ws.conditional_formatting.add(f"{col}{R1}:{col}{RL}", FormulaRule(formula=[f"ABS({col}{R1})>Summary!$C$22"], fill=PatternFill("solid", bgColor=RED_F)))
    status_cf(ws, f"T{R1}:T{RL}", f"T{R1}")
    for col, w in zip("ABCDEFGHIJKLMNOPQRST", (14, 22, 14, 14, 14, 14, 13, 13, 13, 13, 11, 11, 11, 13, 12, 13, 12, 14, 11, 10)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C8"
    ws.auto_filter.ref = f"A7:T{RL}"

    # ---------------- Roll-forward
    ws = wsF
    top_block(ws, "ROLL-FORWARD BY ASSET CLASS")
    put(ws, "A5", "Computed from the register by class. 'Per register' columns are the closing figures the client's register shows.", italic=True, color=GREY)
    sif = lambda col, r: f"SUMIFS({rg(col)},{rg('C')},$A{r})"
    sif2 = lambda col, r: f"SUMIFS('Depn Recalc'!${col}${R1}:${col}${RL},'Depn Recalc'!$B${R1}:$B${RL},$A{r})"
    tables = []
    row = 7

    def table(title, heads, makers, fmts, total_cols):
        nonlocal row
        band(ws, row, title, "A", chr(64 + len(heads)))
        head_row(ws, row + 1, heads, height=44)
        first = row + 2
        for i, cname in enumerate(classes):
            r = first + i
            ws[f"A{r}"].value = cname
            ws[f"A{r}"].font = Font(name=FONT, size=10)
            for j, mk in enumerate(makers):
                c = ws.cell(r, 2 + j, mk(r))
                c.font = Font(name=FONT, size=10)
                c.number_format = fmts[j]
        tr_ = first + NC
        put(ws, f"A{tr_}", "Total", bold=True)
        for j in range(len(makers)):
            L = chr(66 + j)
            if (2 + j) in total_cols:
                put(ws, f"{L}{tr_}", total_cols[2 + j](first, tr_ - 1, L), fmt=fmts[j], bold=True)
            else:
                put(ws, f"{L}{tr_}", f"=SUM({L}{first}:{L}{tr_ - 1})", fmt=fmts[j], bold=True)
        tables.append((first, tr_))
        row = tr_ + 3
        return first, tr_

    # A gross block
    gA = table("A. GROSS BLOCK", ["Asset class", "Opening", "Additions", "Disposals", "Write-off", "Closing (computed)", "Closing per register", "Difference"],
               [lambda r: "=" + sif("F", r), lambda r: "=" + sif("G", r), lambda r: "=" + sif("H", r), lambda r: "=" + sif("I", r),
                lambda r: f"=B{r}+C{r}-D{r}-E{r}", lambda r: "=" + sif("J", r), lambda r: f"=G{r}-F{r}"], [NUM] * 7, {})
    gB = table("B. ACCUMULATED DEPRECIATION", ["Asset class", "Opening", "Depreciation for the year", "Released on disposals / write-off", "Closing (computed)", "Closing per register", "Difference"],
               [lambda r: "=" + sif("K", r), lambda r: "=" + sif("L", r), lambda r: "=" + sif("M", r), lambda r: f"=B{r}+C{r}-D{r}",
                lambda r: "=" + sif("N", r), lambda r: f"=F{r}-E{r}"], [NUM] * 6, {})
    band(ws, row, "C. NET BOOK VALUE", "A", "H")
    head_row(ws, row + 1, ["Asset class", "Prior year (opening gross - opening acc dep)", "Prior year per register", "Difference",
                           "Current year (closing gross - closing acc dep)", "Current year per register", "Difference", "Movement in NBV"], height=56)
    c1 = row + 2
    for i, cname in enumerate(classes):
        r = c1 + i
        a, b = gA[0] + i, gB[0] + i
        vals = {"A": cname, "B": f"=B{a}-B{b}", "C": "=" + sif("P", r), "D": f"=C{r}-B{r}", "E": f"=F{a}-F{b}", "F": "=" + sif("O", r),
                "G": f"=F{r}-E{r}", "H": f"=E{r}-B{r}"}
        for col, v in vals.items():
            c = ws[f"{col}{r}"]
            c.value = v
            c.font = Font(name=FONT, size=10)
            if col != "A":
                c.number_format = NUM
    tC = c1 + NC
    put(ws, f"A{tC}", "Total", bold=True)
    for col in "BCDEFGH":
        put(ws, f"{col}{tC}", f"=SUM({col}{c1}:{col}{tC - 1})", fmt=NUM, bold=True)
    row = tC + 3
    # D depreciation
    band(ws, row, "D. DEPRECIATION - REGISTER vs RECALCULATION", "A", "I")
    head_row(ws, row + 1, ["Asset class", "Depreciation per register", "Recalculated", "Difference", "Difference % of register", "Average gross block",
                           "Effective rate (register dep / average gross)", "Method (Parameters)", "Assets flagged"], height=56)
    d1 = row + 2
    for i, cname in enumerate(classes):
        r = d1 + i
        a = gA[0] + i
        vals = {"A": cname, "B": "=" + sif("L", r), "C": "=" + sif2("N", r), "D": f"=B{r}-C{r}", "E": f"=IF(B{r}=0,0,D{r}/B{r})",
                "F": f"=(B{a}+G{a})/2", "G": f"=IF(F{r}=0,0,B{r}/F{r})",
                "H": f'=IFERROR(INDEX({pm_},MATCH($A{r},{pcls},0)),"Not set")', "I": f"=SUMIFS('Depn Recalc'!$Q${R1}:$Q${RL},'Depn Recalc'!$B${R1}:$B${RL},$A{r})"}
        for col, v in vals.items():
            c = ws[f"{col}{r}"]
            c.value = v
            c.font = Font(name=FONT, size=10)
            c.number_format = {"E": PCT, "G": PCT, "I": "0", "A": "General", "H": "General"}.get(col, NUM)
    tD = d1 + NC
    put(ws, f"A{tD}", "Total", bold=True)
    for col in "BCDFI":
        put(ws, f"{col}{tD}", f"=SUM({col}{d1}:{col}{tD - 1})", fmt=NUM if col != "I" else "0", bold=True)
    put(ws, f"E{tD}", f"=IF(B{tD}=0,0,D{tD}/B{tD})", fmt=PCT, bold=True)
    put(ws, f"G{tD}", f"=IF(F{tD}=0,0,B{tD}/F{tD})", fmt=PCT, bold=True)
    for (f_, l_), col in ((tables[0], "H"), (tables[1], "G")):
        ws.conditional_formatting.add(f"{col}{f_}:{col}{l_}", FormulaRule(formula=[f"ABS({col}{f_})>Summary!$C$22"], fill=PatternFill("solid", bgColor=RED_F)))
    for col, w in zip("ABCDEFGHI", (30, 18, 16, 18, 18, 18, 18, 16, 14)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B7"

    # ---------------- Additions
    ws = wsA
    top_block(ws, "ADDITIONS TESTING")
    put(ws, "A5", "Additions at or above the threshold on the Summary tab are 'Key' and selected for vouching. Complete the blue columns. For the remainder, use the Audit Sampling tool.", italic=True, color=GREY)
    head_row(ws, 7, ["#", "Asset ID", "Description", "Asset class", "Date of purchase", "Cost of addition", "Purchase date within year?", "Key addition?", "Selected for vouching",
                     "Supplier / invoice ref", "Vouched to invoice (Y/N)", "Capital in nature (Y/N)", "Capitalisation date agreed (Y/N)", "Comments / WP ref"], height=56)
    for k, i in enumerate(adds):
        r = A1_ + k
        rr = R1 + i
        vals = {"A": k + 1, "B": f"=Register!A{rr}", "C": f"=Register!B{rr}", "D": f"=Register!C{rr}", "E": f"=Register!D{rr}", "F": f"=Register!G{rr}",
                "G": f'=IF(E{r}="","NO DATE",IF(AND(E{r}>=Summary!$C$8,E{r}<=Summary!$C$7),"Yes","NO"))',
                "H": f'=IF(F{r}>=Summary!$C$21,"Key","")', "I": f'=IF(H{r}="Key","Yes","")'}
        for col, v in vals.items():
            c = ws[f"{col}{r}"]
            c.value = v
            c.font = Font(name=FONT, size=10, color="008000" if col in "BCDEF" else "000000")
            if col == "E":
                c.number_format = DATE
            if col == "F":
                c.number_format = NUM
        for col in "JKLMN":
            ws[f"{col}{r}"].fill = PatternFill("solid", start_color=YELLOW)
            ws[f"{col}{r}"].font = Font(name=FONT, size=10, color=BLUE)
    if adds:
        tr = AL + 1
        put(ws, f"E{tr}", "Total", bold=True)
        put(ws, f"F{tr}", f"=SUM(F{A1_}:F{AL})", fmt=NUM, bold=True)
    else:
        put(ws, "B8", "No additions in the register.", italic=True)
    ws.conditional_formatting.add(f"G{A1_}:G{AL}", FormulaRule(formula=[f'LEFT(G{A1_},2)="NO"'], fill=PatternFill("solid", bgColor=RED_F)))
    ws.conditional_formatting.add(f"H{A1_}:H{AL}", FormulaRule(formula=[f'H{A1_}="Key"'], fill=PatternFill("solid", bgColor=AMBER_F)))
    for col, w in zip("ABCDEFGHIJKLMN", (5, 14, 30, 20, 14, 16, 13, 10, 12, 22, 12, 12, 14, 34)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C8"

    # ---------------- Disposals
    ws = wsX
    top_block(ws, "DISPOSALS AND WRITE-OFFS TESTING")
    put(ws, "A5", "Profit/(loss) = proceeds - net book value disposed. Expected acc dep released comes from the depreciation recalculation.", italic=True, color=GREY)
    head_row(ws, 7, ["#", "Asset ID", "Description", "Asset class", "Type", "Date of sale / write-off", "Gross cost disposed", "Acc dep released (register)", "NBV disposed",
                     "Sale proceeds", "Profit / (loss)", "Expected acc dep released (recalc)", "Difference", "Approved (Y/N)", "Proceeds agreed to bank (Y/N)", "Comments / WP ref"], height=56)
    for k, i in enumerate(disps):
        r = D1_ + k
        rr = R1 + i
        vals = {"A": k + 1, "B": f"=Register!A{rr}", "C": f"=Register!B{rr}", "D": f"=Register!C{rr}",
                "E": f'=IF(AND(Register!I{rr}>0,Register!H{rr}>0),"Sale + write-off",IF(Register!I{rr}>0,"Write-off","Sale"))',
                "F": f"=Register!E{rr}", "G": f"=Register!H{rr}+Register!I{rr}", "H": f"=Register!M{rr}", "I": f"=G{r}-H{r}",
                "J": f"=Register!Q{rr}", "K": f"=J{r}-I{r}", "L": f"='Depn Recalc'!R{rr}", "M": f"=H{r}-L{r}"}
        for col, v in vals.items():
            c = ws[f"{col}{r}"]
            c.value = v
            c.font = Font(name=FONT, size=10, color="008000" if col in "BCDFHJ" else "000000")
            if col == "F":
                c.number_format = DATE
            elif col in "GHIJKLM":
                c.number_format = NUM
        for col in "NOP":
            ws[f"{col}{r}"].fill = PatternFill("solid", start_color=YELLOW)
            ws[f"{col}{r}"].font = Font(name=FONT, size=10, color=BLUE)
    if disps:
        tr = DL + 1
        put(ws, f"F{tr}", "Total", bold=True)
        for col in "GHIJKLM":
            put(ws, f"{col}{tr}", f"=SUM({col}{D1_}:{col}{DL})", fmt=NUM, bold=True)
    else:
        put(ws, "B8", "No disposals or write-offs in the register.", italic=True)
    ws.conditional_formatting.add(f"M{D1_}:M{DL}", FormulaRule(formula=[f"ABS(M{D1_})>Summary!$C$22"], fill=PatternFill("solid", bgColor=RED_F)))
    for col, w in zip("ABCDEFGHIJKLMNOP", (5, 14, 28, 20, 14, 14, 15, 15, 15, 14, 14, 16, 12, 11, 14, 32)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C8"

    # ---------------- Summary
    ws = wsS
    put(ws, "B2", "FIXED ASSET WORKPAPER", bold=True, size=16, color=NAVY)
    put(ws, "B3", "Yellow cells with blue text are inputs. Everything else is a live formula.", italic=True, color=GREY)

    def line(r, label, value=None, note=None, kind="calc", fmt=None, wrap=False):
        put(ws, f"B{r}", label)
        if kind == "in":
            inp(ws, f"C{r}", value, fmt=fmt, wrap=wrap)
        else:
            put(ws, f"C{r}", value, fmt=fmt, border=True, wrap=wrap, align="right" if fmt else None)
        if note:
            put(ws, f"D{r}", note, italic=True, color=GREY)

    band(ws, 5, "A. ENGAGEMENT")
    line(6, "Client", P["client"], kind="in")
    line(7, "Year ended", fy1, kind="in", fmt=DATE)
    line(8, "Financial year start", fy0, "Defaults to one year before year end + 1 day", "in", DATE)
    line(9, "Days in year", "=C7-C8+1", None, "calc", "0")
    line(10, "Currency / units", P.get("currency", ""), kind="in")
    line(11, "Prepared by", P.get("prepared_by", ""), kind="in")
    line(12, "Prepared on", dt.datetime.now(), kind="in", fmt=DATE)
    line(13, "Reviewed by / date", None, kind="in")
    band(ws, 15, "B. MATERIALITY AND THRESHOLDS")
    line(16, "Overall materiality", om, "From your materiality workpaper", "in", NUM)
    line(17, "Performance materiality %", P["pm_pct"], None, "in", "0%")
    line(18, "Performance materiality", "=C16*C17", None, "calc", NUM)
    line(19, "Clearly trivial %", P["trivial_pct"], None, "in", "0%")
    line(20, "Clearly trivial threshold", "=C16*C19", None, "calc", NUM)
    if P.get("add_thr") in (None, ""):
        line(21, "Key additions threshold", "=C18", "Additions at or above are key (defaults to performance materiality)", "in", NUM)
    else:
        line(21, "Key additions threshold", thr, "Additions at or above are key", "in", NUM)
    line(22, "Per-asset tolerance (amount)", tol, "Rounding tolerance used by all checks and recalculation flags", "in", NUM2)

    band(ws, 24, "C. REGISTER OVERVIEW")
    line(25, "Number of assets (rows)", f"=COUNTA({rg('A')})", None, "calc", NUM)
    line(26, "Number of asset classes", NC, None, "calc", NUM)
    for r, lab, col in ((27, "Opening gross value", "F"), (28, "Additions", "G"), (29, "Disposals (gross)", "H"), (30, "Write-off (gross)", "I"),
                        (31, "Closing gross value (register)", "J"), (32, "Closing accumulated depreciation (register)", "N"),
                        (33, "Net book value - current year (register)", "O"), (34, "Net book value - prior year (register)", "P")):
        line(r, lab, f"=SUM({rg(col)})", None, "calc", NUM)

    n_unrec = sum(1 for c in classes if pnotes[c].startswith("Not recognised"))
    line(35, "Classes with no recognised default policy (default applied - confirm on Parameters tab)", n_unrec, None, "calc", NUM)
    band(ws, 36, "D. REGISTER INTEGRITY CHECKS (number of assets with exception)", "B", "D")
    head_row(ws, 37, ["Check", "Exceptions", "Status"], start=2, height=22)
    CK = lambda col: f"Checks!${col}${R1}:${col}${RL}"
    checks = [("Closing gross value does not equal opening + additions - disposals - write-off", f"=SUMPRODUCT(--(ABS({CK('C')})>$C$22))"),
              ("Closing accumulated depreciation does not equal opening + depreciation - released", f"=SUMPRODUCT(--(ABS({CK('D')})>$C$22))"),
              ("NBV current year does not equal closing gross - closing accumulated depreciation", f"=SUMPRODUCT(--(ABS({CK('E')})>$C$22))"),
              ("NBV prior year does not equal opening gross - opening accumulated depreciation", f"=SUMPRODUCT(--(ABS({CK('F')})>$C$22))"),
              ("Purchase date missing or after year end", f"=SUM({CK('G')})"),
              ("Additions inconsistent with purchase date", f"=SUM({CK('H')})"),
              ("Disposal / write-off and sale date mismatch", f"=SUM({CK('I')})"),
              ("Mixed rows (opening gross and additions on one row - split for accurate recalculation)", f"=SUM({CK('J')})"),
              ("Negative values in register", f"=SUM({CK('K')})"),
              ("Duplicate asset IDs", f"=SUM({CK('L')})"),
              ("Closing accumulated depreciation exceeds closing gross value", f"=SUM({CK('M')})"),
              ("Depreciation charged on non-depreciable class (method None)", f"=SUM({CK('N')})"),
              ("Asset class not in Parameters tab", f"=SUM({CK('O')})"),
              ("Depreciation differs from recalculation", f"=SUM({CK('P')})"),
              ("Accumulated depreciation released on disposal differs from expected", f"=SUM({CK('Q')})")]
    for k, (lab, f) in enumerate(checks):
        r = 38 + k
        put(ws, f"B{r}", lab, wrap=True)
        put(ws, f"C{r}", f, fmt=NUM, border=True, align="right")
        put(ws, f"D{r}", f'=IF(C{r}=0,"OK","CHECK")', border=True)
    r_info = 38 + len(checks)           # 53
    put(ws, f"B{r_info}", "Information: fully depreciated assets still on register (no status)")
    put(ws, f"C{r_info}", f"=SUM({CK('R')})", fmt=NUM, border=True, align="right")
    put(ws, f"D{r_info}", "Info", border=True)
    r_all = r_info + 1                  # 54
    put(ws, f"B{r_all}", "Assets with at least one exception", bold=True)
    put(ws, f"C{r_all}", f'=COUNTIF({CK("T")},"CHECK")', fmt=NUM, border=True, bold=True, align="right")
    put(ws, f"D{r_all}", f'=IF(C{r_all}=0,"OK","CHECK")', border=True, bold=True)
    status_cf(ws, f"D38:D{r_all}", "D38")
    assert r_all == 54

    band(ws, 56, "E. DEPRECIATION RECALCULATION")
    line(57, "Depreciation per register", f"=SUM({rg('L')})", None, "calc", NUM)
    line(58, "Depreciation recalculated", f"='Depn Recalc'!N{RL + 1}", "Per Parameters tab policy", "calc", NUM)
    line(59, "Difference (register - recalculated)", "=C57-C58", None, "calc", NUM)
    line(60, "Difference as % of register", "=IF(C57=0,0,C59/C57)", None, "calc", PCT)
    line(61, "Assessment", '=IF(ABS(C59)<=C20,"OK - within clearly trivial",IF(ABS(C59)<=C18,"REVIEW - below performance materiality","CHECK - exceeds performance materiality"))', "Investigate flagged assets on the Depn Recalc tab", "calc", wrap=True)
    status_cf(ws, "C61", "C61")

    band(ws, 63, "F. ADDITIONS TESTING")
    ar = lambda col: f"Additions!${col}${A1_}:${col}${AL}"
    line(64, "Number of additions", f'=COUNT({ar("F")})' if adds else 0, None, "calc", NUM)
    line(65, "Value of additions", f"=SUM({ar('F')})" if adds else 0, None, "calc", NUM)
    line(66, "Key additions - number", f'=COUNTIF({ar("H")},"Key")' if adds else 0, None, "calc", NUM)
    line(67, "Key additions - value", f'=SUMIF({ar("H")},"Key",{ar("F")})' if adds else 0, None, "calc", NUM)
    line(68, "Coverage of additions by key items", "=IF(C65=0,0,C67/C65)", None, "calc", PCT)
    line(69, "Additions vouched (marked Y)", f'=COUNTIF({ar("K")},"Y")' if adds else 0, None, "calc", NUM)
    line(70, "Additions with purchase date outside the year", f'=COUNTIF({ar("G")},"NO*")' if adds else 0, None, "calc", NUM)

    band(ws, 72, "G. DISPOSALS AND WRITE-OFFS")
    dr = lambda col: f"Disposals!${col}${D1_}:${col}${DL}"
    line(73, "Number of disposals / write-offs", f'=COUNT({dr("G")})' if disps else 0, None, "calc", NUM)
    line(74, "Gross cost disposed", f"=SUM({dr('G')})" if disps else 0, None, "calc", NUM)
    line(75, "Accumulated depreciation released", f"=SUM({dr('H')})" if disps else 0, None, "calc", NUM)
    line(76, "Net book value disposed", "=C74-C75", None, "calc", NUM)
    line(77, "Sale proceeds", f"=SUM({dr('J')})" if disps else 0, None, "calc", NUM)
    line(78, "Profit / (loss) on disposal", "=C77-C76", "Compare to the P&L", "calc", NUM)

    band(ws, 80, "H. TIE-OUT TO TRIAL BALANCE / LEAD SCHEDULE (enter the amounts per TB in the yellow cells)", "B", "F")
    head_row(ws, 81, ["Item", "Per register", "Per TB / lead", "Difference", "Status"], start=2, height=22)
    tb = P.get("tb") or {}
    tie = [("Closing gross block", "=C31", tb.get("gross")), ("Closing accumulated depreciation", "=C32", tb.get("ad")),
           ("Closing net book value", "=C33", "FORMULA"), ("Opening net book value (prior year closing)", "=C34", tb.get("nbv_py")),
           ("Depreciation charge for the year", "=C57", tb.get("dep")), ("Profit / (loss) on disposal", "=C78", tb.get("pl"))]
    for k, (lab, f, v) in enumerate(tie):
        r = 82 + k
        put(ws, f"B{r}", lab)
        put(ws, f"C{r}", f, fmt=NUM, border=True, align="right")
        if v == "FORMULA":
            put(ws, f"D{r}", '=IF(AND(D82<>"",D83<>""),D82-D83,"")', fmt=NUM, border=True, align="right")
        else:
            inp(ws, f"D{r}", v, fmt=NUM)
        put(ws, f"E{r}", f'=IF(D{r}="","",C{r}-D{r})', fmt=NUM, border=True, align="right")
        put(ws, f"F{r}", f'=IF(D{r}="","Not entered",IF(ABS(E{r})<=$C$20,"OK","CHECK"))', border=True)
    status_cf(ws, "F82:F87", "F82")
    band(ws, 89, "I. OVERALL", "B", "F")
    line(90, "Integrity checks", f"=D{r_all}", None, "calc")
    line(91, "Depreciation assessment", "=LEFT(C61,FIND(\" \",C61)-1)", None, "calc")
    line(92, "TB tie-out", '=IF(COUNTIF(F82:F87,"CHECK")>0,"CHECK",IF(COUNTIF(F82:F87,"OK")=0,"Not entered","OK"))', None, "calc")
    status_cf(ws, "C90:C92", "C90")
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 78
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 14
    ws.freeze_panes = "A5"

    for w_ in (wsS, wsF, wsD, wsA, wsX, wsC, wsR, wsP):
        landscape(w_)
    wsS.sheet_properties.tabColor = NAVY
    wsF.sheet_properties.tabColor = "70AD47"
    for w_ in (wsD, wsA, wsX, wsC):
        w_.sheet_properties.tabColor = "8EA9DB"
    wsR.sheet_properties.tabColor = "A6A6A6"
    wsP.sheet_properties.tabColor = "A6A6A6"
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    return dict(n_assets=N, n_classes=NC, n_add=len(adds), n_disp=len(disps), exceptions=ana["exceptions"],
                dep_reg=ana["reg_dep"], dep_recalc=ana["recalc"], dep_flags=ana["dep_flags"], rel_flags=ana["rel_flags"],
                fy0=fy0, fy1=fy1, policy=policy, thr=thr)


# ------------------------------------------------------------------ template (fictitious data)
def make_template(path):
    rng = random.Random(11)
    fy0, fy1 = dt.date(2025, 4, 1), dt.date(2026, 3, 31)
    spec = [("Freehold Land", 3, (5e6, 25e6)), ("Buildings", 3, (8e6, 30e6)), ("Plant & Machinery", 10, (1.5e6, 12e6)),
            ("Furniture & Fixtures", 5, (2e5, 1.2e6)), ("Office Equipment", 5, (1e5, 6e5)), ("Computers", 6, (5e4, 9e5)), ("Vehicles", 4, (6e5, 2.5e6))]
    assets, n = [], 0
    for cls, cnt, (lo, hi) in spec:
        for _ in range(cnt):
            n += 1
            assets.append(dict(id=f"FA-{n:04d}", cls=cls, cost=round(rng.uniform(lo, hi), -3), desc=f"{cls[:-1] if cls.endswith('s') else cls} #{n}"))
    new_idx = {5, 7, 9, 12, 16, 19, 22, 26, 29, 33}
    sold_idx = {8: 0.55, 20: 0.12, 24: 0.2, 28: 0.88}     # asset index -> sale proceeds as % of cost
    wo_idx = {21}
    rows = []
    for k, a in enumerate(assets):
        pol = default_policy(a["cls"], "SLM")
        method, life, res = pol[0], pol[1], 0.05 if pol[0] != "None" else 0.0
        if k in new_idx:
            pd_ = fy0 + dt.timedelta(days=rng.randint(10, 330))
            og, add = 0, a["cost"]
        else:
            pd_ = dt.date(rng.randint(2016, 2024), rng.randint(1, 12), rng.randint(1, 28))
            og, add = a["cost"], 0
        r = dict(id=a["id"], desc=a["desc"], cls=a["cls"], pdate=pd_, sdate=None, og=og, add=add, disp=0, wo=0, proc=0)
        cost = og + add
        annual = cost * (1 - res) / life if life else 0
        oad = 0 if og == 0 else min(annual * ((fy0 - pd_).days) / 365.0, cost * (1 - res))
        r["oad"] = round(oad)
        if k in sold_idx or k in wo_idx:
            r["sdate"] = fy0 + dt.timedelta(days=rng.randint(60, 300))
            r["disp" if k in sold_idx else "wo"] = cost
            r["proc"] = round(cost * sold_idx[k]) if k in sold_idx else 0
        dep, _ = recalc_asset(r, (method, life, res), fy0, fy1)
        r["dep"] = round(dep)
        disposed = (r["disp"] + r["wo"]) > 0
        r["adrel"] = round(r["oad"] + r["dep"]) if disposed else 0
        r["cg"] = cost - r["disp"] - r["wo"]
        r["cad"] = r["oad"] + r["dep"] - r["adrel"]
        r["nbvcy"] = r["cg"] - r["cad"]
        r["nbvpy"] = r["og"] - r["oad"]
        rows.append(r)
    # three deliberate exceptions so the checks can be seen working
    rows[10]["dep"] = round(rows[10]["dep"] * 1.12)                       # depreciation overstated (cad/nbv left consistent below)
    rows[10]["cad"] = rows[10]["oad"] + rows[10]["dep"]
    rows[10]["nbvcy"] = rows[10]["cg"] - rows[10]["cad"]
    rows[20]["sdate"] = None                                              # disposal without sale date
    rows[14]["cg"] += 5000                                                # closing gross does not foot
    wb = Workbook()
    ws = wb.active
    ws.title = "Instructions"
    lines = ["FIXED ASSET REGISTER TEMPLATE", "",
             "1. Use the 'Register' tab: one row per asset (or per addition). Keep the header row; delete the sample rows and paste your register.",
             "2. An asset bought in the year has Opening gross = 0 and the cost in Additions. Put additions to an existing asset on their own row.",
             "3. Gross figures are cost; disposals and write-offs are at gross cost; accumulated depreciation released goes in its own column.",
             "4. Date of sale / write-off is needed for every row with a disposal or write-off. Sale proceeds is optional (used for profit/loss).",
             "5. Required columns: Asset class, Opening gross value, Additions, Closing gross value, Opening accumulated depreciation, Depreciation for the year, "
             "Closing accumulated depreciation. Others are optional but recommended.",
             "6. Asset class names must match between the register and the depreciation policy (the tool builds the policy table from your classes).",
             "", "The sample data is FICTITIOUS (year ended 31-Mar-2026) and contains 3 deliberate exceptions so you can see the checks work:",
             "   - FA-0011: depreciation overstated by about 12%",
             "   - FA-0021: disposal recorded without a sale date",
             "   - FA-0015: closing gross value does not foot (+5,000)"]
    for i, t in enumerate(lines, 1):
        ws.cell(i, 1, t).font = Font(name=FONT, bold=(i == 1), size=12 if i == 1 else 10)
    ws.column_dimensions["A"].width = 140
    wr = wb.create_sheet("Register")
    for j, h in enumerate(REG_HEAD, 1):
        c = wr.cell(1, j, h)
        c.font = Font(name=FONT, bold=True, color=WHITE)
        c.fill = PatternFill("solid", start_color=NAVY)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    wr.row_dimensions[1].height = 44
    for i, r in enumerate(rows, 2):
        for j, (_, key) in enumerate(REG_COLS, 1):
            c = wr.cell(i, j, r[key] if r[key] not in ("",) else None)
            c.font = Font(name=FONT, color=BLUE)
            if key in ("pdate", "sdate"):
                c.number_format = DATE
            elif key in NUM_KEYS:
                c.number_format = NUM
    for col, w in zip("ABCDEFGHIJKLMNOPQ", (12, 28, 22, 14, 14, 16, 14, 14, 14, 16, 18, 16, 18, 18, 16, 16, 14)):
        wr.column_dimensions[col].width = w
    wr.freeze_panes = "D2"
    wb.save(path)
