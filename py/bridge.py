"""Shared glue between the web forms and the tool modules.

Used by BOTH the Flask server (app.py) and the in-browser build (Pyodide, see build_static.py).
Every function takes the form fields (dict of strings), the uploaded file name and its bytes, and returns
(download_name, xlsx_bytes, summary_list). User mistakes raise ValueError with a readable message.
"""
import datetime as dt
import re
import tempfile
from pathlib import Path

import fixed_assets as fa
import lead_schedule as ls
import sampling_tool as st


def _safe(name):
    return re.sub(r"[^\w\- ]+", "", name).strip().replace(" ", "_") or "Client"


def _opt(form, k):
    v = str(form.get(k, "")).strip()
    return float(v) if v != "" else None


def _date(form, k):
    return dt.date.fromisoformat(str(form.get(k, "")))


def _read_out(path):
    return Path(path).read_bytes()


# ------------------------------------------------------------------ lead schedule
def lead_schedule(form, filename, data):
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("Please upload an Excel (.xlsx) file.")
    try:
        materiality = float(form.get("materiality", ""))
        year_end = _date(form, "year_end")
        pm = float(form.get("pm_pct") or 75) / 100
        tr = float(form.get("trivial_pct") or 5) / 100
    except ValueError:
        raise ValueError("Please check materiality, year-end date and percentages.")
    client = str(form.get("client", "")).strip() or "Client"
    with tempfile.TemporaryDirectory() as d:
        src = Path(d) / "tb.xlsx"
        src.write_bytes(data)
        try:
            rows, _ = ls.read_tb(src)
        except SystemExit as e:
            raise ValueError(str(e) or "Could not read the trial balance.")
        meta = dict(client=client, year_end=dt.datetime.combine(year_end, dt.time()),
                    currency=str(form.get("currency", "")).strip(), basis=str(form.get("basis", "")).strip(),
                    materiality=materiality, pm_pct=pm, trivial_pct=tr,
                    prepared_by=str(form.get("prepared_by", "")).strip(), prepared_on=dt.datetime.now())
        out = Path(d) / "out.xlsx"
        res = ls.build(rows, meta, str(out))
        summary = [["Accounts read", res["n"]], ["Lead tabs", res["tabs"]],
                   ["Unmapped accounts", len(res["unmapped"])], ["Materiality", f"{materiality:,.0f}"]]
        return f"{_safe(client)}_Lead_Schedules_{year_end:%Y%m%d}.xlsx", _read_out(out), summary


# ------------------------------------------------------------------ audit sampling
def sampling(form, filename, data):
    if not filename.lower().endswith((".xlsx", ".xlsm", ".csv")):
        raise ValueError("Please upload an Excel (.xlsx) or CSV file.")
    try:
        P = dict(
            client=str(form.get("client", "")).strip() or "Client", period_end=_date(form, "period_end"),
            area=str(form.get("area", "")).strip(), assertion=form.get("assertion", "Existence"),
            objective=str(form.get("objective", "")).strip(), prepared_by=str(form.get("prepared_by", "")).strip(),
            om=float(form.get("om", "")), pm_pct=float(form.get("pm_pct") or 75) / 100, tm=_opt(form, "tm"),
            trivial_pct=float(form.get("trivial_pct") or 5) / 100,
            ir=form.get("ir", "Moderate"), toc=form.get("toc", st.TOC_OPTS[0]), toc_desc=str(form.get("toc_desc", "")).strip(),
            sap=form.get("sap", "None"), sap_desc=str(form.get("sap_desc", "")).strip(),
            em=float(form.get("em") or 0), sig=bool(form.get("sig")), key_thr=_opt(form, "key_thr"),
            key_ids=form.get("key_ids", ""), key_neg=bool(form.get("key_neg")),
            method=form.get("method", "MUS"), seed=int(form["seed"]) if str(form.get("seed", "")).strip() else None)
        if P["ir"] not in st.IR_OPTS or P["toc"] not in st.TOC_OPTS or P["sap"] not in st.SAP_OPTS or P["method"] not in st.METHODS:
            raise ValueError("Invalid selection.")
    except (ValueError, KeyError) as e:
        raise ValueError(f"Please check the inputs: {e}")
    with tempfile.TemporaryDirectory() as d:
        src = Path(d) / ("pop" + Path(filename).suffix.lower())
        src.write_bytes(data)
        rows = st.read_population(src)
        out = Path(d) / "out.xlsx"
        res = st.build(rows, P, str(out))
        summary = [["Population items", f"{res['n_items']:,}"], ["Key items + top stratum", res["n_key"] + res["n_top"]],
                   ["Sample size", res["n"] if P["method"] != "Haphazard" else f"{res['n']} (you select)"],
                   ["Assessed RMM", res["rmm"]]]
        return f"{_safe(P['client'])}_Sampling_{P['method']}_{P['period_end']:%Y%m%d}.xlsx", _read_out(out), summary


# ------------------------------------------------------------------ fixed assets
def fixed_assets(form, filename, data):
    if not filename.lower().endswith((".xlsx", ".xlsm", ".csv")):
        raise ValueError("Please upload an Excel (.xlsx) or CSV file.")
    try:
        P = dict(
            client=str(form.get("client", "")).strip() or "Client", year_end=_date(form, "year_end"),
            fy_start=_date(form, "fy_start") if str(form.get("fy_start", "")).strip() else None,
            currency=str(form.get("currency", "")).strip(), om=float(form.get("om", "")),
            pm_pct=float(form.get("pm_pct") or 75) / 100, trivial_pct=float(form.get("trivial_pct") or 5) / 100,
            add_thr=_opt(form, "add_thr"), tol=float(form.get("tol") or 1),
            default_method=form.get("default_method", "SLM"), residual=float(form.get("residual") or 5) / 100,
            prepared_by=str(form.get("prepared_by", "")).strip(),
            tb=dict(gross=_opt(form, "tb_gross"), ad=_opt(form, "tb_ad"), nbv_py=_opt(form, "tb_nbv_py"),
                    dep=_opt(form, "tb_dep"), pl=_opt(form, "tb_pl")))
        if P["default_method"] not in ("SLM", "WDV") or not 0 <= P["residual"] < 1:
            raise ValueError("Invalid method or residual %.")
    except (ValueError, KeyError) as e:
        raise ValueError(f"Please check the inputs: {e}")
    with tempfile.TemporaryDirectory() as d:
        src = Path(d) / ("reg" + Path(filename).suffix.lower())
        src.write_bytes(data)
        rows, notes = fa.read_register(src)
        out = Path(d) / "out.xlsx"
        res = fa.build(rows, P, str(out), notes)
        summary = [["Assets read", f"{res['n_assets']:,}"], ["Assets with exceptions", res["exceptions"]],
                   ["Additions / disposals", f"{res['n_add']} / {res['n_disp']}"],
                   ["Depreciation difference", f"{res['dep_reg'] - res['dep_recalc']:,.0f}"]]
        return f"{_safe(P['client'])}_Fixed_Assets_{P['year_end']:%Y%m%d}.xlsx", _read_out(out), summary


TOOLS = {"lead-schedule": lead_schedule, "sampling": sampling, "fixed-assets": fixed_assets}


def run(tool, form, filename, data):
    """Entry point used by the browser worker. Returns (name, bytes, summary); raises ValueError for user errors."""
    try:
        return TOOLS[tool](form, filename, data)
    except ValueError:
        raise
    except Exception as e:                    # unexpected problem with the uploaded file
        raise ValueError(f"Could not process this file: {e}")
