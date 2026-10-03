"""
MC Capacity vs Order — report generator
========================================
Order Status Report (OSR)  ->  containers per machine line per week, against capacity.

The report is produced from a *template*: the bundled reference_workbook.xlsx is the finished report
(layout, colour palette, Excess / Short Order rows, regional sheets, capacity master, lookups).
Each run only refreshes what changes week to week:

  * the 12 week columns (W # headers and the hidden week-Monday row 500 in ORDER VS WEEK DISTRIBUTION)
  * the WORKING sheet, rebuilt from the new OSR
  * MC ROUTING in WORKING, carried forward by PRODUCT NAME (from the previous report you upload, else
    the routing bundled in the template)
  * Machines / Shifts, carried forward from the previous report when one is uploaded

Everything else (capacity, Excess / Short rows, AFRICA / AUSTRALIA & EUROPE / USA sheets) is live formulas in the template.
To change the look, the masters or the capacity basis, replace reference_workbook.xlsx.
"""

import collections
import datetime
import io
import re

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

OVW = "ORDER VS WEEK DISTRIBUTION"
WORKING = "WORKING"
MC_MASTER = "MC MASTER"
ITEM_CFC = "ITEM CFC PER CONTAINER"
OSR_SHEET = "Order Status-By shipment Date"
FIRST_WEEK_COL = 7                      # G
HELPER_ROW = 500                        # hidden row with the Monday of each week
NON_LINE_LABELS = (None, "TOTAL", "Excess Order", "Short Order")

# Spellings that are accepted in MC ROUTING but are not offered in the dropdown
ROUTE_ALIASES = {"CONSTANTA TAG D", "MD 20 4GM", "PEARL PACK PREMIX"}

WORKING_HEADERS = [
    "BUYER NAME", "NAV DOC NO", "Buyer Contact No", "Buyer Requested Shipment Date", "Month-yy",
    "Plan To Ship", "ITEM ID", "Column1", "PRODUCT NAME", "MC ROUTING", "CONCAT", "BLEND NAME", "SFG NAME",
    "NO OF CTN PER CFC", "NO OF TBGS PER CTN", "ORDER QTY", "STOCK IN PROCESS", "READY STOCK",
    "STOCK IN PROCESS + READY STOCK", "PENDING TO PRODUCE IN CASES", "PENDING TO PRODUCE IN TBGS",
    "WEEK START", "REPORTING WEEK", "CAPACITY GROUP", "PENDING PROD IN CONTAINERS"]
WORKING_WIDTHS = {1: 24, 2: 16, 3: 22, 4: 13, 5: 9, 6: 13, 7: 9, 8: 22, 9: 42, 10: 22, 11: 22, 12: 24, 13: 30,
                  14: 10, 15: 10, 16: 10, 17: 10, 18: 10, 19: 12, 20: 12, 21: 13, 22: 12, 23: 13, 24: 16, 25: 14}


# ── Week numbering: continuous from the Monday of the week holding 1 Jan 2026 ─────────────────────────
_J1 = datetime.date(2026, 1, 1)
_W1 = _J1 - datetime.timedelta(days=_J1.weekday())


def weeknum(d):
    mon = d - datetime.timedelta(days=d.weekday())
    return ((mon - _W1).days // 7) + 1


def week_monday(wn):
    return _W1 + datetime.timedelta(weeks=wn - 1)


def monday_of(d):
    return d - datetime.timedelta(days=d.weekday())


def _norm(x):
    """Product-name key: case, repeated and non-breaking spaces ignored."""
    return re.sub(r"\s+", " ", str(x).replace("\xa0", " ")).strip().upper()


def _to_datetime(v):
    if isinstance(v, datetime.datetime):
        return v
    if isinstance(v, datetime.date):
        return datetime.datetime.combine(v, datetime.time())
    try:
        return datetime.datetime.strptime(str(v).strip(), "%d-%m-%Y")
    except Exception:
        return None


# ── Order Status Report ───────────────────────────────────────────────────────────────────────────────

def extract_osr(osr_wb):
    """OSR rows: (doc, buyer, contact, req_date, plan_date, prod_id, prod_name, order_qty, in_process, ready, pending)."""
    if OSR_SHEET not in osr_wb.sheetnames:
        raise ValueError(f"The Order Status Report needs a sheet named '{OSR_SHEET}' "
                         f"(found: {', '.join(osr_wb.sheetnames)}).")
    ws = osr_wb[OSR_SHEET]
    rows, last = [], {"buyer": None, "doc": None, "contact": None, "reqdate": None, "plan": None}
    for r in range(7, ws.max_row + 1):
        prod_id, prod_name = ws.cell(r, 9).value, ws.cell(r, 10).value
        for key, col in (("buyer", 3), ("doc", 4), ("contact", 5), ("reqdate", 6), ("plan", 7)):
            v = ws.cell(r, col).value
            if v not in (None, " ", ""):
                last[key] = v
        if isinstance(prod_id, (int, float)) and prod_name not in (None, " ", ""):
            rows.append((last["doc"], last["buyer"], last["contact"], last["reqdate"], last["plan"],
                         prod_id, prod_name, ws.cell(r, 11).value, ws.cell(r, 12).value,
                         ws.cell(r, 13).value, ws.cell(r, 15).value))
    if not rows:
        raise ValueError("No product rows were found in the Order Status Report.")
    return rows


# ── Reading routing / inputs from a report workbook ───────────────────────────────────────────────────

def read_routes(wb):
    """{normalised product name: MC ROUTING} from the WORKING sheet of a report (row order is irrelevant)."""
    if WORKING not in wb.sheetnames:
        raise ValueError("This workbook has no WORKING sheet, so no routing can be read from it.")
    ws = wb[WORKING]
    heads = {str(ws.cell(1, c).value).strip().upper(): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}
    cn, cr = heads.get("PRODUCT NAME", 9), heads.get("MC ROUTING", 10)
    out = {}
    for r in range(2, ws.max_row + 1):
        name, route = ws.cell(r, cn).value, ws.cell(r, cr).value
        if name in (None, ""):
            continue
        key = _norm(name)
        if route not in (None, "", "NOT IN ROUTE MASTER") or key not in out:
            out[key] = route if route not in ("", "NOT IN ROUTE MASTER") else None
    return out


def read_line_inputs(ov):
    """{machine line: (machines, shifts)} from the yellow input cells of ORDER VS WEEK DISTRIBUTION."""
    out = {}
    for r in range(3, 120):
        name = ov.cell(r, 3).value
        if name in NON_LINE_LABELS:
            continue
        m, s = ov.cell(r, 4).value, ov.cell(r, 5).value
        if isinstance(m, (int, float)) and isinstance(s, (int, float)):
            out[name] = (m, s)
    return out


def routing_table(ov):
    """{ROUTE UPPER: capacity group} from the hidden lookup in columns T:U of ORDER VS WEEK DISTRIBUTION."""
    out = {}
    for r in range(2, 120):
        k, v = ov.cell(r, 20).value, ov.cell(r, 21).value
        if k and v:
            out[str(k)] = v
    return out


# ── WORKING sheet ─────────────────────────────────────────────────────────────────────────────────────

def build_working(wb, osr_rows, routes, table):
    """Rebuild WORKING from the OSR. Returns (rows written, rows routed)."""
    idx = wb.sheetnames.index(WORKING)
    state = wb[WORKING].sheet_state                              # the template keeps WORKING hidden
    del wb[WORKING]
    ww = wb.create_sheet(WORKING, idx)
    ww.sheet_state = state

    candidates = collections.defaultdict(list)                   # capacity group -> routes offered in the dropdown
    dropdown_group = {}                                          # route (upper) -> group, alias spellings excluded
    for k, v in table.items():
        if k.upper() not in ROUTE_ALIASES:
            candidates[v].append(k)
            dropdown_group[k.upper()] = v

    FN = "Trebuchet MS"
    thin = Side(style="thin", color="FFBFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    norm_font = Font(name=FN, size=10)
    head_font = Font(name=FN, bold=True, color="FFC00000", size=10)
    head_fill = PatternFill(start_color="FFDCE6F1", end_color="FFDCE6F1", fill_type="solid")
    drop_fill = PatternFill(start_color="FFD9EAD3", end_color="FFD9EAD3", fill_type="solid")
    for c, h in enumerate(WORKING_HEADERS, 1):
        x = ww.cell(1, c, h)
        x.font, x.fill, x.border = head_font, head_fill, border
        x.alignment = Alignment(horizontal="center", wrap_text=True)

    wk0 = f"'{OVW}'!$G${HELPER_ROW}"
    dvs, routed = {}, 0
    for i, (doc, buyer, contact, reqdate, plandate, pid, pname, oq, sip, rs, pend) in enumerate(osr_rows):
        r = i + 2
        rd, pdt = _to_datetime(reqdate), _to_datetime(plandate)
        ww.cell(r, 1, buyer)
        ww.cell(r, 2, doc)
        ww.cell(r, 3, contact)
        c4 = ww.cell(r, 4, rd if rd else reqdate)
        if rd:
            c4.number_format = "dd-mmm-yy"
        ww.cell(r, 5, f'=IF(ISNUMBER($D{r}),EOMONTH($D{r},-1)+1,"")').number_format = "mmm-yy"
        c6 = ww.cell(r, 6, pdt if pdt else plandate)
        if pdt:
            c6.number_format = "dd-mmm-yy"
        ww.cell(r, 7, pid)
        ww.cell(r, 8, f"=B{r}&I{r}")
        ww.cell(r, 9, pname)
        route = routes.get(_norm(pname))
        mc = ww.cell(r, 10)
        if route:
            routed += 1
            cands = candidates.get(dropdown_group.get(str(route).upper()), [])
            if len(cands) > 1:
                key = tuple(cands)
                if key not in dvs:
                    dvs[key] = DataValidation(type="list", formula1='"' + ",".join(cands) + '"',
                                              allow_blank=True, showDropDown=False)
                    ww.add_data_validation(dvs[key])
                dvs[key].add(mc)
                mc.fill = drop_fill
            mc.value = route
        ww.cell(r, 11, f'=J{r}&"-"&TEXT(W{r},"dd-mmm-yy")')
        ww.cell(r, 12, f"=IFERROR(VLOOKUP($M{r},'ITEM MASTER'!E:H,4,0),\"\")")
        ww.cell(r, 13, f"=IFERROR(VLOOKUP($I{r},'ITEM MASTER'!B:E,4,0),\"\")")
        ww.cell(r, 14, f"=IFERROR(VLOOKUP($I{r},'ITEM MASTER'!B:F,5,0),0)")
        ww.cell(r, 15, f"=IFERROR(VLOOKUP($M{r},'TBGS PER CTN'!B:C,2,0),0)")
        ww.cell(r, 16, oq if isinstance(oq, (int, float)) else 0)
        ww.cell(r, 17, sip if isinstance(sip, (int, float)) else 0)
        ww.cell(r, 18, rs if isinstance(rs, (int, float)) else 0)
        ww.cell(r, 19, f"=Q{r}+R{r}")
        ww.cell(r, 20, pend if isinstance(pend, (int, float)) else 0).number_format = "#,##0"
        ww.cell(r, 21, f"=N{r}*O{r}*T{r}").number_format = "#,##0"
        ww.cell(r, 22, f"=IF(ISNUMBER($D{r}),$D{r}-WEEKDAY($D{r},3),{wk0})").number_format = "dd-mmm-yy"
        ww.cell(r, 23, f"=MAX($V{r},{wk0})").number_format = "dd-mmm-yy"
        ww.cell(r, 24, f"=IFERROR(VLOOKUP($J{r},'{OVW}'!$T:$U,2,0),\"\")")
        ic = f"VLOOKUP($I{r},'{ITEM_CFC}'!$A:$B,2,0)"
        mf = f"VLOOKUP($X{r},'{MC_MASTER}'!$A:$Q,16,0)"
        ww.cell(r, 25, f'=IFERROR($T{r}/IF(IFERROR({ic},0)>0,{ic},{mf}),"")')
        for c in range(1, 26):
            ww.cell(r, c).border = border
            ww.cell(r, c).font = norm_font
    for col, w in WORKING_WIDTHS.items():
        ww.column_dimensions[L(col)].width = w
    ww.row_dimensions[1].height = 35.05
    ww.freeze_panes = "A2"
    ww.auto_filter.ref = f"A1:Y{len(osr_rows) + 1}"             # filter buttons on the header row, as in the template
    return len(osr_rows), routed


# ── Python-side estimate (shown in the app; mirrors the workbook formulas) ───────────────────────────

def estimate_containers(osr_rows, routes, table, item_cfc, line_cfc, cur_mon, n_weeks):
    """Containers per machine line and week as the workbook will calculate them."""
    group_of = {k.upper(): v for k, v in table.items()}
    per_week = collections.defaultdict(float)                    # (group, week index) -> containers
    after = 0.0
    for (doc, buyer, contact, reqdate, plandate, pid, pname, oq, sip, rs, pend) in osr_rows:
        if not isinstance(pend, (int, float)) or pend <= 0:
            continue
        route = routes.get(_norm(pname))
        group = group_of.get(str(route).upper()) if route else None
        if not group:
            continue
        per = item_cfc.get(str(pname).upper()) or line_cfc.get(group)
        if not per:
            continue
        rd = _to_datetime(reqdate)
        mon = max(monday_of(rd.date()), cur_mon) if rd else cur_mon
        wi = (mon - cur_mon).days // 7
        if wi < n_weeks:
            per_week[(group, wi)] += pend / per
        else:
            after += pend / per
    return per_week, after


# ── Main entry ────────────────────────────────────────────────────────────────────────────────────────

def generate_report(osr_bytes, template_bytes, as_of=None, prev_report_bytes=None):
    """
    osr_bytes         Order Status Report (.xlsx)
    template_bytes    reference_workbook.xlsx (the finished report used as the template)
    as_of             date of the OSR; the first week column is the week holding this date
    prev_report_bytes optional: last week's report (its MC ROUTING and Machines/Shifts are carried forward)

    Returns (xlsx bytes, info dict)
    """
    as_of = as_of or datetime.date.today()
    cur_mon = monday_of(as_of)

    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    for need in (OVW, WORKING, MC_MASTER, ITEM_CFC):
        if need not in wb.sheetnames:
            raise ValueError(f"The template workbook has no sheet '{need}'.")
    ov = wb[OVW]
    table = routing_table(ov)

    # --- routing and inputs: template first, then the previous report on top
    bundled_all = read_routes(wb)                         # every product the template knows (routed or not)
    routes = {k: v for k, v in bundled_all.items() if v}
    known = set(bundled_all)
    source = "bundled with the app"
    inputs_changed = []
    if prev_report_bytes:
        pw = openpyxl.load_workbook(io.BytesIO(prev_report_bytes))
        prev_all = read_routes(pw)
        for k, v in prev_all.items():
            if v:
                routes[k] = v
        known |= set(prev_all)
        source = "previous report"
        if OVW in pw.sheetnames:
            prev_inputs = read_line_inputs(pw[OVW])
            mm = wb[MC_MASTER]
            for r in range(3, 120):
                name = ov.cell(r, 3).value
                if name in NON_LINE_LABELS or name not in prev_inputs:
                    continue
                old = (ov.cell(r, 4).value, ov.cell(r, 5).value)
                new = prev_inputs[name]
                if old != new:
                    inputs_changed.append((name, old, new))
                    ov.cell(r, 4).value, ov.cell(r, 5).value = new
                    for rr in range(3, mm.max_row + 1):
                        if mm.cell(rr, 1).value == name:
                            mm.cell(rr, 5).value, mm.cell(rr, 6).value = new
                            break

    # --- weeks
    n_weeks = 0
    while str(ov.cell(2, FIRST_WEEK_COL + n_weeks).value or "").startswith("W #"):
        n_weeks += 1
    if n_weeks == 0:
        raise ValueError("Could not find the week columns (W #…) in the template.")
    first_wn = weeknum(cur_mon)
    for i in range(n_weeks):
        ov.cell(2, FIRST_WEEK_COL + i).value = f"W #{first_wn + i}"
        c = ov.cell(HELPER_ROW, FIRST_WEEK_COL + i)
        c.value = (f"=DATE({cur_mon.year},{cur_mon.month},{cur_mon.day})" if i == 0
                   else f"={L(FIRST_WEEK_COL + i - 1)}{HELPER_ROW}+7")
    ov.row_dimensions[HELPER_ROW].hidden = True

    # --- WORKING from the new OSR
    osr_rows = extract_osr(openpyxl.load_workbook(io.BytesIO(osr_bytes), data_only=True))
    n_rows, n_routed = build_working(wb, osr_rows, routes, table)

    # --- estimate for the app (same arithmetic as the formulas)
    wv = openpyxl.load_workbook(io.BytesIO(template_bytes), data_only=True)
    item_cfc = {}
    icc = wv[ITEM_CFC]
    for r in range(2, icc.max_row + 1):
        k, v = icc.cell(r, 1).value, icc.cell(r, 2).value
        if k and isinstance(v, (int, float)) and str(k).upper() not in item_cfc:
            item_cfc[str(k).upper()] = v                          # first match wins, like VLOOKUP
    mcv = wv[MC_MASTER]
    line_cfc = {mcv.cell(r, 1).value: mcv.cell(r, 16).value for r in range(3, mcv.max_row + 1)
                if mcv.cell(r, 1).value and isinstance(mcv.cell(r, 16).value, (int, float))}
    per_week, after = estimate_containers(osr_rows, routes, table, item_cfc, line_cfc, cur_mon, n_weeks)

    group_of = {k.upper(): v for k, v in table.items()}
    unrouted, unknown_route, new_products = collections.defaultdict(float), collections.defaultdict(float), set()
    for (doc, buyer, contact, reqdate, plandate, pid, pname, oq, sip, rs, pend) in osr_rows:
        if not isinstance(pend, (int, float)) or pend <= 0:
            continue
        key = _norm(pname)
        route = routes.get(key)
        if not route:
            unrouted[str(pname).strip()] += pend
            if key not in known:
                new_products.add(str(pname).strip())
        elif str(route).upper() not in group_of:
            unknown_route[str(route)] += pend

    # --- only the report tab is selected (otherwise Excel opens with grouped sheets)
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = (ws.title == OVW)
    wb.active = wb.sheetnames.index(OVW)

    buf = io.BytesIO()
    wb.save(buf)
    info = {
        "as_of": as_of,
        "first_week": first_wn,
        "last_week": first_wn + n_weeks - 1,
        "osr_rows": n_rows,
        "routed_rows": n_routed,
        "routing_source": source,
        "containers_in_report": round(sum(per_week.values()), 2),
        "containers_after_last_week": round(after, 2),
        "unrouted": sorted(unrouted.items(), key=lambda kv: -kv[1]),
        "unknown_routes": sorted(unknown_route.items(), key=lambda kv: -kv[1]),
        "new_products": sorted(new_products),
        "inputs_changed": inputs_changed,
    }
    return buf.getvalue(), info
