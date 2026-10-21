"""
MC Capacity vs Order — report generator
========================================
Order Status Report (OSR)  ->  containers per machine line per month/week, against capacity.

The report is produced from a *template*: the bundled reference_workbook.xlsx carries the masters,
capacity lookups and the hidden ORDER VS WEEK DISTRIBUTION sheet, which is the internal "engine" —
it is not shown to the user, but every other sheet pulls its Zone/Region/Sub/Machine Line/Machines/
Shifts/Capacity rows and its two 12-week basis blocks (Based on Pending Production, Based on Order
Qty) from it via live formulas. Each run refreshes:

  * ORDER VS WEEK DISTRIBUTION's 12 week columns (hidden; W # headers + helper row 500)
  * the WORKING sheet, rebuilt from the new OSR
  * MONTH WISE - PENDING PROD and MONTH WISE - ORDER QTY, rebuilt fresh each run: a rolling 4-month
    window starting at the current month, weeks grouped (collapsible) under each month, a part-elapsed
    current month only showing its remaining weeks
  * MC ROUTING in WORKING, carried forward by PRODUCT NAME (from the previous report you upload, else
    the routing bundled in the template)
  * Machines / Shifts, carried forward from the previous report when one is uploaded

AFRICA / AUSTRALIA & EUROPE / USA show only the Order Qty basis, via live formulas into the hidden
ORDER VS WEEK DISTRIBUTION sheet. To change the look, the masters or the capacity basis, replace
reference_workbook.xlsx.
"""

import collections
import copy
import datetime
import io
import re

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

OVW = "ORDER VS WEEK DISTRIBUTION"
WORKING = "WORKING"
MC_MASTER = "MC MASTER"
ITEM_CFC = "ITEM CFC PER CONTAINER"
OSR_SHEET = "Order Status-By shipment Date"
FIRST_WEEK_COL = 8                      # H (A=Zone, B=Region, C=Sub, D=Machine Line, E=Machines, F=Shifts, G=Capacity)
HELPER_ROW = 500                        # hidden row with the Monday of each week
NON_LINE_LABELS = (None, "TOTAL", "Excess Order", "Short Order")

MONTH_SHEETS = [                        # (sheet name, WORKING column holding the basis, in containers)
    ("ORDER QTY", "Z"),
    ("PENDING PROD", "Y"),
]
N_MONTH_WINDOW = 4                      # rolling months shown, starting at the current month
REGION_SHEETS = [                       # (sheet name, Region values from OVW column B to include)
    ("AFR", {"AFRICA"}),
    ("AUS&EU", {"AUSTRALIA", "EUROPE"}),
    ("USA", {"USA"}),
]
# sheets where PERFECTA ENV / C250 AF / MD20 ENV's Week & Month Capacity are shown as one summed figure
MERGE_CAPACITY_SHEETS = {"PENDING PROD", "AFR", "ORDER QTY"}
MERGE_CAPACITY_GROUP = ("PERFECTA ENV", "C250 AF", "MD20 ENV")

# Spellings that are accepted in MC ROUTING but are not offered in the dropdown
ROUTE_ALIASES = {"CONSTANTA TAG D", "MD 20 4GM", "PEARL PACK PREMIX"}

WORKING_HEADERS = [
    "BUYER NAME", "NAV DOC NO", "Buyer Contact No", "Buyer Requested Shipment Date", "Month-yy",
    "Plan To Ship", "ITEM ID", "Column1", "PRODUCT NAME", "MC ROUTING", "CONCAT", "BLEND NAME", "SFG NAME",
    "NO OF CTN PER CFC", "NO OF TBGS PER CTN", "ORDER QTY", "STOCK IN PROCESS", "READY STOCK",
    "STOCK IN PROCESS + READY STOCK", "PENDING TO PRODUCE IN CASES", "PENDING TO PRODUCE IN TBGS",
    "WEEK START", "REPORTING WEEK", "CAPACITY GROUP", "PENDING PROD IN CONTAINERS",
    "PENDING ORDER QTY IN CONTAINERS"]
WORKING_WIDTHS = {1: 24, 2: 16, 3: 22, 4: 13, 5: 9, 6: 13, 7: 9, 8: 22, 9: 42, 10: 22, 11: 22, 12: 24, 13: 30,
                  14: 10, 15: 10, 16: 10, 17: 10, 18: 10, 19: 12, 20: 12, 21: 13, 22: 12, 23: 13, 24: 16, 25: 14,
                  26: 16}


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


def week_of_year(d):
    """Display week number. For the app's anchor year this matches weeknum() (the existing
    continuous numbering); for any later year it resets to 1 at the Monday on/after 1 Jan of
    d's own year, so a block that crosses into a new calendar year starts labelling from W #1."""
    if d.year == _J1.year:
        return weeknum(d)
    j1 = datetime.date(d.year, 1, 1)
    w1 = j1 if j1.weekday() == 0 else j1 + datetime.timedelta(days=7 - j1.weekday())
    mon = monday_of(d)
    return ((mon - w1).days // 7) + 1


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
        name = ov.cell(r, 4).value
        if name in NON_LINE_LABELS:
            continue
        m, s = ov.cell(r, 5).value, ov.cell(r, 6).value
        if isinstance(m, (int, float)) and isinstance(s, (int, float)):
            out[name] = (m, s)
    return out


def routing_table(ov):
    """{ROUTE UPPER: capacity group} from the hidden lookup in columns AH:AI of ORDER VS WEEK DISTRIBUTION."""
    out = {}
    for r in range(2, 120):
        k, v = ov.cell(r, 34).value, ov.cell(r, 35).value
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

    wk0 = f"'{OVW}'!$H${HELPER_ROW}"
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
        base_group = f"IFERROR(VLOOKUP($J{r},'{OVW}'!$AH:$AI,2,0),\"\")"
        ww.cell(r, 24, (
            f'=IF(AND({base_group}="CONSTANTA ENV",'
            f'OR(ISNUMBER(SEARCH("KHAIMA",$I{r})),ISNUMBER(SEARCH("CRANBERRY",$I{r})))),'
            f'"CONSTANTA ENV (HERBAL)",{base_group})'))
        ic = f"VLOOKUP($I{r},'{ITEM_CFC}'!$A:$B,2,0)"
        mf = f"VLOOKUP($X{r},'{MC_MASTER}'!$A:$Q,16,0)"
        ww.cell(r, 25, f'=IFERROR($T{r}/IF(IFERROR({ic},0)>0,{ic},{mf}),"")')
        ww.cell(r, 26, f'=IFERROR($P{r}/IF(IFERROR({ic},0)>0,{ic},{mf}),"")')
        for c in range(1, 27):
            ww.cell(r, c).border = border
            ww.cell(r, c).font = norm_font
    for col, w in WORKING_WIDTHS.items():
        ww.column_dimensions[L(col)].width = w
    ww.row_dimensions[1].height = 35.05
    ww.freeze_panes = "A2"
    ww.auto_filter.ref = f"A1:Z{len(osr_rows) + 1}"             # filter buttons on the header row, as in the template
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
        if group == "CONSTANTA ENV" and ("KHAIMA" in str(pname).upper() or "CRANBERRY" in str(pname).upper()):
            group = "CONSTANTA ENV (HERBAL)"
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


# ── MONTH WISE sheets: rolling N-month window, grouped by week ───────────────────────────────────────

def month_window(cur_mon, n_months=N_MONTH_WINDOW):
    """[(label, [(week_num, monday_date), ...]), ...] for n_months starting at the current month.
    A week belongs to the month holding its Monday; a part-elapsed current month only lists the
    weeks from today's week onward (e.g. "October" on 14 Oct 2026 lists weeks 42, 43, 44 only)."""
    months = []
    y, m = cur_mon.year, cur_mon.month
    for i in range(n_months):
        mm = (m - 1 + i) % 12 + 1
        yy = y + (m - 1 + i) // 12
        start = datetime.date(yy, mm, 1)
        end = datetime.date(yy + 1, 1, 1) if mm == 12 else datetime.date(yy, mm + 1, 1)
        seen, weeks = set(), []
        d = start
        while d < end:
            mon = monday_of(d)
            if mon not in seen and mon >= cur_mon and mon.month == mm and mon.year == yy:
                seen.add(mon)
                weeks.append((weeknum(mon), mon))
            d += datetime.timedelta(days=1)
        weeks.sort(key=lambda t: t[1])
        months.append((start.strftime("%B %Y").upper(), weeks))
    return months


def _style_of(cell, size=None):
    font = copy.copy(cell.font)
    if size is not None:
        font = Font(name=font.name, bold=font.bold, italic=font.italic, color=font.color, size=size)
    return dict(font=font, fill=copy.copy(cell.fill), border=copy.copy(cell.border),
                alignment=copy.copy(cell.alignment), number_format=cell.number_format)


def _apply(cell, st):
    cell.font, cell.fill, cell.border = st["font"], st["fill"], st["border"]
    cell.alignment, cell.number_format = st["alignment"], st["number_format"]


def build_month_sheet(wb, ov, sheet_name, basis_col, cur_mon, as_of, insert_after=OVW, n_months=N_MONTH_WINDOW,
                       region_filter=None, merge_capacity_group=None):
    """(Re)build a month + week wise sheet, grouped by month, summed from WORKING's basis_col
    (containers), mirroring every line from the hidden ORDER VS WEEK DISTRIBUTION sheet.
    region_filter, if given, is a set of Region values (OVW column B) to include, keeping each
    matched block's internal row order but dropping everything else (used for the AFR/AUS&EU/USA
    sheets). merge_capacity_group, if given, is a tuple of Machine Line names that - when they
    land on consecutive rows - get their Week/Month Capacity cells merged into a single summed
    figure instead of being shown one per line."""
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    insert_at = wb.sheetnames.index(insert_after) + 1
    ms = wb.create_sheet(sheet_name, insert_at)

    last_row = 2
    for r in range(3, 300):
        if any(ov.cell(r, c).value not in (None, "") for c in (1, 2, 3, 4)):
            last_row = r
        else:
            break

    def owning_row_map(col_idx):
        m = {}
        for rng in ov.merged_cells.ranges:
            if rng.min_col == col_idx and rng.max_col == col_idx and rng.min_row >= 3:
                for r in range(rng.min_row, rng.max_row + 1):
                    m[r] = rng.min_row
        return m

    zone_owner, region_owner, sub_owner = owning_row_map(1), owning_row_map(2), owning_row_map(3)

    if region_filter is None:
        src_rows = list(range(3, last_row + 1))
    else:
        src_rows = [r for r in range(3, last_row + 1)
                    if ov.cell(region_owner.get(r, r), 2).value in region_filter]
    dest_of = {src_r: i + 3 for i, src_r in enumerate(src_rows)}
    last_dest_row = 2 + len(src_rows)

    st_title = _style_of(ov["A1"])
    st_band = _style_of(ov["H1"])
    st_colhdr = _style_of(ov["A2"], size=10)
    st_colhdr_wk = _style_of(ov["H2"], size=10)
    st_zone, st_region, st_sub = (_style_of(ov["A3"], size=10), _style_of(ov["B3"], size=10),
                                   _style_of(ov["C3"], size=10))
    st_line, st_mach, st_shift, st_cap = (_style_of(ov["D3"], size=10), _style_of(ov["E3"], size=10),
                                           _style_of(ov["F3"], size=10), _style_of(ov["G3"], size=10))
    st_week = _style_of(ov["H3"], size=10)
    st_total_lbl, st_total_oth = _style_of(ov["D8"], size=10), _style_of(ov["A8"], size=10)
    st_total_cap, st_total_wk = _style_of(ov["G8"], size=10), _style_of(ov["H8"], size=10)

    months = month_window(cur_mon, n_months)

    COL_CAP_MONTHLY = 8
    col = COL_CAP_MONTHLY + 1
    blocks = []                                               # (label, first_col, last_col, weeks, total_col)
    for label, weeks in months:
        first = col
        last = first + max(len(weeks), 1) - 1
        total_col = last + 1
        blocks.append((label, first, last, weeks, total_col))
        col = total_col + 2                                   # one spacer column before the next month

    basis_label = {"Z": "ORDER QTY", "Y": "PENDING PROD"}.get(basis_col, basis_col)
    ms.merge_cells(start_row=1, start_column=1, end_row=1, end_column=COL_CAP_MONTHLY)
    _apply(ms.cell(1, 1, f"Month & Week Wise Distribution - Generated on {as_of:%d/%m/%Y} - {basis_label}"), st_title)
    for c in range(2, COL_CAP_MONTHLY + 1):
        _apply(ms.cell(1, c), st_title)
    for label, first, last, weeks, total_col in blocks:
        ms.merge_cells(start_row=1, start_column=first, end_row=1, end_column=total_col)
        abbrev = weeks[0][1].strftime("%b").upper() if weeks else label[:3]
        _apply(ms.cell(1, first, abbrev), st_band)
        for c in range(first, total_col + 1):
            _apply(ms.cell(1, c), st_band)

    headers_ag = ["Zone", "Region", None, "Machine Line", "No. of Machines", "No. of Shifts", "Week Capacity"]
    for c, h in enumerate(headers_ag, 1):
        _apply(ms.cell(2, c, h), st_colhdr)
    _apply(ms.cell(2, COL_CAP_MONTHLY, "Month Capacity"), st_colhdr_wk)
    for label, first, last, weeks, total_col in blocks:
        for i, (wn, wmon) in enumerate(weeks):
            _apply(ms.cell(2, first + i, f"W #{week_of_year(wmon)}"), st_colhdr_wk)
        abbrev = weeks[0][1].strftime("%b").upper() if weeks else label[:3]
        _apply(ms.cell(2, total_col, abbrev), st_colhdr_wk)

    total_re = re.compile(r"=SUM\(([A-Z]+)(\d+):\1(\d+)\)")
    line_dest_row = {}

    for src_r in src_rows:
        r = dest_of[src_r]
        label = ov.cell(src_r, 4).value
        is_total = label == "TOTAL"
        if label not in NON_LINE_LABELS:
            line_dest_row[label] = r
        zo, ro, so = zone_owner.get(src_r, src_r), region_owner.get(src_r, src_r), sub_owner.get(src_r, src_r)
        _apply(ms.cell(r, 1, f"=IF('{OVW}'!A{zo}=\"\",\"\",'{OVW}'!A{zo})"), st_total_oth if is_total else st_zone)
        _apply(ms.cell(r, 2, f"=IF('{OVW}'!B{ro}=\"\",\"\",'{OVW}'!B{ro})"), st_total_oth if is_total else st_region)
        _apply(ms.cell(r, 3, f"=IF('{OVW}'!C{so}=\"\",\"\",'{OVW}'!C{so})"), st_total_oth if is_total else st_sub)
        _apply(ms.cell(r, 4, f"=IF('{OVW}'!D{src_r}=\"\",\"\",'{OVW}'!D{src_r})"), st_total_lbl if is_total else st_line)
        _apply(ms.cell(r, 5, f"=IF('{OVW}'!E{src_r}=\"\",\"\",'{OVW}'!E{src_r})"), st_total_oth if is_total else st_mach)
        _apply(ms.cell(r, 6, f"=IF('{OVW}'!F{src_r}=\"\",\"\",'{OVW}'!F{src_r})"), st_total_oth if is_total else st_shift)
        _apply(ms.cell(r, 7, f"=IF('{OVW}'!G{src_r}=\"\",\"\",'{OVW}'!G{src_r})"), st_total_cap if is_total else st_cap)
        _apply(ms.cell(r, COL_CAP_MONTHLY, f'=IF(ISNUMBER($G{r}),$G{r}*4,$G{r})'),
               st_total_cap if is_total else st_cap)

        for label_, first, last, weeks, total_col in blocks:
            for i, (wn, wmon) in enumerate(weeks):
                c = first + i
                if label == "TOTAL":
                    m = total_re.match(str(ov.cell(src_r, FIRST_WEEK_COL).value))
                    dr1, dr2 = dest_of[int(m.group(2))], dest_of[int(m.group(3))]
                    formula, st = f"=SUM({L(c)}{dr1}:{L(c)}{dr2})", st_total_wk
                elif label == "Excess Order":
                    tr = r - 1
                    formula = f'=IFERROR(IF(ROUND({L(c)}{tr}-$G{tr},3)>0,ROUND({L(c)}{tr}-$G{tr},3),""),"")'
                    st = _style_of(ov.cell(src_r, FIRST_WEEK_COL), size=10)
                elif label == "Short Order":
                    tr = r - 2
                    formula = f'=IFERROR(IF(ROUND($G{tr}-{L(c)}{tr},3)>0,ROUND($G{tr}-{L(c)}{tr},3),""),"")'
                    st = _style_of(ov.cell(src_r, FIRST_WEEK_COL), size=10)
                else:
                    formula = (f'=IFERROR(SUMIFS(WORKING!${basis_col}:${basis_col},WORKING!$X:$X,$D{r},'
                               f'WORKING!$W:$W,DATE({wmon.year},{wmon.month},{wmon.day})),"")')
                    st = st_week
                _apply(ms.cell(r, c, formula), st)
            tot_st = st_total_wk if is_total else st_week
            if weeks:
                _apply(ms.cell(r, total_col, f"=SUM({L(first)}{r}:{L(last)}{r})"), tot_st)
            else:
                _apply(ms.cell(r, total_col, 0), tot_st)

    if merge_capacity_group and all(name in line_dest_row for name in merge_capacity_group):
        rows = sorted(line_dest_row[name] for name in merge_capacity_group)
        if rows == list(range(rows[0], rows[0] + len(rows))):          # only merge if they're consecutive
            r1, r2 = rows[0], rows[-1]
            terms = "+".join(f"IFERROR(ROUND(VLOOKUP($D{rr},'{MC_MASTER}'!$A:$R,18,0)*$E{rr}*$F{rr},3),0)"
                              for rr in rows)
            _apply(ms.cell(r1, 7, f"={terms}"), st_cap)
            for rr in rows[1:]:
                ms.cell(rr, 7).value = None
                ms.cell(rr, COL_CAP_MONTHLY).value = None
            ms.merge_cells(start_row=r1, start_column=7, end_row=r2, end_column=7)
            ms.merge_cells(start_row=r1, start_column=COL_CAP_MONTHLY, end_row=r2, end_column=COL_CAP_MONTHLY)

            for label_, first, last, weeks, total_col in blocks:
                for i, (wn, wmon) in enumerate(weeks):
                    c = first + i
                    terms_wk = "+".join(
                        f'IFERROR(SUMIFS(WORKING!${basis_col}:${basis_col},WORKING!$X:$X,$D{rr},'
                        f'WORKING!$W:$W,DATE({wmon.year},{wmon.month},{wmon.day})),0)'
                        for rr in rows)
                    _apply(ms.cell(r1, c, f"={terms_wk}"), st_week)
                    for rr in rows[1:]:
                        ms.cell(rr, c).value = None
                    ms.merge_cells(start_row=r1, start_column=c, end_row=r2, end_column=c)
                if weeks:
                    _apply(ms.cell(r1, total_col, f"=SUM({L(first)}{r1}:{L(last)}{r1})"), st_week)
                for rr in rows[1:]:
                    ms.cell(rr, total_col).value = None
                ms.merge_cells(start_row=r1, start_column=total_col, end_row=r2, end_column=total_col)

    col_widths = {1: 8, 2: 9.8, 3: 9, 4: 13.8, 5: 6.4, 6: 7, 7: 11.4, COL_CAP_MONTHLY: 9}
    for c, w in col_widths.items():
        ms.column_dimensions[L(c)].width = w
    for bi, (label, first, last, weeks, total_col) in enumerate(blocks):
        for c in range(first, last + 1):
            ms.column_dimensions[L(c)].width = 5.7
            ms.column_dimensions[L(c)].outlineLevel = 1
        ms.column_dimensions[L(total_col)].width = 5.5
        if bi < len(blocks) - 1:
            ms.column_dimensions[L(total_col + 1)].width = 3.5     # spacer before the next month

    ms.sheet_properties.outlinePr.summaryRight = True
    red_font = Font(name="Trebuchet MS", color="FFCC0000", size=11)
    for label, first, last, weeks, total_col in blocks:
        if not weeks:
            continue
        rng = f"{L(first)}3:{L(last)}{last_dest_row}"
        formula = f"AND(ISNUMBER({L(first)}3),ISNUMBER($G3),$G3>0,{L(first)}3>$G3)"
        ms.conditional_formatting.add(rng, FormulaRule(formula=[formula], font=red_font))

    # merge Zone/Region/Sub cells over runs of consecutive dest rows sharing the same source merge
    # (keyed by the source owning row, not the text, so two different blocks that happen to carry
    # the same label - e.g. "DC TAG" in two regions - are never merged together)
    for col, owner in ((1, zone_owner), (2, region_owner), (3, sub_owner)):
        run_start, run_key = None, object()
        dest_of_prev = None
        for src_r in src_rows:
            r = dest_of[src_r]
            key = owner.get(src_r, src_r)
            if key != run_key:
                if run_start is not None and dest_of_prev > run_start:
                    ms.merge_cells(start_row=run_start, start_column=col, end_row=dest_of_prev, end_column=col)
                run_start, run_key = r, key
            dest_of_prev = r
        if run_start is not None and dest_of_prev > run_start:
            ms.merge_cells(start_row=run_start, start_column=col, end_row=dest_of_prev, end_column=col)

    ms.freeze_panes = "A5"
    ms.sheet_view.showGridLines = False
    ms.row_dimensions[1].height = 15
    ms.row_dimensions[2].height = 41.4
    return ms


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
                name = ov.cell(r, 4).value
                if name in NON_LINE_LABELS or name not in prev_inputs:
                    continue
                old = (ov.cell(r, 5).value, ov.cell(r, 6).value)
                new = prev_inputs[name]
                if old != new:
                    inputs_changed.append((name, old, new))
                    ov.cell(r, 5).value, ov.cell(r, 6).value = new
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

    # --- MONTH WISE sheets: rolling N-month window, rebuilt fresh every run
    prev_sheet = OVW
    for sheet_name, basis_col in MONTH_SHEETS:
        mcg = MERGE_CAPACITY_GROUP if sheet_name in MERGE_CAPACITY_SHEETS else None
        build_month_sheet(wb, ov, sheet_name, basis_col, cur_mon, as_of, insert_after=prev_sheet,
                           merge_capacity_group=mcg)
        prev_sheet = sheet_name

    # --- region sheets: same month + week wise, grouped format, filtered to each region (Order Qty basis)
    for sheet_name, region_filter in REGION_SHEETS:
        mcg = MERGE_CAPACITY_GROUP if sheet_name in MERGE_CAPACITY_SHEETS else None
        build_month_sheet(wb, ov, sheet_name, "Z", cur_mon, as_of, insert_after=prev_sheet,
                           region_filter=region_filter, merge_capacity_group=mcg)
        prev_sheet = sheet_name

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

    # --- only the landing tab is selected (otherwise Excel opens with grouped sheets); OVW stays hidden
    landing = MONTH_SHEETS[0][0]
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = (ws.title == landing)
    wb.active = wb.sheetnames.index(landing)

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
