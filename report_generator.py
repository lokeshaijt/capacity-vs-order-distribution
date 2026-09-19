"""
MC Capacity vs Order — Week-Wise Report Generator
All logic self-contained; reads reference data from the bundled workbook.
"""

import datetime
import io
import re
from collections import defaultdict

import openpyxl
from openpyxl.styles import (
    Alignment, Border, Color, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

# ── Style constants ────────────────────────────────────────────────────────────
FONT_NAME = "Trebuchet MS"

def _font(bold=False, size=11, color="FF000000"):
    return Font(name=FONT_NAME, bold=bold, size=size, color=color)

def _fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

def _theme_fill(theme, tint):
    return PatternFill(fgColor=Color(theme=theme, tint=tint), fill_type="solid")

def _border():
    thin = Side(style="thin", color="FF000000")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

TITLE_FILL   = _theme_fill(5, 0.4)
OLIVE_FILL   = _theme_fill(5, -0.5)
HDR_FILL     = _fill("FFB6C6DE")
TOTAL_FILL   = TITLE_FILL
INPUT_FILL   = _fill("FFFFF2CC")
REGION_FILL  = _fill("FFD9D9D9")
PLAIN_FILL   = PatternFill()

H_RED        = _font(bold=True, size=12, color="FF9D360E")
H_WHITE      = _font(bold=True, size=12, color="FFFFFFFF")
HDR_FONT     = _font(bold=True, size=11)
NORM         = _font(size=11)
TOTAL_FONT   = _font(bold=True, size=11)
RED_FONT     = _font(size=11, color="FFCC0000")
NUMFMT       = '_ * #,##0.0_ ;_ * \\-#,##0.0_ ;_ * "-"??_ ;_ @_ '
BORDER       = _border()


# ── Reference-data loaders ─────────────────────────────────────────────────────

def load_reference(ref_wb):
    """Return all lookup tables from the reference workbook."""
    ref = {}

    # ── TBGS PER CTN  (SFG name → tbgs/ctn) ──────────────────────────────
    ws = ref_wb["TBGS PER CTN"]
    ref["tbgs_per_ctn"] = {
        ws.cell(r, 2).value: ws.cell(r, 3).value
        for r in range(2, ws.max_row + 1)
        if ws.cell(r, 2).value and ws.cell(r, 3).value
    }

    # ── ITEM CFC PER CONTAINER  (product name → cfc/container) ──────────
    ws = ref_wb["ITEM CFC PER CONTAINER"]
    ref["item_cfc"] = {
        ws.cell(r, 1).value: ws.cell(r, 2).value
        for r in range(2, ws.max_row + 1)
        if ws.cell(r, 1).value and ws.cell(r, 2).value
    }

    # ── ITEM MASTER  (prod_id → sfg_name, prod_name → ctn_per_cfc) ───────
    ws = ref_wb["ITEM MASTER"]
    prod_id_to_sfg = {}
    prod_name_to_ctn_cfc = {}
    sfg_to_tbgspctn = {}          # SFG item-id → tbgs (col G = tbgs)
    for r in range(2, ws.max_row + 1):
        pid  = ws.cell(r, 1).value   # FG item id
        name = ws.cell(r, 2).value   # FG name
        sid  = ws.cell(r, 4).value   # SFG item id
        sfg  = ws.cell(r, 5).value   # SFG name
        ctn  = ws.cell(r, 6).value   # CTN per CFC
        if pid and sfg:
            prod_id_to_sfg[pid] = sfg
        if name and ctn:
            prod_name_to_ctn_cfc[name] = ctn
    ref["prod_id_to_sfg"]       = prod_id_to_sfg
    ref["prod_name_to_ctn_cfc"] = prod_name_to_ctn_cfc

    # ── MC MASTER  (line → factors) ──────────────────────────────────────
    ws = ref_wb["MC MASTER"]
    mc_master = {}
    for r in range(3, ws.max_row + 1):
        line = ws.cell(r, 1).value
        if not line:
            continue
        mc_master[line] = {
            "no_machines":    ws.cell(r, 5).value,
            "no_shifts":      ws.cell(r, 6).value,
            "tbgs_per_cfc":   ws.cell(r, 14).value,
            "cfc_per_cntr":   ws.cell(r, 16).value,
            "weekly_cap_cntr": ws.cell(r, 17).value,
        }
    ref["mc_master"] = mc_master

    # ── ORDER VS WEEK DISTRIBUTION layout (region/sub/line table) ─────────
    ws = ref_wb["ORDER VS WEEK DISTRIBUTION"]
    layout = []   # list of dicts
    last_region = last_sub = None
    for r in range(3, 41):
        region = ws.cell(r, 1).value
        sub    = ws.cell(r, 2).value
        line   = ws.cell(r, 3).value
        nm     = ws.cell(r, 4).value
        ns     = ws.cell(r, 5).value
        if region:
            last_region = region
        if sub:
            last_sub = sub
        if line == "TOTAL" or line is None:
            layout.append({"type": "total" if line == "TOTAL" else "blank",
                            "region": last_region})
        else:
            layout.append({"type": "line", "region": last_region,
                            "sub": last_sub, "line": line,
                            "no_machines": nm, "no_shifts": ns})
    ref["layout"] = layout

    # ── Routing table: canonical routing name → capacity-group name ────────
    # Stored in hidden cols T,U of ORDER VS WEEK DISTRIBUTION
    ws = ref_wb["ORDER VS WEEK DISTRIBUTION"]
    routing_map = {}
    for r in range(2, 50):
        k = ws.cell(r, 20).value
        v = ws.cell(r, 21).value
        if k and v:
            routing_map[k] = v
    ref["routing_map"] = routing_map

    # ── WORKING sheet: product→MC ROUTING mapping (last-known) ────────────
    ws = ref_wb["WORKING"]
    prod_routing = {}
    for r in range(2, ws.max_row + 1):
        pname = ws.cell(r, 9).value
        route = ws.cell(r, 10).value
        if pname and route:
            prod_routing[pname] = route
    ref["prod_routing"] = prod_routing

    return ref


# ── OSR extraction ─────────────────────────────────────────────────────────────

def extract_osr(osr_wb):
    """Return list of order-row tuples from the Order Status Report."""
    ws = osr_wb["Order Status-By shipment Date"]
    rows = []
    last = {"buyer": None, "doc": None, "contact": None, "reqdate": None, "plan": None}
    for r in range(7, ws.max_row + 1):
        prod_id = ws.cell(r, 9).value
        prod_name = ws.cell(r, 10).value
        for key, col in [("buyer",3),("doc",4),("contact",5),("reqdate",6),("plan",7)]:
            v = ws.cell(r, col).value
            if v not in (None, " ", ""):
                last[key] = v
        if isinstance(prod_id, (int, float)) and prod_name not in (None, " ", ""):
            rows.append((
                last["doc"], last["buyer"], last["contact"],
                last["reqdate"], last["plan"], prod_id, prod_name,
                ws.cell(r, 11).value,  # order qty
                ws.cell(r, 12).value,  # stock wip
                ws.cell(r, 13).value,  # ready stock
                ws.cell(r, 15).value,  # pending prod (CFC)
            ))
    return rows


# ── Production Register extraction ────────────────────────────────────────────

def extract_prod_register(pr_wb, routing_map):
    """Return list of (date, group, item_name, qty, uom) for valid packing lines."""
    ws = pr_wb[pr_wb.sheetnames[0]]
    rows = []
    for r in range(8, ws.max_row + 1):
        pdate = ws.cell(r, 1).value
        wc    = ws.cell(r, 3).value
        item_name = ws.cell(r, 6).value
        qty   = ws.cell(r, 8).value
        uom   = ws.cell(r, 10).value
        if not pdate or not wc:
            continue
        group = routing_map.get(wc)
        if not group:
            continue
        if uom not in ("CTN", "CFC"):
            continue
        d = pdate.date() if isinstance(pdate, datetime.datetime) else pdate
        rows.append((d, group, item_name, qty, uom))
    return rows


# ── Week helpers ───────────────────────────────────────────────────────────────

_JAN1_2026 = datetime.date(2026, 1, 1)
_WEEK1_MON = _JAN1_2026 - datetime.timedelta(days=_JAN1_2026.weekday())


def weeknum(d):
    monday = d - datetime.timedelta(days=d.weekday())
    return ((monday - _WEEK1_MON).days // 7) + 1


def week_monday(wn):
    return _WEEK1_MON + datetime.timedelta(weeks=wn - 1)


# ── Container conversion ───────────────────────────────────────────────────────

def prod_to_containers(prod_rows, ref, target_weeks):
    """Convert production-register rows to container counts per (group, week)."""
    tbgs_per_ctn = ref["tbgs_per_ctn"]
    mc_master    = ref["mc_master"]
    result = defaultdict(float)
    skipped = []

    for d, group, item_name, qty, uom in prod_rows:
        wn = weeknum(d)
        if wn not in target_weeks:
            continue
        factors = mc_master.get(group, {})
        tpc = factors.get("tbgs_per_cfc")
        cpc = factors.get("cfc_per_cntr")
        if not tpc or not cpc:
            skipped.append(("no_factors", group, item_name))
            continue
        if uom == "CTN":
            tbgs_ctn = tbgs_per_ctn.get(item_name)
            if not tbgs_ctn:
                # Derive from product name: e.g. "25 DC ENV" → 25
                m = re.search(r"\b(\d+)\s+DC\b", item_name or "")
                tbgs_ctn = int(m.group(1)) if m else None
            if not tbgs_ctn:
                skipped.append(("no_tbgs", group, item_name))
                continue
            cfc = qty * tbgs_ctn / tpc
        else:
            cfc = qty
        result[(group, wn)] += cfc / cpc

    return dict(result), skipped


# ── Order-to-container conversion ─────────────────────────────────────────────

def orders_to_containers(osr_rows, ref, today):
    """
    Return list of enriched order rows with week bucketing and container counts.
    Each row: (prod_name, group, pending_cfc, containers, reporting_week_monday)
    """
    tbgs_per_ctn  = ref["tbgs_per_ctn"]
    item_cfc      = ref["item_cfc"]
    prod_routing  = ref["prod_routing"]
    routing_map   = ref["routing_map"]
    mc_master     = ref["mc_master"]
    today_mon     = today - datetime.timedelta(days=today.weekday())

    result = []
    for row in osr_rows:
        (doc, buyer, contact, reqdate, plandate, prod_id, prod_name,
         order_qty, stock_wip, ready_stock, pending_cfc) = row
        if not isinstance(pending_cfc, (int, float)) or pending_cfc <= 0:
            continue

        # Routing → capacity group
        route = prod_routing.get(prod_name)
        if not route:
            continue
        group = routing_map.get(route)
        if not group:
            continue

        # CFC → containers
        cfc_per_cntr = item_cfc.get(prod_name)
        if not cfc_per_cntr:
            factors = mc_master.get(group, {})
            cfc_per_cntr = factors.get("cfc_per_cntr")
        if not cfc_per_cntr:
            continue
        containers = pending_cfc / cfc_per_cntr

        # Week bucketing
        if isinstance(reqdate, datetime.datetime):
            req_d = reqdate.date()
        elif isinstance(reqdate, datetime.date):
            req_d = reqdate
        else:
            req_d = None
        if req_d:
            req_mon = req_d - datetime.timedelta(days=req_d.weekday())
            rep_mon = max(req_mon, today_mon)
        else:
            rep_mon = today_mon

        result.append((prod_name, group, pending_cfc, containers, rep_mon))
    return result


# ── Sheet builders ─────────────────────────────────────────────────────────────

def _write_header_rows(ws, title_left, title_right, week_labels, n_fixed_cols=6):
    """Write the two title/header rows common to all distribution sheets."""
    ws.merge_cells(f"A1:{get_column_letter(n_fixed_cols)}1")
    ws["A1"] = title_left
    ws["A1"].font = H_RED; ws["A1"].fill = TITLE_FILL
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

    n_weeks = len(week_labels)
    wfc = n_fixed_cols + 1
    ws.merge_cells(start_row=1, start_column=wfc, end_row=1,
                   end_column=wfc + n_weeks - 1)
    ws.cell(1, wfc, title_right)
    ws.cell(1, wfc).font = H_WHITE; ws.cell(1, wfc).fill = OLIVE_FILL
    ws.cell(1, wfc).alignment = Alignment(horizontal="center", vertical="center")

    hdr_labels = ["Region", None, "Machine Line",
                  "No. of Machines", "No. of Shifts", "Capacity(In containers)"]
    for c, h in enumerate(hdr_labels, 1):
        cell = ws.cell(2, c, h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, lbl in enumerate(week_labels):
        cell = ws.cell(2, wfc + i, lbl)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(wfc + i)].width = 7

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 14


def _apply_merges_and_cf(ws, pending_merges, n_weeks, wfc, first_data_row, last_row):
    for col_letter, ds, de in pending_merges:
        if de > ds:
            ws.merge_cells(f"{col_letter}{ds}:{col_letter}{de}")
        ws[f"{col_letter}{ds}"].alignment = Alignment(
            horizontal="center", vertical="center")

    for i in range(n_weeks):
        col = get_column_letter(wfc + i)
        rng = f"{col}{first_data_row}:{col}{last_row}"
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND(${col}{first_data_row}<>"",'
                     f'$F{first_data_row}<>"",$F{first_data_row}>0,'
                     f'{col}{first_data_row}>$F{first_data_row})'],
            font=RED_FONT))

    ws.freeze_panes = f"{get_column_letter(wfc)}{first_data_row}"
    ws.sheet_view.showGridLines = False


def build_order_sheet(ws, layout, ref, week_labels, week_mondays, order_data,
                      title_right="WEEK WISE ANALYSIS"):
    """Build ORDER VS WEEK DISTRIBUTION or a regional mirror."""
    wfc = 7
    n_weeks = len(week_labels)
    _write_header_rows(ws, "Order Vs Week Distribution", title_right,
                       week_labels)

    # Aggregate orders: (group, week_monday) → total containers
    orders_by_gw = defaultdict(float)
    for prod_name, group, pending_cfc, containers, rep_mon in order_data:
        orders_by_gw[(group, rep_mon)] += containers

    cap_tmpl = (
        "=IFERROR(ROUNDUP((VLOOKUP($C{r},'MC MASTER'!$A:$Q,7,0)"
        "*VLOOKUP($C{r},'MC MASTER'!$A:$Q,10,0)*$D{r}*$E{r}*6)"
        "*(VLOOKUP($C{r},'MC MASTER'!$A:$Q,13,0)"
        "/VLOOKUP($C{r},'MC MASTER'!$A:$Q,12,0))"
        "/VLOOKUP($C{r},'MC MASTER'!$A:$Q,14,0),0)"
        "/VLOOKUP($C{r},'MC MASTER'!$A:$Q,16,0),\"\")"
    )

    FIRST_DATA_ROW = 3
    r = FIRST_DATA_ROW
    pending_merges = []
    cur_a_start = cur_b_start = None
    pending_total_start = r
    cur_sub = None
    _cur_region = None

    mc_master = ref["mc_master"]

    for item in layout:
        if item["type"] == "blank":
            r += 1
            pending_total_start = r
            cur_sub = None
            continue

        if item["type"] == "total":
            if cur_sub is not None:
                merge_end = r - 1
                if cur_b_start is not None and merge_end >= cur_b_start:
                    pending_merges.append(("B", cur_b_start, merge_end))
                cur_b_start = None
                cur_sub = None

            start_r, end_r = pending_total_start, r - 1

            # Close region merge when TOTAL follows REGION lines
            # (region boundary detected: next item has a different region)
            ws.cell(r, 3, "TOTAL")
            ws.cell(r, 6, f"=SUM(F{start_r}:F{end_r})")
            for i, mon in enumerate(week_mondays):
                col = get_column_letter(wfc + i)
                ws[f"{col}{r}"] = f"=SUM({col}{start_r}:{col}{end_r})"
            for c in range(1, 7):
                ws.cell(r, c).font = TOTAL_FONT
                ws.cell(r, c).fill = TOTAL_FILL
                ws.cell(r, c).border = BORDER
            for i in range(n_weeks):
                col = get_column_letter(wfc + i)
                ws[f"{col}{r}"].font = TOTAL_FONT
                ws[f"{col}{r}"].fill = TOTAL_FILL
                ws[f"{col}{r}"].border = BORDER
                ws[f"{col}{r}"].number_format = NUMFMT
            ws.cell(r, 6).number_format = NUMFMT
            r += 1
            pending_total_start = r
            continue

        # data line
        region = item["region"]
        sub    = item.get("sub")
        line   = item["line"]
        nm     = item.get("no_machines") or (mc_master.get(line, {}).get("no_machines") or "")
        ns     = item.get("no_shifts")   or (mc_master.get(line, {}).get("no_shifts")   or "")

        # Detect region change
        if cur_a_start is not None and region != _cur_region:
            pending_merges.append(("A", cur_a_start, r - 1))
            cur_a_start = None

        _cur_region = region

        # Track region block start — write label at the very first row of the region
        if cur_a_start is None:
            cur_a_start = r
            # write region label now; merge applied later
            ws.cell(r, 1, region)
            ws.cell(r, 1).font = Font(name=FONT_NAME, bold=True, size=11)
            ws.cell(r, 1).fill = REGION_FILL
        if sub and sub != cur_sub:
            if cur_sub is not None and cur_b_start is not None:
                pending_merges.append(("B", cur_b_start, r - 1))
            cur_sub = sub
            cur_b_start = r
            ws.cell(r, 2, sub).font = NORM

        ws.cell(r, 3, line).font = NORM
        ws.cell(r, 4, nm); ws.cell(r, 4).fill = INPUT_FILL; ws.cell(r, 4).font = NORM
        ws.cell(r, 5, ns); ws.cell(r, 5).fill = INPUT_FILL; ws.cell(r, 5).font = NORM
        ws.cell(r, 6, cap_tmpl.format(r=r)).font = NORM
        ws.cell(r, 6).number_format = NUMFMT

        for i, mon in enumerate(week_mondays):
            val = orders_by_gw.get((line, mon), 0)
            ws.cell(r, wfc + i, val if val else 0)
            ws.cell(r, wfc + i).font = NORM
            ws.cell(r, wfc + i).number_format = NUMFMT
            ws.cell(r, wfc + i).border = BORDER

        for c in (3, 4, 5, 6):
            ws.cell(r, c).border = BORDER

        r += 1

    LAST_ROW = r - 1

    # close region merge
    if cur_a_start is not None:
        pending_merges.append(("A", cur_a_start, LAST_ROW))
    if cur_b_start is not None:
        pending_merges.append(("B", cur_b_start, LAST_ROW))

    for col_letter, ds, de in pending_merges:
        if de > ds:
            ws.merge_cells(f"{col_letter}{ds}:{col_letter}{de}")
        ws[f"{col_letter}{ds}"].alignment = Alignment(
            horizontal="center", vertical="center")
        if col_letter == "A":
            ws[f"A{ds}"].font = Font(name=FONT_NAME, bold=True, size=11)
            ws[f"A{ds}"].fill = REGION_FILL

    _apply_merges_and_cf(ws, [], n_weeks, wfc, FIRST_DATA_ROW, LAST_ROW)
    return LAST_ROW


def build_achieved_sheet(ws, layout, ref, target_weeks, achieved):
    """Build the Week-Wise Achieved Capacity sheet."""
    week_labels  = [f"W #{wn}" for wn in target_weeks]
    week_mondays = [week_monday(wn) for wn in target_weeks]
    wfc = 7
    n_weeks = len(target_weeks)
    _write_header_rows(ws, "Order Vs Week Distribution",
                       "Week-Wise Achieved Capacity", week_labels)

    FIRST_DATA_ROW = 3
    r = FIRST_DATA_ROW
    pending_merges = []
    cur_a_start = cur_b_start = None
    pending_total_start = r
    cur_sub = None
    _cur_region_a = None

    for item in layout:
        if item["type"] == "blank":
            r += 1; pending_total_start = r; cur_sub = None; continue

        if item["type"] == "total":
            if cur_b_start is not None:
                pending_merges.append(("B", cur_b_start, r - 1))
                cur_b_start = None; cur_sub = None
            start_r, end_r = pending_total_start, r - 1
            ws.cell(r, 3, "TOTAL")
            ws.cell(r, 6, f"=SUM(F{start_r}:F{end_r})")
            for i in range(n_weeks):
                col = get_column_letter(wfc + i)
                ws[f"{col}{r}"] = f"=SUM({col}{start_r}:{col}{end_r})"
            for c in range(1, 7):
                ws.cell(r, c).font = TOTAL_FONT; ws.cell(r, c).fill = TOTAL_FILL
                ws.cell(r, c).border = BORDER
            for i in range(n_weeks):
                col = get_column_letter(wfc + i)
                ws[f"{col}{r}"].font = TOTAL_FONT; ws[f"{col}{r}"].fill = TOTAL_FILL
                ws[f"{col}{r}"].border = BORDER; ws[f"{col}{r}"].number_format = NUMFMT
            ws.cell(r, 6).number_format = NUMFMT
            r += 1; pending_total_start = r; continue

        region = item["region"]; sub = item.get("sub"); line = item["line"]
        nm = item.get("no_machines", ""); ns = item.get("no_shifts", "")

        if cur_a_start is not None and region != _cur_region_a:
            pending_merges.append(("A", cur_a_start, r - 1))
            cur_a_start = None
        _cur_region_a = region

        if cur_a_start is None:
            cur_a_start = r
            ws.cell(r, 1, region)
            ws.cell(r, 1).font = Font(name=FONT_NAME, bold=True, size=11)
            ws.cell(r, 1).fill = REGION_FILL
        if sub and sub != cur_sub:
            if cur_sub is not None and cur_b_start is not None:
                pending_merges.append(("B", cur_b_start, r - 1))
            cur_sub = sub; cur_b_start = r
            ws.cell(r, 2, sub).font = NORM

        ws.cell(r, 3, line).font = NORM
        ws.cell(r, 4, nm).font = NORM; ws.cell(r, 4).border = BORDER
        ws.cell(r, 5, ns).font = NORM; ws.cell(r, 5).border = BORDER
        ws.cell(r, 6, f"=IF('ORDER VS WEEK DISTRIBUTION'!F{r}=\"\",\"\",'ORDER VS WEEK DISTRIBUTION'!F{r})")
        ws.cell(r, 6).font = NORM; ws.cell(r, 6).number_format = NUMFMT
        ws.cell(r, 3).border = BORDER

        for i, wn in enumerate(target_weeks):
            val = achieved.get((line, wn), 0)
            ws.cell(r, wfc + i, val if val else 0)
            ws.cell(r, wfc + i).font = NORM
            ws.cell(r, wfc + i).number_format = NUMFMT
            ws.cell(r, wfc + i).border = BORDER

        r += 1

    LAST_ROW = r - 1

    if cur_a_start is not None:
        pending_merges.append(("A", cur_a_start, LAST_ROW))
    if cur_b_start is not None:
        pending_merges.append(("B", cur_b_start, LAST_ROW))

    for col_letter, ds, de in pending_merges:
        if de > ds:
            ws.merge_cells(f"{col_letter}{ds}:{col_letter}{de}")
        ws[f"{col_letter}{ds}"].alignment = Alignment(
            horizontal="center", vertical="center")
        if col_letter == "A":
            ws[f"A{ds}"].font = Font(name=FONT_NAME, bold=True, size=11)
            ws[f"A{ds}"].fill = REGION_FILL

    _apply_merges_and_cf(ws, [], n_weeks, wfc, FIRST_DATA_ROW, LAST_ROW)
    return LAST_ROW


# ── Hidden lookup table writer ─────────────────────────────────────────────────

def write_lookup_table(ws, routing_map, start_col):
    hf = Font(name=FONT_NAME, size=9, color="FFAAAAAA")
    lc1 = get_column_letter(start_col); lc2 = get_column_letter(start_col + 1)
    ws[f"{lc1}1"] = "Routing Line"; ws[f"{lc1}1"].font = hf
    ws[f"{lc2}1"] = "Capacity Group"; ws[f"{lc2}1"].font = hf
    for i, (k, v) in enumerate(routing_map.items()):
        ws.cell(i + 2, start_col, k).font = hf
        ws.cell(i + 2, start_col + 1, v).font = hf
    ws.column_dimensions[lc1].hidden = True
    ws.column_dimensions[lc2].hidden = True


# ── Main entry point ───────────────────────────────────────────────────────────

def generate_report(ref_bytes, osr_bytes, pr_bytes, today=None,
                    n_weeks_forward=12, n_weeks_achieved=4):
    """
    Parameters
    ----------
    ref_bytes   : bytes — reference workbook
    osr_bytes   : bytes — Order Status Report
    pr_bytes    : bytes — Production Register
    today       : datetime.date (defaults to date.today())
    n_weeks_forward : int — how many forward weeks to show
    n_weeks_achieved: int — how many past weeks to show in achieved sheet

    Returns
    -------
    bytes — the generated .xlsx workbook
    """
    if today is None:
        today = datetime.date.today()

    # ── Load workbooks ────────────────────────────────────────────────────
    ref_wb = openpyxl.load_workbook(io.BytesIO(ref_bytes), data_only=True)
    osr_wb = openpyxl.load_workbook(io.BytesIO(osr_bytes), data_only=True)
    pr_wb  = openpyxl.load_workbook(io.BytesIO(pr_bytes),  data_only=True)

    ref = load_reference(ref_wb)

    # ── Extract data ─────────────────────────────────────────────────────
    osr_rows  = extract_osr(osr_wb)
    prod_rows = extract_prod_register(pr_wb, ref["routing_map"])

    # ── Week ranges ───────────────────────────────────────────────────────
    today_mon     = today - datetime.timedelta(days=today.weekday())
    current_wn    = weeknum(today_mon)

    forward_wns   = list(range(current_wn, current_wn + n_weeks_forward))
    forward_mons  = [week_monday(wn) for wn in forward_wns]
    forward_labels= [f"W #{wn}" for wn in forward_wns]

    achieved_wns  = list(range(current_wn - n_weeks_achieved, current_wn))
    achieved_data, skipped = prod_to_containers(prod_rows, ref, set(achieved_wns))

    order_data    = orders_to_containers(osr_rows, ref, today)

    # ── Build output workbook from reference (keeps all lookup sheets) ────
    out_wb = openpyxl.load_workbook(io.BytesIO(ref_bytes), data_only=False)

    # Remove sheets we will regenerate
    for sn in ["ORDER VS WEEK DISTRIBUTION", "AFRICA", "EUROPE", "USA",
               "Week-Wise Achieved Capacity", "WORKING"]:
        if sn in out_wb.sheetnames:
            del out_wb[sn]

    layout = ref["layout"]

    # ── ORDER VS WEEK DISTRIBUTION ────────────────────────────────────────
    ws_main = out_wb.create_sheet("ORDER VS WEEK DISTRIBUTION")
    write_lookup_table(ws_main, ref["routing_map"], 20)
    main_last_row = build_order_sheet(
        ws_main, layout, ref, forward_labels, forward_mons, order_data)

    # ── Regional sheets ───────────────────────────────────────────────────
    region_row_ranges = {
        "AFRICA":  [],
        "EUROPE":  [],
        "USA":     [],
    }
    r = 3
    for item in layout:
        region = item.get("region", "")
        if region == "AFRICA":
            region_row_ranges["AFRICA"].append(r)
        elif region in ("EUROPE", "RUSSIA"):
            region_row_ranges["EUROPE"].append(r)
        elif region in ("USA", "AUSTRALIA"):
            region_row_ranges["USA"].append(r)
        r += 1 if item["type"] != "blank" else 1

    for sheet_name in ["AFRICA", "EUROPE", "USA"]:
        ws_reg = out_wb.create_sheet(sheet_name)
        _build_regional_mirror(ws_reg, ws_main, region_row_ranges[sheet_name],
                               len(forward_wns), 7)

    # ── Week-Wise Achieved Capacity ───────────────────────────────────────
    ws_ach = out_wb.create_sheet("Week-Wise Achieved Capacity")
    build_achieved_sheet(ws_ach, layout, ref, achieved_wns, achieved_data)

    # ── Update As-Of date in MC SUMMARY - WEEK WISE ───────────────────────
    if "MC SUMMARY - WEEK WISE" in out_wb.sheetnames:
        out_wb["MC SUMMARY - WEEK WISE"]["J1"] = datetime.datetime.combine(
            today, datetime.time())

    # ── Return as bytes ───────────────────────────────────────────────────
    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), skipped


def _build_regional_mirror(ws, ws_main, source_rows, n_weeks, wfc):
    """Build a regional sheet as a live formula mirror of ORDER VS WEEK DISTRIBUTION."""
    FONT_NAME_L = FONT_NAME

    # title + header rows (row 1 & 2)
    ws.merge_cells(f"A1:F1")
    ws["A1"] = "Order Vs Week Distribution"
    ws["A1"].font = H_RED; ws["A1"].fill = TITLE_FILL
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=1, start_column=wfc, end_row=1,
                   end_column=wfc + n_weeks - 1)
    ws.cell(1, wfc, "WEEK WISE ANALYSIS")
    ws.cell(1, wfc).font = H_WHITE; ws.cell(1, wfc).fill = OLIVE_FILL
    ws.cell(1, wfc).alignment = Alignment(horizontal="center", vertical="center")

    for c in range(1, wfc + n_weeks):
        cell = ws.cell(2, c,
            f"='ORDER VS WEEK DISTRIBUTION'!{get_column_letter(c)}2")
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        if c >= wfc:
            ws.column_dimensions[get_column_letter(c)].width = 7
    ws.column_dimensions["A"].width = 12; ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 22; ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 10; ws.column_dimensions["F"].width = 14

    # data rows
    dest_row = 3
    pending_merges = []
    cur_a_start = cur_b_start = None

    for src_row in source_rows:
        a_val = ws_main.cell(src_row, 1).value
        b_val = ws_main.cell(src_row, 2).value
        if a_val not in (None, ""):
            if cur_a_start is not None:
                pending_merges.append(("A", cur_a_start, dest_row - 1))
            cur_a_start = dest_row
        if b_val not in (None, ""):
            if cur_b_start is not None:
                pending_merges.append(("B", cur_b_start, dest_row - 1))
            cur_b_start = dest_row

        is_total = ws_main.cell(src_row, 3).value == "TOTAL"
        for c in range(1, wfc + n_weeks):
            col = get_column_letter(c)
            ref_val = f"'ORDER VS WEEK DISTRIBUTION'!{col}{src_row}"
            cell = ws.cell(dest_row, c, f'=IF({ref_val}="","",{ref_val})')
            cell.font = TOTAL_FONT if is_total else NORM
            cell.fill = TOTAL_FILL if is_total else PatternFill()
            cell.border = BORDER
            if c >= 4:
                cell.number_format = NUMFMT

        if a_val:
            ws.cell(dest_row, 1).fill = REGION_FILL
            ws.cell(dest_row, 1).font = Font(name=FONT_NAME_L, bold=True, size=11)

        dest_row += 1

    LAST_ROW = dest_row - 1
    if cur_a_start is not None:
        pending_merges.append(("A", cur_a_start, LAST_ROW))
    if cur_b_start is not None:
        pending_merges.append(("B", cur_b_start, LAST_ROW))

    for col_letter, ds, de in pending_merges:
        if de > ds:
            ws.merge_cells(f"{col_letter}{ds}:{col_letter}{de}")
        ws[f"{col_letter}{ds}"].alignment = Alignment(
            horizontal="center", vertical="center")

    for i in range(n_weeks):
        col = get_column_letter(wfc + i)
        rng = f"{col}3:{col}{LAST_ROW}"
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND(${col}3<>"",$F3<>"",$F3>0,{col}3>$F3)'],
            font=RED_FONT))

    ws.freeze_panes = f"{get_column_letter(wfc)}3"
    ws.sheet_view.showGridLines = False
