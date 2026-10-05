# Capacity vs Order Distribution — Report Generator

Streamlit app: upload the **Order Status Report (OSR)** → download the **Capacity vs Order Distribution** workbook
(New capacity — Arul Sir).

## What it produces

The finalized report, with these sheets:

| Sheet | Visible |
|---|---|
| `ORDER VS WEEK DISTRIBUTION` | yes — the main report, grouped **Zone → Region → Sub → Machine Line** |
| `Zone 1` | yes — live mirror of the Zone 1 rows |
| `Zone 2` | yes — live mirror of the Zone 2 rows |
| `Zone 3` | yes — live mirror of the Zone 3 rows |
| `ITEM MASTER` · `TBGS PER CTN` · `MC CAPACITY MASTER` · `MC MASTER` · `ITEM CFC PER CONTAINER` · `WORKING` · `MACHINE CAPACITY MASTER` | hidden (right-click a tab → Unhide) |

- 12 week columns from the OSR week; pending orders in containers per machine line.
- **Excess Order** / **Short Order** rows under every TOTAL (all four visible sheets).
- Capacity is live (unit capacity × Machines × Shifts); change the yellow cells and everything recalculates.
- Over-capacity weeks are red.
- Each machine line belongs to exactly one Zone; regions can span more than one zone (e.g. AFRICA appears under
  Zone 1, Zone 2 and Zone 3 for its different product lines).
- Fonts, colours, column widths, hidden sheets and sheet names are exactly those of the template.

## How it works

`reference_workbook.xlsx` is the finished report and is used as the **template**: layout, fonts, colour palette,
zone sheets, hidden sheets, capacity master, lookups. Each run refreshes only what changes weekly:

| Refreshed | From |
|---|---|
| Week headers (`W #…`) and the hidden week-Monday row 500 | the OSR date you pick |
| `WORKING` sheet | the new OSR |
| `MC ROUTING` (column J of `WORKING`) | **by product name** — from last week's report if you upload it, else the routing bundled in the template |
| Machines / Shifts | last week's report if uploaded, else the template |

Because routing is matched by product name, it does not matter how `WORKING` is sorted in the report you upload.
(Sort the *whole* table, though — if only some columns are sorted, names detach from quantities in that file.)

## After generating

- **New products** with pending orders and no route are listed. Unhide `WORKING`, type a route in column J; they flow
  into the report. Upload that report next week to keep the routing.
- **Routes not in the routing table** (e.g. a product routed to PAKONA, which has no line in the report) are listed
  and not counted. Accepted alias spellings: `CONSTANTA TAG D`, `MD 20 4GM`, `Pearl Pack Premix`.
- Products with no route at all (e.g. bulk lines) are not counted.
- Orders due before the OSR week are counted in the first week; orders beyond the 12th week are not shown
  (the app tells you how many containers that is).

## Updating the masters, look, or capacity

Replace `reference_workbook.xlsx` with an updated finished report (keep the sheet names `ORDER VS WEEK DISTRIBUTION`,
`WORKING`, `MC MASTER`, `ITEM CFC PER CONTAINER`). That is how to change the fonts, the colour palette, sheet names,
which sheets are hidden, the capacity basis (`MC MASTER` column R = unit capacity per machine per shift), the layout,
`ITEM CFC PER CONTAINER` or the routing lookup.

## Files

```
app.py                  — Streamlit UI
report_generator.py     — generation logic
reference_workbook.xlsx — template: the finished report
requirements.txt
README.md
```

## Deploy on Streamlit Cloud

1. Push this repo to GitHub
2. [share.streamlit.io](https://share.streamlit.io) → **New app** → branch `main`, main file `app.py`
3. **Deploy** — no secrets needed
