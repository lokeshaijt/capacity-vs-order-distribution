# MC Capacity vs Order — Week-Wise Report Generator

Streamlit app that takes two weekly uploads and produces the full
MC Capacity vs Order workbook automatically.

## What it generates

| Sheet | Description |
|---|---|
| `ORDER VS WEEK DISTRIBUTION` | 12-week forward order vs capacity by region/machine line |
| `AFRICA` / `EUROPE` / `USA` | Regional mirrors (live formulas) |
| `Week-Wise Achieved Capacity` | Last 4 weeks actual production vs achievable capacity |

## Files in this repo

```
app.py                  — Streamlit UI
report_generator.py     — All generation logic (no external API)
reference_workbook.xlsx — Master data: MC MASTER, ITEM MASTER, TBGS PER CTN,
                          routing map, layout structure
requirements.txt
```

## How to deploy on Streamlit Cloud

1. Fork / push this repo to GitHub (make it **public** or connect your account).
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Select your repo, branch `main`, main file `app.py`.
4. Click **Deploy** — no secrets or environment variables needed.

## How to use

1. Open the app URL.
2. Upload today's **Order Status Report** (`.xlsx`).
3. Upload today's **Production Register** (`.xlsx`).
4. Click **Generate Report** → download the `.xlsx`.

## Updating master data

The reference workbook is bundled with the app.
When MC MASTER, ITEM MASTER, TBGS PER CTN, routing layout, or
the ORDER VS WEEK DISTRIBUTION structure changes:

1. Save the latest workbook as `reference_workbook.xlsx`.
2. Replace the file in this repo and push — Streamlit Cloud
   will auto-redeploy within a minute.

## How the conversion works

```
Production Register (CTN)
  × TBGS per CTN  (from TBGS PER CTN sheet, or derived from product name)
  ÷ TBGS per CFC  (from MC MASTER)
  ÷ CFC per Container  (from MC MASTER)
= Achieved containers/week

Order Status (Pending Prod in CFC)
  ÷ CFC per Container  (item-specific from ITEM CFC PER CONTAINER,
                         or machine-level from MC MASTER as fallback)
= Pending containers/week  →  bucketed by Buyer Requested Shipment Date
  (overdue orders pulled into current week)
```
