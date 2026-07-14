# Voucher Eligibility Automation Tool

Streamlit application for automated voucher code (VC) eligibility processing across **Lazada**, **Zalora**, and **Shopee** marketplaces for **SG**, **MY**, and **PH** regions.

---

## Changelog — Latest Enhancements

- **Shopee SellerSKU handling & SKU length validation.** The Shopee pipeline now explicitly uses **SellerSKU** as the SKU reference throughout the entire Shopee filtration process (column resolved via `SellerSKU` → `Seller SKU` → `SKU` → `seller_sku`, case-insensitive). Before any eligibility or voucher filtration runs, rows are validated against Shopee's fixed SKU format: **only rows where the SellerSKU is exactly 13 characters long are kept**; every other row (blank, truncated, malformed, or otherwise mis-lengthed SellerSKU) is dropped up front so it can never leak into the Product-ID-level flag propagation, PID exclusion columns, or voucher eligibility results. A summary of how many rows were removed is shown in the "Column Mapping Info" panel after processing.
- **ZIP upload support (all file inputs).** SC Report, Ecom Tracker, Content File, and AM Exclusion Sheet uploaders now all accept `.zip` archives in addition to `.xlsx` / `.xls` / `.csv`. When a `.zip` is uploaded, every supported data file inside it is automatically extracted and the resulting data is **consolidated** (concatenated) into a single dataset before processing — so a zip containing multiple SC Report exports, multiple Ecom Tracker files, etc. is treated exactly as if it were one combined file. Hidden/system files (`.DS_Store`, `__MACOSX/…`, Excel lock files) are skipped automatically. A corrupted zip, an empty zip, or a zip whose Ecom Tracker member is missing the required region tab all surface a clear validation error naming the specific file inside the archive.
- **Region-based Ecom Tracker worksheet resolution.** The Ecom Tracker is now read exclusively from the worksheet whose name exactly matches the selected Region (case-insensitive) — e.g. Region = `PH` reads only the `PH` tab. Every downstream field, lookup, mapping, and validation (filtering, eligibility checks, Launch Date lookup, RRP, SRP, exclusions, pricing) is sourced from that worksheet alone. If the required tab is missing from the uploaded workbook, the app shows a clear validation error instead of silently falling back to another sheet.
- **Shopee marketplace support — Product ID (PID) level eligibility.** Shopee uses a completely different eligibility model from Lazada/Zalora: every check runs at the **Product ID** level rather than the ALU level (a Product ID can have multiple SC Report rows/variations and can be tied to multiple ALUs).
  - Four mandatory PID flag columns are always added: `Future_PID`, `Ecom NO_PID`, `Low Price_PID`, `AM Exclude ALU_PID`. A Product ID flagged in *any* of these is immediately ineligible for every voucher.
  - One dynamic `<Exclusion Remark>_PID` column is generated automatically per unique Exclusion Remarks value found in the Ecom Tracker.
  - Voucher columns are appended after all PID flag/dynamic columns.
  - Inclusion Keyword logic: `Open for all` ignores all dynamic exclusion columns (only the 4 mandatory flags apply); any other keyword requires a match in its corresponding dynamic `_PID` column.
  - The "Export - Eligible List" button exports eligible **Product IDs** (deduplicated) for Shopee.
- **PH Region** added alongside SG and MY.
- **Region-specific RRP / SRP thresholds:**
  - SG: 16
  - MY: 38.9
  - PH: 649
  - Rule (all regions): RRP must be `> threshold` (RRP ≤ threshold, RRP = 0, or blank RRP are excluded). SRP is eligible when `0` or `> threshold`; excluded when `0 < SRP ≤ threshold`.
- **"NA" Launch Date values** in the selected Launch Date column are now treated as eligible (cutoff check is skipped for these rows) across all regions.
- **Launch Date Column Selection** — a new sidebar dropdown, populated dynamically from the uploaded Ecom Tracker's column headers, replaces the previous hardcoded Launch Date column detection. Launch Date Cutoff validation always uses the column selected here (falling back to auto-detection if not yet selected).
- **AM Exclusion Sheet — marketplace-aware matching + fallback rule:**
  - Matching is by ALU.
  - If the AM Exclusion Sheet has a Marketplace column, rows are scoped to the currently selected Marketplace using case-insensitive, multi-delimiter matching (e.g. `Lazada, Zalora` / `Lazada & Zalora`). Rows with a blank Marketplace value apply to all marketplaces.
  - If an ALU is present in the AM Exclusion Sheet (and applicable to the selected marketplace) but its Exclusion Type text does not resolve into a recognised rule, the ALU now defaults to **"Exclude from all Vouchers"** rather than being silently treated as not excluded.
- **Export — Eligible List button:** exports only eligible records per voucher to `.xlsx`.
  - Lazada → exports the eligible **Shop SKU** values.
  - Zalora → exports the eligible **SellerSku** values.
  - File naming: `<Marketplace>_<Region>_<Marketplace>_<Voucher Name>.xlsx` (e.g. `Lazada_PH_Lazada_30% NMS.xlsx`). The worksheet tab is renamed to exactly match the Voucher Name.
  - If more than one voucher is configured, all per-voucher files are bundled into a single `.zip` download.

---

## Project Structure

```
voucher_app/
├── app.py              # Streamlit dashboard UI
├── processor.py        # All business logic & data processing
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

---

## Setup & Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Deploy to Streamlit Community Cloud

1. Push this folder to a **GitHub repository**.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in.
3. Click **New app** → select your repo and branch.
4. Set **Main file path** to `app.py`.
5. Click **Deploy**.

> `requirements.txt` must be present at the repository root.

---

## Input Files

| File | Required | Format | Notes |
|------|----------|--------|-------|
| SC Report | ✅ | `.xlsx` / `.xls` / `.csv` | Seller Center export for selected marketplace |
| Ecom Tracker | ✅ | `.xlsx` / `.xls` / `.csv` | Headers auto-detected (scans for STYLE# row) |
| Content File | ✅ | `.xlsx` / `.xls` / `.csv` | EAN → ALU mapping source |
| AM Exclusion Sheet | ⬜ Optional | `.xlsx` / `.xls` / `.csv` | Region-specific exclusion list |

---

## Dashboard Inputs

| Input | Type | Description |
|-------|------|-------------|
| Select Marketplace | Dropdown | `Lazada` or `Zalora` |
| Region | Dropdown | `SG` or `MY` |
| Price Tier Reference | Text | Column range e.g. `AX-BA` |
| Launch Date Cutoff | Date Picker | Products launched after this date are excluded |
| Voucher Configurations | Dynamic cards | Add multiple vouchers; each evaluated independently |

---

## Key Column Mappings

### ALU Mapping
| Source | Column | Purpose |
|--------|--------|---------|
| SC Report | `SellerSKU` | Matched against Content File EAN |
| Content File | `EAN` | Linked to SellerSKU |
| Content File | `Color No` | Source of ALU value |
| Ecom Tracker | `STYLE#` | Matched to ALU to retrieve product data |

### Ecom Tracker Column Lookups
| Field | Column Name |
|-------|-------------|
| Launch Date | `Launch Dates` (format: **DD-MM-YYYY**) |
| Ecom Status (Lazada) | `Lazada` |
| Ecom Status (Zalora) | `Zalora` |

### Price Tier Reference — Positional Mapping

Example reference: `AX-BA`

| Position | Excel Col | Field |
|----------|-----------|-------|
| START | AX | RRP |
| START+1 | AY | SRP |
| START+2 | AZ | DISC % |
| END | BA | Exclusion Remarks |

Column names are resolved **by header name first**, then by position as fallback.

### AM Exclusion Sheet Tabs
| Region | Tab |
|--------|-----|
| SG | `SG VC Exclusions` |
| MY | `MY VC Exclusions` |

---

## Inclusion Keywords — Dropdown

When the Ecom Tracker is uploaded, the app reads all unique non-blank values from the **Exclusion Remarks** column and populates a multiselect dropdown for each voucher.

- Select one or more values per voucher.
- A SKU must **exactly match** at least one selected keyword to pass the filter.
- Matching is **case-sensitive and exact** — `"OPEN FOR ALL"` does not match `"OPEN FOR ALL (10days max)"`.
- Leave the dropdown empty to skip keyword filtering for that voucher.

If the Ecom Tracker is not yet uploaded, a text input is shown as a fallback (comma-separated values).

---

## Date Handling

The Ecom Tracker `Launch Dates` column stores dates in **DD-MM-YYYY** format. The tool preserves this format throughout:

| Input type | Example | Handled? |
|------------|---------|----------|
| DD-MM-YYYY string (native) | `13-05-2025` | ✅ Primary format |
| pandas Timestamp (Excel read) | `Timestamp('2025-05-13')` | ✅ Converted to DD-MM-YYYY |
| datetime object | `datetime(2025, 5, 13)` | ✅ Converted to DD-MM-YYYY |
| DD/MM/YYYY slash | `13/05/2025` | ✅ Converted to DD-MM-YYYY |
| ISO / pandas string | `2025-05-13` or `2025-05-13 00:00:00` | ✅ Converted to DD-MM-YYYY |
| `Past Season`, `TBC` | — | ❌ Excluded |
| `00-Jan-1900` / blank | — | ❌ Excluded |
| Any text without digits | — | ❌ Excluded |

**Output format is always DD-MM-YYYY**, regardless of how pandas read the source cell.

---

## Voucher Eligibility Validation Logic

Each voucher is evaluated **completely independently**. Filters are applied in order:

| # | Filter | Rule |
|---|--------|------|
| 1 | Launch Date | Must be a valid **DD-MM-YYYY** date ≤ Cutoff Date. Excludes: blank, `00-Jan-1900`, `Past Season`, `TBC`, any non-date text. All date inputs are normalised to DD-MM-YYYY in the output. |
| 2 | Ecom Status | Must be exactly `YES` (marketplace-specific column: `Lazada` or `Zalora`). |
| 3 | RRP | Must be > 16. Excludes: blank, zero, ≤ 16. |
| 4 | SRP | Must be 0 or > 16. Excludes: SRP values between 1 and 16 inclusive. |
| 5 | Inclusion Keyword | Exclusion Remarks column must exactly match a configured keyword. Skip if no keywords selected. |
| 6 | AM Exclusion | ALU must not be excluded for this voucher's percentage. |

**Output:** `Yes-Eligible` if all filters pass. Blank if not eligible.

### SRP Examples
| SRP | Result |
|-----|--------|
| 0 | ✅ Eligible |
| 18 | ✅ Eligible |
| 16 | ❌ Excluded |
| 10 | ❌ Excluded |

---

## AM Exclusion Logic

Exclusion text is parsed case-insensitively with flexible pattern matching. Multiple rules can apply to a single ALU — the most restrictive combination is used.

### Rule 1 — Global Exclusion
Triggered by any of these phrases (case-insensitive):

- `Exclude from all Voucher`
- `Exclude from platform voucher`
- `Exclude`
- `Exclude from VC`
- `Exclude from Voucher`
- `Voucher exclusion`

**Action:** ALU is excluded from **every** voucher.

---

### Rule 2 — Exact Percentage Exclusion
Triggered by patterns like:
- `Exclude from 10% VC`
- `Exclude from 15%`
- `Exclude from 20% Voucher`

**Action:** ALU is excluded only from vouchers whose name contains the **same percentage**.

**Example:**

| Vouchers | AM Exclusion | Result |
|----------|-------------|--------|
| 10% ABC, 15% DEF | Exclude from 10% Voucher | Excluded from 10% ABC; Eligible for 15% DEF |

---

### Rule 3 — Percentage and Above Exclusion
Triggered by patterns like:
- `Exclude from 30% and above`
- `Exclude from 40% above`

**Action:** ALU is excluded from all vouchers with a percentage **≥ threshold**.

**Examples:**

| Vouchers | AM Exclusion | Result |
|----------|-------------|--------|
| 30% NMS, 40% NMS, 45% MS80 | Exclude from 30% and above | All three excluded |
| 30% NMS, 40% NMS, 45% MS80 | Exclude from 40% | Only 40% NMS excluded |
| 30% NMS, 40% NMS, 45% MS80 | Exclude from 40% and above | 40% NMS and 45% MS80 excluded; 30% NMS eligible |

---

## Output

**File name:** `Voucher_Eligibility_Output_{Marketplace}.xlsx`

| Section | Columns |
|---------|---------|
| Original SC Report | All columns from the uploaded SC Report |
| Working Columns | `ALU`, `Launch Date`, `Ecom Status`, `RRP`, `SRP`, `DISC %`, `Exclusion`, `AM Exclude` |
| Voucher Columns | One column per configured voucher — `Yes-Eligible` or blank |

**Excel formatting:**
- 🔵 Dark blue headers = SC Report columns
- 🟢 Dark green headers = Working columns
- 🟣 Purple headers = Voucher columns
- Green cell highlight = `Yes-Eligible`
- Alternating row shading
- Frozen header row
