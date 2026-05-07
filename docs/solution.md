# Product Costing Workbook — Steps & How the Problem Is Solved

The original assignment is summarized in **[`problem.md`](problem.md)**. This document explains **how that problem is solved** in **[`workbook/Product_Costing_Vendor_Sourcing.xlsx`](../workbook/Product_Costing_Vendor_Sourcing.xlsx)** and the **steps** to use and maintain it.

---

## 1. What the problem asks for

- Three vendor-specific price lists (**Part_ID**, **Part_Name**, **Unit_Price**).
- One **Product_BOM** sheet where you pick a vendor per line, enter quantities, and see **unit price** and **total cost**.
- **Unit prices on the BOM must come from the vendor sheets**, not from static numbers typed only on the BOM.
- Quantities must drive **total quantity** and **line cost** automatically.
- The workbook should stay **easy to use** and reasonably **scalable** (more parts, more build quantity).

---

## 2. How the solution is designed (high level)

| Idea | Why it helps |
|------|----------------|
| **Excel Tables** on each vendor sheet (`tblVendorA`, `tblVendorB`, `tblVendorC`) | Clean formatting and easy row entry for new parts. |
| **Excel Table** on **Product_BOM** (`tblBOM`) | Clean filtering, formatting, dropdowns, and a bounded entry area for BOM lines. |
| **Nested `IF` + `VLOOKUP` + `IFERROR`** for **Part_Name** and **Unit_Price** | Chooses which vendor worksheet to read, then looks up **Part_ID** in that sheet—satisfying “price from vendor sheet by vendor + part” while staying compatible with older Excel versions. |
| **`Units_to_build` in I2** | **Total_Qty** = per-product quantity × build count, so one cell scales the whole BOM for batch costing. |
| **Data validation** on **Vendor** | Reduces typos; vendor names must match what the lookup formulas expect. |
| **Generator script** [`scripts/build_vendor_workbook.py`](../scripts/build_vendor_workbook.py) | Recreates the `.xlsx` after you edit sample data or table size in code. |

---

## 3. Workbook layout

### 3.1 Sheet order

1. **Vendor_A** — pricing table `tblVendorA`  
2. **Vendor_B** — pricing table `tblVendorB`  
3. **Vendor_C** — pricing table `tblVendorC`  
4. **Product_BOM** — BOM table `tblBOM` and parameters **I1:I2**

### 3.2 Product_BOM columns (table `tblBOM`)

| Column | Role |
|--------|------|
| **Part_ID** | You enter the part key; must exist on the chosen vendor’s table for a valid price. |
| **Part_Name** | **Formula** — looks up the name from the selected vendor table (or shows a message if the part is not on that vendor). |
| **Vendor** | **Dropdown** — `Vendor_A`, `Vendor_B`, or `Vendor_C`. |
| **Qty_Per_Product** | How many of this part go into **one** finished product. |
| **Total_Qty** | **Formula** — `Qty_Per_Product × $I$2` (see below). |
| **Unit_Price** | **Formula** — `VLOOKUP` on the vendor sheet selected in **Vendor**. |
| **Total_Cost** | **Formula** — `Total_Qty × Unit_Price` when both part and price are present. |

### 3.3 Parameter cell

- **I1** — label: `Units_to_build`  
- **I2** — number of products to cost (default **1**).  
  Changing **I2** updates **Total_Qty** for every BOM line, then **Total_Cost**.

---

## 4. How each requirement is met (formulas)

**Excel version:** **Excel 2010+** or **Microsoft 365** is supported. The workbook intentionally uses broadly compatible functions: **`VLOOKUP`**, **`IFERROR`**, and **`IF`**.

### 4.1 Unit price from the correct vendor sheet

For each BOM row, **Unit_Price** is:

- If **Part_ID** is blank → blank.  
- Else → nested **`IF`** checks **Vendor**:
  - `Vendor_A` → `IFERROR(VLOOKUP($A2, Vendor_A!$A:$C, 3, FALSE), "")`
  - `Vendor_B` → same pattern on `Vendor_B`
  - `Vendor_C` → same pattern on `Vendor_C`
- If the part is missing on that vendor → **`IFERROR`** returns an empty **Unit_Price** (and **Total_Cost** stays blank).

So every price path points at **columns on the vendor worksheets**.

### 4.2 Part name aligned with vendor + Part_ID

**Part_Name** uses the same nested **`IF` + `VLOOKUP` + `IFERROR`** pattern against the **Part_Name** column on the selected vendor sheet. If the part is not on that vendor’s list, the cell shows **-- not on vendor --** (so you notice a mismatch between **Vendor** and **Part_ID**).

### 4.3 Automatic costs from quantities

- **Total_Qty** = `Qty_Per_Product × $I$2`
- **Total_Cost** = `Total_Qty × Unit_Price` only when **Part_ID** is non-empty and **Unit_Price** resolved (avoids multiplying by blank text).

### 4.4 Ease of use and scalability

- **More parts for a vendor:** add rows **inside** that vendor’s table (immediately under the last data row in the table).  
- **More BOM lines in the generated file:** use the blank rows already inside `tblBOM`; increase `table_last_row` if you need more prebuilt formula rows.
- **Larger BOM table in a regenerated file:** increase `table_last_row` in `scripts/build_vendor_workbook.py` and rerun the script.  
- **Another vendor:** add a sheet + table, extend the nested **`IF`** formulas in the script for **Part_Name** and **Unit_Price**, extend the **Vendor** validation list, then regenerate.

---

## 5. Steps — day-to-day use

1. Open **`workbook/Product_Costing_Vendor_Sourcing.xlsx`**.  
2. On **Vendor_A** / **Vendor_B** / **Vendor_C**, add or edit parts **inside** each vendor table. Keep **Part_ID** unique within that vendor.  
3. Go to **Product_BOM**.  
4. Set **Units_to_build** in **I2** (e.g. `100` for a hundred-unit build).  
5. For each line: enter **Part_ID**, choose **Vendor**, enter **Qty_Per_Product**.  
6. Confirm **Part_Name** and **Unit_Price** look right; fix **Vendor** or **Part_ID** if you see **-- not on vendor --** or blank price.
7. Sum **Total_Cost** as needed (e.g. a total row below the table or `=SUM(tblBOM[Total_Cost])` in a spare cell—add manually if you want a grand total).

---

## 6. Steps — rebuild the Excel file from source

1. Install Python, then dependencies: `python -m pip install -r requirements.txt` (from the repository root).  
2. Edit **`scripts/build_vendor_workbook.py`** if you change sample vendor rows, BOM samples, or `table_last_row`.  
3. From the **repository root** (the folder that contains `scripts/` and `workbook/`), run:

   ```bash
   python scripts/build_vendor_workbook.py
   ```

4. This overwrites **`workbook/Product_Costing_Vendor_Sourcing.xlsx`** with a fresh copy (any manual changes in the old file are lost unless you back it up first).

---

## 7. Repository layout (quick map)

| Path | Purpose |
|------|--------|
| [`README.md`](../README.md) | Project overview, quick start, remote URL. |
| [`docs/problem.md`](problem.md) | Problem / requirements text. |
| `docs/solution.md` | This document — steps and how the solution works. |
| [`requirements.txt`](../requirements.txt) | Python dependencies for `scripts/build_vendor_workbook.py`. |
| [`workbook/Product_Costing_Vendor_Sourcing.xlsx`](../workbook/Product_Costing_Vendor_Sourcing.xlsx) | Deliverable workbook. |
| [`scripts/build_vendor_workbook.py`](../scripts/build_vendor_workbook.py) | Programmatic build of the workbook. |
| [`LICENSE`](../LICENSE) | MIT license. |
