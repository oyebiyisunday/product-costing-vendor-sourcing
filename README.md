# Product costing & vendor sourcing (Excel)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Excel workbook for a **multi-vendor bill of materials (BOM)**: pick a vendor per line, pull **unit price** from that vendor’s sheet, and roll up **total cost** from quantities and build volume.

**Repository:** [github.com/oyebiyisunday/product-costing-vendor-sourcing](https://github.com/oyebiyisunday/product-costing-vendor-sourcing)

---

## About this project

This repository is a **small, production-minded template**: three vendor price tables, one **Product_BOM** sheet with **dynamic lookups** and **batch scaling** via **Units_to_build**. It is intended for learning and portfolio use, and as a starting point you can adapt to your own part numbering, vendors, and approval workflow.

**Maintainer:** [@oyebiyisunday](https://github.com/oyebiyisunday)

---

## License

Released under the **[MIT License](LICENSE)**.

---

## Layout

| Path | Description |
|------|-------------|
| [`workbook/Product_Costing_Vendor_Sourcing.xlsx`](workbook/Product_Costing_Vendor_Sourcing.xlsx) | **Vendor_A**, **Vendor_B**, **Vendor_C** pricing tables + **Product_BOM** with formulas. |
| [`scripts/build_vendor_workbook.py`](scripts/build_vendor_workbook.py) | Regenerates the workbook (Python + [openpyxl](https://openpyxl.readthedocs.io/)). |
| [`docs/problem.md`](docs/problem.md) | Original problem / requirements. |
| [`docs/solution.md`](docs/solution.md) | Design, formulas, and step-by-step usage. |
| [`requirements.txt`](requirements.txt) | Python dependency pin for rebuilding the workbook. |

---

## Quick start

1. Open **`workbook/Product_Costing_Vendor_Sourcing.xlsx`** in **Microsoft 365** or **Excel 2021+** (**XLOOKUP**, **SWITCH**, **IFNA**).
2. Edit vendor parts on **Vendor_A** / **Vendor_B** / **Vendor_C** (inside each table).
3. On **Product_BOM**, set **Units_to_build** in **I2**, then enter **Part_ID**, **Vendor**, and **Qty_Per_Product**.

More detail: **[`docs/solution.md`](docs/solution.md)**.

---

## Rebuild the workbook

From the repository root:

```bash
python -m pip install -r requirements.txt
python scripts/build_vendor_workbook.py
```

This overwrites **`workbook/Product_Costing_Vendor_Sourcing.xlsx`**.

---

## Git clone and remote

Clone (creates `origin` automatically):

```bash
git clone https://github.com/oyebiyisunday/product-costing-vendor-sourcing.git
cd product-costing-vendor-sourcing
```

If you created the project locally with `git init` instead of cloning, add the remote once, then push:

```bash
git remote add origin https://github.com/oyebiyisunday/product-costing-vendor-sourcing.git
git branch -M main
git push -u origin main
```

---

## Requirements

- **Excel:** Microsoft 365 or Excel 2021+.
- **Regenerate script:** Python 3.x; install deps with `requirements.txt` (includes `openpyxl`).

