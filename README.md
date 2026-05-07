# Product Costing & Vendor Sourcing

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Status](https://img.shields.io/badge/Status-Production%20Ready-brightgreen.svg)
![Excel](https://img.shields.io/badge/Excel-2021%2B-217346.svg)
![Python](https://img.shields.io/badge/Python-3.x-3776ab.svg)

**A production-ready Excel template for multi-vendor bill of materials (BOM) management with dynamic pricing, cost rollup, and batch scaling.**

Pick a vendor per line item, pull unit prices from vendor-specific sheets, and automatically calculate total costs based on quantities and build volume—all with zero manual calculation.

**[Live Repository](https://github.com/oyebiyisunday/product-costing-vendor-sourcing)** · **[Documentation](docs/)** · **[Report Issue](https://github.com/oyebiyisunday/product-costing-vendor-sourcing/issues)**

---

## Table of Contents

- [Features](#features)
- [Quick Start](#quick-start)
- [Project Structure](#project-structure)
- [Usage Guide](#usage-guide)
- [Rebuild & Customize](#rebuild--customize)
- [Requirements](#requirements)
- [Git Setup](#git-setup)
- [Documentation](#documentation)
- [Support](#support)
- [License](#license)

---

## Features

✅ **Multi-Vendor Support** – Compare pricing across three vendor tables (easily extensible)  
✅ **Dynamic Lookups** – XLOOKUP formulas automatically pull unit prices based on selected vendor  
✅ **Batch Scaling** – Single `Units_to_build` input scales all line-item quantities  
✅ **Total Cost Rollup** – Automatic cost calculation per part and overall BOM  
✅ **Production-Ready** – Uses Excel 2021+ native functions (XLOOKUP, SWITCH, IFNA)  
✅ **Rebuild Script** – Python script to regenerate or customize the workbook  
✅ **Fully Documented** – Problem definition, solution design, and usage guides included  

---

## Quick Start

### 1. Open the Workbook

Open `workbook/Product_Costing_Vendor_Sourcing.xlsx` in **Microsoft 365** or **Excel 2021+** (required for `XLOOKUP`, `SWITCH`, and `IFNA`).

### 2. Review Vendor Pricing

Visit **Vendor_A**, **Vendor_B**, or **Vendor_C** sheets to see current pricing tables.

### 3. Create Your BOM

On the **Product_BOM** sheet, fill the user-input cells; formula cells recalculate automatically:

| Cell | Field | Type | Example |
|------|-------|------|---------|
| `I2` | `Units_to_build` | Input | `100` |
| Column A | `Part_ID` | Input | `P001` |
| Column B | `Part_Name` | Formula (auto) | _looked up_ |
| Column C | `Vendor` | Input (dropdown) | `Vendor_A` |
| Column D | `Qty_Per_Product` | Input | `10` |
| Column E | `Total_Qty` | Formula (auto) | `D × $I$2` |
| Column F | `Unit_Price` | Formula (auto) | _looked up_ |
| Column G | `Total_Cost` | Formula (auto) | `E × F` |

→ **Total cost updates automatically!**

For detailed guidance: **[docs/solution.md](docs/solution.md)**

---

## Project Structure

```
product-costing-vendor-sourcing/
├── workbook/
│   └── Product_Costing_Vendor_Sourcing.xlsx    # Main Excel workbook with vendor sheets & BOM
├── scripts/
│   ├── build_vendor_workbook.py                # Regenerate workbook (Python + openpyxl)
│   └── verify_workbook.py                      # CI check: rebuilds and diffs against committed xlsx
├── docs/
│   ├── problem.md                              # Original problem statement & requirements
│   └── solution.md                             # Design, formulas, and step-by-step usage
├── .github/workflows/
│   └── ci.yml                                  # GitHub Actions workflow (lint + workbook verify)
├── requirements.txt                            # Python dependencies
├── LICENSE                                     # MIT License
└── README.md                                   # This file
```

### Key Files

| Path | Purpose |
|------|---------|
| `workbook/Product_Costing_Vendor_Sourcing.xlsx` | **Vendor_A**, **Vendor_B**, **Vendor_C** pricing tables + **Product_BOM** with XLOOKUP formulas |
| `scripts/build_vendor_workbook.py` | Regenerates the workbook using Python (openpyxl) |
| `scripts/verify_workbook.py` | Rebuilds to a temp file and asserts byte-for-cell parity with the committed `.xlsx` |
| `.github/workflows/ci.yml` | Runs compile check + `verify_workbook.py` on every push and pull request to `main` |
| `docs/problem.md` | Detailed problem definition and original requirements |
| `docs/solution.md` | Architecture, formula explanations, and full usage walkthrough |
| `requirements.txt` | Python dependency pin (`openpyxl`) |

---

## Usage Guide

### For End Users (Non-Technical)

1. **Open** `workbook/Product_Costing_Vendor_Sourcing.xlsx`
2. **Check vendor pricing** on Vendor_A, Vendor_B, Vendor_C sheets (update as needed)
3. **Set `Units_to_build`** in cell I2 on the Product_BOM sheet
4. **Enter your parts** (only the input columns; formula columns fill themselves):
   - Column A: `Part_ID`
   - Column C: `Vendor` (dropdown — Vendor_A, Vendor_B, or Vendor_C)
   - Column D: `Qty_Per_Product`
5. **View results:** `Part_Name`, `Total_Qty`, `Unit_Price`, and `Total_Cost` roll up automatically

### For Developers (Customization)

See **[docs/solution.md](docs/solution.md)** for:
- Formula breakdowns (XLOOKUP, SWITCH, IFNA)
- Extending to more vendors
- Adding new columns or calculations

---

## Rebuild & Customize

### Prerequisites

- Python 3.x
- `openpyxl` (install via requirements.txt)

### Rebuild the Workbook

From the repository root:

```bash
# Install dependencies
python -m pip install -r requirements.txt

# Regenerate workbook
python scripts/build_vendor_workbook.py
```

This will overwrite `workbook/Product_Costing_Vendor_Sourcing.xlsx` with the current Python-generated version.

### Customization

Edit `scripts/build_vendor_workbook.py` to:
- Add more vendor sheets
- Modify table structure or columns
- Update vendor pricing data programmatically
- Include additional calculations

---

## Requirements

| Requirement | Version | Notes |
|-------------|---------|-------|
| **Excel** | 2021 or Microsoft 365 | Supports XLOOKUP, SWITCH, IFNA functions |
| **Python** (rebuild only) | 3.x | Required only if regenerating the workbook |
| **openpyxl** (rebuild only) | Latest | Installed via `requirements.txt` |

---

## Git Setup

### Clone This Repository

```bash
git clone https://github.com/oyebiyisunday/product-costing-vendor-sourcing.git
cd product-costing-vendor-sourcing
```

### Add Remote (If Starting from Local)

If you created the project locally with `git init`:

```bash
git remote add origin https://github.com/oyebiyisunday/product-costing-vendor-sourcing.git
git branch -M main
git push -u origin main
```

---

## Documentation

| Document | Purpose |
|----------|---------|
| **[docs/problem.md](docs/problem.md)** | Original problem statement, business requirements, and constraints |
| **[docs/solution.md](docs/solution.md)** | Detailed solution architecture, formula logic, and step-by-step usage |
| **[LICENSE](LICENSE)** | MIT License text |

---

## Support

- 📋 **Questions or issues?** Open an [issue on GitHub](https://github.com/oyebiyisunday/product-costing-vendor-sourcing/issues)
- 🚀 **Want to contribute?** Pull requests welcome!
- 📧 **Contact:** [@oyebiyisunday](https://github.com/oyebiyisunday)

---

## License

This project is released under the **[MIT License](LICENSE)**.

**You are free to:**
- ✅ Use commercially
- ✅ Modify and distribute
- ✅ Use privately

**With conditions:**
- ⚖️ Provide a copy of the license
- ⚖️ State changes made to the original

---

## Changelog

**v1.0.0** (2026-05-06)
- ✨ Initial release
- 📦 Multi-vendor BOM template with dynamic pricing
- 🐍 Python rebuild script included
- 📚 Comprehensive documentation

---

**Maintained by [@oyebiyisunday](https://github.com/oyebiyisunday)**
