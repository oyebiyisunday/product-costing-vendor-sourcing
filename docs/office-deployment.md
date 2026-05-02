# Office deployment and collaboration

This note is for **rolling the workbook out in a typical organization**—file placement, co-editing, roles, and technical prerequisites.

---

## 1. Where to store the workbook

For teams that **share one live file**, host it on **Microsoft 365** storage:

- **SharePoint** (team site document library), or  
- **OneDrive for Business** (with a clear owner or shared library).

**Why:** Central version history, permissions, and (where licensed) **real-time co-authoring** in Excel for the web or desktop. Email attachments quickly fork into conflicting copies unless you enforce a single source of truth.

---

## 2. Co-authoring and Excel tables

The deliverable uses **Excel Tables** (structured references). Co-authoring generally works well, but **concurrent edits to the same row or overlapping table resize** can cause merge conflicts or confusion.

**Practical guardrails**

- **Agree on roles:** e.g. *procurement or supply chain* maintains **Vendor_A / Vendor_B / Vendor_C**; *engineering or program* maintains **Product_BOM** lines and **Units_to_build**.  
- **Avoid two people editing the same BOM row at the same time** when possible; split by section or by product variant if the file grows.  
- For heavy parallel work, consider **splitting** into one workbook per program, or a **read-only vendor master** workbook linked or copied on a schedule—only if your process requires it.

---

## 3. Python build script vs. day-to-day users

| Audience | Needs |
|----------|--------|
| **Analysts, planners, buyers** | **Excel only.** Open `workbook/Product_Costing_Vendor_Sourcing.xlsx`, edit tables, refresh if your process uses external links (this template does not require them). |
| **Developers or template owners** | **Python 3** + `requirements.txt` to run `scripts/build_vendor_workbook.py` when you want a **clean regenerated** file from source (e.g. after changing sample data or table size in code). |

Regenerating **overwrites** the committed workbook path; treat the script as **template maintenance**, not a daily end-user tool unless you adopt that workflow deliberately.

---

## 4. Excel version and IT standards

This workbook relies on **XLOOKUP**, **SWITCH**, and **IFNA**. Those functions require a **current-generation** Excel build:

- **Microsoft 365** (recommended in enterprise), or  
- **Excel 2021** / **Office LTSC 2021** or newer.

**Before a company-wide rollout:** confirm the **oldest supported Excel version** in your estate. If you must support **Excel 2019 or earlier**, you will need a **legacy-compatible formula design** (for example, **INDEX/MATCH** and nested **IF** or **CHOOSE** instead of **XLOOKUP/SWITCH**). That is a separate template variant, not the file checked into this repository by default.

---

## 5. Security and compliance posture

- **No VBA macros** and **no external data connections** in the generated template—fewer security review blockers than macro-enabled workbooks.  
- **Sensitive pricing:** apply **SharePoint sensitivity labels**, **DLP**, and **library permissions** as your organization requires; this repository cannot enforce those controls.  
- **Audit trail:** rely on **SharePoint version history** and, if required, **Purview** / change-management process for who may alter vendor master data.

---

## 6. Summary

| Topic | Recommendation |
|-------|----------------|
| **Hosting** | SharePoint or OneDrive; single authoritative copy. |
| **Editing** | Separate responsibilities for vendor sheets vs. BOM where feasible. |
| **Excel** | Microsoft 365 or Excel 2021+; validate against corporate standard. |
| **Python** | Optional; for regenerating the template from source, not for routine costing users. |

For functional behavior of the workbook itself, see **[`solution.md`](solution.md)**.
