# Security policy

## Supported versions

| Component | Supported |
|-----------|-----------|
| `scripts/build_vendor_workbook.py` (latest on `main`) | Yes |
| Older commits | Best-effort only |

The **Excel workbook** is generated output. Treat workbook content you distribute according to your own data-classification policy; this repository does not embed credentials.

## Reporting a vulnerability

If you believe you have found a **security vulnerability** in this repository (for example, unsafe patterns in the Python generator that could affect users who run it):

1. **Do not** open a public issue with exploit details.
2. Use **[GitHub private vulnerability reporting](https://github.com/oyebiyisunday/product-costing-vendor-sourcing/security/advisories/new)** for this repository (if enabled in repository settings), or contact the maintainer ([@oyebiyisunday](https://github.com/oyebiyisunday)) through a private channel if you have one.

Maintainers will acknowledge receipt as soon as practical and coordinate a fix and disclosure timeline.

## Scope

- **In scope:** the Python build script, repository automation, and supply-chain concerns tied to this repo.  
- **Out of scope:** Microsoft Excel product security, third-party `openpyxl` defects (report upstream to [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl)), or misuse of spreadsheets in your organization (handled by your IT and process controls).
