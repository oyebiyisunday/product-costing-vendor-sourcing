# Contributing

Thank you for helping improve this project. The goal is to keep the workbook **clear**, **reproducible**, and **safe to run** in typical office environments.

## Before you start

- Read **[`README.md`](README.md)** and **[`docs/solution.md`](docs/solution.md)** so changes stay aligned with the BOM / vendor model.
- Office-oriented guidance lives in **[`docs/office-deployment.md`](docs/office-deployment.md)**.

## How to contribute

1. **Fork** the repository and create a **feature branch** from `main`.
2. Make focused changes (one logical concern per pull request where possible).
3. Ensure **CI passes** (see workflow in `.github/workflows/ci.yml`): dependencies install, the script runs, and the workbook is produced.
4. Open a **pull request** with a clear description of *what* changed and *why*.
5. If you change user-visible behavior of the workbook or rebuild steps, update **`README.md`** and the relevant **`docs/`** file in the same PR.

## Standards

- **Python:** follow existing style in `scripts/` (type hints, `pathlib`, clear names). Prefer `openpyxl` public APIs documented for the pinned version range in `requirements.txt`.
- **Excel compatibility:** do not regress below **Microsoft 365 / Excel 2021** for formula features unless you introduce a clearly labeled **legacy** workbook or sheet variant and document it.
- **Documentation:** keep language precise; link paths relative to the repository root where helpful.

## Code of conduct

All contributors are expected to follow the **[Code of Conduct](CODE_OF_CONDUCT.md)**.
