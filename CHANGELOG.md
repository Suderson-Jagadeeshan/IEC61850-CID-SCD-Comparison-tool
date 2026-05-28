# Changelog

All notable changes to the IEC 61850 CID / SCD Comparison Tool are documented here.

---

## [2.0] — 2024

### Added
- Excel export (.xlsx) with Summary sheet, full Differences sheet, and per-category tabs
- PDF export with wrapped lines, metadata, and logo support
- Progress bar with background threading — UI stays responsive on large SCD files
- Side-by-side diff view: Category | Element | Old Value | New Value
- Summary count at top of results, broken down by category
- License & Copyright dialog accessible via Help menu
- Colour-coded results in the View pane (green=added, red=removed, blue=modified)

### Changed
- Technical key prefix is now auto-derived from the selected IED name — no manual input
- Metadata fields (Name, Designation etc.) now correctly enable for both PDF and Excel modes
- Excel sheet names now strip invalid characters (/, \, ?, *, [, ], :)

### Fixed
- `ValueError: Invalid character / found in sheet title` when exporting Excel
- Metadata fields inaccessible when switching to PDF mode
- GUI crash when `DejaVuSans.ttf`, `your_logo.png`, or `cid_compare_core.py` were missing

---

## [1.1] — 2024

### Added
- SCD file support with per-file IED selector dropdown
- IED list auto-populated on file browse
- Technical key prefix auto-filled from first IED name
- Horizontal scroll in result box
- Run button disabled until both files are selected
- Metadata fields hidden when in View mode
- Report date pre-filled with today's date

### Fixed
- GUI crash on missing `reportlab` or `cid_compare_core` at import time
- GUI crash when `your_logo.png` not found

---

## [1.0] — 2024

### Initial Release
- Basic CID file comparison via `cid_compare_core.py`
- Tkinter GUI with file browse, technical key prefix input
- View results in scrolled text box
- PDF export via reportlab
- Categorised differences: DataSet, Signal, RCB, IP, Config Revision
