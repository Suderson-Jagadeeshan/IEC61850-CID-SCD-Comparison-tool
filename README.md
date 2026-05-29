# IEC 61850 CID / SCD Comparison Tool

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![License](https://img.shields.io/badge/License-Proprietary-red)
![Version](https://img.shields.io/badge/Version-2.0-green)
![Domain](https://img.shields.io/badge/Domain-IEC%2061850-orange)

A desktop GUI tool for comparing IEC 61850 **CID** and **SCD** configuration files.  
Built for protection and automation engineers who need a fast, reliable way to track  
configuration changes across substation IEDs — without manually diffing XML.

Please check the release note v3.1 - EXE File added for user to access the tool easily 
1. Download the source file >> Extract and keep it in desire folder
2. download IEC61850 CIDTool .v3.Zip
3. extract the Zip to the same repo of Source file and enjoy using the tool

---

## Screenshot

> _Run the tool and load two CID/SCD files to see results like this:_

```
  TOTAL DIFFERENCES FOUND: 7
  ────────────────────────────────────────────
  Config Revision            2 changes
  DataSet                    3 changes
  ReportControl              2 changes
  ────────────────────────────────────────────

  CATEGORY               ELEMENT                              OLD VALUE         NEW VALUE
  ───────────────────────────────────────────────────────────────────────────────────────
  Config Revision        IED.configVersion                    1                 2
  DataSet                DS 'DS_PROT_01' → PDIS/Op fc=ST      EXISTS            REMOVED
  Signal / FCDA          DS 'DS_PROT_01' → PTOV/Op fc=ST      NOT PRESENT       ADDED
  ReportControl          RCB 'RCB_PROT_01' → intgPd           1000              2000
  IP / Network           IP Address                           192.168.1.101     192.168.1.105
```

---

## Features

- **CID & SCD support** — works with single-IED CID files and multi-IED SCD files
- **Auto IED detection** — browses the file and populates an IED selector dropdown automatically
- **Auto technical key** — derives the technical key prefix from the IED name, no manual input needed
- **Side-by-side diff view** — Old Value vs New Value in a clean table layout
- **Colour-coded results** — green for additions, red for removals, blue for modifications
- **Summary count** — total differences broken down by category at the top of every result
- **Excel export (.xlsx)** — colour-coded workbook with a Summary sheet, full Differences sheet, and one tab per category
- **PDF export** — formatted report with your name, project details, and logo
- **Progress bar** — background threading keeps the UI responsive on large SCD files
- **License notice** — accessible via Help → License & Copyright

### What gets compared

| Category | Details |
|---|---|
| Config Revision | `configVersion`, `desc`, `type`, `manufacturer` on the IED element |
| IP / Network | IP address from `ConnectedAP` |
| DataSet | Added / removed `DataSet` elements |
| Signal / FCDA | Added / removed `FCDA` signals within DataSets |
| ReportControl | `confRev`, `buffered`, `intgPd`, `datSet`, `rptID` on RCBs |
| Logical Node | Added / removed `LN` elements across Logical Devices |

---

## Requirements

### Python
Python 3.8 or higher — download from [python.org](https://www.python.org/downloads/)

### Dependencies

| Package | Purpose | Install |
|---|---|---|
| `openpyxl` | Excel export | `pip install openpyxl` |
| `reportlab` | PDF export | `pip install reportlab` |
| `Pillow` | Logo display in GUI | `pip install Pillow` |

> The tool runs without any of the above — Excel/PDF buttons will show an install prompt if the library is missing.

Install all at once:
```bash
pip install openpyxl reportlab Pillow
```

---

## Getting Started

### 1. Clone the repository
```bash
git clone https://github.com/YOUR_USERNAME/iec61850-cid-compare.git
cd iec61850-cid-compare
```

### 2. Install dependencies
```bash
pip install openpyxl reportlab Pillow
```

### 3. Run the tool
```bash
python cid_gui.py
```

### 4. Optional — add your logo
Place a file named `your_logo.png` in the same folder as `cid_gui.py`.  
Recommended size: 100 × 50 px. The tool works fine without it.

### 5. Optional — add DejaVu font for PDF
Download `DejaVuSans.ttf` from [dejavu-fonts.org](https://dejavu-fonts.github.io/)  
and place it in the same folder. Falls back to Helvetica if not found.

---

## How to Use

1. **Browse** the Old CID/SCD file — the IED dropdown fills automatically
2. **Browse** the New CID/SCD file — same
3. **Select the IED** you want to compare from each dropdown (important for SCD files with multiple IEDs)
4. The **Technical Key** is shown automatically — no typing needed
5. Choose your **Output Mode**: View Results / Generate PDF / Export Excel
6. For PDF or Excel, fill in your Name, Designation, Project Name, Voltage Level
7. Click **▶ Run Comparison**

---

## File Structure

```
iec61850-cid-compare/
│
├── cid_gui.py              # Main application (GUI + built-in comparison engine)
├── cid_compare_core.py     # Optional: plug in your own comparison logic here
├── your_logo.png           # Optional: your company logo
├── DejaVuSans.ttf          # Optional: font for PDF export
│
├── sample_files/
│   ├── sample_old.cid      # Sample CID file for testing
│   └── sample_new.cid      # Sample CID file with deliberate differences
│
├── README.md
├── CHANGELOG.md
└── LICENSE
```

---

## Sample Files

Two sample CID files are included in `sample_files/` for testing.  
They contain 7 deliberate differences covering every comparison category:

| # | Change | Category |
|---|---|---|
| 1 | `configVersion` 1 → 2 | Config Revision |
| 2 | `PDIS/Op` removed from DS_PROT_01 | DataSet + Signal |
| 3 | `PTOV/Op` added to DS_PROT_01 | DataSet + Signal |
| 4 | RCB `confRev` 1 → 3, `intgPd` 1000 → 2000 | ReportControl |
| 5 | `TotW` removed, `TotVAr` added in DS_MEAS_01 | Signal |
| 6 | Measurement RCB `intgPd` 5000 → 10000 | ReportControl |
| 7 | IP address `192.168.1.101` → `192.168.1.105` | IP / Network |

---

## Extending with cid_compare_core.py

The tool works standalone with its built-in engine.  
If you have your own comparison logic, create `cid_compare_core.py` with:

```python
def compare_cid_files(old_path, new_path, tech_prefix,
                      old_ied=None, new_ied=None):
    """
    Returns a list of dicts:
      { 'category': str, 'element': str, 'old': str, 'new': str }
    """
    ...
```

The GUI will automatically detect and use it instead of the built-in engine.

---

## Roadmap

- [ ] Filter results by category (checkboxes)
- [ ] Recent files list
- [ ] Dark mode
- [ ] Compare multiple IED pairs at once
- [ ] GOOSE and SMV subscription comparison
- [ ] SCL schema validation before comparison

---

## Author

**Suderson Jagadeeshan**  
Protection & Automation Engineer  
📧 suderson123@gmail.com

---

## License

© 2024 Suderson Jagadeeshan. All rights reserved.

This software is proprietary. You may use it for internal engineering work and  
generate reports for your projects. You may not copy, redistribute, modify, or  
use it for commercial resale. See `LICENSE` for full terms.
