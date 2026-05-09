# ICS Visits Manager

A desktop tool for parsing nurse home-visit `.ics` calendar files, computing visit amounts, flagging OASIS requirements, detecting duplicates, and exporting to Excel.

📖 [中文文档 / Chinese version](README.zh-CN.md)

## Features

- **Parse `.ics` files** — extract every `VEVENT` (date, time, patient, visit type, location)
- **Auto-pricing** — apply per-type rates (SOC $100, DC $60, RECERT/ROC $70, IV/FU $65)
- **OASIS detection** — auto-flag SOC / DC / RECERT / ROC visits as OASIS-required
- **Duplicate & cancellation report** — same UID multiple versions, `STATUS:CANCELLED` events, missing UIDs, same-patient-same-time clashes
- **Date range filtering** — today / this week / this month / last month / all / custom range
- **Excel export** — Clean (with amounts + color-coded OASIS column), Raw (full original fields), Bundle (both sheets), and a separate Duplicate Report
- **Native macOS GUI** built with PyQt6 — no browser required

## Two flavors

| File | UI | Notes |
|------|----|----|
| `nurseapp_qt.py` | Native PyQt6 window | Recommended — packageable into a `.app` |
| `nurseapp.py` | Streamlit (browser) | Original prototype, still works |

## Requirements

- macOS (tested on Apple Silicon; build script targets `universal2` so Intel works too)
- Python 3.10+
- Packages: `pandas`, `PyQt6`, `icalendar`, `openpyxl`

## Quick start (run from source)

```bash
pip install pandas PyQt6 icalendar openpyxl
python nurseapp_qt.py
```

The Streamlit version:

```bash
pip install streamlit pandas icalendar openpyxl
streamlit run nurseapp.py
```

## Build a distributable `.app` (macOS)

The build script creates a clean virtual environment, installs only what's needed, builds with PyInstaller, fixes the Qt cocoa plugin, and zips the result.

**Prerequisites:** install [Python from python.org](https://www.python.org/downloads/macos/) (the **macOS 64-bit universal2 installer**). Adjust the version path in `build_app.sh` if you didn't install 3.13.

```bash
bash build_app.sh
```

Output:
- `dist/ICS Visits 管理.app` — double-click to run
- `dist/ICS Visits 管理.zip` — send this to others

The resulting `.app` is `universal2`, runs on both Intel and Apple Silicon Macs.

## Distributing to non-technical users

The app is unsigned, so on first launch macOS will block it. Tell recipients:

> Right-click the `.app` → **Open** → click **Open** in the dialog. After the first time, double-click works normally.

## Project structure

```
.
├── nurseapp_qt.py     # Main PyQt6 app
├── nurseapp.py        # Streamlit version (legacy)
├── run_app.py         # Streamlit launcher
├── build_app.sh       # macOS .app build script
├── rthook_qt.py       # PyInstaller runtime hook for Qt plugin path
└── README.md
```

## Visit type mapping

| Type    | Amount | OASIS required |
|---------|--------|----------------|
| SOC     | $100   | YES            |
| DC      | $60    | YES            |
| RECERT  | $70    | YES            |
| ROC     | $70    | YES            |
| IV      | $65    | NO             |
| FU      | $65    | NO             |

The visit type is parsed from the calendar event summary in `Patient Name - Visit Type` format. Variants like `Follow Up`, `F/U`, `FU` all normalize to `FU`.

## Notes

- Cancelled events (`STATUS:CANCELLED` or summaries starting with `cancel`) are excluded from Clean output by default but kept in Raw — toggle the "include cancelled" checkbox to include them.
- Deduplication: same `UID + RECURRENCE-ID + DTSTART` keeps only the latest by `SEQUENCE` / `LAST-MODIFIED`.
- Conda's PyQt6 ships a Qt5-linked `libqcocoa.dylib` by mistake; the build script overwrites it with the correct Qt6 version after PyInstaller runs.
