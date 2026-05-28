# ICS Visits Manager

> ## ⚖️ Disclaimer
>
> This software is a **fully local, offline desktop application**. It does not transmit, upload, or share any user data with any remote server, cloud service, or third party.
>
> All data processing, file reading, and Excel exports happen entirely on the user's own computer. The user is solely responsible for:
>
> - **Backing up** any input (`.ics`) and output (`.xlsx`) files before use;
> - **Verifying** the correctness of all computed amounts, OASIS flags, and exported data;
> - **Safeguarding** any patient or personally identifiable information (PII / PHI) handled by this tool, including ensuring it is shared only with authorized parties through secure channels.
>
> The software is provided **"as is"**, without warranty of any kind, express or implied. The developer assumes **no liability** for any data loss, file corruption, information disclosure, regulatory non-compliance (including but not limited to HIPAA), financial loss, or any other damages arising from the use, misuse, or inability to use this software, whether caused by user error, system failure, or any other cause.
>
> By using this software, you acknowledge that you have read, understood, and accepted this disclaimer.

A desktop tool for parsing nurse home-visit `.ics` calendar files, computing visit amounts, flagging OASIS requirements, detecting duplicates, and exporting to Excel.

📖 [中文文档 / Chinese version](README.zh-CN.md)

> ⚠️ **Known Issue (under investigation)**
> During testing, the packaged `.app` may fail to launch on Macs with a different chip architecture than the one it was built on (e.g. an Intel Mac trying to run an ARM build, or vice versa). We're still working on a reliable cross-architecture build. If you hit this, please open an [Issue](https://github.com/1229391595max-oss/nsapp/issues) or join the [Discussion](https://github.com/1229391595max-oss/nsapp/discussions) and tell us your Mac model + macOS version + the exact error message.
> Workaround: use the single-file HTML version (`ics_visits_manager.html`) or run from source (`python nurseapp_qt.py`) — both avoid the packaged-app chip issue.

## Features

- **Parse `.ics` files** — extract every `VEVENT` (date, time, patient, visit type, location)
- **Auto-pricing** — apply per-type rates (user-customizable, default $x per type)
- **Rule backup / transfer** — export and import visit-type settings as a local `.json` file
- **OASIS detection** — auto-flag SOC / DC / RECERT / ROC visits as OASIS-required
- **Duplicate & cancellation report** — same UID multiple versions, `STATUS:CANCELLED` events, missing UIDs, same-patient-same-time clashes
- **Date range filtering** — today / this week / this month / last month / all / custom range
- **Excel export** — Clean (with amounts + color-coded OASIS column), Raw (full original fields), Bundle (both sheets), and a separate Duplicate Report
- **Native macOS GUI** built with PyQt6 — no browser required
- **Single-file HTML version** — fully offline, no install, no Python, and no macOS app-signing / chip-architecture issue

## Available versions

| File | UI | Notes |
|------|----|----|
| `nurseapp_qt.py` | Native PyQt6 window | Recommended — packageable into a `.app` |
| `ics_visits_manager.html` | Browser-based local file | Easiest for non-technical users — double-click to open, fully offline |
| `nurseapp.py` | Streamlit (browser) | Original prototype, still works |

## Use the HTML version

Download or copy `ics_visits_manager.html`, then double-click it. It opens in your browser, but all parsing and Excel generation happen locally on your computer.

This version is the easiest option for sharing because it does not require Python, PyInstaller, app signing, Terminal commands, or separate Intel / Apple Silicon builds.

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

The app is unsigned, so on first launch macOS will block it. Two scenarios:

**Scenario A — "Unidentified developer" warning**
Right-click the `.app` → **Open** → click **Open** in the dialog. After the first time, double-click works normally.

**Scenario B — `"…" is damaged and can't be opened. You should move it to the Trash.`**
This is a macOS **false positive** caused by the browser's quarantine flag (it shows up when downloading unsigned apps through Chrome / Safari / etc.). The file is **not** actually corrupted.

Fix: open **Terminal** and run the command below (replace the path with the actual location of your `.app`):

```bash
xattr -cr "/path/to/ICS Visits 管理.app"
```

Common examples:

```bash
# If the .app is in Downloads
xattr -cr ~/Downloads/"ICS Visits 管理.app"

# If it's on the Desktop
xattr -cr ~/Desktop/"ICS Visits 管理.app"
```

After running it, double-click works normally.

> 💡 **Tip for distributors:** to avoid this for your users entirely, ship the `.app` inside a `.dmg` or `.zip` that you build *locally* (not downloaded by the recipient through a browser), or sign + notarize it with an Apple Developer account.

## Project structure

```
.
├── nurseapp_qt.py     # Main PyQt6 app
├── ics_visits_manager.html # Single-file offline HTML app
├── nurseapp.py        # Streamlit version (legacy)
├── run_app.py         # Streamlit launcher
├── build_app.sh       # macOS .app build script
├── rthook_qt.py       # PyInstaller runtime hook for Qt plugin path
└── README.md
```

## Visit type mapping

| Type    | Amount | OASIS required |
|---------|--------|----------------|
| SOC     | $x     | YES            |
| DC      | $x     | YES            |
| RECERT  | $x     | YES            |
| ROC     | $x     | YES            |
| IV      | $x     | NO             |
| FU      | $x     | NO             |

Amounts are **fully customizable** per user via the in-app ⚙️ Visit Types editor. Each user can define their own codes, matching keywords, prices, and OASIS flags; settings are saved locally and persist across sessions.

Users can also export their visit-type settings to `ics-visits-settings.json` and import the file later on another browser or computer.

The visit type is parsed from the calendar event summary in `Patient Name - Visit Type` format. Variants like `Follow Up`, `F/U`, `FU` all normalize to `FU`.

## Notes

- Cancelled events (`STATUS:CANCELLED` or summaries starting with `cancel`) are excluded from Clean output by default but kept in Raw — toggle the "include cancelled" checkbox to include them.
- Deduplication: same `UID + RECURRENCE-ID + DTSTART` keeps only the latest by `SEQUENCE` / `LAST-MODIFIED`.
- Conda's PyQt6 ships a Qt5-linked `libqcocoa.dylib` by mistake; the build script overwrites it with the correct Qt6 version after PyInstaller runs.
