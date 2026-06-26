# GSTR 2B Reconciliation — Desktop App

A standalone Windows desktop application for GSTR 2B reconciliation. All
processing (Stage 1 cleaning, Stage 2 merge, Stage 3 reconciliation) is done
in Python — the window you see is just the interface. Nothing is sent over the
internet; everything runs on your machine.

## What's in this folder

| File | Purpose |
|------|---------|
| `app.py` | The desktop window host (pywebview). Bridges the UI to Python. |
| `engine.py` | All the reconciliation logic. The single source of truth. |
| `ui.html` | The interface (upload zones, the three panels). |
| `requirements.txt` | Python dependencies. |
| `build.bat` | One-click builder that produces the `.exe`. |
| `README.md` | This file. |

## Building the .exe (one time, on Windows)

The executable must be built on a Windows machine — that's just how Windows
executables work.

### If you don't have Python yet
1. Download Python 3.10 or newer from <https://www.python.org/downloads/>.
2. Run the installer and **tick "Add python.exe to PATH"** on the first screen.
3. Finish the install.

### Build
1. Put all the files in this folder together on your Windows PC.
2. Double-click **`build.bat`**.
3. Wait a few minutes while it installs dependencies and compiles.
4. When it finishes, your app is at **`dist\GSTR2BRecon.exe`**.

That single `.exe` carries all its dependencies inside it. You can copy it
anywhere — Desktop, a USB stick, another PC — and it runs on its own.

## Using the app

1. **Double-click `GSTR2BRecon.exe`.** A window opens.
2. **Upload files** — drag your `2B Raw` and `Tally Raw` `.xlsx` files into the
   two drop zones, set the month (e.g. `2026-Jan`), and click *Proceed*.
3. **Run processing** — work down the three steps:
   - *Stage 1* cleans both raw files.
   - *Stage 2* optionally merges previous-month Carry Forward files (drop them
     in if you have them; skip if you don't).
   - *Stage 3* reconciles. You can optionally add a GST Exclusion List.
4. **Save outputs** — after Stage 3, choose where to save (or use the default)
   and click *Save all outputs*. Use *Open folder* to view them.

## Where outputs go

By default the app saves to:

```
Documents\GSTR2B Outputs\<month>\
```

You can change this any time with the **Change location…** button before saving.
Each run creates a sub-folder named after the month, keeping months organised:

```
GSTR2B Outputs\
└── 2026-Jan\
    ├── GSTR 2B Summary 2026-Jan.xlsx      (Reconciled + Not to be Claimed tabs)
    ├── 2B Carry Forward 2026-Jan.xlsx
    ├── Tally Carry Forward 2026-Jan.xlsx
    ├── 2B Latest 2026-Jan.xlsx
    └── Tally Latest 2026-Jan.xlsx
```

## Notes

- The matching rules (exact date, fuzzy invoice number with a numeric-exact
  guard, ±₹1 amount tolerance, per-component GST checks) and the near-miss
  remarks all live in `engine.py`. Change them there and rebuild.
- The app needs the Microsoft **WebView2 runtime**, which is already present on
  Windows 10/11. If a machine lacks it, install the Evergreen runtime from
  Microsoft (free): search "WebView2 Runtime download".
