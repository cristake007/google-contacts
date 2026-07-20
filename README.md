# Google Contact Finder

## First-time setup

Python 3.10 or newer is required. From PowerShell in this repository, run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\bootstrap.ps1
```

The bootstrap creates `.venv`, installs the pinned Python packages, and installs
the matching Playwright Chromium build. These generated files are not committed.

## Run the finder

Start or restart from the original workbook, processing 20 companies:

```powershell
& ".\.venv\Scripts\python.exe" .\google_contact_finder_v4.py .\all_cleaned.xlsx --limit 20
```

Resume from `all_cleaned_google_contacts_v4.xlsx` and skip completed rows:

```powershell
& ".\.venv\Scripts\python.exe" .\google_contact_finder_v4.py .\all_cleaned.xlsx --resume --limit 20
```

Resume and retry rows previously marked as blocked, failed, or requiring review:

```powershell
& ".\.venv\Scripts\python.exe" .\google_contact_finder_v4.py .\all_cleaned.xlsx --resume --retry-failed --limit 20
```

Process one company by CUI:

```powershell
& ".\.venv\Scripts\python.exe" .\google_contact_finder_v4.py .\all_cleaned.xlsx --only-cui 12345678
```

Use `--limit 0` to process all remaining rows. Press `Ctrl+C` to stop safely;
the workbook is saved before the browser closes.
