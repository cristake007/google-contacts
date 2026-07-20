# Google Contact Finder

## First-time setup

Python 3.10 or newer is required. From PowerShell in this repository, run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\bootstrap.ps1
```

The bootstrap creates and activates `.venv`, installs the pinned Python packages,
and installs the matching Playwright Chromium build. These generated files are
not committed. Run the activation command again when you open a new PowerShell
session.

To install everything and immediately start a fresh 20-company run:

```powershell
.\bootstrap.ps1 -RunMode Fresh -Limit 20
```

`-RunMode` also accepts `Resume` and `RetryFailed`; omit it when you only want
to install and activate the environment.

To activate the virtual environment in the current PowerShell session:

```powershell
.\.venv\Scripts\Activate.ps1
```

Activation is optional when using `run_contacts.ps1`; the runner always uses the
Python executable inside `.venv` directly.

## Run the finder

Start or restart from the original workbook, processing 20 companies:

```powershell
.\run_contacts.ps1 -Mode Fresh -Limit 20
```

Resume from `all_cleaned_google_contacts_v4.xlsx` and skip completed rows:

```powershell
.\run_contacts.ps1 -Mode Resume -Limit 20
```

Resume and retry rows previously marked as blocked, failed, or requiring review:

```powershell
.\run_contacts.ps1 -Mode RetryFailed -Limit 20
```

Process one company by CUI:

```powershell
.\run_contacts.ps1 -Mode Resume -OnlyCui 12345678
```

Use `-Limit 0` to process all applicable rows. Extra finder arguments can be
placed after the runner options, for example:

```powershell
.\run_contacts.ps1 -Mode Resume -Limit 20 --no-manual-captcha
```

Press `Ctrl+C` to stop safely; the workbook is saved before the browser closes.
