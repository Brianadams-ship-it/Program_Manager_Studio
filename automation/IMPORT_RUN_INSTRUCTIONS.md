# Automation Import — Run Instructions

Prerequisites
- Microsoft Excel installed on the machine.
- Excel Trust Center: enable **Trust access to the VBA project object model** (File → Options → Trust Center → Trust Center Settings → Macro Settings).
- Pause OneDrive (or move files out of OneDrive) to avoid cloud-lock/save conflicts.
- Close all Excel windows before running the import scripts.
- Run PowerShell as your normal user. If script execution policy blocks running the provided scripts, you can run PowerShell with a temporary bypass:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Quick commands (run from repository root)

- Run the full import orchestration (processes all `.xls*` files and appends into the run log):

```powershell
Set-Location -Path .\automation
.\run_all_imports.ps1
```

- Retry only the previously failed imports (parses `import_run_full.txt` and retries):

```powershell
Set-Location -Path .\automation
.\retry_failed_imports.ps1
```

- Import a specific workbook (single-file run):

```powershell
Set-Location -Path .\automation
.\import_vba_runner.ps1 -Workbooks @('C:\path\to\your\file.xlsx') -AutoRun $true
```

Where to check results
- `automation/import_run_full.txt` — consolidated log of import attempts (success/failure details).
- The converted macro-enabled files will be saved as `.xlsm` alongside the originals when conversion is necessary.

Troubleshooting
- If you see `Microsoft Excel cannot access the file` or SaveAs/Resolve-Path errors:
  - Ensure OneDrive is paused and the file isn't opened by another process.
  - Ensure Excel is closed before running the scripts.
  - If the file is read-only or blocked, check file properties and unblock if necessary.
- If the script fails with PowerShell parse errors, confirm the repository files are intact and that no Markdown fences were left inside `.ps1` files.

Next steps
1. Pause OneDrive and close Excel.
2. Re-run `retry_failed_imports.ps1`.
3. If failures persist, open `automation/import_run_full.txt` and share the failing file paths; I can parse and suggest per-file remediation or re-run the importer for specific paths.

If you want, I can run the retry step now (I will attempt it and report results). If you prefer to pause OneDrive first, let me know and I'll wait.