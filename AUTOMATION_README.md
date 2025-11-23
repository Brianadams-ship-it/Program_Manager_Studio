'Optimum Upgrade — Automation README'

This folder contains automation helper scripts and a skeleton VBA module to assist with importing and managing automation macros.

Files added by automation draft:
- `import_vba.ps1`: PowerShell script to import `.bas` files into an Excel workbook's VBProject.
- `OptimumUpgrade_Automation.bas`: Skeleton VBA module with entry points for test-case generation, verification summary, and ProjectInfo sync.

Quick steps to use:
1. Place your `.bas` files in a folder, e.g. `vba/` inside the repo.
2. Ensure the target workbook is an `.xlsm` (macro-enabled) file and backed up.
3. In Excel: File → Options → Trust Center → Trust Center Settings… → Macro Settings → enable "Trust access to the VBA project object model".
4. Run the importer from PowerShell (run as a user who can start Excel):

```powershell
Set-Location "C:\Users\badams\OneDrive - Intellisense Systems\Desktop\repo\Program_Manager_Studio"
.\import_vba.ps1 -SourceFolder .\vba -TargetWorkbook .\Automation_ControlPanel_Template.xlsm -Overwrite
```

Notes and caveats:
- Programmatic access to the VBProject is controlled by Excel security; if not enabled the script will fail.
- The script automates Excel via COM; Excel may prompt or require an interactive session depending on policy.
- Always keep a backup of your workbook before importing modules.

Next steps I can take for you:
- Review actual `.bas` modules (if you upload/unzip them) and integrate into a unified module set.
- Extend the skeleton VBA with real generation logic based on your templates/workbook structure.
