# Optimum Upgrade – Automation Setup Guide

This folder contains the files that add real automation to the Optimum Upgrade Excel toolkits.

## Files

- **OptimumUpgrade_Automation.bas**  
  VBA module that contains:
  - GenerateTestCasesFromRequirements
  - BuildVerificationSummary
  - ExportVerificationSummaryCSV
  - EnsureProjectInfoSheet
  - UpdateSheetHeadersFromProjectInfo

- **import_vba.ps1**  
  PowerShell script that automatically imports the VBA module into one or more Excel workbooks and converts them to macro-enabled `.xlsm` files if needed.

- **Automation_ControlPanel_Template.xlsx**  
  Example Excel file with a `ControlPanel` sheet that shows how to present automation actions to users. You can copy this sheet into any workbook that has the VBA module imported.

---

## Requirements

- Windows with Microsoft Excel installed.
- In Excel, under **Trust Center → Macro Settings**:
  - Enable macros for the workbook you want to use.
  - Check **“Trust access to the VBA project object model”** so the import script can add the VBA module.
- PowerShell (installed by default on Windows 10+).

---

## How to Import Automation into a Workbook

1. Extract the entire ZIP (including the `automation` folder) to a local folder, e.g.:
   `C:\OptimumUpgrade_Site\`

2. Open PowerShell and change directory to the automation folder, for example:
   ```powershell
   cd C:\OptimumUpgrade_Site\automation
   ```

3. (Optional, first time only) Allow scripts to run in your user scope:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   ```

4. Open `import_vba.ps1` in a text editor and check the `Workbooks` parameter. For example:
   ```powershell
   param(
       [string[]]$Workbooks = @(
           "sample_project\FalconEye_Compliance_Matrix.xlsx"
       )
   )
   ```
   - Add or remove workbook paths as needed (relative to the root of the extracted folder).

5. Run the script:
   ```powershell
   .\import_vba.ps1
   ```

The script will:
- Open each listed workbook.
- Convert it to `.xlsm` (macro-enabled) if it is `.xlsx`.
- Import `OptimumUpgrade_Automation.bas` into its VBA project.
- Save and close the workbook.

You can then open the `.xlsm` file in Excel and run the macros.

---

## How to Use the Macros

Once the module is imported into a workbook, you will see these macros under **Developer → Macros**:

- `GenerateTestCasesFromRequirements`  
  - Looks at the `Requirements` sheet.  
  - For requirements whose verification method includes `Test` or `Demonstration` and do not have a Test Case ID, it:
    - Creates the next `TC-###` ID.
    - Writes it into the `TestCases` sheet with a default description and placeholders.

- `BuildVerificationSummary`  
  - Creates or refreshes a `VerificationSummary` sheet.  
  - Pulls Req ID, Requirement, Verification Method, Test Case ID, and Status from `Requirements`.

- `ExportVerificationSummaryCSV`  
  - Exports the `VerificationSummary` sheet to a CSV file so it can be imported into a Word report or other tools.

- `EnsureProjectInfoSheet`  
  - Creates a `ProjectInfo` sheet if it does not exist, with standard fields such as `ProjectName`, `ProjectAcronym`, `Customer`, `ContractNumber`, etc.

- `UpdateSheetHeadersFromProjectInfo`  
  - Reads values from `ProjectInfo` and updates each sheet's header (cells A1 and A2) so the project name, acronym, customer, and contract number are consistent across the workbook.

---

## Using the Automation Control Panel Template

The file `Automation_ControlPanel_Template.xlsx` contains a `ControlPanel` sheet with labels and structure for automation actions.

To integrate it with a specific workbook:

1. Import the VBA module using `import_vba.ps1` (as described above) so the macros exist in that workbook.
2. Open `Automation_ControlPanel_Template.xlsx` and your target workbook side by side.
3. Right-click the `ControlPanel` sheet tab in the template → **Move or Copy…** → select your target workbook → check **Create a copy** → **OK**.
4. In your target workbook, on the `ControlPanel` sheet:
   - Insert **Form Control Buttons** (Developer → Insert → Button (Form Control)).
   - Assign each button to one of the macros, for example:
     - “Generate Test Cases” → `GenerateTestCasesFromRequirements`
     - “Build Verification Summary” → `BuildVerificationSummary`
     - “Export Verification Summary CSV” → `ExportVerificationSummaryCSV`
     - “Create ProjectInfo Sheet” → `EnsureProjectInfoSheet`
     - “Update Headers From ProjectInfo” → `UpdateSheetHeadersFromProjectInfo`

Users can then click the buttons on the ControlPanel sheet to run automation without going into the Macro dialog.

---

## Recommended Flow for a New Project

1. Start from your base toolkit workbook (e.g., Compliance / RTM master).  
2. Run `EnsureProjectInfoSheet` and fill in Program Name, Acronym, Customer, etc.  
3. Enter or import requirements into the `Requirements` sheet.  
4. Run `GenerateTestCasesFromRequirements` to auto-build the `TestCases` sheet.  
5. Run `BuildVerificationSummary` to create the `VerificationSummary` sheet.  
6. (Optional) Run `ExportVerificationSummaryCSV` to feed a Word Verification Report.  
7. Use `UpdateSheetHeadersFromProjectInfo` whenever project metadata changes.

This gives you real automation on top of structured templates, and you can extend the module over time with additional macros.
