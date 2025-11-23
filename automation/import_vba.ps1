$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$automationModule = Join-Path $scriptDir "OptimumUpgrade_Automation.bas"

param(
    [string[]]$Workbooks = @("sample_project\FalconEye_Compliance_Matrix.xlsx")
)

if (!(Test-Path $automationModule)) {
    Write-Error "Automation module not found: $automationModule"
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($wbPathRel in $Workbooks) {
    $wbFullPath = Join-Path $scriptDir $wbPathRel
    if (!(Test-Path $wbFullPath)) {
        Write-Warning "Workbook not found: $wbFullPath"
        continue
    }

    Write-Host "Opening workbook: $wbFullPath"
    $wb = $excel.Workbooks.Open($wbFullPath)

    # Save as macro-enabled if needed
    if ($wb.FileFormat -ne 52) {  # 52 = xlOpenXMLWorkbookMacroEnabled
        $xlsmPath = [System.IO.Path]::ChangeExtension($wbFullPath, ".xlsm")
        Write-Host "Converting to macro-enabled workbook: $xlsmPath"
        $wb.SaveAs($xlsmPath, 52)
        $wb.Close($true)
        $wb = $excel.Workbooks.Open($xlsmPath)
        $wbFullPath = $xlsmPath
    }

    try {
        $vbProj = $wb.VBProject
        Write-Host "Importing VBA module into: $wbFullPath"
        $vbProj.VBComponents.Import($automationModule) | Out-Null
        $wb.Save()
    } catch {
        Write-Error "Failed to import VBA module into $wbFullPath. Check 'Trust access to VBA project' setting in Excel."
    } finally {
        $wb.Close($true)
    }
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$automationModule = Join-Path $scriptDir "OptimumUpgrade_Automation.bas"

<#
This script imports the `OptimumUpgrade_Automation.bas` module into one or more Excel workbooks.

Usage:
  - Edit the `$Workbooks` array below (paths are relative to this `automation/` folder).
  - Ensure Excel Trust Center has "Trust access to the VBA project object model" enabled.
  - Run from PowerShell: `.\import_vba.ps1`
#>

param(
    [string[]]$Workbooks = @(
        '..\OptimumUpgrade_ALL_v15_full_package\sample_project\FalconEye_Compliance_Matrix.xlsx'
    )
)

# By default we will insert a Workbook_Open handler that calls Install_Automation
[switch]$AutoRun = $true

if (!(Test-Path $automationModule)) {
    Write-Error "Automation module not found: $automationModule"
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($wbPathRel in $Workbooks) {
    $wbFullPath = Join-Path $scriptDir $wbPathRel
    if (!(Test-Path $wbFullPath)) {
        Write-Warning "Workbook not found: $wbFullPath"
        continue
    }

    Write-Host "Opening workbook: $wbFullPath"
    $wb = $excel.Workbooks.Open((Resolve-Path $wbFullPath).Path)

    # Save as macro-enabled if needed
    if ($wb.FileFormat -ne 52) {  # 52 = xlOpenXMLWorkbookMacroEnabled
        $xlsmPath = [System.IO.Path]::ChangeExtension($wbFullPath, ".xlsm")
        Write-Host "Converting to macro-enabled workbook: $xlsmPath"
        $wb.SaveAs((Resolve-Path $xlsmPath).Path, 52)
        $wb.Close($true)
        $wb = $excel.Workbooks.Open((Resolve-Path $xlsmPath).Path)
        $wbFullPath = $xlsmPath
    }

    try {
        $vbProj = $wb.VBProject
        Write-Host "Importing VBA module into: $wbFullPath"
        $vbProj.VBComponents.Import((Resolve-Path $automationModule).Path) | Out-Null

        if ($AutoRun) {
            try {
                $thisComp = $vbProj.VBComponents.Item("ThisWorkbook")
            } catch {
                $thisComp = $null
            }

            if ($thisComp -ne $null) {
                $codeMod = $thisComp.CodeModule
                $lines = 0
                try { $lines = $codeMod.CountOfLines } catch { $lines = 0 }

                $hasOpen = $false
                if ($lines -gt 0) {
                    $fullText = $codeMod.Lines(1, $lines)
                    if ($fullText -match 'Sub\s+Workbook_Open' -or $fullText -match 'Sub\s+Auto_Open') { $hasOpen = $true }
                }

                if (-not $hasOpen) {
                    $insertAt = $lines + 1
                    $codeToInsert = "Private Sub Workbook_Open()`r`n    On Error Resume Next`r`n    Call Install_Automation`r`nEnd Sub"
                    $codeMod.InsertLines($insertAt, $codeToInsert)
                    Write-Host "Inserted Workbook_Open into ThisWorkbook to call Install_Automation"
                } else {
                    Write-Host "Workbook_Open or Auto_Open already present; skipping insertion"
                }
            } else {
                Write-Warning "ThisWorkbook VBComponent not found; cannot insert Workbook_Open handler."
            }
        }

        $wb.Save()
        Write-Host "Import succeeded: $wbFullPath"
    } catch {
        Write-Error "Failed to import VBA module into $wbFullPath. Check 'Trust access to VBA project' setting in Excel. Error: $_"
    } finally {
        $wb.Close($true)
    }
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$automationModule = Join-Path $scriptDir "OptimumUpgrade_Automation.bas"

param(
    [string[]]$Workbooks = @("sample_project\FalconEye_Compliance_Matrix.xlsx")
)

if (!(Test-Path $automationModule)) {
    Write-Error "Automation module not found: $automationModule"
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($wbPathRel in $Workbooks) {
    $wbFullPath = Join-Path $scriptDir $wbPathRel
    if (!(Test-Path $wbFullPath)) {
        Write-Warning "Workbook not found: $wbFullPath"
        continue
    }

    Write-Host "Opening workbook: $wbFullPath"
    $wb = $excel.Workbooks.Open($wbFullPath)

    # Save as macro-enabled if needed
    if ($wb.FileFormat -ne 52) {  # 52 = xlOpenXMLWorkbookMacroEnabled
        $xlsmPath = [System.IO.Path]::ChangeExtension($wbFullPath, ".xlsm")
        Write-Host "Converting to macro-enabled workbook: $xlsmPath"
        $wb.SaveAs($xlsmPath, 52)
        $wb.Close($true)
        $wb = $excel.Workbooks.Open($xlsmPath)
        $wbFullPath = $xlsmPath
    }

    try {
        $vbProj = $wb.VBProject
        Write-Host "Importing VBA module into: $wbFullPath"
        $vbProj.VBComponents.Import($automationModule) | Out-Null
        $wb.Save()
    } catch {
        Write-Error "Failed to import VBA module into $wbFullPath. Check 'Trust access to VBA project' setting in Excel."
    } finally {
        $wb.Close($true)
    }
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
