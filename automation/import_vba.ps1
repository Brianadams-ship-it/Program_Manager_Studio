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
