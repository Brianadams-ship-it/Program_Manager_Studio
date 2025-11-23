'SYNOPSIS: Imports VBA modules (.bas) from a folder into an Excel workbook's VBProject.
'
'NOTES: Requires Excel installed and 'Trust access to the VBA project object model' enabled
'in Excel Trust Center (File → Options → Trust Center → Trust Center Settings… → Macro Settings).
'
'Usage examples:
'  .\import_vba.ps1 -SourceFolder .\vba -TargetWorkbook .\Automation_ControlPanel_Template.xlsm
'  .\import_vba.ps1 -SourceFolder .\ -TargetWorkbook .\Automation_ControlPanel_Template.xlsm -Overwrite

param(
    [Parameter(Mandatory=$true)][string]$SourceFolder,
    [Parameter(Mandatory=$true)][string]$TargetWorkbook,
    [switch]$Overwrite
)

function Import-BasModule {
    param(
        [string]$WorkbookPath,
        [string]$ModulePath,
        [switch]$Overwrite
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $wb = $excel.Workbooks.Open((Resolve-Path $WorkbookPath).Path)
        $vbProj = $wb.VBProject

        # Module name from file name
        $modName = [System.IO.Path]::GetFileNameWithoutExtension($ModulePath)

        # Remove existing module if overwrite requested
        if ($Overwrite) {
            foreach ($comp in $vbProj.VBComponents) {
                if ($comp.Name -eq $modName) {
                    $vbProj.VBComponents.Remove($comp) | Out-Null
                    break
                }
            }
        }

        Write-Host "Importing $ModulePath as $modName into $WorkbookPath"
        $vbProj.VBComponents.Import((Resolve-Path $ModulePath).Path) | Out-Null

        # Save workbook
        $wb.Save()
    }
    catch {
        Write-Error "Error importing $ModulePath: $_"
        throw
    }
    finally {
        if ($wb) { $wb.Close($true) }
        $excel.Quit()
        if ($vbProj -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbProj) | Out-Null }
        if ($wb -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
        if ($excel -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

if (-not (Test-Path $SourceFolder)) { Write-Error "Source folder not found: $SourceFolder"; exit 2 }
if (-not (Test-Path $TargetWorkbook)) { Write-Error "Target workbook not found: $TargetWorkbook"; exit 2 }

# warn about VBProject access
Write-Host "WARNING: Ensure Excel Trust Center → 'Trust access to the VBA project object model' is enabled." -ForegroundColor Yellow

$files = Get-ChildItem -Path $SourceFolder -Filter *.bas -File -Recurse
if ($files.Count -eq 0) { Write-Host "No .bas files found in $SourceFolder"; exit 0 }

foreach ($f in $files) {
    Import-BasModule -WorkbookPath $TargetWorkbook -ModulePath $f.FullName -Overwrite:$Overwrite
}

Write-Host "Import complete. Open the workbook in Excel and verify macros." -ForegroundColor Green
