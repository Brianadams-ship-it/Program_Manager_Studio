# Parse import_run_full.txt for failed files and retry them using import_vba_runner.ps1
$log = Join-Path $PSScriptRoot 'import_run_full.txt'
if (-not (Test-Path $log)) { Write-Error "Log not found: $log"; exit 1 }

$content = Get-Content $log
$failed = @()

foreach ($line in $content) {
    if ($line -match '^Runner error for (.+?): Microsoft Excel') {
        $failed += $matches[1]
    }
}

$failed = $failed | Select-Object -Unique
if ($failed.Count -eq 0) { Write-Output "No failed files found in log."; exit 0 }

Write-Output "Retrying $($failed.Count) failed files..."
foreach ($f in $failed) {
    Write-Output "--- Retrying: $f ---" | Tee-Object -FilePath $log -Append
    try {
        & "$PSScriptRoot\import_vba_runner.ps1" -Workbooks @($f) -AutoRun $true 2>&1 | Tee-Object -FilePath $log -Append
    } catch {
        Write-Output ("Retry runner error for {0}: {1}" -f $f, $_.Exception.Message) | Tee-Object -FilePath $log -Append
    }
}

Write-Output 'Retry pass complete' | Tee-Object -FilePath $log -Append
Get-Content $log -Tail 200 | Write-Output
