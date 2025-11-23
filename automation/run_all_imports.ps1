# Run import_vba_runner.ps1 across all .xls* files under the repo root and log results
$repoRoot = (Resolve-Path "..\").Path
$log = Join-Path $PSScriptRoot 'import_run_full.txt'
if (Test-Path $log) { Remove-Item $log -Force }

Write-Output "Repository root: $repoRoot" | Tee-Object -FilePath $log
$files = Get-ChildItem -Path $repoRoot -Filter '*.xls*' -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notmatch '\\node_modules\\' }
Write-Output "Found $($files.Count) Excel files" | Tee-Object -FilePath $log -Append

foreach ($f in $files) {
    $path = $f.FullName
    Write-Output "--- Processing: $path ---" | Tee-Object -FilePath $log -Append
    try {
        & "$PSScriptRoot\import_vba_runner.ps1" -Workbooks @($path) -AutoRun $true 2>&1 | Tee-Object -FilePath $log -Append
    } catch {
        $msg = $_.Exception.Message
        Write-Output ("Runner error for {0}: {1}" -f $path, $msg) | Tee-Object -FilePath $log -Append
    }
}

Write-Output 'Run complete' | Tee-Object -FilePath $log -Append
