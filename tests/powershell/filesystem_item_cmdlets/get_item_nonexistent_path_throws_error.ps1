# vybe-test: powershell/filesystem_item_cmdlets/get_item_nonexistent_path_throws_error
# Get-Item on a path that does not exist throws a terminating error with -ErrorAction Stop
$fakePath = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_gi_nonexistent_$PID"
$threw = $false

try {
    Get-Item $fakePath -ErrorAction Stop
} catch {
    $threw = $true
}

if (-not $threw) {
    Write-Host "FAIL: Get-Item on non-existent path did not throw"
    exit 1
}

Write-Host "PASS"
exit 0
