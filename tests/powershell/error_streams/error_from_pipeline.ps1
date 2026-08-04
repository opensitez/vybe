# vybe-test: powershell/error_streams/error_from_pipeline
$ErrorActionPreference = 'Continue'
Get-Item nonexistent | Out-Null
if ($Error.Count -lt 1) {
    Write-Host "FAIL: expected pipeline error"
    exit 1
}
Write-Host 'PASS'
exit 0
