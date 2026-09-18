# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_timespan
# System.TimeSpan instances serialize with <TS> tags preserving total duration
$tmp = [System.IO.Path]::GetTempFileName()
$orig = [TimeSpan]::FromSeconds(7325) # 2 hours, 2 minutes, 5 seconds

$orig | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored.TotalSeconds -ne 7325) {
    Write-Host "FAIL: TimeSpan TotalSeconds mismatch, got: $($restored.TotalSeconds)"
    exit 1
}

if ($restored.Hours -ne 2 -or $restored.Minutes -ne 2 -or $restored.Seconds -ne 5) {
    Write-Host "FAIL: TimeSpan duration components mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
