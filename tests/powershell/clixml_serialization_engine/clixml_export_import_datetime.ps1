# vybe-test: powershell/clixml_serialization_engine/clixml_export_import_datetime
# DateTime instances serialize with <DT> tags preserving date and time components
$tmp = [System.IO.Path]::GetTempFileName()
$orig = [DateTime]::Parse("2026-11-25 18:45:30")

$orig | Export-Clixml -Path $tmp
$restored = Import-Clixml -Path $tmp
Remove-Item -Force $tmp

if ($restored.Year -ne 2026 -or $restored.Month -ne 11 -or $restored.Day -ne 25) {
    Write-Host "FAIL: date components mismatch, got: $($restored.ToString('yyyy-MM-dd'))"
    exit 1
}

if ($restored.Hour -ne 18 -or $restored.Minute -ne 45 -or $restored.Second -ne 30) {
    Write-Host "FAIL: time components mismatch, got: $($restored.ToString('HH:mm:ss'))"
    exit 1
}

Write-Host "PASS"
exit 0
