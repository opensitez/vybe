# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_time_components_override
# Passing -Hour, -Minute, and -Second overrides the time portions of the DateTime
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$altered = Get-Date -Date $date -Hour 10 -Minute 20 -Second 30

if ($altered.Hour -ne 10 -or $altered.Minute -ne 20 -or $altered.Second -ne 30) {
    Write-Host "FAIL: time components mismatch: $($altered.ToString('HH:mm:ss'))"
    exit 1
}

Write-Host "PASS"
exit 0
