# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_day_parameter_override
# Passing -Day overrides the day while preserving year, month, and time components
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$altered = Get-Date -Date $date -Day 28

if ($altered.Day -ne 28) {
    Write-Host "FAIL: Day was not updated to 28, got: $($altered.Day)"
    exit 1
}

if ($altered.Month -ne 5 -or $altered.Year -ne 2026) {
    Write-Host "FAIL: non-day components were unexpectedly altered"
    exit 1
}

Write-Host "PASS"
exit 0
