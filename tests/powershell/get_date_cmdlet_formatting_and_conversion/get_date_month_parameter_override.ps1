# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_month_parameter_override
# Passing -Month overrides the month while preserving year, day, and time components
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$altered = Get-Date -Date $date -Month 11

if ($altered.Month -ne 11) {
    Write-Host "FAIL: Month was not updated to 11, got: $($altered.Month)"
    exit 1
}

if ($altered.Year -ne 2026 -or $altered.Day -ne 10) {
    Write-Host "FAIL: non-month components were unexpectedly altered"
    exit 1
}

Write-Host "PASS"
exit 0
