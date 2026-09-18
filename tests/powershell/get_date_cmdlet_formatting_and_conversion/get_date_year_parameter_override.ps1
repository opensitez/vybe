# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_year_parameter_override
# Passing -Year overrides the year while preserving the other date and time components
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$altered = Get-Date -Date $date -Year 2035

if ($altered.Year -ne 2035) {
    Write-Host "FAIL: Year was not updated to 2035, got: $($altered.Year)"
    exit 1
}

if ($altered.Month -ne 5 -or $altered.Day -ne 10 -or $altered.Hour -ne 14) {
    Write-Host "FAIL: non-year components were unexpectedly altered"
    exit 1
}

Write-Host "PASS"
exit 0
