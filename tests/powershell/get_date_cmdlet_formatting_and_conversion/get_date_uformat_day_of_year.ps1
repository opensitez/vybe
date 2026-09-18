# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_uformat_day_of_year
# Get-Date -UFormat '%j' yields a 3-digit padded day-of-year integer string (001-366)
$date = [DateTime]::Parse("2026-05-10")
$dayOfYear = Get-Date -Date $date -UFormat "%j"

# May 10 in 2026 is day: 31 (Jan) + 28 (Feb) + 31 (Mar) + 30 (Apr) + 10 = 130
if ($dayOfYear -ne "130") {
    Write-Host "FAIL: day of year mismatch, expected '130', got: '$dayOfYear'"
    exit 1
}

Write-Host "PASS"
exit 0
