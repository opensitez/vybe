# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_uformat_full_weekday_name
# Get-Date -UFormat '%A' produces the full locale weekday name
$date = [DateTime]::Parse("2026-05-10")
$weekday = Get-Date -Date $date -UFormat "%A"

if ($weekday -ne "Sunday") {
    Write-Host "FAIL: weekday name mismatch, expected 'Sunday', got: '$weekday'"
    exit 1
}

Write-Host "PASS"
exit 0
