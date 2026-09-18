# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_uformat_short_date_representation
# Get-Date -UFormat '%D' formats the date as %m/%d/%y
$date = [DateTime]::Parse("2026-05-10")
$shortDate = Get-Date -Date $date -UFormat "%D"

if ($shortDate -ne "05/10/26") {
    Write-Host "FAIL: UFormat %D mismatch, expected '05/10/26', got: '$shortDate'"
    exit 1
}

Write-Host "PASS"
exit 0
