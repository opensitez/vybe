# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_uformat_composite_datetime_string
# Get-Date -UFormat formats timestamps using standard strftime specifiers (%Y, %m, %d, %H, %M, %S)
$date = [DateTime]::Parse("2026-05-10 14:30:45")
$ufmt = Get-Date -Date $date -UFormat "%Y/%m/%d %H:%M:%S"

if ($ufmt -ne "2026/05/10 14:30:45") {
    Write-Host "FAIL: UFormat composite mismatch, got: '$ufmt'"
    exit 1
}

Write-Host "PASS"
exit 0
