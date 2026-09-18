# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_uformat_unix_epoch_seconds
# Get-Date -UFormat '%s' calculates the total seconds elapsed since the Unix epoch (1970-01-01T00:00:00Z)
$epochStart = [DateTime]::Parse("1970-01-01 00:00:00Z")
$epochString = Get-Date -Date $epochStart -UFormat "%s"

if ($epochString -ne "0") {
    Write-Host "FAIL: expected epoch timestamp '0', got: '$epochString'"
    exit 1
}

$epoch2026 = Get-Date -Date ([DateTime]::Parse("2026-01-01 00:00:00Z")) -UFormat "%s"
if ($epoch2026 -ne "1767225600") {
    Write-Host "FAIL: expected 2026 epoch timestamp '1767225600', got: '$epoch2026'"
    exit 1
}

Write-Host "PASS"
exit 0
