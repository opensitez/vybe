# vybe-test: powershell/string_substitution_rules/substitute_cmdlet_property
$date = Get-Date
$result = "$($date.DayOfWeek)"
if ([string]::IsNullOrWhiteSpace($result)) {
    Write-Host 'FAIL: expected day of week text from cmdlet object property'
    exit 1
}

Write-Host 'PASS'
exit 0
