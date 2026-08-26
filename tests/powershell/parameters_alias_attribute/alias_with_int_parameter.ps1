# vybe-test: powershell/parameters_alias_attribute/alias_with_int_parameter
function Set-MaxCount {
    param([Alias("Limit", "Top")][int]$Count)
    return $Count
}
$r1 = Set-MaxCount -Limit 10
$r2 = Set-MaxCount -Top 20
if ($r1 -ne 10 -or $r2 -ne 20) {
    Write-Host "FAIL: Multiple numeric aliases failed"
    exit 1
}
Write-Host "PASS"
exit 0
