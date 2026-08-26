# vybe-test: powershell/parameters_validate_set/validateset_case_insensitive_match
function Get-Fruit {
    param([ValidateSet("Apple", "Banana")][string]$Item)
    return $Item
}
$res = Get-Fruit -Item "apple"
if ($res -ne "apple" -and $res -ne "Apple") {
    Write-Host "FAIL: ValidateSet case-insensitive argument failed"
    exit 1
}
Write-Host "PASS"
exit 0
