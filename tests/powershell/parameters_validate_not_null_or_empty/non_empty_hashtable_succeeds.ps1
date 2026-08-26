# vybe-test: powershell/parameters_validate_not_null_or_empty/non_empty_hashtable_succeeds
function Set-ConfigHt2 {
    param([ValidateNotNullOrEmpty()][hashtable]$Config)
    return $Config.Count
}
$res = Set-ConfigHt2 -Config @{ key = "val" }
if ($res -ne 1) {
    Write-Host "FAIL: Non-empty hashtable failed ValidateNotNullOrEmpty"
    exit 1
}
Write-Host "PASS"
exit 0
