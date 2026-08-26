# vybe-test: powershell/parameters_validate_set/validateset_integers_as_strings
function Select-Port {
    param([ValidateSet("80", "443", "8080")][string]$Port)
    return $Port
}
$res = Select-Port -Port "443"
if ($res -ne "443") {
    Write-Host "FAIL: String port in ValidateSet failed"
    exit 1
}
Write-Host "PASS"
exit 0
