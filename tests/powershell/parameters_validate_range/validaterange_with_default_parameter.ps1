# vybe-test: powershell/parameters_validate_range/validaterange_with_default_parameter
function Get-CountDefault {
    param([ValidateRange(1, 10)][int]$Count = 5)
    return $Count
}
$res = Get-CountDefault
if ($res -ne 5) {
    Write-Host "FAIL: ValidateRange with default value failed"
    exit 1
}
Write-Host "PASS"
exit 0
