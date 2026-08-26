# vybe-test: powershell/parameters_validate_range/validaterange_double_precision_bounds
function Set-Rate {
    param([ValidateRange(0.0, 1.0)][double]$Rate)
    return $Rate
}
$res = Set-Rate -Rate 0.75
if ($res -ne 0.75) {
    Write-Host "FAIL: ValidateRange double failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
