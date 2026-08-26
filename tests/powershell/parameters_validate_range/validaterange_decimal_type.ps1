# vybe-test: powershell/parameters_validate_range/validaterange_decimal_type
function Test-DecRange {
    param([ValidateRange(0, 100)][decimal]$Amount)
    return $Amount
}
$res = Test-DecRange -Amount ([decimal]50.5)
if ($res -ne [decimal]50.5) {
    Write-Host "FAIL: ValidateRange decimal type failed"
    exit 1
}
Write-Host "PASS"
exit 0
