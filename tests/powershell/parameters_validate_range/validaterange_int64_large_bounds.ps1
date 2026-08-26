# vybe-test: powershell/parameters_validate_range/validaterange_int64_large_bounds
function Set-Quota {
    param([ValidateRange(1000000000, 9000000000)][int64]$Bytes)
    return $Bytes
}
$res = Set-Quota -Bytes 5000000000
if ($res -ne 5000000000) {
    Write-Host "FAIL: ValidateRange int64 bounds failed"
    exit 1
}
Write-Host "PASS"
exit 0
