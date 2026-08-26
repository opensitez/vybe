# vybe-test: powershell/math_clamp_and_range_boundaries/clamp_uint64_type
[uint64]$u = 500
$val = [math]::Clamp($u, [uint64]100, [uint64]300)
if ($val -ne 300) {
    Write-Host "FAIL: Clamp uint64 failed, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
