# vybe-test: powershell/math_rounding_midpoint_modes/round_with_decimal_places_and_to_even
$val = 1.245
$rounded = [math]::Round($val, 2, [System.MidpointRounding]::ToEven)
if ($rounded -ne 1.24 -and $rounded -ne 1.25) {
    Write-Host "FAIL: Round with decimals and ToEven failed, got $rounded"
    exit 1
}
Write-Host "PASS"
exit 0
