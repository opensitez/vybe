# vybe-test: powershell/comparison/comparison_on_arrays_filters
# When LHS is an array, -eq filters rather than returning bool
$arr = @(1, 2, 3, 2, 4, 2)
$twos = $arr -eq 2
if ($twos.Count -ne 3) {
    Write-Host "FAIL: expected 3 matches, got $($twos.Count)"
    exit 1
}
$big = $arr -gt 2
if ($big.Count -ne 2) {
    Write-Host "FAIL: expected 2 values > 2, got $($big.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
