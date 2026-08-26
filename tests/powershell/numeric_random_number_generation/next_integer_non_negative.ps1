# vybe-test: powershell/numeric_random_number_generation/next_integer_non_negative
$rnd = [System.Random]::new()
$val = $rnd.Next()
if ($val -lt 0) {
    Write-Host "FAIL: Next() must be non-negative, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
