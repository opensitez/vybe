# vybe-test: powershell/math_clamp_and_range_boundaries/clamp_min_greater_than_max_throws_argument_exception
$caught = $false
try {
    $x = [math]::Clamp(5, 10, 2)
} catch [System.ArgumentException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected ArgumentException when min > max"
    exit 1
}
Write-Host "PASS"
exit 0
