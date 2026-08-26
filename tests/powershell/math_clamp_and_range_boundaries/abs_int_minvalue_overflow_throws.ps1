# vybe-test: powershell/math_clamp_and_range_boundaries/abs_int_minvalue_overflow_throws
$caught = $false
try {
    $x = [math]::Abs([int]::MinValue)
} catch [System.OverflowException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected OverflowException on Abs(Int32.MinValue)"
    exit 1
}
Write-Host "PASS"
exit 0
