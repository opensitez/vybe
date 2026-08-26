# vybe-test: powershell/math_rounding_midpoint_modes/negative_precision_arguments_throw
$caught = $false
try {
    $x = [math]::Round(123.45, -1)
} catch [System.ArgumentOutOfRangeException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected ArgumentOutOfRangeException for negative precision"
    exit 1
}
Write-Host "PASS"
exit 0
