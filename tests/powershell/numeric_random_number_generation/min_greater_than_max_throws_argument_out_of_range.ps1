# vybe-test: powershell/numeric_random_number_generation/min_greater_than_max_throws_argument_out_of_range
$rnd = [System.Random]::new()
$caught = $false
try {
    $x = $rnd.Next(10, 5)
} catch [System.ArgumentOutOfRangeException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected ArgumentOutOfRangeException when min > max"
    exit 1
}
Write-Host "PASS"
exit 0
