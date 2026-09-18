# vybe-test: powershell/get_unique_cmdlet/get_unique_floating_point_comparisons
# Consecutive floating-point numbers with identical values are deduplicated
$floats = @(1.25, 1.25, 2.75, 3.5, 3.5)
$unique = @($floats | Get-Unique)

if ($unique.Count -ne 3) {
    Write-Host "FAIL: expected 3 unique float elements, got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne 1.25 -or $unique[1] -ne 2.75 -or $unique[2] -ne 3.5) {
    Write-Host "FAIL: float sequence mismatch: @($($unique -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
