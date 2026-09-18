# vybe-test: powershell/get_unique_cmdlet/get_unique_negative_numbers
# Consecutive negative numbers with identical values are deduplicated correctly
$numbers = @(-50, -50, -25, -25, 0)
$unique = @($numbers | Get-Unique)

if ($unique.Count -ne 3) {
    Write-Host "FAIL: expected 3 unique numbers, got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne -50 -or $unique[1] -ne -25 -or $unique[2] -ne 0) {
    Write-Host "FAIL: negative number sequence mismatch: @($($unique -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
