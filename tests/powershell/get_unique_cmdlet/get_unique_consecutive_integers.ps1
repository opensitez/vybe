# vybe-test: powershell/get_unique_cmdlet/get_unique_consecutive_integers
# Get-Unique eliminates consecutive duplicate values from a sorted integer stream
$items = @(1, 1, 2, 2, 2, 3)
$unique = @($items | Get-Unique)

if ($unique.Count -ne 3) {
    Write-Host "FAIL: expected 3 unique items, got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne 1 -or $unique[1] -ne 2 -or $unique[2] -ne 3) {
    Write-Host "FAIL: unexpected items: @($($unique -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
