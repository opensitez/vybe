# vybe-test: powershell/get_unique_cmdlet/get_unique_chained_with_sort_object
# Idiomatic PowerShell deduplication pairs Sort-Object with Get-Unique
$unsorted = @(40, 20, 40, 10, 20, 30)
$unique = @($unsorted | Sort-Object | Get-Unique)

if ($unique.Count -ne 4) {
    Write-Host "FAIL: expected 4 deduplicated sorted elements, got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne 10 -or $unique[1] -ne 20 -or $unique[2] -ne 30 -or $unique[3] -ne 40) {
    Write-Host "FAIL: sorted unique elements mismatch: @($($unique -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
