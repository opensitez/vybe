# vybe-test: powershell/get_unique_cmdlet/get_unique_unsorted_duplicates_retained
# Like UNIX uniq, Get-Unique only compares against the immediately preceding item; non-consecutive duplicates remain
$items = @(1, 2, 1)
$unique = @($items | Get-Unique)

if ($unique.Count -ne 3) {
    Write-Host "FAIL: non-consecutive duplicates should be retained, got count $($unique.Count)"
    exit 1
}

if ($unique[0] -ne 1 -or $unique[1] -ne 2 -or $unique[2] -ne 1) {
    Write-Host "FAIL: unsorted items sequence mismatch: @($($unique -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
