# vybe-test: powershell/get_unique_cmdlet/get_unique_ontype_filters_consecutive_types
# -OnType emits an object only if its .NET type differs from the preceding object's type
$items = @(1, 2, "first", "second", 3.14)
$unique = @($items | Get-Unique -OnType)

if ($unique.Count -ne 3) {
    Write-Host "FAIL: expected 3 unique types (int, string, double), got $($unique.Count)"
    exit 1
}

if ($unique[0] -isnot [int] -or $unique[1] -isnot [string] -or $unique[2] -isnot [double]) {
    Write-Host "FAIL: -OnType output types mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
