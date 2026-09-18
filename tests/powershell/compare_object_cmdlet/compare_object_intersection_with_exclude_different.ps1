# vybe-test: powershell/compare_object_cmdlet/compare_object_intersection_with_exclude_different
# Combining -IncludeEqual and -ExcludeDifferent extracts the mathematical intersection
$intersection = Compare-Object @("a", "b", "c") @("b", "c", "d") -IncludeEqual -ExcludeDifferent

if ($intersection.Count -ne 2) {
    Write-Host "FAIL: expected 2 intersection elements, got $($intersection.Count)"
    exit 1
}

$vals = @($intersection | ForEach-Object { $_.InputObject })
if ($vals[0] -ne "b" -or $vals[1] -ne "c") {
    Write-Host "FAIL: expected @('b', 'c'), got: @($($vals -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
