# vybe-test: powershell/cmdlets/compare_object
$a = @(1, 2, 3)
$b = @(2, 3, 4)
$diff = Compare-Object $a $b
if ($diff.Count -ne 2) {
    Write-Host "FAIL: expected 2 differences, got $($diff.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
