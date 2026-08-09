# vybe-test: powershell/null_coalescing_assignment/null_assignment_zero_not_null
$num = 0
$num ??= 999
if ($num -ne 0) {
    Write-Host "FAIL: zero integer should NOT be treated as null by ??=, got $num"
    exit 1
}
Write-Host "PASS"
exit 0
