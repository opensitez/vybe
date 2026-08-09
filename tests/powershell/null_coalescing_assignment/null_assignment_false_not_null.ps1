# vybe-test: powershell/null_coalescing_assignment/null_assignment_false_not_null
$b = $false
$b ??= $true
if ($b -ne $false) {
    Write-Host "FAIL: boolean false should NOT be treated as null by ??=, got $b"
    exit 1
}
Write-Host "PASS"
exit 0
