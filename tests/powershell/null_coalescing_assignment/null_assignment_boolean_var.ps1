# vybe-test: powershell/null_coalescing_assignment/null_assignment_boolean_var
$flag = $null
$flag ??= $true
if ($flag -ne $true) {
    Write-Host "FAIL: boolean variable ??= expected true, got $flag"
    exit 1
}
Write-Host "PASS"
exit 0
