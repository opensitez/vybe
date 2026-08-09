# vybe-test: powershell/null_coalescing_assignment/null_assignment_array_var
$arr = $null
$arr ??= @(10, 20)
if ($arr.Count -ne 2 -or $arr[1] -ne 20) {
    Write-Host "FAIL: array variable ??= expected @(10, 20)"
    exit 1
}
Write-Host "PASS"
exit 0
