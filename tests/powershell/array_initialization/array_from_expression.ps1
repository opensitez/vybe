# vybe-test: powershell/array_initialization/array_from_expression
$arr = @(1..5)
if ($arr.Length -ne 5 -or $arr[4] -ne 5) {
    Write-Host "FAIL: Array init from expression failed"
    exit 1
}
Write-Host "PASS"
exit 0
