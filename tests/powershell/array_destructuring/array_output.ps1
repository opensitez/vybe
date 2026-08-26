# vybe-test: powershell/array_destructuring/array_output
$arr = @(1, 2, 3)
$a, $b, $c = $arr
if ($a -ne 1 -or $b -ne 2 -or $c -ne 3) {
    Write-Host "FAIL: Array destructuring failed"
    exit 1
}
Write-Host "PASS"
exit 0
