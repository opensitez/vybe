# vybe-test: powershell/types/is_operator
$x = 42
$result = $x -is [int]
if ($result -ne $true) {
    Write-Host "FAIL: expected True for int type check"
    exit 1
}
$str = "hello"
$result2 = $str -is [string]
if ($result2 -ne $true) {
    Write-Host "FAIL: expected True for string type check"
    exit 1
}
Write-Host "PASS"
exit 0
