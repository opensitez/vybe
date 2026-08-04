# vybe-test: powershell/types/type_array
$arr = @(1, 2, 3)
$result = $arr -is [array]
if ($result -ne $true) {
    Write-Host "FAIL: expected True for array type check"
    exit 1
}
Write-Host "PASS"
exit 0
