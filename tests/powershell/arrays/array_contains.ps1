# vybe-test: powershell/arrays/array_contains
$arr = @(1, 2, 3, 4, 5)
$result = $arr -contains 3
if ($result -ne $true) {
    Write-Host "FAIL: expected True, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
