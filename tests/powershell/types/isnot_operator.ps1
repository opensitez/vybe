# vybe-test: powershell/types/isnot_operator
$x = 42
$result = $x -isnot [string]
if ($result -ne $true) {
    Write-Host "FAIL: expected True for isnot string check"
    exit 1
}
Write-Host "PASS"
exit 0
