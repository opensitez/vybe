# vybe-test: powershell/math/math_operations
$result = [Math]::PI
if ($result -lt 3.14 -or $result -gt 3.15) {
    Write-Host "FAIL: expected PI around 3.14, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
