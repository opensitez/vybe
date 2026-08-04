# vybe-test: powershell/variables/null_variable
$x = $null
if ($null -ne $x) {
    Write-Host "FAIL: expected null, got $x"
    exit 1
}
Write-Host "PASS"
exit 0
