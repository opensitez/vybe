# vybe-test: powershell/variables/automatic_variables
$result = $true
if ($result -ne $true) {
    Write-Host "FAIL: expected $true to be True"
    exit 1
}
$result2 = $false
if ($result2 -ne $false) {
    Write-Host "FAIL: expected $false to be False"
    exit 1
}
Write-Host "PASS"
exit 0
