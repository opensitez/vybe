# vybe-test: powershell/ternary/ternary_with_boolean
$result = $true ? 'yes' : 'no'
if ($result -ne 'yes') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
