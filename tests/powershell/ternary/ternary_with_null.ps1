# vybe-test: powershell/ternary/ternary_with_null
$value = $null
$result = ($value -eq $null) ? 'null' : 'notnull'
if ($result -ne 'null') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
