# vybe-test: powershell/variable_expansion/expand_null_in_string
$value = $null
if ("$value" -ne '') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
