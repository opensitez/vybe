# vybe-test: powershell/variable_expansion/expand_array_in_string
$arr = 1,2
if ("$arr" -ne '1 2') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
