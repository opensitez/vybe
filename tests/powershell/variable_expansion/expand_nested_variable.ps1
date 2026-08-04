# vybe-test: powershell/variable_expansion/expand_nested_variable
$a = 'z'
$b = '$a'
if ("$b" -ne '$a') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
