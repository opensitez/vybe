# vybe-test: powershell/variable_expansion/expand_variable_in_string
$name = 'x'
if ("Hello $name" -ne 'Hello x') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
