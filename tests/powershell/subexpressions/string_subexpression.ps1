# vybe-test: powershell/subexpressions/string_subexpression
$value = $('hello'.ToUpper())
if ($value -ne 'HELLO') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
