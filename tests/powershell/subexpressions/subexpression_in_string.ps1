# vybe-test: powershell/subexpressions/subexpression_in_string
$text = "Value: $($(1+1))"
if ($text -ne 'Value: 2') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
