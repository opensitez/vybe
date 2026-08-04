# vybe-test: powershell/subexpressions/nested_subexpression
$value = $($(1 + $(2)))
if ($value -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
