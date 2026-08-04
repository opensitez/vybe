# vybe-test: powershell/subexpressions/command_output_subexpression
$value = $(Write-Output 'ok')
if ($value -ne 'ok') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
