# vybe-test: powershell/subexpressions/command_subexpression
$value = $(Get-Command Write-Host).Name
if ($value -ne 'Write-Host') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
