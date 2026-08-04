# vybe-test: powershell/parentheses/function_call
$result = (Get-Date).Year
if ($result -ge 2026) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'FAIL'
exit 1
