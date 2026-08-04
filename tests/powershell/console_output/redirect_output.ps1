# vybe-test: powershell/console_output/redirect_output
$result = @(Write-Output 'a')
if ($result[0] -ne 'a') {
    Write-Host "FAIL: expected a"
    exit 1
}
Write-Host 'PASS'
exit 0
