# vybe-test: powershell/redirection/pipe_redirection
$content = 'x' | ForEach-Object { $_.ToUpper() }
if ($content -ne 'X') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
