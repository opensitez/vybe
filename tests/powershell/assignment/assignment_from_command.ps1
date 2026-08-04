# vybe-test: powershell/assignment/assignment_from_command
$result = Get-Process | Select-Object -First 1
if ($result -eq $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
