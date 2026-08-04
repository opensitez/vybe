# vybe-test: powershell/assignment/command_assignment
$result = Get-Date
if ($result -eq $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
