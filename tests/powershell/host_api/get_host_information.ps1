# vybe-test: powershell/host_api/get_host_information
$host = Get-Host
if ($host.Name -eq $null) {
    Write-Host "FAIL: expected host name"
    exit 1
}
Write-Host 'PASS'
exit 0
