# vybe-test: powershell/dsc_resources/get_dsc_resource_command
if (-not (Get-Command Get-DscResource -ErrorAction SilentlyContinue)) {
    Write-Host "FAIL: Get-DscResource command unavailable"
    exit 1
}
Write-Host 'PASS'
exit 0
